require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { execFile } = require('child_process');

const db = require('./db');

const app = express();
const PYTHON_BIN = process.env.PYTHON_BIN || 'python3';
const RECALC_SCRIPT = process.env.RECALC_SCRIPT || ''; // optional LibreOffice recalc.py

const UPLOAD_DIR = path.join(__dirname, '..', 'uploads');
const OUTPUT_DIR = path.join(__dirname, '..', 'outputs');

[UPLOAD_DIR, OUTPUT_DIR].forEach((d) => {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
});

const allowedOrigins = [
  'https://adidas-dpr-frontend-wzur.onrender.com',
  'http://localhost:4200',
  'http://127.0.0.1:4200',
  process.env.FRONTEND_URL,
].filter(Boolean);

app.use(
  cors({
    origin: (origin, callback) => {
      if (!origin || allowedOrigins.includes(origin)) {
        return callback(null, true);
      }
      return callback(new Error(`Origin ${origin} not allowed by CORS`));
    },
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'Accept', 'Origin', 'X-Requested-With'],
    credentials: true,
  })
);
app.use(express.json({ limit: '256kb' }));

const upload = multer({
  dest: UPLOAD_DIR,
  limits: { fileSize: 25 * 1024 * 1024 }, // 25MB
  fileFilter: (req, file, cb) => {
    const ok = /\.(xlsx|xls)$/i.test(file.originalname);
    cb(ok ? null : new Error('Only .xlsx or .xls files are allowed.'), ok);
  },
});

app.get('/health', (req, res) => {
  res.status(200).json({ status: 'ok', uptime: process.uptime(), timestamp: new Date().toISOString() });
});

app.get('/api/health', async (req, res) => {
  res.status(200).json({ ok: true, database: await db.isAvailable() });
});

app.get('/api/jobs', async (req, res) => {
  res.json({ jobs: await db.listRecentJobs() });
});

// In-memory map of planId -> generated workbook paths on disk. This lets
// step 2 (upload the linewise performance report) find the exact workbook
// that step 1 (upload Book1) produced, without needing the client to
// re-upload it. Independent of MySQL so this works even if the DB
// is unreachable.
const planFiles = new Map();

app.post('/api/process', (req, res) => {
  upload.single('file')(req, res, async (uploadErr) => {
    if (uploadErr) {
      return res.status(400).json({ error: uploadErr.message });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    const inputWithExt = req.file.path + path.extname(req.file.originalname);
    fs.renameSync(req.file.path, inputWithExt);

    const planId = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    let jobId = await db.createJob(req.file.originalname, planId);
    const outFileName = `Generated_Plan_Vs_Actuals_${planId}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR, outFileName);
    const scriptPath = path.join(__dirname, 'generate.py');

    execFile(PYTHON_BIN, [scriptPath, inputWithExt, outputPath], async (err, stdout, stderr) => {
      fs.unlink(inputWithExt, () => {});

      if (err) {
        let details = (stderr || err.message || '').trim();
        try {
          const parsed = JSON.parse(details.split('\n').pop());
          if (parsed.error) details = parsed.error;
        } catch (_) { /* keep raw details */ }
        console.error('generate.py failed:', details);
        if (jobId == null) jobId = await db.createJob(req.file.originalname, planId);
        await db.markJobFailed(jobId, details);
        return res.status(422).json({ error: details || 'Processing failed.' });
      }

      if (!fs.existsSync(outputPath)) {
        if (jobId == null) jobId = await db.createJob(req.file.originalname, planId);
        await db.markJobFailed(jobId, 'Output file was not generated.');
        return res.status(500).json({ error: 'Output file was not generated.' });
      }

      const finish = async () => {
        let summary = {};
        try { summary = JSON.parse(stdout.trim().split('\n').pop()); } catch (_) { /* ignore */ }

        if (jobId == null) jobId = await db.createJob(req.file.originalname, planId);
        await db.markJobComplete(jobId, {
          stylesCount: summary.styles ?? null,
          rowsCount: summary.rows ?? null,
          datesCount: summary.dates ?? null,
          outputFile: outFileName,
        });

        planFiles.set(planId, { originalPath: outputPath, currentPath: outputPath, originalName: req.file.originalname });

        return res.json({
          success: true,
          jobId,
          planId,
          downloadUrl: `/api/download/${outFileName}?name=${encodeURIComponent(req.file.originalname)}`,
          summary,
        });
      };

      if (RECALC_SCRIPT && fs.existsSync(RECALC_SCRIPT)) {
        execFile(PYTHON_BIN, [RECALC_SCRIPT, outputPath, '60'], () => finish());
      } else {
        finish();
      }
    });
  });
});

// Step 1.5: upload an Order Sheet to aggregate order summary (Style, X-fty, Qty, Sewing, Wash, Pack)
// into the workbook that /api/process generated for the given planId.
app.post('/api/fill-order-summary/:planId', (req, res) => {
  upload.single('file')(req, res, async (uploadErr) => {
    if (uploadErr) {
      return res.status(400).json({ error: uploadErr.message });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    const { planId } = req.params;
    const planRecord = planFiles.get(planId);
    const sourcePath = planRecord?.currentPath;
    if (!sourcePath || !fs.existsSync(sourcePath)) {
      fs.unlink(req.file.path, () => {});
      return res.status(404).json({
        error: 'Could not find the generated workbook for this session. Please process your Book1 file again first.',
      });
    }

    const inputWithExt = req.file.path + path.extname(req.file.originalname);
    fs.renameSync(req.file.path, inputWithExt);

    const outFileName = `Generated_Plan_Vs_Actuals_${planId}_orders_${Date.now()}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR, outFileName);
    const scriptPath = path.join(__dirname, 'fill_order_summary.py');

    execFile(PYTHON_BIN, [scriptPath, sourcePath, inputWithExt, outputPath], (err, stdout, stderr) => {
      fs.unlink(inputWithExt, () => {});

      if (err) {
        let details = (stderr || err.message || '').trim();
        try {
          const parsed = JSON.parse(details.split('\n').pop());
          if (parsed.error) details = parsed.error;
        } catch (_) { /* keep raw details */ }
        console.error('fill_order_summary.py failed:', details);
        return res.status(422).json({ error: details || 'Processing failed.' });
      }

      if (!fs.existsSync(outputPath)) {
        return res.status(500).json({ error: 'Output file was not generated.' });
      }

      let summary = {};
      try { summary = JSON.parse(stdout.trim().split('\n').pop()); } catch (_) { /* ignore */ }

      planFiles.set(planId, { ...planRecord, currentPath: outputPath });

      const origName = planRecord?.originalName || 'Plan_Vs_Actuals_Style_wise.xlsx';
      return res.json({
        success: true,
        downloadUrl: `/api/download/${outFileName}?name=${encodeURIComponent(origName)}`,
        summary,
      });
    });
  });
});

// Step 2: upload a linewise performance report to fill Sewing Actuals into
// the workbook that /api/process generated for the given planId.
app.post('/api/fill-actuals/:planId', (req, res) => {
  upload.single('file')(req, res, async (uploadErr) => {
    if (uploadErr) {
      return res.status(400).json({ error: uploadErr.message });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    const { planId } = req.params;
    const planRecord = planFiles.get(planId);
    const sourcePath = planRecord?.currentPath;
    if (!sourcePath || !fs.existsSync(sourcePath)) {
      fs.unlink(req.file.path, () => {});
      return res.status(404).json({
        error: 'Could not find the generated workbook for this session. Please process your Book1 file again first.',
      });
    }

    const inputWithExt = req.file.path + path.extname(req.file.originalname);
    fs.renameSync(req.file.path, inputWithExt);

    const outFileName = `Plan_Vs_Actuals_Style_wise_updated_${Date.now()}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR, outFileName);
    const scriptPath = path.join(__dirname, 'fill_actuals.py');

    execFile(PYTHON_BIN, [scriptPath, sourcePath, inputWithExt, outputPath], (err, stdout, stderr) => {
      fs.unlink(inputWithExt, () => {});

      if (err) {
        let details = (stderr || err.message || '').trim();
        try {
          const parsed = JSON.parse(details.split('\n').pop());
          if (parsed.error) details = parsed.error;
        } catch (_) { /* keep raw details */ }
        console.error('fill_actuals.py failed:', details);
        return res.status(422).json({ error: details || 'Processing failed.' });
      }

      if (!fs.existsSync(outputPath)) {
        return res.status(500).json({ error: 'Output file was not generated.' });
      }

      let summary = {};
      try { summary = JSON.parse(stdout.trim().split('\n').pop()); } catch (_) { /* ignore */ }

      planFiles.set(planId, { ...planRecord, currentPath: outputPath });

      const origName = planRecord?.originalName || 'Plan_Vs_Actuals_Style_wise.xlsx';
      return res.json({
        success: true,
        downloadUrl: `/api/download/${outFileName}?name=${encodeURIComponent(origName)}`,
        summary,
      });
    });
  });
});

// Direct fill-actuals endpoint for daily performance updates (Sewing and/or Wash & Pack)
app.post('/api/fill-actuals-direct', (req, res) => {
  upload.any()(req, res, async (uploadErr) => {
    if (uploadErr) {
      return res.status(400).json({ error: uploadErr.message });
    }
    const files = Array.isArray(req.files) ? req.files : [];
    const wbFile = files.find((f) => f.fieldname === 'workbook') || files[0];
    const sewingFile = files.find((f) => f.fieldname === 'sewingReport' || f.fieldname === 'report');
    const wpFile = files.find((f) => f.fieldname === 'washPackReport');

    if (!wbFile) {
      files.forEach((f) => fs.unlink(f.path, () => {}));
      return res.status(400).json({ error: 'Please upload existing workbook.' });
    }

    const wbExt = wbFile.path + path.extname(wbFile.originalname);
    fs.renameSync(wbFile.path, wbExt);

    const sewingExt = sewingFile ? sewingFile.path + path.extname(sewingFile.originalname) : null;
    if (sewingFile && sewingExt) fs.renameSync(sewingFile.path, sewingExt);

    const wpExt = wpFile ? wpFile.path + path.extname(wpFile.originalname) : null;
    if (wpFile && wpExt) fs.renameSync(wpFile.path, wpExt);

    const outFileName = `Plan_Vs_Actuals_Style_wise_updated_${Date.now()}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR, outFileName);

    const cleanup = () => {
      fs.unlink(wbExt, () => {});
      if (sewingExt) fs.unlink(sewingExt, () => {});
      if (wpExt) fs.unlink(wpExt, () => {});
    };

    const runScript = (scriptName, inputWb, reportFile, targetOut) => {
      return new Promise((resolve, reject) => {
        const scriptPath = path.join(__dirname, scriptName);
        execFile(PYTHON_BIN, [scriptPath, inputWb, reportFile, targetOut], (err, stdout, stderr) => {
          if (err) {
            let details = (stderr || err.message || '').trim();
            try {
              const parsed = JSON.parse(details.split('\n').pop());
              if (parsed.error) details = parsed.error;
            } catch (_) {}
            return reject(new Error(details || `${scriptName} processing failed.`));
          }
          let summary = {};
          try { summary = JSON.parse(stdout.trim().split('\n').pop()); } catch (_) {}
          resolve(summary);
        });
      });
    };

    try {
      let currentSource = wbExt;
      let combinedSummary = { sheetsUpdated: 0, cellsFilled: 0 };

      if (sewingExt) {
        const intermediatePath = path.join(OUTPUT_DIR, `temp_sewing_${Date.now()}.xlsx`);
        const sSummary = await runScript('fill_actuals.py', currentSource, sewingExt, intermediatePath);
        combinedSummary.sheetsUpdated = Math.max(combinedSummary.sheetsUpdated, sSummary.sheetsUpdated || 0);
        combinedSummary.cellsFilled += (sSummary.cellsFilled || 0);
        currentSource = intermediatePath;
      }

      if (wpExt) {
        const wpSummary = await runScript('fill_wash_pack.py', currentSource, wpExt, outputPath);
        combinedSummary.sheetsUpdated = Math.max(combinedSummary.sheetsUpdated, wpSummary.sheetsUpdated || 0);
        combinedSummary.cellsFilled += (wpSummary.cellsFilled || 0);
        if (currentSource !== wbExt && fs.existsSync(currentSource)) {
          fs.unlink(currentSource, () => {});
        }
      } else if (sewingExt) {
        fs.renameSync(currentSource, outputPath);
      } else {
        fs.copyFileSync(wbExt, outputPath);
      }

      cleanup();

      return res.json({
        success: true,
        downloadUrl: `/api/download/${outFileName}?name=${encodeURIComponent(wbFile.originalname)}`,
        summary: combinedSummary,
      });
    } catch (err) {
      cleanup();
      return res.status(422).json({ error: err.message || 'Processing failed.' });
    }
  });
});

// Step 3: upload a Wash & Pack daily performance report for plan session
app.post('/api/fill-wash-pack/:planId', (req, res) => {
  upload.single('file')(req, res, async (uploadErr) => {
    if (uploadErr) {
      return res.status(400).json({ error: uploadErr.message });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    const { planId } = req.params;
    const planRecord = planFiles.get(planId);
    const sourcePath = planRecord?.currentPath;
    if (!sourcePath || !fs.existsSync(sourcePath)) {
      fs.unlink(req.file.path, () => {});
      return res.status(404).json({
        error: 'Could not find the generated workbook for this session. Please process your Book1 file again first.',
      });
    }

    const inputWithExt = req.file.path + path.extname(req.file.originalname);
    fs.renameSync(req.file.path, inputWithExt);

    const outFileName = `Plan_Vs_Actuals_Style_wise_updated_${Date.now()}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR, outFileName);
    const scriptPath = path.join(__dirname, 'fill_wash_pack.py');

    execFile(PYTHON_BIN, [scriptPath, sourcePath, inputWithExt, outputPath], (err, stdout, stderr) => {
      fs.unlink(inputWithExt, () => {});

      if (err) {
        let details = (stderr || err.message || '').trim();
        try {
          const parsed = JSON.parse(details.split('\n').pop());
          if (parsed.error) details = parsed.error;
        } catch (_) { /* keep raw details */ }
        console.error('fill_wash_pack.py failed:', details);
        return res.status(422).json({ error: details || 'Processing failed.' });
      }

      if (!fs.existsSync(outputPath)) {
        return res.status(500).json({ error: 'Output file was not generated.' });
      }

      let summary = {};
      try { summary = JSON.parse(stdout.trim().split('\n').pop()); } catch (_) { /* ignore */ }

      planFiles.set(planId, { ...planRecord, currentPath: outputPath });

      const origName = planRecord?.originalName || 'Plan_Vs_Actuals_Style_wise.xlsx';
      return res.json({
        success: true,
        downloadUrl: `/api/download/${outFileName}?name=${encodeURIComponent(origName)}`,
        summary,
      });
    });
  });
});

// Direct Daily Wash & Pack Actuals Upload Endpoint
app.post('/api/fill-wash-pack-direct', (req, res) => {
  upload.any()(req, res, async (uploadErr) => {
    if (uploadErr) {
      return res.status(400).json({ error: uploadErr.message });
    }
    const files = Array.isArray(req.files) ? req.files : [];
    const wbFile = files.find((f) => f.fieldname === 'workbook') || files[0];
    const rptFile = files.find((f) => f.fieldname === 'report' || f.fieldname === 'washPackReport');

    if (!wbFile) {
      files.forEach((f) => fs.unlink(f.path, () => {}));
      return res.status(400).json({ error: 'Please upload existing workbook.' });
    }

    const wbExt = wbFile.path + path.extname(wbFile.originalname);
    fs.renameSync(wbFile.path, wbExt);

    const outFileName = `Plan_Vs_Actuals_Style_wise_updated_${Date.now()}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR, outFileName);

    // If only workbook uploaded (no report), return workbook unchanged
    if (!rptFile) {
      fs.copyFileSync(wbExt, outputPath);
      fs.unlink(wbExt, () => {});
      return res.json({
        success: true,
        downloadUrl: `/api/download/${outFileName}?name=${encodeURIComponent(wbFile.originalname)}`,
        summary: { sheetsUpdated: 0, cellsFilled: 0 },
      });
    }

    const rptExt = rptFile.path + path.extname(rptFile.originalname);
    fs.renameSync(rptFile.path, rptExt);

    const scriptPath = path.join(__dirname, 'fill_wash_pack.py');

    execFile(PYTHON_BIN, [scriptPath, wbExt, rptExt, outputPath], (err, stdout, stderr) => {
      fs.unlink(wbExt, () => {});
      fs.unlink(rptExt, () => {});

      if (err) {
        let details = (stderr || err.message || '').trim();
        try {
          const parsed = JSON.parse(details.split('\n').pop());
          if (parsed.error) details = parsed.error;
        } catch (_) { /* keep raw */ }
        return res.status(422).json({ error: details || 'Processing failed.' });
      }

      let summary = {};
      try { summary = JSON.parse(stdout.trim().split('\n').pop()); } catch (_) { /* ignore */ }

      return res.json({
        success: true,
        downloadUrl: `/api/download/${outFileName}?name=${encodeURIComponent(wbFile.originalname)}`,
        summary,
      });
    });
  });
});


// ==================== UNIT 23 ENDPOINTS ====================

app.post('/api/unit23/process', (req, res) => {
  upload.single('file')(req, res, async (uploadErr) => {
    if (uploadErr) {
      return res.status(400).json({ error: uploadErr.message });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    const inputWithExt = req.file.path + path.extname(req.file.originalname);
    fs.renameSync(req.file.path, inputWithExt);

    const planId = `U23-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    let jobId = await db.createJob(`U-23_${req.file.originalname}`, planId);
    const outFileName = `U-23_Generated_Plan_Vs_Actuals_${planId}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR, outFileName);
    const scriptPath = path.join(__dirname, 'generate_unit23.py');


    execFile(PYTHON_BIN, [scriptPath, inputWithExt, outputPath], async (err, stdout, stderr) => {
      fs.unlink(inputWithExt, () => {});

      if (err) {
        let details = (stderr || err.message || '').trim();
        try {
          const parsed = JSON.parse(details.split('\n').pop());
          if (parsed.error) details = parsed.error;
        } catch (_) { /* keep raw details */ }
        console.error('Unit 23 generate.py failed:', details);
        if (jobId == null) jobId = await db.createJob(`U-23_${req.file.originalname}`, planId);
        await db.markJobFailed(jobId, details);
        return res.status(422).json({ error: details || 'Processing failed.' });
      }

      if (!fs.existsSync(outputPath)) {
        if (jobId == null) jobId = await db.createJob(`U-23_${req.file.originalname}`, planId);
        await db.markJobFailed(jobId, 'Output file was not generated.');
        return res.status(500).json({ error: 'Output file was not generated.' });
      }

      const finish = async () => {
        let summary = {};
        try { summary = JSON.parse(stdout.trim().split('\n').pop()); } catch (_) { /* ignore */ }

        if (jobId == null) jobId = await db.createJob(`U-23_${req.file.originalname}`, planId);
        await db.markJobComplete(jobId, {
          stylesCount: summary.styles ?? null,
          rowsCount: summary.rows ?? null,
          datesCount: summary.dates ?? null,
          outputFile: outFileName,
        });

        planFiles.set(planId, { originalPath: outputPath, currentPath: outputPath, originalName: req.file.originalname });

        return res.json({
          success: true,
          jobId,
          planId,
          downloadUrl: `/api/download/${outFileName}?name=${encodeURIComponent(req.file.originalname)}`,
          summary,
        });
      };

      if (RECALC_SCRIPT && fs.existsSync(RECALC_SCRIPT)) {
        execFile(PYTHON_BIN, [RECALC_SCRIPT, outputPath, '60'], () => finish());
      } else {
        finish();
      }
    });
  });
});

app.post('/api/unit23/fill-actuals-direct', (req, res) => {
  upload.any()(req, res, async (uploadErr) => {
    if (uploadErr) {
      return res.status(400).json({ error: uploadErr.message });
    }
    const files = Array.isArray(req.files) ? req.files : [];
    const wbFile = files.find((f) => f.fieldname === 'workbook') || files[0];
    const sewingFile = files.find((f) => f.fieldname === 'sewingReport' || f.fieldname === 'report');
    const indexedDprReports = files
      .filter((f) => /^dprReport_\d+$/.test(f.fieldname))
      .map((file) => {
        const index = file.fieldname.slice('dprReport_'.length);
        return { file, reportDate: req.body?.[`dprDate_${index}`] || '' };
      });
    const legacyDprFile = files.find((f) => f.fieldname === 'washPackReport' || f.fieldname === 'dprReport');
    const dprReports = indexedDprReports;
    if (legacyDprFile) {
      dprReports.push({ file: legacyDprFile, reportDate: req.body?.reportDate || '' });
    }

    if (!wbFile) {
      files.forEach((f) => fs.unlink(f.path, () => {}));
      return res.status(400).json({ error: 'Please upload existing workbook.' });
    }

    const reportDates = dprReports.map((report) => report.reportDate).filter(Boolean);
    const invalidDprReport = dprReports.find(({ file, reportDate }) => {
      if (!reportDate) return indexedDprReports.some((report) => report.file === file);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(reportDate)) return true;
      const parsedDate = new Date(`${reportDate}T00:00:00.000Z`);
      return Number.isNaN(parsedDate.getTime()) || parsedDate.toISOString().slice(0, 10) !== reportDate;
    });
    if (invalidDprReport) {
      files.forEach((f) => fs.unlink(f.path, () => {}));
      return res.status(400).json({
        error: `Select a valid report date for DPR file "${invalidDprReport.file.originalname}".`,
      });
    }
    if (new Set(reportDates).size !== reportDates.length) {
      files.forEach((f) => fs.unlink(f.path, () => {}));
      return res.status(400).json({ error: 'Each Wash & Pack DPR must have a different report date.' });
    }

    const wbExt = wbFile.path + path.extname(wbFile.originalname);
    fs.renameSync(wbFile.path, wbExt);

    const sewingExt = sewingFile ? sewingFile.path + path.extname(sewingFile.originalname) : null;
    if (sewingFile && sewingExt) fs.renameSync(sewingFile.path, sewingExt);

    const dprExts = dprReports.map(({ file }) => {
      const fileWithExt = file.path + path.extname(file.originalname);
      fs.renameSync(file.path, fileWithExt);
      return fileWithExt;
    });

    const outFileName = `U-23_Plan_Vs_Actuals_Style_wise_updated_${Date.now()}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR, outFileName);
    const temporaryPaths = [];

    const cleanup = () => {
      fs.unlink(wbExt, () => {});
      if (sewingExt) fs.unlink(sewingExt, () => {});
      dprExts.forEach((filePath) => fs.unlink(filePath, () => {}));
      temporaryPaths.forEach((filePath) => fs.unlink(filePath, () => {}));
    };

    const runScript = (scriptName, inputWb, reportFile, targetOut, extraArg = null) => {
      return new Promise((resolve, reject) => {
        const scriptPath = path.join(__dirname, scriptName);
        const args = [scriptPath, inputWb, reportFile, targetOut];
        if (extraArg) args.push(extraArg);

        execFile(PYTHON_BIN, args, (err, stdout, stderr) => {
          if (err) {
            let details = (stderr || err.message || '').trim();
            try {
              const parsed = JSON.parse(details.split('\n').pop());
              if (parsed.error) details = parsed.error;
            } catch (_) {}
            return reject(new Error(details || `${scriptName} processing failed.`));
          }
          let summary = {};
          try { summary = JSON.parse(stdout.trim().split('\n').pop()); } catch (_) {}
          resolve(summary);
        });
      });
    };

    try {
      let currentSource = wbExt;
      let combinedSummary = { sheetsUpdated: 0, cellsFilled: 0 };

      if (sewingExt) {
        const intermediatePath = path.join(OUTPUT_DIR, `temp_sewing_${Date.now()}.xlsx`);
        temporaryPaths.push(intermediatePath);
        const sSummary = await runScript('fill_actuals.py', currentSource, sewingExt, intermediatePath);
        combinedSummary.sheetsUpdated = Math.max(combinedSummary.sheetsUpdated, sSummary.sheetsUpdated || 0);
        combinedSummary.cellsFilled += (sSummary.cellsFilled || 0);
        currentSource = intermediatePath;
      }

      if (dprReports.length) {
        const orderedDprReports = dprReports
          .map((report, index) => ({ ...report, filePath: dprExts[index] }))
          .sort((a, b) => a.reportDate.localeCompare(b.reportDate));

        for (let index = 0; index < orderedDprReports.length; index += 1) {
          const report = orderedDprReports[index];
          const targetOut = index === orderedDprReports.length - 1
            ? outputPath
            : path.join(OUTPUT_DIR, `temp_dpr_${Date.now()}_${index}.xlsx`);
          if (targetOut !== outputPath) temporaryPaths.push(targetOut);

          const dprSummary = await runScript(
            'fill_dpr_wash_pack.py',
            currentSource,
            report.filePath,
            targetOut,
            report.reportDate
          );
          combinedSummary.sheetsUpdated = Math.max(combinedSummary.sheetsUpdated, dprSummary.sheetsUpdated || 0);
          combinedSummary.cellsFilled += (dprSummary.cellsFilled || 0);
          currentSource = targetOut;
        }
      } else if (sewingExt) {
        fs.renameSync(currentSource, outputPath);
      } else {
        fs.copyFileSync(wbExt, outputPath);
      }

      cleanup();

      return res.json({
        success: true,
        downloadUrl: `/api/download/${outFileName}?name=${encodeURIComponent(wbFile.originalname)}`,
        summary: combinedSummary,
      });
    } catch (err) {
      cleanup();
      return res.status(422).json({ error: err.message || 'Processing failed.' });
    }
  });
});

app.post('/api/apply-exclusions/:planId', (req, res) => {
  const { planId } = req.params;
  const planRecord = planFiles.get(planId);
  const sourcePath = planRecord?.currentPath || planRecord?.originalPath;
  if (!sourcePath || !fs.existsSync(sourcePath)) {
    return res.status(404).json({
      error: 'Could not find the generated workbook for this session. Please process your Book1 file again first.',
    });
  }

  const excludedDates = Array.isArray(req.body?.excludedDates) ? req.body.excludedDates : [];
  const invalidDate = excludedDates.find((d) => !/^\d{4}-\d{2}-\d{2}$/.test(String(d)));
  if (invalidDate) {
    return res.status(400).json({ error: `Invalid excluded date: ${invalidDate}` });
  }

  const outFileName = `Generated_Plan_Vs_Actuals_${planId}_excluded_${Date.now()}.xlsx`;
  const outputPath = path.join(OUTPUT_DIR, outFileName);
  const scriptPath = path.join(__dirname, 'apply_exclusions.py');

  execFile(PYTHON_BIN, [scriptPath, sourcePath, outputPath, JSON.stringify(excludedDates)], (err, stdout, stderr) => {
    if (err) {
      let details = (stderr || err.message || '').trim();
      try {
        const parsed = JSON.parse(details.split('\n').pop());
        if (parsed.error) details = parsed.error;
      } catch (_) { /* keep raw details */ }
      console.error('apply_exclusions.py failed:', details);
      return res.status(422).json({ error: details || 'Failed to apply exclusions.' });
    }

    if (!fs.existsSync(outputPath)) {
      return res.status(500).json({ error: 'Output file was not generated.' });
    }

    let summary = {};
    try { summary = JSON.parse(stdout.trim().split('\n').pop()); } catch (_) { /* ignore */ }

    planFiles.set(planId, { ...planRecord, currentPath: outputPath });

    const origName = planRecord?.originalName || 'Plan_Vs_Actuals_Style_wise.xlsx';
    return res.json({
      success: true,
      downloadUrl: `/api/download/${outFileName}?name=${encodeURIComponent(origName)}`,
      summary,
    });
  });
});

app.get('/api/download/:file', (req, res) => {
  const file = req.params.file;
  if (!/^[\w.-]+\.xlsx$/.test(file)) return res.status(400).send('Invalid file name.');
  const filePath = path.join(OUTPUT_DIR, file);
  if (!fs.existsSync(filePath)) return res.status(404).send('File not found or already downloaded.');

  let downloadName = req.query.name ? String(req.query.name) : file;
  downloadName = path.basename(downloadName);
  if (!/\.xlsx$/i.test(downloadName)) {
    downloadName += '.xlsx';
  }
  res.download(filePath, downloadName);
});

module.exports = app;
