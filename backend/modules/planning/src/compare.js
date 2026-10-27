function parseDate(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return null;
  const trimmed = dateStr.trim();
  if (!trimmed || trimmed === '-' || trimmed.toLowerCase() === 'n/a') return null;

  const parts = trimmed.split(/[/.-]/);
  if (parts.length === 3) {
    let year, month, day;
    if (parts[0].length === 4) {
      // YYYY-MM-DD
      year = parseInt(parts[0], 10);
      month = parseInt(parts[1], 10) - 1;
      day = parseInt(parts[2], 10);
    } else if (isNaN(parts[1])) {
      // 23-Sep-2026
      day = parseInt(parts[0], 10);
      const months = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
      const mStr = parts[1].toLowerCase().slice(0, 3);
      month = months[mStr] !== undefined ? months[mStr] : -1;
      year = parseInt(parts[2], 10);
      if (year < 100) year += 2000;
    } else {
      // MM/DD/YYYY or M/D/YYYY 
      month = parseInt(parts[0], 10) - 1;
      day = parseInt(parts[1], 10);
      year = parseInt(parts[2], 10);
      if (year < 100) year += 2000;
    }

    if (!isNaN(year) && !isNaN(month) && month >= 0 && month < 12 && !isNaN(day) && day >= 1 && day <= 31) {
      return new Date(Date.UTC(year, month, day));
    }
  }

  const timestamp = Date.parse(trimmed);
  if (!isNaN(timestamp)) {
    const d = new Date(timestamp);
    return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  }

  return null;
}

function calculateDateDiff(julyDateStr, augustDateStr) {
  const jHas = julyDateStr && julyDateStr !== '-';
  const aHas = augustDateStr && augustDateStr !== '-';
  


  if (!jHas && aHas) return 'New';
  if (jHas && !aHas) return 'Removed';
  if (!jHas && !aHas) return '0 days';

  const d1 = parseDate(julyDateStr);
  const d2 = parseDate(augustDateStr);

  if (!d1 || !d2) {
    if (julyDateStr === augustDateStr) return '0 days';
    return 'Changed';
  }

  const diffMs = d2.getTime() - d1.getTime();
  const diffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));

  if (diffDays > 0) return `+${diffDays} days`;
  if (diffDays < 0) return `${diffDays} days`;
  return '0 days';
}

function groupByStyle(rows) {
  const map = new Map();
  for (const row of rows) {
    if (!map.has(row.style)) {
      map.set(row.style, {
        style: row.style,
        description: row.description || '',
        totalQty: 0,
        rows: [],
      });
    }
    const entry = map.get(row.style);
    entry.totalQty += row.qty;
    if (!entry.description && row.description) {
      entry.description = row.description;
    }
    entry.rows.push(row);
  }
  return map;
}

function matchSchedulesForStyle(style, defaultDesc, julyRows = [], augustRows = []) {
  const julyUsed = new Array(julyRows.length).fill(false);
  const augustUsed = new Array(augustRows.length).fill(false);
  const schedulePairs = [];

  // Pass 1: Exact match on (pcd, exFactory, qty)
  for (let i = 0; i < julyRows.length; i++) {
    if (julyUsed[i]) continue;
    const jRow = julyRows[i];
    for (let j = 0; j < augustRows.length; j++) {
      if (augustUsed[j]) continue;
      const aRow = augustRows[j];
      if (
        jRow.pcd && jRow.pcd === aRow.pcd &&
        jRow.exFactory && jRow.exFactory === aRow.exFactory &&
        jRow.qty === aRow.qty &&
        (!jRow.buyDate || !aRow.buyDate || jRow.buyDate === aRow.buyDate)
      ) {
        julyUsed[i] = true;
        augustUsed[j] = true;
        schedulePairs.push({
          style,
          description: aRow.description || jRow.description || defaultDesc,
          julyPcd: jRow.pcd,
          augustPcd: aRow.pcd,
          pcdDiff: '0 days',
          julyExFactory: jRow.exFactory,
          augustExFactory: aRow.exFactory,
          exFactoryDiff: '0 days',
          julyQty: jRow.qty,
          augustQty: aRow.qty,
          qtyDiff: 0,
          qtyDiffStr: '0',
          status: 'Same',
          buyDate: aRow.buyDate || jRow.buyDate || '',
          itemNumber: aRow.itemNumber || jRow.itemNumber || '',
        });
        break;
      }
    }
  }

  // Pass 2: Match by itemNumber or buyDate if present and non-empty
  for (let i = 0; i < julyRows.length; i++) {
    if (julyUsed[i]) continue;
    const jRow = julyRows[i];
    if (!jRow.itemNumber && !jRow.buyDate) continue;

    for (let j = 0; j < augustRows.length; j++) {
      if (augustUsed[j]) continue;
      const aRow = augustRows[j];

      const itemMatch = jRow.itemNumber && aRow.itemNumber && jRow.itemNumber === aRow.itemNumber;
      const buyDateMatch = jRow.buyDate && aRow.buyDate && jRow.buyDate === aRow.buyDate;

      if (itemMatch || buyDateMatch) {
        julyUsed[i] = true;
        augustUsed[j] = true;

        const pcdDiff = calculateDateDiff(jRow.pcd, aRow.pcd);
        const exFactoryDiff = calculateDateDiff(jRow.exFactory, aRow.exFactory);
        const qtyDiff = aRow.qty - jRow.qty;
        const qtyDiffStr = qtyDiff > 0 ? `+${qtyDiff}` : `${qtyDiff}`;
        const isSame = pcdDiff === '0 days' && exFactoryDiff === '0 days' && qtyDiff === 0;

        schedulePairs.push({
          style,
          description: aRow.description || jRow.description || defaultDesc,
          julyPcd: jRow.pcd || '-',
          augustPcd: aRow.pcd || '-',
          pcdDiff,
          julyExFactory: jRow.exFactory || '-',
          augustExFactory: aRow.exFactory || '-',
          exFactoryDiff,
          julyQty: jRow.qty,
          augustQty: aRow.qty,
          qtyDiff,
          qtyDiffStr,
          status: isSame ? 'Same' : 'Changed',
          buyDate: aRow.buyDate || jRow.buyDate || '',
          itemNumber: aRow.itemNumber || jRow.itemNumber || '',
        });
        break;
      }
    }
  }

  // Pass 3: Match by PCD and EX FACTORY (Quantity changed)
  for (let i = 0; i < julyRows.length; i++) {
    if (julyUsed[i]) continue;
    const jRow = julyRows[i];

    for (let j = 0; j < augustRows.length; j++) {
      if (augustUsed[j]) continue;
      const aRow = augustRows[j];

      if (jRow.pcd && jRow.pcd === aRow.pcd && jRow.exFactory && jRow.exFactory === aRow.exFactory) {
        julyUsed[i] = true;
        augustUsed[j] = true;

        const pcdDiff = '0 days';
        const exFactoryDiff = '0 days';
        const qtyDiff = aRow.qty - jRow.qty;
        const qtyDiffStr = qtyDiff > 0 ? `+${qtyDiff}` : `${qtyDiff}`;

        schedulePairs.push({
          style,
          description: aRow.description || jRow.description || defaultDesc,
          julyPcd: jRow.pcd,
          augustPcd: aRow.pcd,
          pcdDiff,
          julyExFactory: jRow.exFactory,
          augustExFactory: aRow.exFactory,
          exFactoryDiff,
          julyQty: jRow.qty,
          augustQty: aRow.qty,
          qtyDiff,
          qtyDiffStr,
          status: qtyDiff === 0 ? 'Same' : 'Changed',
          buyDate: aRow.buyDate || jRow.buyDate || '',
          itemNumber: aRow.itemNumber || jRow.itemNumber || '',
        });
        break;
      }
    }
  }

  // Pass 4: Match by PCD or EX FACTORY (Date shifted)
  for (let i = 0; i < julyRows.length; i++) {
    if (julyUsed[i]) continue;
    const jRow = julyRows[i];

    for (let j = 0; j < augustRows.length; j++) {
      if (augustUsed[j]) continue;
      const aRow = augustRows[j];

      if (
        (jRow.pcd && jRow.pcd === aRow.pcd) ||
        (jRow.exFactory && jRow.exFactory === aRow.exFactory)
      ) {
        julyUsed[i] = true;
        augustUsed[j] = true;

        const pcdDiff = calculateDateDiff(jRow.pcd, aRow.pcd);
        const exFactoryDiff = calculateDateDiff(jRow.exFactory, aRow.exFactory);
        const qtyDiff = aRow.qty - jRow.qty;
        const qtyDiffStr = qtyDiff > 0 ? `+${qtyDiff}` : `${qtyDiff}`;
        const isSame = pcdDiff === '0 days' && exFactoryDiff === '0 days' && qtyDiff === 0;

        schedulePairs.push({
          style,
          description: aRow.description || jRow.description || defaultDesc,
          julyPcd: jRow.pcd || '-',
          augustPcd: aRow.pcd || '-',
          pcdDiff,
          julyExFactory: jRow.exFactory || '-',
          augustExFactory: aRow.exFactory || '-',
          exFactoryDiff,
          julyQty: jRow.qty,
          augustQty: aRow.qty,
          qtyDiff,
          qtyDiffStr,
          status: isSame ? 'Same' : 'Changed',
          buyDate: aRow.buyDate || jRow.buyDate || '',
          itemNumber: aRow.itemNumber || jRow.itemNumber || '',
        });
        break;
      }
    }
  }

  // Pass 5: Proximity date matching for remaining rows
  for (let i = 0; i < julyRows.length; i++) {
    if (julyUsed[i]) continue;
    const jRow = julyRows[i];

    let bestJ = -1;
    let minDistance = Infinity;

    for (let j = 0; j < augustRows.length; j++) {
      if (augustUsed[j]) continue;
      const aRow = augustRows[j];

      const pcd1 = parseDate(jRow.pcd);
      const pcd2 = parseDate(aRow.pcd);
      const ex1 = parseDate(jRow.exFactory);
      const ex2 = parseDate(aRow.exFactory);

      let dist = 0;
      let count = 0;
      if (pcd1 && pcd2) {
        dist += Math.abs(pcd2.getTime() - pcd1.getTime());
        count++;
      }
      if (ex1 && ex2) {
        dist += Math.abs(ex2.getTime() - ex1.getTime());
        count++;
      }

      if (count > 0) {
        const avgDistDays = (dist / count) / (1000 * 60 * 60 * 24);
        if (avgDistDays < 90 && avgDistDays < minDistance) {
          minDistance = avgDistDays;
          bestJ = j;
        }
      }
    }

    if (bestJ !== -1) {
      julyUsed[i] = true;
      augustUsed[bestJ] = true;
      const aRow = augustRows[bestJ];

      const pcdDiff = calculateDateDiff(jRow.pcd, aRow.pcd);
      const exFactoryDiff = calculateDateDiff(jRow.exFactory, aRow.exFactory);
      const qtyDiff = aRow.qty - jRow.qty;
      const qtyDiffStr = qtyDiff > 0 ? `+${qtyDiff}` : `${qtyDiff}`;
      const isSame = pcdDiff === '0 days' && exFactoryDiff === '0 days' && qtyDiff === 0;

      schedulePairs.push({
        style,
        description: aRow.description || jRow.description || defaultDesc,
        julyPcd: jRow.pcd || '-',
        augustPcd: aRow.pcd || '-',
        pcdDiff,
        julyExFactory: jRow.exFactory || '-',
        augustExFactory: aRow.exFactory || '-',
        exFactoryDiff,
        julyQty: jRow.qty,
        augustQty: aRow.qty,
        qtyDiff,
        qtyDiffStr,
        status: isSame ? 'Same' : 'Changed',
        buyDate: aRow.buyDate || jRow.buyDate || '',
        itemNumber: aRow.itemNumber || jRow.itemNumber || '',
      });
    }
  }

  // Pass 6: Unmatched July rows -> REMOVED
  for (let i = 0; i < julyRows.length; i++) {
    if (julyUsed[i]) continue;
    const jRow = julyRows[i];
    const qtyDiff = -jRow.qty;
    schedulePairs.push({
      style,
      description: jRow.description || defaultDesc,
      julyPcd: jRow.pcd || '-',
      augustPcd: '-',
      pcdDiff: 'Removed',
      julyExFactory: jRow.exFactory || '-',
      augustExFactory: '-',
      exFactoryDiff: 'Removed',
      julyQty: jRow.qty,
      augustQty: null,
      qtyDiff,
      qtyDiffStr: `${qtyDiff}`,
      status: 'Removed',
      buyDate: jRow.buyDate || '',
      itemNumber: jRow.itemNumber || '',
    });
  }

  // Pass 7: Unmatched August rows -> ADDED
  for (let j = 0; j < augustRows.length; j++) {
    if (augustUsed[j]) continue;
    const aRow = augustRows[j];
    const qtyDiff = aRow.qty;
    schedulePairs.push({
      style,
      description: aRow.description || defaultDesc,
      julyPcd: '-',
      augustPcd: aRow.pcd || '-',
      pcdDiff: 'New',
      julyExFactory: '-',
      augustExFactory: aRow.exFactory || '-',
      exFactoryDiff: 'New',
      julyQty: null,
      augustQty: aRow.qty,
      qtyDiff,
      qtyDiffStr: `+${qtyDiff}`,
      status: 'Added',
      buyDate: aRow.buyDate || '',
      itemNumber: aRow.itemNumber || '',
    });
  }

  return schedulePairs;
}

function classifyStyle({ qtyDelta, julyQty, pcdChanged, exFactoryChanged }) {
  if (!pcdChanged && !exFactoryChanged && qtyDelta === 0) {
    return 'LOW';
  }
  const percentChange = julyQty > 0 ? Math.abs(qtyDelta) / julyQty : (qtyDelta !== 0 ? 1 : 0);
  if (exFactoryChanged || percentChange >= 0.2) {
    return 'HIGH';
  }
  if (pcdChanged || qtyDelta !== 0) {
    return 'MEDIUM';
  }
  return 'LOW';
}

function compareSnapshots(julyRows, augustRows) {
  const julyMap = groupByStyle(julyRows);
  const augustMap = groupByStyle(augustRows);

  const allStyles = new Set([...julyMap.keys(), ...augustMap.keys()]);
  const styleResults = [];
  const allSchedules = [];

  let totalSameSchedules = 0;
  let totalChangedSchedules = 0;
  let totalAddedSchedules = 0;
  let totalRemovedSchedules = 0;

  for (const style of allStyles) {
    const jul = julyMap.get(style);
    const aug = augustMap.get(style);

    const julyQty = jul ? jul.totalQty : 0;
    const augustQty = aug ? aug.totalQty : 0;
    const qtyDelta = augustQty - julyQty;
    const description = (aug && aug.description) || (jul && jul.description) || '';

    const julRows = jul ? jul.rows : [];
    const augRows = aug ? aug.rows : [];

    const schedules = matchSchedulesForStyle(style, description, julRows, augRows);
    allSchedules.push(...schedules);

    let stylePcdChanged = false;
    let styleExFactoryChanged = false;

    for (const sched of schedules) {
      if (sched.status === 'Same') totalSameSchedules++;
      else if (sched.status === 'Changed') totalChangedSchedules++;
      else if (sched.status === 'Added') totalAddedSchedules++;
      else if (sched.status === 'Removed') totalRemovedSchedules++;

      if (sched.pcdDiff !== '0 days') stylePcdChanged = true;
      if (sched.exFactoryDiff !== '0 days') styleExFactoryChanged = true;
    }

    const presenceChange = !jul ? 'NEW_IN_AUGUST' : !aug ? 'REMOVED_IN_AUGUST' : 'PRESENT_IN_BOTH';
    const priorityStatus = classifyStyle({
      qtyDelta,
      julyQty,
      pcdChanged: stylePcdChanged,
      exFactoryChanged: styleExFactoryChanged,
    });

    styleResults.push({
      style,
      description,
      julyQty,
      augustQty,
      qtyDelta,
      pcdChanged: stylePcdChanged,
      exFactoryChanged: styleExFactoryChanged,
      presenceChange,
      status: priorityStatus,
      julySchedule: julRows,
      augustSchedule: augRows,
      schedules,
    });
  }

  // Sort styles alphabetically by name
  styleResults.sort((a, b) => a.style.localeCompare(b.style));

  const summary = {
    totalStyles: allStyles.size,
    changedStyles: styleResults.filter((s) => s.qtyDelta !== 0 || s.pcdChanged || s.exFactoryChanged || s.schedules.some(sched => sched.status !== 'Same')).length,
    qtyChanged: styleResults.filter((s) => s.qtyDelta !== 0).length,
    pcdChanged: styleResults.filter((s) => s.pcdChanged).length,
    exFactoryChanged: styleResults.filter((s) => s.exFactoryChanged).length,
    highPriority: styleResults.filter((s) => s.status === 'HIGH').length,
    mediumPriority: styleResults.filter((s) => s.status === 'MEDIUM').length,
    lowPriority: styleResults.filter((s) => s.status === 'LOW').length,
    totalSchedules: allSchedules.length,
    sameSchedules: totalSameSchedules,
    changedSchedules: totalChangedSchedules,
    addedSchedules: totalAddedSchedules,
    removedSchedules: totalRemovedSchedules,
  };

  return { summary, styles: styleResults, allSchedules, dimensionRows: buildDimensionRows(julyRows, augustRows) };
}

function buildDimensionRows(julyRows, augustRows) {
  const totals = new Map();
  const addRows = (rows, snapshot) => {
    for (const row of rows) {
      const season = row.season || 'All';
      const month = row.month || 'Unknown';
      const key = `${row.style}\u0000${season}\u0000${month}`;
      if (!totals.has(key)) totals.set(key, { style: row.style, season, month, julyQty: 0, augustQty: 0 });
      totals.get(key)[snapshot] += row.qty;
    }
  };
  addRows(julyRows, 'julyQty');
  addRows(augustRows, 'augustQty');
  return Array.from(totals.values()).map((row) => ({ ...row, variation: row.augustQty - row.julyQty }));
}

module.exports = { compareSnapshots, calculateDateDiff, parseDate };

