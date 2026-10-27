/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ['./src/**/*.{html,ts}'],
  theme: {
    extend: {
      colors: {
        shahi: {
          navy: '#102a43',
          blue: '#1e4e76',
          teal: '#168c8c',
          canvas: '#f3f6f8'
        }
      }
    }
  },
  plugins: []
};
