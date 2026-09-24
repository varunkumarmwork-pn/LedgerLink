# LedgerLink (front end)

Screens: client list, add client, processing, workbook review.
Sample data lives in `src/data/sample.js`. The Python backend comes next.

## Run it
1. Install Node.js LTS from https://nodejs.org
2. Open this folder in VS Code
3. Terminal > New Terminal, then:
   npm install
   npm run dev
4. The browser opens at http://localhost:5173

## Where things are
- src/styles.css            colours, fonts, spacing (all tokens at the top)
- src/components/Header.jsx top bar
- src/pages/Clients.jsx     client table, hover delete, type DELETE dialog
- src/pages/AddClient.jsx   two choice boxes, file picker
- src/pages/Processing.jsx  progress steps (simulated for now)
- src/pages/Workbook.jsx    Excel-style grid, formula bar, review panel
