const XLSX = require('xlsx');

const round = (num) => Math.round(num * 1000) / 1000;

function extractWedstrijdData(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const dataRows = rows.slice(2);

  const colIdx = {
    name: 1,
    club: 2,
    category: 3, 
    exercise: 4,
    dAScore: 5,
    dBScore: 6,
    aScore: 12,
    eScore: 17,
    aftrek: 20,
    subtotal: 22,
    total: 23,
  };

  const categories = {};

  dataRows.forEach(row => {
    const name = row[colIdx.name];
    const exercise = row[colIdx.exercise].toLowerCase();
    const category = row[colIdx.category];

    if (!name || name.trim() === '-' || !exercise || exercise.trim() === '-' || !category || category.trim() === '-') return;

    const gymnastKey = `${name}_${row[colIdx.club]}`;

    const gymnastEntry = categories[category]?.[gymnastKey] || {
      name: name,
      club: row[colIdx.club],
      exercises: {},
      total: round(parseFloat(String(row[colIdx.total]).replace(',', '.')) || 0),
    };

    gymnastEntry.exercises[exercise] = {
      dAScore: round(parseFloat(String(row[colIdx.dAScore]).replace(',', '.')) || 0),
      dBScore: round(parseFloat(String(row[colIdx.dBScore]).replace(',', '.')) || 0),
      aScore: round(parseFloat(String(row[colIdx.aScore]).replace(',', '.')) || 0),
      eScore: round(parseFloat(String(row[colIdx.eScore]).replace(',', '.')) || 0),
      aftrek: round(parseFloat(String(row[colIdx.aftrek]).replace(',', '.')) || 0),
      subtotal: round(parseFloat(String(row[colIdx.subtotal]).replace(',', '.')) || 0),
    };

    if (!categories[category]) categories[category] = {};
    categories[category][gymnastKey] = gymnastEntry;
  });

  Object.keys(categories).forEach(category => {
    const gymnasts = Object.values(categories[category]);

    gymnasts.sort((a, b) => b.total - a.total);

    gymnasts.forEach((gymnast, index) => {
      gymnast.rank = index + 1;
    });

    categories[category] = gymnasts;
  });

  return categories;
}

function extractWorkbookData(filepath) {
  const workbook = XLSX.readFile(filepath);
  const wedstrijdSheet = workbook.Sheets['Wedstrijd'];

  if (!wedstrijdSheet) {
    throw new Error("No 'Wedstrijd' sheet found!");
  }

  return extractWedstrijdData(wedstrijdSheet);
}

const pathToExcel = process.argv[2];

if (!pathToExcel) {
  console.error("Usage: node script.js <excel-file>");
  process.exit(1);
}
const matchData = extractWorkbookData(pathToExcel);
//console.log(JSON.stringify(matchData, null, 2));

function generateHTML(wedstrijdData) {
  let html = `<html><head>
  <style>
    table { border-collapse: collapse; margin-bottom: 20px; width: 100%; }
    th, td { border: 1px solid #ccc; padding: 8px; text-align: center; }
    th { background-color: #f0f0f0; }
  </style></head><body>`;

  Object.entries(wedstrijdData).forEach(([category, gymnasts]) => {
    // Determine unique exercises for the header
    const exercisesSet = new Set();
    gymnasts.forEach(g => Object.keys(g.exercises).forEach(e => exercisesSet.add(e)));
    const exercises = Array.from(exercisesSet);

    html += `<h2>${category}</h2><table><tr>
      <th></th>
      <th>Deelnemer</th>`;
    exercises.forEach(exercise => {
      html += `<th>${exercise}</th>`;
    });
    html += `<th>Total</th></tr>`;

    gymnasts.forEach(gymnast => {
      html += `<tr><td>${gymnast.rank}</td>`;
      html += `<td>${gymnast.name}<br/><small>${gymnast.club}</small></td>`;

      exercises.forEach(exercise => {
        const ex = gymnast.exercises[exercise];
        if (ex) {
          html += `<td style="white-space: pre-line;">
            D: ${ex.dAScore?.toFixed(3)}/${ex.dBScore?.toFixed(3)}
            A: ${ex.aScore.toFixed(3)}
            E: ${ex.eScore.toFixed(3)}
            P: ${ex.aftrek.toFixed(3)}
            ${ex.subtotal.toFixed(3)}
          </td>`;
        } else {
          html += `<td>-</td>`;
        }
      });

      html += `<td>${gymnast.total.toFixed(3)}</td></tr>`;
    });

    html += `</table>`;
  });

  html += `</body></html>`;
  return html;
}

const fs = require('fs');
const htmlContent = generateHTML(matchData);
fs.writeFileSync('wedstrijd_results.html', htmlContent);

