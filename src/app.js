import * as fs from 'fs';
import xlsx from 'xlsx';

const data = [
  ['name', 'email', 'km', 'coefficient', 'total_value', 'phone', 'status'],
  ['Loan Matteus', 'loan@email.com', 3021.510, 1.2, 0.0, '99 9 9999-9999', 'active'],
  ['Niulanio Bezerra', 'niuxp@email.com', 3120.054, 1.2, 0.0, '99 9 9999-9999', 'active'],
  ['Andressa Pereira', 'dessa@email.com', 3030.05, 1.2, 0.0, '99 9 9999-9999', 'active'],
  ['Pedro Felipe', 'ped@email.com', 240.074, 1.2, 0.0, '99 9 9999-9999', 'active'],
  ['Brencarla Medeiros', 'brenm@email.com', 3000.220114, 1.2, 0.0, '99 9 9999-9999', 'active'],
  ['Matheus Henrique', 'matheus_777@email.com', 30.08464, 1.2, 0.0, '99 9 9999-9999', 'active'],
  ['Naiara Lopes', 'nai@email.com', 580.049, 1.2, 0.0, '99 9 9999-9999', 'active'],
  ['Pedro Antônio', 'ksbushink@email.com', 7.01897, 1.2, 0.0, '99 9 9999-9999', 'active'],
  ['Antoniel Pereira', 'antoin@email.com', 3570.015104, 1.2, 0.0, '99 9 9999-9999', 'active'],
  ['Mateus Alex', 'fakedojao@email.com', 470.015104, 1.2, 0.0, '99 9 9999-9999', 'active'],
];

data.forEach(data => {
  if (typeof data[4] === 'number') {
    data[4] = data[2] * data[3];
  }
});

const workbook = xlsx.utils.book_new();
workbook.SheetNames.push('refunds');

const ws = xlsx.utils.aoa_to_sheet(data);
const collumnE = data.map((_, index) => (`E${index+1}`));
collumnE.shift(); // Item #1 é o header do xlsx

for (const property in ws) {
  if (collumnE.includes(property)) {
    ws[property].z = '0.00';
  }
}

workbook.Sheets["refunds"] = ws;

const wbuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
fs.writeFileSync('refunds.xlsx', wbuffer);

console.log({ ws });
