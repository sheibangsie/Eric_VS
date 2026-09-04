const { createServer } = require('node:http');
const { existsSync, mkdirSync } = require('node:fs');
const { readFile, writeFile } = require('node:fs/promises');
const path = require('node:path');
const XLSX = require('xlsx');

const port = 3000;
const workbookPath = path.join(__dirname, 'data', 'transactions.xlsx');
const seed = [
  { id: 1, date: '2026-09-04', description: 'Salary deposit', category: 'Income', type: 'Income', amount: 5200 },
  { id: 2, date: '2026-09-03', description: 'Apartment rent', category: 'Housing', type: 'Expense', amount: 1650 },
  { id: 3, date: '2026-09-02', description: 'Weekly groceries', category: 'Food', type: 'Expense', amount: 86.45 },
  { id: 4, date: '2026-09-01', description: 'Metro pass', category: 'Transport', type: 'Expense', amount: 72 },
  { id: 5, date: '2026-08-30', description: 'Streaming bundle', category: 'Lifestyle', type: 'Expense', amount: 24.99 },
];

function send(response, status, body) {
  response.writeHead(status, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': 'http://localhost:4200', 'Access-Control-Allow-Headers': 'Content-Type', 'Access-Control-Allow-Methods': 'GET,PUT,DELETE,OPTIONS' });
  response.end(JSON.stringify(body));
}

async function readTransactions() {
  if (!existsSync(workbookPath)) {
    await writeTransactions(seed);
    return seed;
  }
  const workbook = XLSX.read(await readFile(workbookPath), { type: 'buffer' });
  return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
}

async function writeTransactions(transactions) {
  mkdirSync(path.dirname(workbookPath), { recursive: true });
  const sheet = XLSX.utils.json_to_sheet(transactions);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'Transactions');
  await writeFile(workbookPath, XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }));
}

const server = createServer(async (request, response) => {
  if (request.method === 'OPTIONS') return send(response, 204, {});
  if (!request.url?.startsWith('/api/transactions')) return send(response, 404, { error: 'Not found' });
  try {
    let transactions = await readTransactions();
    if (request.method === 'GET') return send(response, 200, transactions);
    if (request.method === 'PUT') {
      let body = '';
      for await (const chunk of request) body += chunk;
      transactions = JSON.parse(body);
      await writeTransactions(transactions);
      return send(response, 200, transactions);
    }
    if (request.method === 'DELETE') {
      const id = Number(request.url.split('/').pop());
      await writeTransactions(transactions.filter((item) => Number(item.id) !== id));
      return send(response, 204, {});
    }
    return send(response, 405, { error: 'Method not allowed' });
  } catch (error) {
    console.error(error);
    return send(response, 500, { error: 'Could not read or write the workbook' });
  }
});

server.listen(port, () => console.log(`Ledger API running at http://localhost:${port}`));
