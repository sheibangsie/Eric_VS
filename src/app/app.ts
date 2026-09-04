import { CommonModule } from '@angular/common';
import { Component, computed, signal } from '@angular/core';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';

type TransactionType = 'Income' | 'Expense';

interface Transaction {
  id: number;
  date: string;
  description: string;
  category: string;
  type: TransactionType;
  amount: number;
}

interface NewTransaction {
  date: string;
  description: string;
  category: string;
  type: TransactionType;
  amount: number | null;
}

@Component({
  imports: [CommonModule, FormsModule],
  selector: 'app-root',
  styleUrl: './app.css',
  templateUrl: './app.html',
})
export class App {
  private readonly apiUrl = ['localhost', '127.0.0.1'].includes(window.location.hostname)
    ? 'http://localhost:3000/api/transactions'
    : null;
  protected readonly activeSection = signal('Overview');
  protected readonly monthLabel = 'September 2026';
  protected readonly categories = ['Housing', 'Food', 'Transport', 'Lifestyle', 'Bills', 'Health'];
  protected readonly transactions = signal<Transaction[]>([
    { id: 1, date: '2026-09-04', description: 'Salary deposit', category: 'Income', type: 'Income', amount: 5200 },
    { id: 2, date: '2026-09-03', description: 'Apartment rent', category: 'Housing', type: 'Expense', amount: 1650 },
    { id: 3, date: '2026-09-02', description: 'Weekly groceries', category: 'Food', type: 'Expense', amount: 86.45 },
    { id: 4, date: '2026-09-01', description: 'Metro pass', category: 'Transport', type: 'Expense', amount: 72 },
    { id: 5, date: '2026-08-30', description: 'Streaming bundle', category: 'Lifestyle', type: 'Expense', amount: 24.99 },
  ]);
  protected readonly newTransaction = signal<NewTransaction>(this.emptyTransaction());
  protected readonly budget = signal(3800);
  protected readonly totalIncome = computed(() => this.transactions().filter((item) => item.type === 'Income').reduce((sum, item) => sum + item.amount, 0));
  protected readonly totalSpent = computed(() => this.transactions().filter((item) => item.type === 'Expense').reduce((sum, item) => sum + item.amount, 0));
  protected readonly balance = computed(() => this.totalIncome() - this.totalSpent());
  protected readonly budgetProgress = computed(() => Math.min(100, (this.totalSpent() / this.budget()) * 100));

  constructor() {
    const saved = localStorage.getItem('ledger-transactions');
    if (saved) {
      try { this.transactions.set(JSON.parse(saved)); } catch { localStorage.removeItem('ledger-transactions'); }
    }
    const savedBudget = Number(localStorage.getItem('ledger-budget'));
    if (savedBudget > 0) this.budget.set(savedBudget);
    this.loadFromApi();
  }

  protected updateBudget(value: string | number): void {
    const amount = Number(value);
    if (amount > 0) {
      this.budget.set(amount);
      localStorage.setItem('ledger-budget', String(amount));
    }
  }

  protected editBudget(): void {
    const value = window.prompt('Enter your monthly budget', String(this.budget()));
    if (value !== null) this.updateBudget(value);
  }

  protected updateField(field: keyof NewTransaction, value: string | number | null): void {
    this.newTransaction.update((form) => {
      if (field === 'type' && value === 'Income') return { ...form, type: 'Income', category: 'Income' };
      if (field === 'type' && value === 'Expense') return { ...form, type: 'Expense', category: form.category === 'Income' ? 'Food' : form.category };
      return { ...form, [field]: value };
    });
  }

  protected selectSection(section: string): void {
    this.activeSection.set(section);
  }

  protected addTransaction(): void {
    const entry = this.newTransaction();
    if (!entry.description.trim() || !entry.date || !entry.category || !entry.amount || entry.amount <= 0) return;
    this.transactions.update((items) => [{ ...entry, id: Date.now(), description: entry.description.trim(), amount: Number(entry.amount) }, ...items]);
    this.persist();
    this.syncToApi();
    this.newTransaction.set(this.emptyTransaction());
  }

  protected removeTransaction(id: number): void {
    this.transactions.update((items) => items.filter((item) => item.id !== id));
    this.persist();
    if (this.apiUrl) void fetch(`${this.apiUrl}/${id}`, { method: 'DELETE' }).catch(() => undefined);
  }

  protected importWorkbook(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const workbook = XLSX.read(reader.result, { type: 'array' });
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[workbook.SheetNames[0]]);
      const imported: Transaction[] = rows.map((row, index) => ({
        id: Date.now() + index,
        date: String(row['Date'] ?? new Date().toISOString().slice(0, 10)),
        description: String(row['Description'] ?? 'Imported transaction'),
        category: String(row['Category'] ?? 'Other'),
        type: String(row['Type'] ?? 'Expense') === 'Income' ? 'Income' as TransactionType : 'Expense' as TransactionType,
        amount: Number(row['Amount'] ?? 0),
      })).filter((item) => item.amount > 0);
      this.transactions.set([...imported, ...this.transactions()]);
      this.persist();
      this.syncToApi();
      input.value = '';
    };
    reader.readAsArrayBuffer(file);
  }

  protected exportWorkbook(): void {
    const rows = this.transactions().map(({ id, ...item }) => ({ Date: item.date, Description: item.description, Category: item.category, Type: item.type, Amount: item.amount }));
    const sheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, 'Transactions');
    XLSX.writeFile(workbook, 'ledger-transactions.xlsx');
  }

  protected categoryTotal(category: string): number {
    return this.transactions().filter((item) => item.category === category && item.type === 'Expense').reduce((sum, item) => sum + item.amount, 0);
  }

  private persist(): void { localStorage.setItem('ledger-transactions', JSON.stringify(this.transactions())); }
  private loadFromApi(): void {
    if (!this.apiUrl) return;
    void fetch(this.apiUrl).then((response) => response.ok ? response.json() : Promise.reject()).then((items: Transaction[]) => {
      this.transactions.set(items);
      this.persist();
    }).catch(() => undefined);
  }
  private syncToApi(): void {
    if (!this.apiUrl) return;
    void fetch(this.apiUrl, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(this.transactions()) }).catch(() => undefined);
  }
  private emptyTransaction(): NewTransaction { return { date: new Date().toISOString().slice(0, 10), description: '', category: 'Food', type: 'Expense', amount: null }; }
}
