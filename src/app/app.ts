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
  subcategory: string;
  type: TransactionType;
  amount: number;
}

interface NewTransaction {
  date: string;
  description: string;
  category: string;
  subcategory: string;
  type: TransactionType;
  amount: number | null;
}

interface CategoryGroup {
  name: string;
  subcategories: string[];
}

interface SubcategoryTrend {
  name: string;
  total: number;
  points: number[];
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
  protected readonly selectedMonth = signal('2026-09');
  protected readonly monthOptions = computed(() => [...new Set(this.transactions().map((item) => item.date.slice(0, 7)))].sort().reverse());
  protected readonly monthLabel = computed(() => this.formatMonth(this.selectedMonth()));
  protected readonly categoryGroups = signal<CategoryGroup[]>([
    { name: 'Housing', subcategories: ['Rent', 'Utilities', 'Repairs'] },
    { name: 'Food', subcategories: ['Groceries', 'Restaurants', 'Coffee'] },
    { name: 'Transport', subcategories: ['Commute', 'Fuel', 'Parking'] },
    { name: 'Lifestyle', subcategories: ['Entertainment', 'Shopping', 'Subscriptions'] },
    { name: 'Bills', subcategories: ['Phone', 'Internet', 'Insurance'] },
    { name: 'Health', subcategories: ['Medicine', 'Appointments', 'Fitness'] },
  ]);
  protected get categories(): string[] { return this.categoryGroups().map((group) => group.name); }
  protected readonly transactions = signal<Transaction[]>([]);
  protected readonly newTransaction = signal<NewTransaction>(this.emptyTransaction());
  protected readonly budget = signal(3800);
  protected readonly selectedTransactions = computed(() => this.transactions().filter((item) => item.date.startsWith(this.selectedMonth())));
  protected readonly totalIncome = computed(() => this.selectedTransactions().filter((item) => item.type === 'Income').reduce((sum, item) => sum + item.amount, 0));
  protected readonly totalSpent = computed(() => this.selectedTransactions().filter((item) => item.type === 'Expense').reduce((sum, item) => sum + item.amount, 0));
  protected readonly balance = computed(() => this.totalIncome() - this.totalSpent());
  protected readonly budgetProgress = computed(() => Math.min(100, (this.totalSpent() / this.budget()) * 100));
  protected readonly reportCategories = computed(() => this.categories
    .map((category) => ({ category, total: this.categoryTotal(category) }))
    .filter((item) => item.total > 0)
    .sort((first, second) => second.total - first.total));
  protected readonly reportExpenseCount = computed(() => this.transactions().filter((item) => item.type === 'Expense').length);
  protected readonly reportSavingsRate = computed(() => this.totalIncome() > 0 ? (this.balance() / this.totalIncome()) * 100 : 0);
  protected readonly budgetLeft = computed(() => Math.max(0, this.budget() - this.totalSpent()));
  protected readonly savingsProgress = computed(() => Math.min(100, this.reportSavingsRate()));
  protected readonly savingsFundGoal = 10000;
  protected readonly savingsFundProgress = computed(() => Math.min(100, Math.max(0, (this.balance() / this.savingsFundGoal) * 100)));
  protected readonly trendData = computed(() => this.monthOptions().slice(0, 6).reverse().map((month) => ({
    month: new Intl.DateTimeFormat('en-US', { month: 'short', timeZone: 'UTC' }).format(new Date(`${month}-01T00:00:00Z`)),
    amount: this.transactions().filter((item) => item.type === 'Expense' && item.date.startsWith(month)).reduce((sum, item) => sum + item.amount, 0),
  })));
  protected readonly subcategoryTrends = computed<SubcategoryTrend[]>(() => {
    const months = this.monthOptions().slice(0, 6).reverse();
    const totals = new Map<string, number[]>();
    this.transactions().filter((item) => item.type === 'Expense' && item.subcategory).forEach((item) => {
      const points = totals.get(item.subcategory) ?? months.map(() => 0);
      const monthIndex = months.indexOf(item.date.slice(0, 7));
      if (monthIndex >= 0) points[monthIndex] += item.amount;
      totals.set(item.subcategory, points);
    });
    return [...totals.entries()]
      .map(([name, points]) => ({ name, total: points.reduce((sum, amount) => sum + amount, 0), points }))
      .sort((first, second) => second.total - first.total)
      .slice(0, 2)
      .map((trend) => {
        const maximum = Math.max(...trend.points, 1);
        return { ...trend, points: trend.points.map((point) => point ? Math.max(12, (point / maximum) * 100) : 4) };
      });
  });
  protected readonly newCategoryName = signal('');
  protected readonly newCategorySubcategory = signal('');
  protected readonly selectedCategoryForSubcategory = signal('');
  protected readonly newSubcategoryName = signal('');

  constructor() {
    const saved = localStorage.getItem('ledger-transactions');
    if (saved) {
      try { this.transactions.set(JSON.parse(saved)); } catch { localStorage.removeItem('ledger-transactions'); }
    }
    const savedBudget = Number(localStorage.getItem('ledger-budget'));
    if (savedBudget > 0) this.budget.set(savedBudget);
    const savedCategories = localStorage.getItem('ledger-categories');
    if (savedCategories) {
      try { this.categoryGroups.set(JSON.parse(savedCategories)); } catch { localStorage.removeItem('ledger-categories'); }
    }
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
      if (field === 'type' && value === 'Expense') return { ...form, type: 'Expense', category: form.category === 'Income' ? this.categories[0] : form.category, subcategory: form.subcategory || this.subcategoriesFor(form.category === 'Income' ? this.categories[0] : form.category)[0] || '' };
      if (field === 'category') return { ...form, category: String(value), subcategory: this.subcategoriesFor(String(value))[0] || '' };
      return { ...form, [field]: value };
    });
  }

  protected subcategoriesFor(category: string): string[] {
    return this.categoryGroups().find((group) => group.name === category)?.subcategories ?? [];
  }

  protected addCategory(name: string, subcategory: string): void {
    const categoryName = name.trim();
    const subcategoryName = subcategory.trim();
    if (!categoryName || this.categories.some((category) => category.toLowerCase() === categoryName.toLowerCase())) return;
    this.categoryGroups.update((groups) => [...groups, { name: categoryName, subcategories: subcategoryName ? [subcategoryName] : [] }]);
    this.persistCategories();
  }

  protected addSubcategory(category: string, subcategory: string): void {
    const subcategoryName = subcategory.trim();
    if (!category || !subcategoryName) return;
    this.categoryGroups.update((groups) => groups.map((group) => group.name === category && !group.subcategories.some((item) => item.toLowerCase() === subcategoryName.toLowerCase()) ? { ...group, subcategories: [...group.subcategories, subcategoryName] } : group));
    this.persistCategories();
  }

  protected createCategory(): void {
    this.addCategory(this.newCategoryName(), this.newCategorySubcategory());
    this.newCategoryName.set('');
    this.newCategorySubcategory.set('');
  }

  protected createSubcategory(): void {
    this.addSubcategory(this.selectedCategoryForSubcategory(), this.newSubcategoryName());
    this.newSubcategoryName.set('');
  }

  protected selectSection(section: string): void {
    this.activeSection.set(section);
  }

  protected selectMonth(month: string): void {
    this.selectedMonth.set(month);
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
        subcategory: String(row['Subcategory'] ?? ''),
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
    const rows = this.transactions().map(({ id, ...item }) => ({ Date: item.date, Description: item.description, Category: item.category, Subcategory: item.subcategory, Type: item.type, Amount: item.amount }));
    const sheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, 'Transactions');
    XLSX.writeFile(workbook, 'ledger-transactions.xlsx');
  }

  protected categoryTotal(category: string): number {
    return this.selectedTransactions().filter((item) => item.category === category && item.type === 'Expense').reduce((sum, item) => sum + item.amount, 0);
  }

  protected formatMonth(month: string): string {
    return new Intl.DateTimeFormat('en-US', { month: 'long', year: 'numeric', timeZone: 'UTC' }).format(new Date(`${month}-01T00:00:00Z`));
  }

  private persist(): void { localStorage.setItem('ledger-transactions', JSON.stringify(this.transactions())); }
  private persistCategories(): void { localStorage.setItem('ledger-categories', JSON.stringify(this.categoryGroups())); }
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
  private emptyTransaction(): NewTransaction { return { date: new Date().toISOString().slice(0, 10), description: '', category: 'Food', subcategory: 'Groceries', type: 'Expense', amount: null }; }
}
