import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx-js-style";

type Tab = "dashboard" | "expenses" | "income" | "savings";

interface Transaction {
  id: string;
  date: string;
  amount: number;
  category: string;
  description: string;
  note: string;
}

interface SavingsGoal {
  id: string;
  name: string;
  target: number;
  saved: number;
  deadline: string;
  color: string;
}

const EXPENSE_CATEGORIES = [
  "Makanan & Minuman",
  "Transportasi",
  "Belanja",
  "Tagihan & Utilitas",
  "Hiburan",
  "Kesehatan",
  "Pendidikan",
  "Tempat Tinggal",
  "Pakaian",
  "Lainnya",
];

const INCOME_CATEGORIES = [
  "Gaji",
  "Freelance",
  "Bisnis",
  "Investasi",
  "Hadiah",
  "Bonus",
  "Sewa",
  "Lainnya",
];

const GOAL_COLORS = [
  "#1A4A3A",
  "#C4623A",
  "#C8963E",
  "#4A6741",
  "#6B4A3A",
  "#3A4A6B",
];

function formatRupiah(amount: number) {
  return new Intl.NumberFormat("id-ID", {
    style: "currency",
    currency: "IDR",
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(amount);
}

function formatDate(dateStr: string) {
  return new Date(dateStr).toLocaleDateString("id-ID", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

function getCurrentMonth() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
}

function getMonthLabel(monthStr: string) {
  const [year, month] = monthStr.split("-");
  return new Date(parseInt(year), parseInt(month) - 1).toLocaleDateString("id-ID", {
    month: "long",
    year: "numeric",
  });
}

function downloadFinanceReport(
  expenses: Transaction[],
  income: Transaction[],
  selectedMonth: string
) {
  const transactions = [
    ...income.map((transaction) => ({ ...transaction, type: "Penghasilan", sign: 1 })),
    ...expenses.map((transaction) => ({ ...transaction, type: "Pengeluaran", sign: -1 })),
  ].sort((a, b) => a.date.localeCompare(b.date));
  const totalExpenses = expenses.reduce((sum, transaction) => sum + transaction.amount, 0);
  const totalIncome = income.reduce((sum, transaction) => sum + transaction.amount, 0);
  const balance = totalIncome - totalExpenses;
  const expenseByCategory = EXPENSE_CATEGORIES.map((category) => [
    category,
    expenses.filter((transaction) => transaction.category === category)
      .reduce((sum, transaction) => sum + transaction.amount, 0),
  ]).filter(([, amount]) => amount > 0);
  const incomeByCategory = INCOME_CATEGORIES.map((category) => [
    category,
    income.filter((transaction) => transaction.category === category)
      .reduce((sum, transaction) => sum + transaction.amount, 0),
  ]).filter(([, amount]) => amount > 0);

  const navy = "30445C";
  const lime = "D5E648";
  const paleGreen = "DDE9D8";
  const white = "FFFFFF";
  const muted = "F1F4F0";
  const border = { style: "thin", color: { rgb: "C8D1C5" } };
  const titleStyle = { font: { bold: true, color: { rgb: navy }, sz: 20 }, fill: { fgColor: { rgb: lime } }, alignment: { vertical: "center" } };
  const sectionStyle = { font: { bold: true, color: { rgb: white } }, fill: { fgColor: { rgb: navy } }, alignment: { vertical: "center" } };
  const headerStyle = { font: { bold: true, color: { rgb: navy } }, fill: { fgColor: { rgb: lime } }, border, alignment: { horizontal: "center", vertical: "center", wrapText: true } };
  const labelStyle = { font: { bold: true, color: { rgb: navy } }, fill: { fgColor: { rgb: paleGreen } }, border };
  const bodyStyle = { border, alignment: { vertical: "center" } };
  const moneyStyle = { ...bodyStyle, numFmt: '"Rp" #,##0;[Red]-"Rp" #,##0' };

  const reportRows: (string | number)[][] = [
    ["LAPORAN KEUANGAN PRIBADI"],
    [`Periode: ${getMonthLabel(selectedMonth)}`],
    [],
    ["RINGKASAN KEUANGAN"],
    ["Total Penghasilan", totalIncome],
    ["Total Pengeluaran", totalExpenses],
    ["Saldo Terakhir", balance],
    [],
    ["PENGHASILAN PER KATEGORI"],
    ["Kategori", "Jumlah"],
    ...incomeByCategory,
    ["Total Penghasilan", totalIncome],
    [],
    ["PENGELUARAN PER KATEGORI"],
    ["Kategori", "Jumlah"],
    ...expenseByCategory,
    ["Total Pengeluaran", totalExpenses],
  ];
  const reportSheet = XLSX.utils.aoa_to_sheet(reportRows);
  reportSheet["A1"].s = titleStyle;
  reportSheet["A2"].s = { font: { italic: true, color: { rgb: "5F6D63" } } };
  ["A4", "A9", `A${10 + incomeByCategory.length + 2}`].forEach((cell) => {
    if (reportSheet[cell]) reportSheet[cell].s = sectionStyle;
  });
  ["A10", `A${11 + incomeByCategory.length + 2}`].forEach((cell) => {
    if (reportSheet[cell]) reportSheet[cell].s = headerStyle;
  });
  ["A5", "A6", "A7"].forEach((cell) => { reportSheet[cell].s = labelStyle; });
  ["B5", "B6", "B7"].forEach((cell) => { reportSheet[cell].s = { ...moneyStyle, font: { bold: true, color: { rgb: navy } } }; });
  for (let row = 11; row <= 11 + incomeByCategory.length; row++) {
    reportSheet[`A${row}`].s = row === 11 + incomeByCategory.length ? labelStyle : bodyStyle;
    reportSheet[`B${row}`].s = { ...moneyStyle, font: row === 11 + incomeByCategory.length ? { bold: true, color: { rgb: navy } } : undefined };
  }
  const expenseHeaderRow = 12 + incomeByCategory.length;
  const expenseStartRow = expenseHeaderRow + 1;
  for (let row = expenseHeaderRow + 1; row <= expenseStartRow + expenseByCategory.length; row++) {
    reportSheet[`A${row}`].s = row === expenseStartRow + expenseByCategory.length ? labelStyle : bodyStyle;
    reportSheet[`B${row}`].s = { ...moneyStyle, font: row === expenseStartRow + expenseByCategory.length ? { bold: true, color: { rgb: navy } } : undefined };
  }
  reportSheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }, { s: { r: 3, c: 0 }, e: { r: 3, c: 1 } }, { s: { r: 8, c: 0 }, e: { r: 8, c: 1 } }, { s: { r: expenseHeaderRow - 2, c: 0 }, e: { r: expenseHeaderRow - 2, c: 1 } }];
  reportSheet["!cols"] = [{ wch: 29 }, { wch: 19 }];
  reportSheet["!rows"] = [{ hpt: 34 }, { hpt: 22 }];

  let runningBalance = 0;
  const rawRows = [
    ["Tanggal", "Jenis", "Kategori", "Deskripsi", "Catatan", "Jumlah (Rp)", "Saldo Berjalan (Rp)"],
    ...transactions.map((transaction) => {
      runningBalance += transaction.amount * transaction.sign;
      return [transaction.date, transaction.type, transaction.category, transaction.description, transaction.note, transaction.amount * transaction.sign, runningBalance];
    }),
  ];
  const rawSheet = XLSX.utils.aoa_to_sheet(rawRows);
  for (let column = 0; column < 7; column++) rawSheet[XLSX.utils.encode_cell({ r: 0, c: column })].s = headerStyle;
  for (let row = 1; row < rawRows.length; row++) {
    for (let column = 0; column < 5; column++) rawSheet[XLSX.utils.encode_cell({ r: row, c: column })].s = { ...bodyStyle, fill: { fgColor: { rgb: row % 2 ? white : muted } } };
    rawSheet[XLSX.utils.encode_cell({ r: row, c: 5 })].s = { ...moneyStyle, fill: { fgColor: { rgb: row % 2 ? white : muted } } };
    rawSheet[XLSX.utils.encode_cell({ r: row, c: 6 })].s = { ...moneyStyle, fill: { fgColor: { rgb: row % 2 ? white : muted } } };
  }
  rawSheet["!cols"] = [{ wch: 13 }, { wch: 16 }, { wch: 24 }, { wch: 30 }, { wch: 34 }, { wch: 16 }, { wch: 20 }];
  rawSheet["!autofilter"] = { ref: `A1:G${rawRows.length}` };

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, reportSheet, "Laporan Keuangan");
  XLSX.utils.book_append_sheet(workbook, rawSheet, "Data Mentah");
  XLSX.writeFile(workbook, `laporan-keuangan-${selectedMonth}.xlsx`);
}

function useLocalStorage<T>(key: string, initial: T): [T, (v: T) => void] {
  const [state, setState] = useState<T>(() => {
    try {
      const stored = localStorage.getItem(key);
      return stored ? JSON.parse(stored) : initial;
    } catch {
      return initial;
    }
  });
  const set = useCallback(
    (v: T) => {
      setState(v);
      localStorage.setItem(key, JSON.stringify(v));
    },
    [key]
  );
  return [state, set];
}

function uid() {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

export default function App() {
  const [tab, setTab] = useState<Tab>("dashboard");
  const [expenses, setExpenses] = useLocalStorage<Transaction[]>("expenses", SAMPLE_EXPENSES);
  const [income, setIncome] = useLocalStorage<Transaction[]>("income", SAMPLE_INCOME);
  const [goals, setGoals] = useLocalStorage<SavingsGoal[]>("goals", SAMPLE_GOALS);
  const [selectedMonth, setSelectedMonth] = useState(getCurrentMonth());
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  const monthExpenses = expenses.filter((e) => e.date.startsWith(selectedMonth));
  const monthIncome = income.filter((i) => i.date.startsWith(selectedMonth));

  const totalExpenses = monthExpenses.reduce((s, e) => s + e.amount, 0);
  const totalIncome = monthIncome.reduce((s, i) => s + i.amount, 0);
  const balance = totalIncome - totalExpenses;

  const NAV_ITEMS: { id: Tab; label: string; icon: string }[] = [
    { id: "dashboard", label: "Dashboard", icon: "◈" },
    { id: "expenses", label: "Pengeluaran", icon: "↑" },
    { id: "income", label: "Pemasukan", icon: "↓" },
    { id: "savings", label: "Tabungan", icon: "◎" },
  ];

  return (
    <div className="min-h-screen bg-[#F5F0E8] flex flex-col lg:flex-row">
      {/* Sidebar */}
      <aside className="hidden lg:flex flex-col w-56 min-h-screen bg-[#1C1917] text-[#F5F0E8] px-5 py-8 shrink-0">
        <div className="mb-10">
          <div
            className="text-2xl mb-1"
            style={{ fontFamily: "DM Serif Display, serif" }}
          >
            Keuangan
          </div>
          <div className="text-xs text-[#6B5E55]" style={{ fontFamily: "JetBrains Mono, monospace" }}>
            Personal Finance
          </div>
        </div>
        <nav className="flex flex-col gap-1 flex-1">
          {NAV_ITEMS.map((item) => (
            <button
              key={item.id}
              onClick={() => setTab(item.id)}
              className={`flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-all text-left ${
                tab === item.id
                  ? "bg-[#F5F0E8] text-[#1C1917]"
                  : "text-[#A89888] hover:text-[#F5F0E8] hover:bg-[#2A2520]"
              }`}
            >
              <span className="text-base w-4">{item.icon}</span>
              {item.label}
            </button>
          ))}
        </nav>
        <div className="mt-auto pt-6 border-t border-[#2A2520]">
          <div className="text-xs text-[#4A4038]" style={{ fontFamily: "JetBrains Mono, monospace" }}>
            {getMonthLabel(selectedMonth)}
          </div>
          <input
            type="month"
            value={selectedMonth}
            onChange={(e) => setSelectedMonth(e.target.value)}
            className="mt-1 w-full bg-[#2A2520] text-[#A89888] text-xs rounded px-2 py-1.5 border border-[#3A332C] focus:outline-none focus:border-[#C8963E]"
            style={{ fontFamily: "JetBrains Mono, monospace", colorScheme: "dark" }}
          />
        </div>
      </aside>

      {/* Mobile header */}
      <header className="lg:hidden bg-[#1C1917] text-[#F5F0E8] flex items-center justify-between px-4 py-4">
        <div style={{ fontFamily: "DM Serif Display, serif" }} className="text-xl">
          Keuangan
        </div>
        <button
          onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
          className="text-[#A89888] hover:text-[#F5F0E8] transition-colors"
        >
          {mobileMenuOpen ? "✕" : "☰"}
        </button>
      </header>

      {mobileMenuOpen && (
        <div className="lg:hidden bg-[#1C1917] text-[#F5F0E8] px-4 pb-4">
          <div className="flex flex-col gap-1 mb-3">
            {NAV_ITEMS.map((item) => (
              <button
                key={item.id}
                onClick={() => { setTab(item.id); setMobileMenuOpen(false); }}
                className={`flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium text-left ${
                  tab === item.id ? "bg-[#F5F0E8] text-[#1C1917]" : "text-[#A89888]"
                }`}
              >
                <span>{item.icon}</span>
                {item.label}
              </button>
            ))}
          </div>
          <input
            type="month"
            value={selectedMonth}
            onChange={(e) => setSelectedMonth(e.target.value)}
            className="w-full bg-[#2A2520] text-[#A89888] text-xs rounded px-2 py-1.5 border border-[#3A332C] focus:outline-none"
            style={{ fontFamily: "JetBrains Mono, monospace", colorScheme: "dark" }}
          />
        </div>
      )}

      {/* Main content */}
      <main className="flex-1 p-5 lg:p-8 overflow-auto">
        {tab === "dashboard" && (
          <Dashboard
            totalIncome={totalIncome}
            totalExpenses={totalExpenses}
            balance={balance}
            monthExpenses={monthExpenses}
            monthIncome={monthIncome}
            selectedMonth={selectedMonth}
            goals={goals}
            onDownload={() => downloadFinanceReport(monthExpenses, monthIncome, selectedMonth)}
          />
        )}
        {tab === "expenses" && (
          <ExpenseTracker
            expenses={expenses}
            setExpenses={setExpenses}
            selectedMonth={selectedMonth}
          />
        )}
        {tab === "income" && (
          <IncomeLog
            income={income}
            setIncome={setIncome}
            selectedMonth={selectedMonth}
          />
        )}
        {tab === "savings" && (
          <SavingsSection goals={goals} setGoals={setGoals} />
        )}
      </main>
    </div>
  );
}

/* ─────────────────────────────────────────────
   DASHBOARD
───────────────────────────────────────────── */
function Dashboard({
  totalIncome,
  totalExpenses,
  balance,
  monthExpenses,
  monthIncome,
  selectedMonth,
  goals,
  onDownload,
}: {
  totalIncome: number;
  totalExpenses: number;
  balance: number;
  monthExpenses: Transaction[];
  monthIncome: Transaction[];
  selectedMonth: string;
  goals: SavingsGoal[];
  onDownload: () => void;
}) {
  const spendRatio = totalIncome > 0 ? (totalExpenses / totalIncome) * 100 : 0;

  // group expenses by category
  const expByCategory: Record<string, number> = {};
  monthExpenses.forEach((e) => {
    expByCategory[e.category] = (expByCategory[e.category] || 0) + e.amount;
  });
  const sortedCategories = Object.entries(expByCategory)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  return (
    <div>
      <div className="mb-8 flex items-start justify-between gap-4 flex-wrap">
        <div>
          <h1
            className="text-3xl lg:text-4xl text-[#1C1917] mb-1"
            style={{ fontFamily: "DM Serif Display, serif" }}
          >
            Ringkasan Bulan Ini
          </h1>
          <p className="text-sm text-[#A89888]">
            {getMonthLabel(selectedMonth)} — semua angka terupdate otomatis
          </p>
        </div>
        <button
          onClick={onDownload}
          className="flex items-center gap-2 px-4 py-2.5 bg-[#1C1917] text-[#F5F0E8] rounded-xl text-sm font-medium hover:bg-[#3A332C] transition-colors"
          title="Unduh laporan Excel"
        >
          <span>↓</span> Unduh Excel
        </button>
      </div>

      {/* Summary cards */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-8">
        <SummaryCard
          label="Total Pemasukan"
          amount={totalIncome}
          color="#1A4A3A"
          bgColor="#E6F0EC"
          sign="+"
        />
        <SummaryCard
          label="Total Pengeluaran"
          amount={totalExpenses}
          color="#C4623A"
          bgColor="#F5E8E2"
          sign="−"
        />
        <div className={`rounded-2xl p-5 border ${balance >= 0 ? "bg-[#FFFDF8] border-[#DDD5C8]" : "bg-[#F5E8E2] border-[#E8C8B8]"}`}>
          <div className="text-xs font-medium text-[#A89888] mb-3 uppercase tracking-wider">
            Sisa Saldo
          </div>
          <div
            className={`text-2xl font-semibold mb-1 ${balance >= 0 ? "text-[#1C1917]" : "text-[#C4623A]"}`}
            style={{ fontFamily: "JetBrains Mono, monospace" }}
          >
            {formatRupiah(Math.abs(balance))}
          </div>
          <div className="text-xs text-[#A89888]">
            {balance >= 0 ? "Surplus — keuangan sehat" : "Defisit — waspada pengeluaran"}
          </div>
        </div>
      </div>

      {/* Spend ratio bar */}
      <div className="bg-[#FFFDF8] rounded-2xl border border-[#DDD5C8] p-5 mb-6">
        <div className="flex justify-between items-center mb-3">
          <span className="text-sm font-medium text-[#1C1917]">Rasio Pengeluaran</span>
          <span
            className="text-sm font-semibold"
            style={{ fontFamily: "JetBrains Mono, monospace", color: spendRatio > 80 ? "#C4623A" : "#1A4A3A" }}
          >
            {spendRatio.toFixed(1)}%
          </span>
        </div>
        <div className="h-2.5 bg-[#EDE6D9] rounded-full overflow-hidden">
          <div
            className="h-full rounded-full transition-all duration-500"
            style={{
              width: `${Math.min(spendRatio, 100)}%`,
              background: spendRatio > 80 ? "#C4623A" : spendRatio > 60 ? "#C8963E" : "#1A4A3A",
            }}
          />
        </div>
        <div className="flex justify-between mt-1.5 text-xs text-[#A89888]">
          <span>Ideal &lt; 70%</span>
          <span>{spendRatio > 100 ? "Melebihi pemasukan!" : spendRatio > 80 ? "Hati-hati" : "Bagus!"}</span>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Category breakdown */}
        <div className="bg-[#FFFDF8] rounded-2xl border border-[#DDD5C8] p-5">
          <h2 className="text-sm font-semibold text-[#1C1917] mb-4">
            Pengeluaran per Kategori
          </h2>
          {sortedCategories.length === 0 ? (
            <p className="text-sm text-[#A89888]">Belum ada pengeluaran bulan ini.</p>
          ) : (
            <div className="space-y-3">
              {sortedCategories.map(([cat, amount]) => {
                const pct = totalExpenses > 0 ? (amount / totalExpenses) * 100 : 0;
                return (
                  <div key={cat}>
                    <div className="flex justify-between text-xs mb-1">
                      <span className="text-[#1C1917] font-medium">{cat}</span>
                      <span className="text-[#A89888]" style={{ fontFamily: "JetBrains Mono, monospace" }}>
                        {formatRupiah(amount)}
                      </span>
                    </div>
                    <div className="h-1.5 bg-[#EDE6D9] rounded-full overflow-hidden">
                      <div
                        className="h-full rounded-full"
                        style={{ width: `${pct}%`, background: "#C4623A", opacity: 0.7 + pct / 300 }}
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* Recent transactions */}
        <div className="bg-[#FFFDF8] rounded-2xl border border-[#DDD5C8] p-5">
          <h2 className="text-sm font-semibold text-[#1C1917] mb-4">
            Transaksi Terbaru
          </h2>
          <div className="space-y-2.5 max-h-64 overflow-y-auto pr-1">
            {[
              ...monthExpenses.map((e) => ({ ...e, type: "expense" as const })),
              ...monthIncome.map((i) => ({ ...i, type: "income" as const })),
            ]
              .sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime())
              .slice(0, 8)
              .map((t) => (
                <div key={t.id} className="flex items-center justify-between py-1.5 border-b border-[#EDE6D9] last:border-0">
                  <div className="flex items-center gap-2.5">
                    <div
                      className="w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold shrink-0"
                      style={{
                        background: t.type === "income" ? "#E6F0EC" : "#F5E8E2",
                        color: t.type === "income" ? "#1A4A3A" : "#C4623A",
                      }}
                    >
                      {t.type === "income" ? "↓" : "↑"}
                    </div>
                    <div>
                      <div className="text-xs font-medium text-[#1C1917] leading-tight">{t.description}</div>
                      <div className="text-xs text-[#A89888]">{t.category}</div>
                    </div>
                  </div>
                  <div
                    className="text-xs font-semibold shrink-0"
                    style={{
                      fontFamily: "JetBrains Mono, monospace",
                      color: t.type === "income" ? "#1A4A3A" : "#C4623A",
                    }}
                  >
                    {t.type === "income" ? "+" : "−"}{formatRupiah(t.amount)}
                  </div>
                </div>
              ))}
            {monthExpenses.length === 0 && monthIncome.length === 0 && (
              <p className="text-sm text-[#A89888]">Belum ada transaksi bulan ini.</p>
            )}
          </div>
        </div>
      </div>

      {/* Goals quick view */}
      {goals.length > 0 && (
        <div className="mt-6 bg-[#FFFDF8] rounded-2xl border border-[#DDD5C8] p-5">
          <h2 className="text-sm font-semibold text-[#1C1917] mb-4">Target Tabungan</h2>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
            {goals.slice(0, 3).map((g) => {
              const pct = Math.min((g.saved / g.target) * 100, 100);
              return (
                <div key={g.id} className="p-3 rounded-xl bg-[#F5F0E8] border border-[#DDD5C8]">
                  <div className="flex items-center gap-2 mb-2">
                    <div className="w-2 h-2 rounded-full" style={{ background: g.color }} />
                    <span className="text-xs font-medium text-[#1C1917]">{g.name}</span>
                  </div>
                  <div className="text-xs text-[#A89888] mb-2" style={{ fontFamily: "JetBrains Mono, monospace" }}>
                    {formatRupiah(g.saved)} / {formatRupiah(g.target)}
                  </div>
                  <div className="h-1.5 bg-[#DDD5C8] rounded-full overflow-hidden">
                    <div className="h-full rounded-full transition-all" style={{ width: `${pct}%`, background: g.color }} />
                  </div>
                  <div className="text-xs text-[#A89888] mt-1">{pct.toFixed(0)}%</div>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

function SummaryCard({ label, amount, color, bgColor, sign }: {
  label: string; amount: number; color: string; bgColor: string; sign: string;
}) {
  return (
    <div className="rounded-2xl p-5 border border-[#DDD5C8]" style={{ background: bgColor }}>
      <div className="text-xs font-medium text-[#A89888] mb-3 uppercase tracking-wider">{label}</div>
      <div className="text-2xl font-semibold mb-1" style={{ fontFamily: "JetBrains Mono, monospace", color }}>
        {sign}{formatRupiah(amount)}
      </div>
      <div className="w-8 h-0.5 rounded-full" style={{ background: color, opacity: 0.4 }} />
    </div>
  );
}

/* ─────────────────────────────────────────────
   EXPENSE TRACKER
───────────────────────────────────────────── */
function ExpenseTracker({
  expenses,
  setExpenses,
  selectedMonth,
}: {
  expenses: Transaction[];
  setExpenses: (v: Transaction[]) => void;
  selectedMonth: string;
}) {
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ date: new Date().toISOString().slice(0, 10), amount: "", category: EXPENSE_CATEGORIES[0], description: "", note: "" });
  const [filterCat, setFilterCat] = useState("Semua");
  const [editId, setEditId] = useState<string | null>(null);

  const monthExpenses = expenses.filter((e) => e.date.startsWith(selectedMonth));
  const filtered = filterCat === "Semua" ? monthExpenses : monthExpenses.filter((e) => e.category === filterCat);
  const total = filtered.reduce((s, e) => s + e.amount, 0);

  function submit() {
    if (!form.amount || !form.description) return;
    if (editId) {
      setExpenses(expenses.map((e) => e.id === editId ? { ...e, ...form, amount: parseFloat(form.amount) } : e));
      setEditId(null);
    } else {
      setExpenses([{ id: uid(), ...form, amount: parseFloat(form.amount) }, ...expenses]);
    }
    setForm({ date: new Date().toISOString().slice(0, 10), amount: "", category: EXPENSE_CATEGORIES[0], description: "", note: "" });
    setShowForm(false);
  }

  function startEdit(e: Transaction) {
    setForm({ date: e.date, amount: String(e.amount), category: e.category, description: e.description, note: e.note });
    setEditId(e.id);
    setShowForm(true);
  }

  function remove(id: string) {
    setExpenses(expenses.filter((e) => e.id !== id));
  }

  return (
    <div>
      <div className="flex items-start justify-between mb-8 flex-wrap gap-3">
        <div>
          <h1 className="text-3xl lg:text-4xl text-[#1C1917] mb-1" style={{ fontFamily: "DM Serif Display, serif" }}>
            Pengeluaran
          </h1>
          <p className="text-sm text-[#A89888]">{getMonthLabel(selectedMonth)}</p>
        </div>
        <button
          onClick={() => { setShowForm(!showForm); setEditId(null); setForm({ date: new Date().toISOString().slice(0, 10), amount: "", category: EXPENSE_CATEGORIES[0], description: "", note: "" }); }}
          className="flex items-center gap-2 px-4 py-2.5 bg-[#C4623A] text-white rounded-xl text-sm font-medium hover:bg-[#D97B58] transition-colors"
        >
          <span>+</span> Tambah Pengeluaran
        </button>
      </div>

      {showForm && (
        <TransactionForm
          form={form}
          setForm={setForm}
          categories={EXPENSE_CATEGORIES}
          onSubmit={submit}
          onCancel={() => { setShowForm(false); setEditId(null); }}
          isEdit={!!editId}
          accentColor="#C4623A"
        />
      )}

      <div className="flex gap-2 flex-wrap mb-4">
        {["Semua", ...EXPENSE_CATEGORIES].map((cat) => (
          <button
            key={cat}
            onClick={() => setFilterCat(cat)}
            className={`px-3 py-1 rounded-full text-xs font-medium transition-colors ${filterCat === cat ? "bg-[#C4623A] text-white" : "bg-[#EDE6D9] text-[#6B5E55] hover:bg-[#DDD5C8]"}`}
          >
            {cat}
          </button>
        ))}
      </div>

      <div className="mb-3 flex justify-between items-center">
        <span className="text-xs text-[#A89888]">{filtered.length} transaksi</span>
        <span className="text-sm font-semibold text-[#C4623A]" style={{ fontFamily: "JetBrains Mono, monospace" }}>
          Total: {formatRupiah(total)}
        </span>
      </div>

      <div className="space-y-2">
        {filtered.length === 0 && (
          <div className="text-center py-16 text-[#A89888]">
            <div className="text-4xl mb-3 opacity-30">↑</div>
            <p className="text-sm">Belum ada pengeluaran untuk bulan ini.</p>
          </div>
        )}
        {filtered.map((e) => (
          <TransactionRow
            key={e.id}
            t={e}
            type="expense"
            onEdit={() => startEdit(e)}
            onDelete={() => remove(e.id)}
          />
        ))}
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────────
   INCOME LOG
───────────────────────────────────────────── */
function IncomeLog({
  income,
  setIncome,
  selectedMonth,
}: {
  income: Transaction[];
  setIncome: (v: Transaction[]) => void;
  selectedMonth: string;
}) {
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ date: new Date().toISOString().slice(0, 10), amount: "", category: INCOME_CATEGORIES[0], description: "", note: "" });
  const [editId, setEditId] = useState<string | null>(null);

  const monthIncome = income.filter((i) => i.date.startsWith(selectedMonth));
  const total = monthIncome.reduce((s, i) => s + i.amount, 0);

  function submit() {
    if (!form.amount || !form.description) return;
    if (editId) {
      setIncome(income.map((i) => i.id === editId ? { ...i, ...form, amount: parseFloat(form.amount) } : i));
      setEditId(null);
    } else {
      setIncome([{ id: uid(), ...form, amount: parseFloat(form.amount) }, ...income]);
    }
    setForm({ date: new Date().toISOString().slice(0, 10), amount: "", category: INCOME_CATEGORIES[0], description: "", note: "" });
    setShowForm(false);
  }

  function startEdit(t: Transaction) {
    setForm({ date: t.date, amount: String(t.amount), category: t.category, description: t.description, note: t.note });
    setEditId(t.id);
    setShowForm(true);
  }

  function remove(id: string) {
    setIncome(income.filter((i) => i.id !== id));
  }

  return (
    <div>
      <div className="flex items-start justify-between mb-8 flex-wrap gap-3">
        <div>
          <h1 className="text-3xl lg:text-4xl text-[#1C1917] mb-1" style={{ fontFamily: "DM Serif Display, serif" }}>
            Pemasukan
          </h1>
          <p className="text-sm text-[#A89888]">{getMonthLabel(selectedMonth)}</p>
        </div>
        <button
          onClick={() => { setShowForm(!showForm); setEditId(null); setForm({ date: new Date().toISOString().slice(0, 10), amount: "", category: INCOME_CATEGORIES[0], description: "", note: "" }); }}
          className="flex items-center gap-2 px-4 py-2.5 bg-[#1A4A3A] text-white rounded-xl text-sm font-medium hover:bg-[#2D6B56] transition-colors"
        >
          <span>+</span> Tambah Pemasukan
        </button>
      </div>

      {showForm && (
        <TransactionForm
          form={form}
          setForm={setForm}
          categories={INCOME_CATEGORIES}
          onSubmit={submit}
          onCancel={() => { setShowForm(false); setEditId(null); }}
          isEdit={!!editId}
          accentColor="#1A4A3A"
        />
      )}

      <div className="mb-3 flex justify-between items-center">
        <span className="text-xs text-[#A89888]">{monthIncome.length} transaksi</span>
        <span className="text-sm font-semibold text-[#1A4A3A]" style={{ fontFamily: "JetBrains Mono, monospace" }}>
          Total: {formatRupiah(total)}
        </span>
      </div>

      <div className="space-y-2">
        {monthIncome.length === 0 && (
          <div className="text-center py-16 text-[#A89888]">
            <div className="text-4xl mb-3 opacity-30">↓</div>
            <p className="text-sm">Belum ada pemasukan untuk bulan ini.</p>
          </div>
        )}
        {monthIncome.map((t) => (
          <TransactionRow
            key={t.id}
            t={t}
            type="income"
            onEdit={() => startEdit(t)}
            onDelete={() => remove(t.id)}
          />
        ))}
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────────
   SAVINGS GOALS
───────────────────────────────────────────── */
function SavingsSection({ goals, setGoals }: { goals: SavingsGoal[]; setGoals: (v: SavingsGoal[]) => void }) {
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  const [form, setForm] = useState({ name: "", target: "", saved: "", deadline: "", color: GOAL_COLORS[0] });
  const [depositId, setDepositId] = useState<string | null>(null);
  const [depositAmount, setDepositAmount] = useState("");

  const totalTarget = goals.reduce((s, g) => s + g.target, 0);
  const totalSaved = goals.reduce((s, g) => s + g.saved, 0);

  function submitGoal() {
    if (!form.name || !form.target) return;
    if (editId) {
      setGoals(goals.map((g) => g.id === editId ? { ...g, name: form.name, target: parseFloat(form.target), saved: parseFloat(form.saved) || g.saved, deadline: form.deadline, color: form.color } : g));
      setEditId(null);
    } else {
      setGoals([...goals, { id: uid(), name: form.name, target: parseFloat(form.target), saved: parseFloat(form.saved) || 0, deadline: form.deadline, color: form.color }]);
    }
    setForm({ name: "", target: "", saved: "", deadline: "", color: GOAL_COLORS[0] });
    setShowForm(false);
  }

  function startEdit(g: SavingsGoal) {
    setForm({ name: g.name, target: String(g.target), saved: String(g.saved), deadline: g.deadline, color: g.color });
    setEditId(g.id);
    setShowForm(true);
  }

  function remove(id: string) {
    setGoals(goals.filter((g) => g.id !== id));
  }

  function deposit(id: string) {
    const amt = parseFloat(depositAmount);
    if (!amt) return;
    setGoals(goals.map((g) => g.id === id ? { ...g, saved: g.saved + amt } : g));
    setDepositId(null);
    setDepositAmount("");
  }

  return (
    <div>
      <div className="flex items-start justify-between mb-8 flex-wrap gap-3">
        <div>
          <h1 className="text-3xl lg:text-4xl text-[#1C1917] mb-1" style={{ fontFamily: "DM Serif Display, serif" }}>
            Tabungan & Tujuan
          </h1>
          <p className="text-sm text-[#A89888]">Tetapkan target dan lacak progresmu</p>
        </div>
        <button
          onClick={() => { setShowForm(!showForm); setEditId(null); setForm({ name: "", target: "", saved: "", deadline: "", color: GOAL_COLORS[0] }); }}
          className="flex items-center gap-2 px-4 py-2.5 bg-[#C8963E] text-white rounded-xl text-sm font-medium hover:bg-[#D9A84E] transition-colors"
        >
          <span>+</span> Tambah Tujuan
        </button>
      </div>

      {/* Summary */}
      {goals.length > 0 && (
        <div className="grid grid-cols-2 gap-4 mb-6">
          <div className="bg-[#F5EDD9] rounded-2xl border border-[#DDD5C8] p-4">
            <div className="text-xs text-[#A89888] mb-2 uppercase tracking-wider">Total Ditabung</div>
            <div className="text-xl font-semibold text-[#C8963E]" style={{ fontFamily: "JetBrains Mono, monospace" }}>
              {formatRupiah(totalSaved)}
            </div>
          </div>
          <div className="bg-[#FFFDF8] rounded-2xl border border-[#DDD5C8] p-4">
            <div className="text-xs text-[#A89888] mb-2 uppercase tracking-wider">Total Target</div>
            <div className="text-xl font-semibold text-[#1C1917]" style={{ fontFamily: "JetBrains Mono, monospace" }}>
              {formatRupiah(totalTarget)}
            </div>
          </div>
        </div>
      )}

      {showForm && (
        <div className="bg-[#FFFDF8] rounded-2xl border border-[#DDD5C8] p-5 mb-6">
          <h3 className="text-sm font-semibold text-[#1C1917] mb-4">{editId ? "Edit Tujuan" : "Tujuan Baru"}</h3>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <FormField label="Nama Tujuan">
              <input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="Liburan ke Bali" className={INPUT_CLASS} />
            </FormField>
            <FormField label="Target (Rp)">
              <input type="number" value={form.target} onChange={(e) => setForm({ ...form, target: e.target.value })} placeholder="10000000" className={INPUT_CLASS} style={{ fontFamily: "JetBrains Mono, monospace" }} />
            </FormField>
            <FormField label="Sudah Ditabung (Rp)">
              <input type="number" value={form.saved} onChange={(e) => setForm({ ...form, saved: e.target.value })} placeholder="0" className={INPUT_CLASS} style={{ fontFamily: "JetBrains Mono, monospace" }} />
            </FormField>
            <FormField label="Target Tanggal">
              <input type="date" value={form.deadline} onChange={(e) => setForm({ ...form, deadline: e.target.value })} className={INPUT_CLASS} />
            </FormField>
            <FormField label="Warna">
              <div className="flex gap-2 mt-1">
                {GOAL_COLORS.map((c) => (
                  <button
                    key={c}
                    onClick={() => setForm({ ...form, color: c })}
                    className={`w-7 h-7 rounded-full border-2 transition-all ${form.color === c ? "border-[#1C1917] scale-110" : "border-transparent"}`}
                    style={{ background: c }}
                  />
                ))}
              </div>
            </FormField>
          </div>
          <div className="flex gap-2 mt-4">
            <button onClick={submitGoal} className="px-4 py-2 bg-[#C8963E] text-white rounded-lg text-sm font-medium hover:bg-[#D9A84E] transition-colors">
              {editId ? "Simpan" : "Tambah"}
            </button>
            <button onClick={() => { setShowForm(false); setEditId(null); }} className="px-4 py-2 bg-[#EDE6D9] text-[#6B5E55] rounded-lg text-sm font-medium hover:bg-[#DDD5C8] transition-colors">
              Batal
            </button>
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        {goals.length === 0 && (
          <div className="col-span-2 text-center py-16 text-[#A89888]">
            <div className="text-4xl mb-3 opacity-30">◎</div>
            <p className="text-sm">Belum ada tujuan tabungan. Yuk mulai!</p>
          </div>
        )}
        {goals.map((g) => {
          const pct = Math.min((g.saved / g.target) * 100, 100);
          const remaining = g.target - g.saved;
          return (
            <div key={g.id} className="bg-[#FFFDF8] rounded-2xl border border-[#DDD5C8] p-5">
              <div className="flex items-start justify-between mb-4">
                <div className="flex items-center gap-2.5">
                  <div className="w-3 h-3 rounded-full shrink-0" style={{ background: g.color }} />
                  <h3 className="text-sm font-semibold text-[#1C1917]">{g.name}</h3>
                </div>
                <div className="flex gap-1">
                  <button onClick={() => startEdit(g)} className="p-1.5 text-[#A89888] hover:text-[#1C1917] transition-colors text-xs">✎</button>
                  <button onClick={() => remove(g.id)} className="p-1.5 text-[#A89888] hover:text-[#C4623A] transition-colors text-xs">✕</button>
                </div>
              </div>

              <div className="mb-1.5 flex justify-between text-xs">
                <span className="text-[#A89888]">Progress</span>
                <span className="font-semibold" style={{ fontFamily: "JetBrains Mono, monospace", color: g.color }}>
                  {pct.toFixed(0)}%
                </span>
              </div>
              <div className="h-2 bg-[#EDE6D9] rounded-full overflow-hidden mb-3">
                <div className="h-full rounded-full transition-all duration-500" style={{ width: `${pct}%`, background: g.color }} />
              </div>

              <div className="space-y-1 mb-4">
                <div className="flex justify-between text-xs">
                  <span className="text-[#A89888]">Ditabung</span>
                  <span className="font-medium text-[#1C1917]" style={{ fontFamily: "JetBrains Mono, monospace" }}>{formatRupiah(g.saved)}</span>
                </div>
                <div className="flex justify-between text-xs">
                  <span className="text-[#A89888]">Target</span>
                  <span className="font-medium text-[#1C1917]" style={{ fontFamily: "JetBrains Mono, monospace" }}>{formatRupiah(g.target)}</span>
                </div>
                <div className="flex justify-between text-xs">
                  <span className="text-[#A89888]">Kurang</span>
                  <span className="font-medium text-[#C4623A]" style={{ fontFamily: "JetBrains Mono, monospace" }}>{formatRupiah(Math.max(remaining, 0))}</span>
                </div>
                {g.deadline && (
                  <div className="flex justify-between text-xs">
                    <span className="text-[#A89888]">Target Tanggal</span>
                    <span className="text-[#1C1917]">{formatDate(g.deadline)}</span>
                  </div>
                )}
              </div>

              {depositId === g.id ? (
                <div className="flex gap-2">
                  <input
                    type="number"
                    value={depositAmount}
                    onChange={(e) => setDepositAmount(e.target.value)}
                    placeholder="Jumlah (Rp)"
                    className={`${INPUT_CLASS} flex-1 text-xs`}
                    style={{ fontFamily: "JetBrains Mono, monospace" }}
                    autoFocus
                  />
                  <button onClick={() => deposit(g.id)} className="px-3 py-1.5 bg-[#C8963E] text-white rounded-lg text-xs font-medium hover:bg-[#D9A84E]">OK</button>
                  <button onClick={() => setDepositId(null)} className="px-3 py-1.5 bg-[#EDE6D9] text-[#6B5E55] rounded-lg text-xs font-medium">✕</button>
                </div>
              ) : (
                <button
                  onClick={() => { setDepositId(g.id); setDepositAmount(""); }}
                  className="w-full py-2 rounded-lg text-xs font-medium border transition-colors hover:text-white"
                  style={{ borderColor: g.color, color: g.color }}
                  onMouseEnter={(e) => { (e.currentTarget as HTMLButtonElement).style.background = g.color; }}
                  onMouseLeave={(e) => { (e.currentTarget as HTMLButtonElement).style.background = "transparent"; }}
                >
                  + Tambah Tabungan
                </button>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────────
   SHARED COMPONENTS
───────────────────────────────────────────── */
const INPUT_CLASS =
  "w-full px-3 py-2 rounded-lg bg-[#F5F0E8] border border-[#DDD5C8] text-sm text-[#1C1917] placeholder-[#C8B8A8] focus:outline-none focus:border-[#C8963E] transition-colors";

function FormField({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-xs font-medium text-[#6B5E55] mb-1">{label}</label>
      {children}
    </div>
  );
}

function TransactionForm({
  form, setForm, categories, onSubmit, onCancel, isEdit, accentColor,
}: {
  form: { date: string; amount: string; category: string; description: string; note: string };
  setForm: (f: { date: string; amount: string; category: string; description: string; note: string }) => void;
  categories: string[];
  onSubmit: () => void;
  onCancel: () => void;
  isEdit: boolean;
  accentColor: string;
}) {
  return (
    <div className="bg-[#FFFDF8] rounded-2xl border border-[#DDD5C8] p-5 mb-6">
      <h3 className="text-sm font-semibold text-[#1C1917] mb-4">{isEdit ? "Edit Transaksi" : "Transaksi Baru"}</h3>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <FormField label="Tanggal">
          <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} className={INPUT_CLASS} />
        </FormField>
        <FormField label="Jumlah (Rp)">
          <input type="number" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} placeholder="250000" className={INPUT_CLASS} style={{ fontFamily: "JetBrains Mono, monospace" }} />
        </FormField>
        <FormField label="Kategori">
          <select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className={INPUT_CLASS}>
            {categories.map((c) => <option key={c}>{c}</option>)}
          </select>
        </FormField>
        <FormField label="Deskripsi">
          <input type="text" value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} placeholder="Makan siang di warung" className={INPUT_CLASS} />
        </FormField>
        <FormField label="Catatan (opsional)">
          <input type="text" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} placeholder="Info tambahan..." className={INPUT_CLASS} />
        </FormField>
      </div>
      <div className="flex gap-2 mt-4">
        <button
          onClick={onSubmit}
          className="px-4 py-2 text-white rounded-lg text-sm font-medium transition-colors"
          style={{ background: accentColor }}
        >
          {isEdit ? "Simpan" : "Tambah"}
        </button>
        <button onClick={onCancel} className="px-4 py-2 bg-[#EDE6D9] text-[#6B5E55] rounded-lg text-sm font-medium hover:bg-[#DDD5C8] transition-colors">
          Batal
        </button>
      </div>
    </div>
  );
}

function TransactionRow({ t, type, onEdit, onDelete }: {
  t: Transaction; type: "income" | "expense"; onEdit: () => void; onDelete: () => void;
}) {
  const [expanded, setExpanded] = useState(false);
  const isIncome = type === "income";
  return (
    <div className="bg-[#FFFDF8] rounded-xl border border-[#DDD5C8] px-4 py-3 hover:border-[#C8B8A8] transition-colors">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-3 flex-1 min-w-0" onClick={() => setExpanded(!expanded)} role="button">
          <div
            className="w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold shrink-0"
            style={{ background: isIncome ? "#E6F0EC" : "#F5E8E2", color: isIncome ? "#1A4A3A" : "#C4623A" }}
          >
            {isIncome ? "↓" : "↑"}
          </div>
          <div className="flex-1 min-w-0">
            <div className="text-sm font-medium text-[#1C1917] truncate">{t.description}</div>
            <div className="text-xs text-[#A89888]">{t.category} · {formatDate(t.date)}</div>
          </div>
        </div>
        <div className="flex items-center gap-3 shrink-0 ml-3">
          <span
            className="text-sm font-semibold"
            style={{ fontFamily: "JetBrains Mono, monospace", color: isIncome ? "#1A4A3A" : "#C4623A" }}
          >
            {isIncome ? "+" : "−"}{formatRupiah(t.amount)}
          </span>
          <button onClick={onEdit} className="text-[#A89888] hover:text-[#1C1917] transition-colors text-xs p-1">✎</button>
          <button onClick={onDelete} className="text-[#A89888] hover:text-[#C4623A] transition-colors text-xs p-1">✕</button>
        </div>
      </div>
      {expanded && t.note && (
        <div className="mt-2 ml-11 text-xs text-[#A89888] bg-[#F5F0E8] rounded-lg px-3 py-2">
          {t.note}
        </div>
      )}
    </div>
  );
}

/* ─────────────────────────────────────────────
   SAMPLE DATA
───────────────────────────────────────────── */
const thisMonth = getCurrentMonth();
const SAMPLE_EXPENSES: Transaction[] = [
  { id: uid(), date: `${thisMonth}-03`, amount: 85000, category: "Makanan & Minuman", description: "Makan siang bersama teman", note: "Di warung nasi Padang dekat kantor" },
  { id: uid(), date: `${thisMonth}-05`, amount: 250000, category: "Transportasi", description: "Bensin motor", note: "" },
  { id: uid(), date: `${thisMonth}-07`, amount: 450000, category: "Tagihan & Utilitas", description: "Tagihan listrik bulanan", note: "PLN token" },
  { id: uid(), date: `${thisMonth}-10`, amount: 120000, category: "Hiburan", description: "Langganan Netflix", note: "" },
  { id: uid(), date: `${thisMonth}-12`, amount: 350000, category: "Belanja", description: "Beli baju kerja baru", note: "Diskon 30% di Matahari" },
  { id: uid(), date: `${thisMonth}-15`, amount: 75000, category: "Makanan & Minuman", description: "Kopi dan snack", note: "" },
];
const SAMPLE_INCOME: Transaction[] = [
  { id: uid(), date: `${thisMonth}-01`, amount: 8500000, category: "Gaji", description: "Gaji bulan ini", note: "Sudah dipotong BPJS dan pajak" },
  { id: uid(), date: `${thisMonth}-08`, amount: 1500000, category: "Freelance", description: "Desain logo klien", note: "Proyek dari referral teman" },
  { id: uid(), date: `${thisMonth}-14`, amount: 200000, category: "Lainnya", description: "Transfer dari orang tua", note: "" },
];
const SAMPLE_GOALS: SavingsGoal[] = [
  { id: uid(), name: "Dana Darurat", target: 30000000, saved: 12500000, deadline: "2027-06-01", color: "#1A4A3A" },
  { id: uid(), name: "Liburan ke Jepang", target: 15000000, saved: 3200000, deadline: "2027-03-01", color: "#C4623A" },
  { id: uid(), name: "Laptop Baru", target: 20000000, saved: 7800000, deadline: "2026-12-01", color: "#C8963E" },
];
