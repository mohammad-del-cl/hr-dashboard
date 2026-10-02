import { useMemo, useState } from "react"

import {
  DollarSign,
  TrendingDown,
  TrendingUp,
  X,
} from "lucide-react"

import SectionTree from "./FinanceSectionTree"

import {
  filterFinanceRecords,
  formatMoney,
  formatNumber,
  getSectionPath,
  normalizeAmount,
  normalizeId,
} from "./financeTree"

/* =========================================================
   Records Table (جدول هزینه یا درآمد)
========================================================= */

function SectionRecordsTable({
  title,
  type,
  records,
  sections,
}) {
  const isIncome = type === "income"

  const total = records.reduce(
    (sum, record) => sum + normalizeAmount(record.amount),
    0
  )

  const amountClass = isIncome
    ? "text-green-400"
    : "text-red-400"

  return (
    <div className="overflow-hidden rounded-2xl border border-white/10">
      <div className="flex items-center justify-between border-b border-white/10 bg-[#181818] px-5 py-4">
        <div className="flex items-center gap-2">
          {isIncome ? (
            <TrendingUp size={18} className="text-green-400" />
          ) : (
            <TrendingDown size={18} className="text-red-400" />
          )}

          <h3 className="font-bold">{title}</h3>
        </div>

        <span className="text-xs text-gray-500">
          {formatNumber(records.length)} تراکنش
        </span>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full min-w-[800px]">
          <thead>
            <tr className="border-b border-white/10 bg-[#181818] text-right text-xs text-gray-500">
              <th className="px-5 py-3">تاریخ</th>
              <th className="px-5 py-3">بخش</th>
              <th className="px-5 py-3">شرح</th>
              <th className="px-5 py-3">مبلغ</th>
              <th className="px-5 py-3">دسته‌بندی</th>
              <th className="px-5 py-3">وضعیت</th>
            </tr>
          </thead>

          <tbody>
            {records.length === 0 ? (
              <tr>
                <td
                  colSpan="6"
                  className="px-5 py-10 text-center text-sm text-gray-600"
                >
                  تراکنشی برای این انتخاب و بازه‌ی تاریخ یافت نشد.
                </td>
              </tr>
            ) : (
              records.map((record) => (
                <tr
                  key={normalizeId(record.id)}
                  className="border-b border-white/5 transition hover:bg-white/[0.02]"
                >
                  <td className="px-5 py-3 text-sm text-gray-400">
                    {record.date || "—"}
                  </td>

                  <td className="px-5 py-3 text-sm text-gray-400">
                    {getSectionPath(sections, record.sectionId) ||
                      record.sectionName ||
                      "—"}
                  </td>

                  <td className="px-5 py-3 text-sm text-white">
                    {record.description || "—"}
                  </td>

                  <td className={`px-5 py-3 text-sm font-bold ${amountClass}`}>
                    {formatMoney(record.amount)}
                  </td>

                  <td className="px-5 py-3 text-sm text-gray-400">
                    {record.category || "—"}
                  </td>

                  <td className="px-5 py-3">
                    <span className="rounded-lg bg-white/5 px-3 py-1 text-xs text-gray-400">
                      {record.status || "—"}
                    </span>
                  </td>
                </tr>
              ))
            )}
          </tbody>

          {records.length > 0 && (
            <tfoot>
              <tr className="border-t border-white/10 bg-[#181818]">
                <td
                  colSpan="3"
                  className="px-5 py-3 text-sm font-bold"
                >
                  {isIncome ? "جمع درآمد" : "جمع هزینه"}
                </td>

                <td className={`px-5 py-3 text-sm font-bold ${amountClass}`}>
                  {formatMoney(total)}
                </td>

                <td colSpan="2" />
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  )
}


/* =========================================================
   Profit / Loss Ledger Table (جدول سود و زیان)
========================================================= */

function ProfitLossLedgerTable({
  title,
  rows,
  incomeTotal,
  expenseTotal,
  incomeSections,
  expenseSections,
}) {
  const net = incomeTotal - expenseTotal

  return (
    <div className="overflow-hidden rounded-2xl border border-white/10">
      <div className="flex items-center justify-between border-b border-white/10 bg-[#181818] px-5 py-4">
        <div className="flex items-center gap-2">
          <DollarSign size={18} className="text-[#f0c040]" />
          <h3 className="font-bold">{title}</h3>
        </div>

        <span
          className={`rounded-lg px-3 py-1 text-xs font-bold ${
            net >= 0
              ? "bg-green-500/10 text-green-400"
              : "bg-red-500/10 text-red-400"
          }`}
        >
          {net >= 0 ? "سود" : "زیان"}: {formatMoney(Math.abs(net))}
        </span>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full min-w-[900px]">
          <thead>
            <tr className="border-b border-white/10 bg-[#181818] text-right text-xs text-gray-500">
              <th className="px-5 py-3">تاریخ</th>
              <th className="px-5 py-3">نوع</th>
              <th className="px-5 py-3">بخش</th>
              <th className="px-5 py-3">شرح</th>
              <th className="px-5 py-3">درآمد</th>
              <th className="px-5 py-3">هزینه</th>
              <th className="px-5 py-3">مانده تجمعی</th>
            </tr>
          </thead>

          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td
                  colSpan="7"
                  className="px-5 py-10 text-center text-sm text-gray-600"
                >
                  تراکنشی برای این انتخاب و بازه‌ی تاریخ یافت نشد.
                </td>
              </tr>
            ) : (
              rows.map((row) => {
                const isIncome = row.kind === "income"

                const sections = isIncome
                  ? incomeSections
                  : expenseSections

                return (
                  <tr
                    key={`${row.kind}-${normalizeId(row.record.id)}`}
                    className="border-b border-white/5 transition hover:bg-white/[0.02]"
                  >
                    <td className="px-5 py-3 text-sm text-gray-400">
                      {row.record.date || "—"}
                    </td>

                    <td className="px-5 py-3">
                      <span
                        className={`rounded-lg px-3 py-1 text-xs ${
                          isIncome
                            ? "bg-green-500/10 text-green-400"
                            : "bg-red-500/10 text-red-400"
                        }`}
                      >
                        {isIncome ? "درآمد" : "هزینه"}
                      </span>
                    </td>

                    <td className="px-5 py-3 text-sm text-gray-400">
                      {getSectionPath(sections, row.record.sectionId) ||
                        row.record.sectionName ||
                        "—"}
                    </td>

                    <td className="px-5 py-3 text-sm text-white">
                      {row.record.description || "—"}
                    </td>

                    <td className="px-5 py-3 text-sm font-bold text-green-400">
                      {isIncome ? formatMoney(row.amount) : "—"}
                    </td>

                    <td className="px-5 py-3 text-sm font-bold text-red-400">
                      {!isIncome ? formatMoney(row.amount) : "—"}
                    </td>

                    <td
                      className={`px-5 py-3 text-sm font-bold ${
                        row.balance >= 0
                          ? "text-green-400"
                          : "text-red-400"
                      }`}
                    >
                      {formatMoney(row.balance)}
                    </td>
                  </tr>
                )
              })
            )}
          </tbody>

          {rows.length > 0 && (
            <tfoot>
              <tr className="border-t border-white/10 bg-[#181818]">
                <td
                  colSpan="4"
                  className="px-5 py-3 text-sm font-bold"
                >
                  جمع کل
                </td>

                <td className="px-5 py-3 text-sm font-bold text-green-400">
                  {formatMoney(incomeTotal)}
                </td>

                <td className="px-5 py-3 text-sm font-bold text-red-400">
                  {formatMoney(expenseTotal)}
                </td>

                <td
                  className={`px-5 py-3 text-sm font-bold ${
                    net >= 0 ? "text-green-400" : "text-red-400"
                  }`}
                >
                  {net >= 0 ? "سود" : "زیان"}: {formatMoney(Math.abs(net))}
                </td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  )
}


/* =========================================================
   Profit / Loss Panel
   سه دکمه: هزینه | درآمد | سود و زیان
   انتخاب بخش/زیربخش + فیلتر تاریخ + ۳ جدول
   (پایین صفحه گزارش کلی قبلی بدون تغییر نمایش داده می‌شود)
========================================================= */

export default function ProfitLossPanel({
  OverviewComponent,
  ...props
}) {
  const { records, config } = props

  const [view, setView] = useState("loss")
  const [expenseSectionId, setExpenseSectionId] = useState("")
  const [incomeSectionId, setIncomeSectionId] = useState("")
  const [includeChildren, setIncludeChildren] = useState(true)
  const [dateFrom, setDateFrom] = useState("")
  const [dateTo, setDateTo] = useState("")

  const expenseSections = useMemo(
    () => (Array.isArray(config?.expense) ? config.expense : []),
    [config]
  )

  const incomeSections = useMemo(
    () => (Array.isArray(config?.income) ? config.income : []),
    [config]
  )

  const expenseRecords = useMemo(
    () =>
      filterFinanceRecords({
        records,
        config,
        type: "expense",
        sections: expenseSections,
        sectionId: expenseSectionId,
        includeChildren,
        from: dateFrom,
        to: dateTo,
      }),
    [records, config, expenseSections, expenseSectionId, includeChildren, dateFrom, dateTo]
  )

  const incomeRecords = useMemo(
    () =>
      filterFinanceRecords({
        records,
        config,
        type: "income",
        sections: incomeSections,
        sectionId: incomeSectionId,
        includeChildren,
        from: dateFrom,
        to: dateTo,
      }),
    [records, config, incomeSections, incomeSectionId, includeChildren, dateFrom, dateTo]
  )

  const incomeTotal = useMemo(
    () =>
      incomeRecords.reduce(
        (sum, record) => sum + normalizeAmount(record.amount),
        0
      ),
    [incomeRecords]
  )

  const expenseTotal = useMemo(
    () =>
      expenseRecords.reduce(
        (sum, record) => sum + normalizeAmount(record.amount),
        0
      ),
    [expenseRecords]
  )

  /*
    دفتر سود و زیان: درآمدها و هزینه‌ها به ترتیب تاریخ
    همراه با مانده‌ی تجمعی
  */
  const ledgerRows = useMemo(() => {
    const rows = [
      ...incomeRecords.map((record) => ({
        kind: "income",
        record,
      })),
      ...expenseRecords.map((record) => ({
        kind: "expense",
        record,
      })),
    ].sort((a, b) =>
      String(a.record.date || "").localeCompare(
        String(b.record.date || "")
      )
    )

    let balance = 0

    return rows.map((row) => {
      const amount = normalizeAmount(row.record.amount)

      balance += row.kind === "income" ? amount : -amount

      return { ...row, amount, balance }
    })
  }, [incomeRecords, expenseRecords])

  const expenseLabel = expenseSectionId
    ? getSectionPath(expenseSections, expenseSectionId)
    : "همه بخش‌های هزینه"

  const incomeLabel = incomeSectionId
    ? getSectionPath(incomeSections, incomeSectionId)
    : "همه بخش‌های درآمد"

  const invalidRange =
    dateFrom && dateTo && dateFrom > dateTo

  const viewButtons = [
    {
      id: "expense",
      label: "هزینه",
      icon: TrendingDown,
      iconClass: "bg-red-500/10 text-red-400",
    },
    {
      id: "income",
      label: "درآمد",
      icon: TrendingUp,
      iconClass: "bg-green-500/10 text-green-400",
    },
    {
      id: "loss",
      label: "سود و زیان",
      icon: DollarSign,
      iconClass: "bg-[#d4a017]/10 text-[#f0c040]",
    },
  ]

  function renderPicker(type) {
    const isIncome = type === "income"

    return (
      <div className="rounded-2xl border border-white/10 bg-[#181818] p-4">
        <div className="mb-3 flex items-center gap-2">
          {isIncome ? (
            <TrendingUp size={18} className="text-green-400" />
          ) : (
            <TrendingDown size={18} className="text-red-400" />
          )}

          <h3 className="text-sm font-bold">
            {isIncome
              ? "انتخاب بخش یا زیربخش درآمد"
              : "انتخاب بخش یا زیربخش هزینه"}
          </h3>
        </div>

        <div className="max-h-72 overflow-y-auto">
          <SectionTree
            sections={isIncome ? incomeSections : expenseSections}
            selectedId={
              isIncome ? incomeSectionId : expenseSectionId
            }
            onSelect={
              isIncome ? setIncomeSectionId : setExpenseSectionId
            }
            allLabel={
              isIncome
                ? "همه بخش‌های درآمد"
                : "همه بخش‌های هزینه"
            }
          />
        </div>
      </div>
    )
  }

  return (
    <div>
      {/* Buttons */}
      <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
        {viewButtons.map((button) => {
          const Icon = button.icon

          return (
            <button
              key={button.id}
              onClick={() => setView(button.id)}
              className={`flex items-center gap-3 rounded-2xl border p-4 text-right transition ${
                view === button.id
                  ? "border-[#d4a017] bg-[#d4a017]/10"
                  : "border-white/10 bg-[#181818] hover:border-[#d4a017]/50"
              }`}
            >
              <div
                className={`flex h-10 w-10 items-center justify-center rounded-xl ${button.iconClass}`}
              >
                <Icon size={20} />
              </div>

              <span className="font-bold">{button.label}</span>
            </button>
          )
        })}
      </div>

      {/* Section pickers */}
      <div
        className={`mt-6 grid grid-cols-1 gap-4 ${
          view === "loss" ? "lg:grid-cols-2" : ""
        }`}
      >
        {(view === "expense" || view === "loss") &&
          renderPicker("expense")}

        {(view === "income" || view === "loss") &&
          renderPicker("income")}
      </div>

      {/* Filters */}
      <div className="mt-4 flex flex-col gap-3 md:flex-row md:items-end">
        <div>
          <label className="mb-2 block text-xs text-gray-500">
            از تاریخ
          </label>

          <input
            type="date"
            value={dateFrom}
            onChange={(event) => setDateFrom(event.target.value)}
            className="w-full rounded-xl border border-white/10 bg-[#151515] px-4 py-3 text-sm text-white outline-none focus:border-[#d4a017]"
          />
        </div>

        <div>
          <label className="mb-2 block text-xs text-gray-500">
            تا تاریخ
          </label>

          <input
            type="date"
            value={dateTo}
            onChange={(event) => setDateTo(event.target.value)}
            className="w-full rounded-xl border border-white/10 bg-[#151515] px-4 py-3 text-sm text-white outline-none focus:border-[#d4a017]"
          />
        </div>

        {(dateFrom || dateTo) && (
          <button
            onClick={() => {
              setDateFrom("")
              setDateTo("")
            }}
            className="flex items-center justify-center gap-2 rounded-xl border border-white/10 px-4 py-3 text-sm text-gray-400 transition hover:bg-white/5 hover:text-white"
          >
            <X size={16} />
            حذف فیلتر تاریخ
          </button>
        )}

        <label className="flex items-center gap-2 py-3 text-xs text-gray-400 md:mr-auto">
          <input
            type="checkbox"
            checked={includeChildren}
            onChange={(event) =>
              setIncludeChildren(event.target.checked)
            }
            className="accent-[#d4a017]"
          />
          شامل تراکنش‌های زیربخش‌ها
        </label>
      </div>

      {invalidRange && (
        <p className="mt-2 text-xs text-red-400">
          تاریخ شروع بعد از تاریخ پایان است.
        </p>
      )}

      {/* Three tables */}
      <div className="mt-6 space-y-6">
        <SectionRecordsTable
          title={`جدول هزینه‌ها — ${expenseLabel}`}
          type="expense"
          records={expenseRecords}
          sections={expenseSections}
        />

        <SectionRecordsTable
          title={`جدول درآمدها — ${incomeLabel}`}
          type="income"
          records={incomeRecords}
          sections={incomeSections}
        />

        <ProfitLossLedgerTable
          title="جدول سود و زیان"
          rows={ledgerRows}
          incomeTotal={incomeTotal}
          expenseTotal={expenseTotal}
          incomeSections={incomeSections}
          expenseSections={expenseSections}
        />
      </div>

      {/* گزارش کلی قبلی (بدون تغییر) */}
      <div className="mt-10 border-t border-white/10 pt-8">
        {OverviewComponent && <OverviewComponent {...props} />}
      </div>
    </div>
  )
}
