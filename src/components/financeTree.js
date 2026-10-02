/* =========================================================
   financeTree.js
   توابع کمکی درخت بخش‌ها (بخش / زیربخش) و فیلتر تاریخ
   برای امور مالی.
========================================================= */


/* ---------- توابع پایه (کپی از FinancialAffairs.jsx) ---------- */

export function normalizeId(value) {
  if (
    value === undefined ||
    value === null ||
    value === ""
  ) {
    return ""
  }

  return String(value).trim()
}


export function normalizeName(value) {
  return String(value ?? "")
    .trim()
    .replace(/\u200c/g, " ")
    .replace(/ي/g, "ی")
    .replace(/ى/g, "ی")
    .replace(/ك/g, "ک")
    .replace(/\s+/g, " ")
    .toLowerCase()
}


export function getExcelSectionName(record) {
  const value =
    record?.section ??
    record?.sectionName ??
    record?.Section ??
    record?.SectionName ??
    ""

  return String(value)
    .trim()
    .replace(/\u200c/g, " ")
    .replace(/\s+/g, " ")
}


export function normalizeFinanceType(value) {
  const type = normalizeName(value)

  if (
    type === "expense" ||
    type === "expenses" ||
    type === "هزینه" ||
    type === "هزینه‌ها" ||
    type === "هزینه ها" ||
    type === "خرج" ||
    type === "cost"
  ) {
    return "expense"
  }

  return "income"
}


export function resolveFinanceType(record, config) {
  const safeRecord = record || {}
  const safeConfig = config || {}

  const sectionId = normalizeId(safeRecord.sectionId)
  const sectionName = normalizeName(
    getExcelSectionName(safeRecord)
  )

  const incomeSections = Array.isArray(safeConfig.income)
    ? safeConfig.income
    : []

  const expenseSections = Array.isArray(safeConfig.expense)
    ? safeConfig.expense
    : []

  const matchesSection = (section) => {
    if (!section) return false

    const idMatches =
      sectionId &&
      normalizeId(section.id) === sectionId

    const nameMatches =
      sectionName &&
      normalizeName(section.name) === sectionName

    return Boolean(idMatches || nameMatches)
  }

  const isIncomeSection = incomeSections.some(matchesSection)
  const isExpenseSection = expenseSections.some(matchesSection)

  // Config بخش مالی منبع دقیق‌تری برای تشخیص نوع تراکنش است.
  // این موضوع مخصوصاً برای رکوردهای Excel قدیمی یا رکوردهایی
  // که type آن‌ها خالی/اشتباه است، جلوی ورود هزینه به درآمد را می‌گیرد.
  if (isIncomeSection && !isExpenseSection) {
    return "income"
  }

  if (isExpenseSection && !isIncomeSection) {
    return "expense"
  }

  const normalizedType = normalizeName(safeRecord.type)

  if (
    normalizedType === "expense" ||
    normalizedType === "expenses" ||
    normalizedType === "هزینه" ||
    normalizedType === "هزینه ها" ||
    normalizedType === "هزینه‌ها" ||
    normalizedType === "خرج" ||
    normalizedType === "cost"
  ) {
    return "expense"
  }

  if (
    normalizedType === "income" ||
    normalizedType === "incomes" ||
    normalizedType === "درآمد"
  ) {
    return "income"
  }

  // برای سازگاری با منطق قبلی، مقدار ناشناخته income محسوب می‌شود.
  return normalizeFinanceType(safeRecord.type)
}


export function convertPersianDigits(value) {
  return String(value ?? "")
    .replace(/[۰-۹]/g, (digit) =>
      "۰۱۲۳۴۵۶۷۸۹".indexOf(digit)
    )
    .replace(/[٠-٩]/g, (digit) =>
      "٠١٢٣٤٥٦٧٨٩".indexOf(digit)
    )
}


export function normalizeAmount(value) {
  if (
    value === undefined ||
    value === null ||
    value === ""
  ) {
    return 0
  }

  const cleaned = convertPersianDigits(value)
    .replace(/,/g, "")
    .replace(/٬/g, "")
    .replace(/،/g, "")
    .replace(/\s/g, "")
    .replace(/[^\d.-]/g, "")

  return Number(cleaned) || 0
}


export function formatNumber(value) {
  return new Intl.NumberFormat(
    "fa-IR"
  ).format(
    Number(value) || 0
  )
}


export function formatMoney(value) {
  return `${formatNumber(value)} تومان`
}



/* =========================================================
   Section Tree Helpers (بخش‌ها و زیربخش‌ها)

   ساختار ذخیره‌سازی همان آرایه‌ی تخت قبلی است
   (config.income / config.expense)؛ فقط هر بخش می‌تواند
   یک فیلد اختیاری parentId داشته باشد.
   parentId خالی = بخش اصلی.
========================================================= */

/*
  آیا تاریخ داخل بازه [from, to] هست؟
  تاریخ‌ها به شکل YYYY-MM-DD ذخیره می‌شوند و مقایسه‌ی
  رشته‌ای برای آن‌ها درست کار می‌کند.
*/
export function isDateInRange(date, from, to) {
  if (!from && !to) {
    return true
  }

  const value = String(date || "")

  if (!value) {
    return false
  }

  if (from && value < from) {
    return false
  }

  if (to && value > to) {
    return false
  }

  return true
}


/*
  تبدیل آرایه‌ی تخت بخش‌ها به درخت
*/
export function buildSectionTree(sections) {
  const list = Array.isArray(sections) ? sections : []

  const nodes = new Map()

  list.forEach((section) => {
    nodes.set(normalizeId(section.id), {
      ...section,
      children: [],
    })
  })

  const roots = []

  nodes.forEach((node) => {
    const parentId = normalizeId(node.parentId)

    if (
      parentId &&
      parentId !== normalizeId(node.id) &&
      nodes.has(parentId)
    ) {
      nodes.get(parentId).children.push(node)
    } else {
      roots.push(node)
    }
  })

  return roots
}


/*
  شناسه‌ی یک بخش + همه‌ی زیربخش‌های آن (در هر عمق)
*/
export function getSectionIdsWithDescendants(sections, rootId) {
  const list = Array.isArray(sections) ? sections : []

  const result = new Set([normalizeId(rootId)])

  let changed = true

  while (changed) {
    changed = false

    list.forEach((section) => {
      const sectionId = normalizeId(section.id)
      const parentId = normalizeId(section.parentId)

      if (
        parentId &&
        result.has(parentId) &&
        !result.has(sectionId)
      ) {
        result.add(sectionId)
        changed = true
      }
    })
  }

  return result
}


/*
  مسیر کامل بخش، مثلاً: غذاخوری / رستوران / صبحانه
*/
export function getSectionPath(sections, sectionId) {
  const list = Array.isArray(sections) ? sections : []

  const names = []
  const seen = new Set()

  let current = list.find(
    (section) =>
      normalizeId(section.id) === normalizeId(sectionId)
  )

  while (
    current &&
    !seen.has(normalizeId(current.id))
  ) {
    seen.add(normalizeId(current.id))
    names.unshift(current.name)

    const parentId = normalizeId(current.parentId)

    current = parentId
      ? list.find(
          (section) =>
            normalizeId(section.id) === parentId
        )
      : null
  }

  return names.join(" / ")
}


/*
  فیلتر تراکنش‌های یک نوع (income / expense) بر اساس
  بخش انتخاب‌شده (اختیاری) و بازه‌ی تاریخ
*/
export function filterFinanceRecords({
  records,
  config,
  type,
  sections,
  sectionId,
  includeChildren,
  from,
  to,
}) {
  const ids = sectionId
    ? includeChildren
      ? getSectionIdsWithDescendants(sections, sectionId)
      : new Set([normalizeId(sectionId)])
    : null

  return (records || [])
    .filter(
      (record) =>
        resolveFinanceType(record, config) === type &&
        (!ids || ids.has(normalizeId(record.sectionId))) &&
        isDateInRange(record.date, from, to)
    )
    .sort((a, b) =>
      String(a.date || "").localeCompare(String(b.date || ""))
    )
}


/* =========================================================
   Section Path Helpers (برای ورود از Excel)
========================================================= */

/*
  "غذاخوری / رستوران > ناهار"  →  ["غذاخوری","رستوران","ناهار"]
*/
export function splitSectionPath(value) {
  return String(value ?? "")
    .replace(/\u200c/g, " ")
    .split(/[\/\\>»]/)
    .map((part) => part.replace(/\s+/g, " ").trim())
    .filter(Boolean)
}


/*
  پیدا کردن بخش بر اساس مسیر کامل.
  اگر مسیر فقط یک نام بود و بخش اصلی با آن نام نبود،
  بخشی که نامش همان است (در هر سطح) پیدا می‌شود.
*/
export function findSectionByPath(sections, pathText) {
  const list = Array.isArray(sections) ? sections : []
  const parts = splitSectionPath(pathText)

  if (parts.length === 0) {
    return null
  }

  let parentId = ""
  let found = null

  for (const part of parts) {
    found =
      list.find(
        (section) =>
          normalizeId(section.parentId) === parentId &&
          normalizeName(section.name) === normalizeName(part)
      ) || null

    if (!found) {
      break
    }

    parentId = normalizeId(found.id)
  }

  if (found) {
    return found
  }

  if (parts.length === 1) {
    return (
      list.find(
        (section) =>
          normalizeName(section.name) ===
          normalizeName(parts[0])
      ) || null
    )
  }

  return null
}
