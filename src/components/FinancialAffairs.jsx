import {
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react"

import {
  ArrowRight,
  Building2,
  CalendarDays,
  DollarSign,
  FileText,
  FolderPlus,
  Pencil,
  Plus,
  Receipt,
  Search,
  Trash2,
  TrendingDown,
  TrendingUp,
  Upload,
  Download,
  Wallet,
  X,
} from "lucide-react"

import {
  getFinanceConfig,
  saveFinanceConfig,
  getFinanceRecords,
  saveFinanceRecord,
  deleteFinanceRecord,
} from "../services/dataService"

import {
  readFinanceExcelFile,
  downloadFinanceExcelTemplate,
} from "../services/excelService"


/* =========================================================
   Default Finance Config
========================================================= */

const DEFAULT_FINANCE_CONFIG = {
  id: "finance-config",

  income: [
    {
      id: "income-cafeteria",
      name: "درآمد غذاخوری",
    },
    {
      id: "income-pool",
      name: "درآمد استخر",
    },
  ],

  expense: [
    {
      id: "expense-cafeteria",
      name: "هزینه غذاخوری",
    },
    {
      id: "expense-pool",
      name: "هزینه استخر",
    },
  ],
}


/* =========================================================
   Helpers
========================================================= */

function createId(prefix = "finance") {
  return `${prefix}-${Date.now()}-${Math.random()
    .toString(36)
    .substring(2, 8)}`
}


/*
  همه IDها را String می‌کنیم.
*/
function normalizeId(value) {
  if (
    value === undefined ||
    value === null ||
    value === ""
  ) {
    return ""
  }

  return String(value).trim()
}


/*
  برای مقایسه مطمئن نام بخش‌ها.
  مثال:
  ك = ک
  ي = ی
*/
function normalizeName(value) {
  return String(value ?? "")
    .trim()
    .replace(/\u200c/g, " ")
    .replace(/ي/g, "ی")
    .replace(/ى/g, "ی")
    .replace(/ك/g, "ک")
    .replace(/\s+/g, " ")
    .toLowerCase()
}


/*
  تبدیل نوع Excel به income / expense
*/
function normalizeFinanceType(value) {
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


/*
  نام بخش را از Excel می‌خواند.

  هر دو مورد قابل قبول هستند:

  section
  sectionName
*/
function getExcelSectionName(record) {
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


/*
  تبدیل اعداد فارسی و عربی به انگلیسی
*/
function convertPersianDigits(value) {
  return String(value ?? "")
    .replace(/[۰-۹]/g, (digit) =>
      "۰۱۲۳۴۵۶۷۸۹".indexOf(digit)
    )
    .replace(/[٠-٩]/g, (digit) =>
      "٠١٢٣٤٥٦٧٨٩".indexOf(digit)
    )
}


/*
  تبدیل مبلغ Excel
*/
function normalizeAmount(value) {
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


function formatNumber(value) {
  return new Intl.NumberFormat(
    "fa-IR"
  ).format(
    Number(value) || 0
  )
}


function formatMoney(value) {
  return `${formatNumber(value)} تومان`
}


/*
  تاریخ امروز
*/
function getToday() {
  const date = new Date()

  const year =
    date.getFullYear()

  const month =
    String(
      date.getMonth() + 1
    ).padStart(2, "0")

  const day =
    String(
      date.getDate()
    ).padStart(2, "0")

  return `${year}-${month}-${day}`
}


/*
  تبدیل تاریخ Excel به YYYY-MM-DD
*/
function normalizeDate(value) {
  if (!value) {
    return getToday()
  }

  /*
    اگر Date باشد
  */
  if (value instanceof Date) {
    if (
      Number.isNaN(
        value.getTime()
      )
    ) {
      return getToday()
    }

    const year =
      value.getFullYear()

    const month =
      String(
        value.getMonth() + 1
      ).padStart(2, "0")

    const day =
      String(
        value.getDate()
      ).padStart(2, "0")

    return `${year}-${month}-${day}`
  }

  let text =
    convertPersianDigits(
      String(value).trim()
    )

  if (!text) {
    return getToday()
  }

  /*
    YYYY/MM/DD
  */
  if (
    /^\d{4}\/\d{1,2}\/\d{1,2}$/.test(
      text
    )
  ) {
    const [
      year,
      month,
      day,
    ] = text.split("/")

    return `${year}-${String(
      month
    ).padStart(2, "0")}-${String(
      day
    ).padStart(2, "0")}`
  }

  /*
    YYYY-MM-DD
  */
  if (
    /^\d{4}-\d{1,2}-\d{1,2}$/.test(
      text
    )
  ) {
    const [
      year,
      month,
      day,
    ] = text.split("-")

    return `${year}-${String(
      month
    ).padStart(2, "0")}-${String(
      day
    ).padStart(2, "0")}`
  }

  /*
    اگر Excel تاریخ را به شکل
    عدد ذخیره کرده باشد.
  */
  if (
    /^\d+(\.\d+)?$/.test(
      text
    )
  ) {
    const excelSerial =
      Number(text)

    if (
      excelSerial > 20000 &&
      excelSerial < 100000
    ) {
      const excelEpoch =
        new Date(
          Date.UTC(
            1899,
            11,
            30
          )
        )

      const result =
        new Date(
          excelEpoch.getTime() +
            excelSerial *
              86400000
        )

      const year =
        result.getUTCFullYear()

      const month =
        String(
          result.getUTCMonth() + 1
        ).padStart(2, "0")

      const day =
        String(
          result.getUTCDate()
        ).padStart(2, "0")

      return `${year}-${month}-${day}`
    }
  }

  return text
}


/* =========================================================
   Normalize Finance Config
========================================================= */

function normalizeFinanceConfig(
  config
) {
  function normalizeSections(
    sections,
    type
  ) {
    if (
      !Array.isArray(
        sections
      )
    ) {
      return []
    }

    return sections
      .map(
        (
          section,
          index
        ) => {
          const name =
            String(
              section?.name ?? ""
            ).trim()

          if (!name) {
            return null
          }

          return {
            id:
              normalizeId(
                section?.id
              ) ||
              `${type}-${index}-${normalizeName(
                name
              )
                .replace(
                  /\s+/g,
                  "-"
                )
                .replace(
                  /[^a-zA-Z0-9\u0600-\u06FF-]/g,
                  ""
                )}`,

            name,
          }
        }
      )
      .filter(Boolean)
  }

  const income =
    normalizeSections(
      config?.income,
      "income"
    )

  const expense =
    normalizeSections(
      config?.expense,
      "expense"
    )

  return {
    id: "finance-config",

    income:
      income.length > 0
        ? income
        : DEFAULT_FINANCE_CONFIG.income.map(
            (section) => ({
              ...section,
              id: normalizeId(
                section.id
              ),
            })
          ),

    expense:
      expense.length > 0
        ? expense
        : DEFAULT_FINANCE_CONFIG.expense.map(
            (section) => ({
              ...section,
              id: normalizeId(
                section.id
              ),
            })
          ),
  }
}


/* =========================================================
   Main Component
========================================================= */

export default function FinancialAffairs({
  onBack,
}) {
  const [
    config,
    setConfig,
  ] = useState(
    DEFAULT_FINANCE_CONFIG
  )

  const [
    records,
    setRecords,
  ] = useState([])

  const [
    activeSection,
    setActiveSection,
  ] = useState("income")

  const [
    selectedSubSectionId,
    setSelectedSubSectionId,
  ] = useState("")

  const [
    showAddModal,
    setShowAddModal,
  ] = useState(false)

  const [
    showRecordModal,
    setShowRecordModal,
  ] = useState(false)

  const [
    editingRecord,
    setEditingRecord,
  ] = useState(null)

  const [
    newSectionName,
    setNewSectionName,
  ] = useState("")

  const [
    searchText,
    setSearchText,
  ] = useState("")

  const [
    loading,
    setLoading,
  ] = useState(true)

  const [
    saving,
    setSaving,
  ] = useState(false)

  const [
    importingExcel,
    setImportingExcel,
  ] = useState(false)

  const excelInputRef =
    useRef(null)

  const [
    recordForm,
    setRecordForm,
  ] = useState({
    date: getToday(),
    description: "",
    amount: "",
    category: "",
    status: "تسویه شده",
    note: "",
  })


  /* =========================================================
     Current Sections
  ========================================================= */

  const currentSections =
    useMemo(() => {
      if (
        activeSection ===
        "profit-loss"
      ) {
        return []
      }

      return Array.isArray(
        config?.[activeSection]
      )
        ? config[
            activeSection
          ]
        : []
    }, [
      config,
      activeSection,
    ])


  /* =========================================================
     Selected Section
  ========================================================= */

  const selectedSubSection =
    useMemo(() => {
      if (
        activeSection ===
        "profit-loss"
      ) {
        return null
      }

      const selectedId =
        normalizeId(
          selectedSubSectionId
        )

      if (!selectedId) {
        return null
      }

      return (
        currentSections.find(
          (section) =>
            normalizeId(
              section.id
            ) === selectedId
        ) || null
      )
    }, [
      currentSections,
      activeSection,
      selectedSubSectionId,
    ])


  /* =========================================================
     Load Finance Data
  ========================================================= */

  useEffect(() => {
    loadFinanceData()
  }, [])


  async function loadFinanceData() {
    try {
      const [
        savedConfig,
        savedRecords,
      ] = await Promise.all([
        getFinanceConfig(),
        getFinanceRecords(),
      ])

      const normalizedConfig =
        normalizeFinanceConfig(
          savedConfig
        )

      const normalizedRecords = (
        savedRecords || []
      ).map(
        (record) => ({
          ...record,

          id:
            normalizeId(
              record.id
            ) ||
            createId(
              "finance"
            ),

          type:
            normalizeFinanceType(
              record.type
            ),

          sectionId:
            normalizeId(
              record.sectionId
            ),

          sectionName:
            getExcelSectionName(
              record
            ),

          section:
            getExcelSectionName(
              record
            ),

          amount:
            normalizeAmount(
              record.amount
            ),

          date:
            normalizeDate(
              record.date
            ),

          description:
            String(
              record.description ??
                ""
            ).trim(),

          category:
            String(
              record.category ??
                ""
            ).trim(),

          status:
            String(
              record.status ??
                "تسویه شده"
            ).trim(),

          note:
            String(
              record.note ??
                ""
            ).trim(),
        })
      )

      setConfig(
        normalizedConfig
      )

      setRecords(
        normalizedRecords
      )

      if (
        normalizedConfig.income
          .length > 0
      ) {
        setActiveSection(
          "income"
        )

        setSelectedSubSectionId(
          normalizeId(
            normalizedConfig
              .income[0]
              .id
          )
        )
      } else if (
        normalizedConfig.expense
          .length > 0
      ) {
        setActiveSection(
          "expense"
        )

        setSelectedSubSectionId(
          normalizeId(
            normalizedConfig
              .expense[0]
              .id
          )
        )
      }
    } catch (error) {
      console.error(
        "خطا در خواندن اطلاعات مالی:",
        error
      )

      alert(
        "خواندن اطلاعات مالی با خطا مواجه شد."
      )
    } finally {
      setLoading(false)
    }
  }


  /* =========================================================
     Save Config
  ========================================================= */

  async function saveConfig(
    newConfig
  ) {
    const normalizedConfig =
      normalizeFinanceConfig(
        newConfig
      )

    try {
      await saveFinanceConfig(
        normalizedConfig
      )

      setConfig(
        normalizedConfig
      )

      return true
    } catch (error) {
      console.error(
        "خطا در ذخیره تنظیمات مالی:",
        error
      )

      alert(
        "ذخیره تنظیمات مالی انجام نشد."
      )

      return false
    }
  }


  /* =========================================================
     Main Section Change
  ========================================================= */

  function handleMainSectionChange(
    section
  ) {
    setActiveSection(
      section
    )

    setSearchText("")

    const sections =
      Array.isArray(
        config?.[section]
      )
        ? config[section]
        : []

    if (
      sections.length > 0
    ) {
      setSelectedSubSectionId(
        normalizeId(
          sections[0].id
        )
      )
    } else {
      setSelectedSubSectionId("")
    }
  }


  /* =========================================================
     Add Financial Section
  ========================================================= */

  async function handleAddSection() {
    if (
      activeSection !==
        "income" &&
      activeSection !==
        "expense"
    ) {
      return
    }

    const name =
      newSectionName.trim()

    if (!name) {
      alert(
        "نام بخش را وارد کنید."
      )
      return
    }

    const existingSections =
      Array.isArray(
        config?.[activeSection]
      )
        ? config[
            activeSection
          ]
        : []

    const exists =
      existingSections.some(
        (section) =>
          normalizeName(
            section.name
          ) ===
          normalizeName(name)
      )

    if (exists) {
      alert(
        "این بخش قبلاً وجود دارد."
      )
      return
    }

    const newSection = {
      id: createId(
        activeSection
      ),
      name,
    }

    const newConfig = {
      ...config,

      [activeSection]: [
        ...existingSections,
        newSection,
      ],
    }

    const saved =
      await saveConfig(
        newConfig
      )

    if (!saved) {
      return
    }

    setActiveSection(
      activeSection
    )

    setSelectedSubSectionId(
      normalizeId(
        newSection.id
      )
    )

    setSearchText("")
    setNewSectionName("")
    setShowAddModal(false)
  }


  /* =========================================================
     Delete Financial Section
  ========================================================= */

  async function handleDeleteSection(
    section
  ) {
    const sectionId =
      normalizeId(
        section.id
      )

    const relatedRecords =
      records.filter(
        (record) =>
          normalizeId(
            record.sectionId
          ) === sectionId
      )

    if (
      relatedRecords.length > 0
    ) {
      alert(
        `امکان حذف «${section.name}» وجود ندارد.\n\n${formatNumber(
          relatedRecords.length
        )} تراکنش برای این بخش ثبت شده است.\nابتدا تراکنش‌های این بخش را حذف کنید.`
      )

      return
    }

    const confirmed =
      window.confirm(
        `آیا از حذف «${section.name}» مطمئن هستید؟`
      )

    if (!confirmed) {
      return
    }

    const newSections =
      (
        config[
          activeSection
        ] || []
      ).filter(
        (item) =>
          normalizeId(
            item.id
          ) !== sectionId
      )

    const newConfig = {
      ...config,

      [activeSection]:
        newSections,
    }

    const saved =
      await saveConfig(
        newConfig
      )

    if (!saved) {
      return
    }

    if (
      normalizeId(
        selectedSubSectionId
      ) === sectionId
    ) {
      if (
        newSections.length > 0
      ) {
        setSelectedSubSectionId(
          normalizeId(
            newSections[0].id
          )
        )
      } else {
        setSelectedSubSectionId("")
      }
    }
  }


  /* =========================================================
     Excel Picker
  ========================================================= */

  function openExcelFilePicker() {
    if (
      importingExcel
    ) {
      return
    }

    excelInputRef.current?.click()
  }


  /* =========================================================
     Find / Create Section
  ========================================================= */

  function findOrCreateSection(
    workingConfig,
    type,
    sectionName
  ) {
    const normalizedSectionName =
      normalizeName(
        sectionName
      )

    if (
      !normalizedSectionName
    ) {
      return null
    }

    const sections =
      Array.isArray(
        workingConfig[type]
      )
        ? workingConfig[type]
        : []

    const existingSection =
      sections.find(
        (section) =>
          normalizeName(
            section.name
          ) ===
          normalizedSectionName
      )

    if (
      existingSection
    ) {
      return existingSection
    }

    const newSection = {
      id: createId(type),
      name:
        String(
          sectionName
        ).trim(),
    }

    workingConfig[type] = [
      ...sections,
      newSection,
    ]

    return newSection
  }


  /* =========================================================
     Excel Import
  ========================================================= */

  async function handleExcelImport(
    event
  ) {
    const file =
      event.target.files?.[0]

    if (!file) {
      return
    }

    setImportingExcel(true)

    try {
      const importedRecords =
        await readFinanceExcelFile(
          file
        )

      if (
        !Array.isArray(
          importedRecords
        ) ||
        importedRecords.length ===
          0
      ) {
        alert(
          "هیچ تراکنشی از فایل Excel خوانده نشد."
        )

        return
      }

      /*
        ساخت یک Config مستقل برای
        پردازش Excel
      */
      const newConfig = {
        id: "finance-config",

        income: [
          ...(config.income || []),
        ].map(
          (section) => ({
            id:
              normalizeId(
                section.id
              ),
            name:
              String(
                section.name
              ).trim(),
          })
        ),

        expense: [
          ...(config.expense || []),
        ].map(
          (section) => ({
            id:
              normalizeId(
                section.id
              ),
            name:
              String(
                section.name
              ).trim(),
          })
        ),
      }


      /* =====================================================
         پیدا کردن یا ساخت بخش‌ها
      ===================================================== */

      const preparedRows =
        importedRecords.map(
          (record) => {
            const type =
              normalizeFinanceType(
                record.type
              )

            const sectionName =
              getExcelSectionName(
                record
              )

            const section =
              findOrCreateSection(
                newConfig,
                type,
                sectionName
              )

            return {
              original:
                record,

              type,

              sectionName,

              section,
            }
          }
        )


      /*
        ذخیره Config جدید
      */
      const configSaved =
        await saveConfig(
          newConfig
        )

      if (!configSaved) {
        return
      }


      /* =====================================================
         ساخت رکوردهای نهایی
      ===================================================== */

      const finalRecords =
        preparedRows.map(
          ({
            original,
            type,
            sectionName,
            section,
          },
          index) => {
            const recordId =
              normalizeId(
                original.id
              ) ||
              createId(
                `excel-${index}`
              )

            return {
              ...original,

              id: recordId,

              date:
                normalizeDate(
                  original.date
                ),

              type,

              sectionId:
                normalizeId(
                  section?.id
                ),

              sectionName:
                section?.name ||
                sectionName,

              section:
                section?.name ||
                sectionName,

              description:
                String(
                  original.description ??
                    ""
                ).trim(),

              amount:
                normalizeAmount(
                  original.amount
                ),

              category:
                String(
                  original.category ??
                    ""
                ).trim(),

              status:
                String(
                  original.status ??
                    "تسویه شده"
                ).trim() ||
                "تسویه شده",

              note:
                String(
                  original.note ??
                    ""
                ).trim(),
            }
          }
        )


      /* =====================================================
         ذخیره رکوردها
      ===================================================== */

      for (
        const record of finalRecords
      ) {
        await saveFinanceRecord(
          record
        )
      }


      /* =====================================================
         Update State
      ===================================================== */

      setRecords(
        (previous) => {
          const updated = [
            ...previous,
          ]

          for (
            const record of finalRecords
          ) {
            const index =
              updated.findIndex(
                (item) =>
                  normalizeId(
                    item.id
                  ) ===
                  normalizeId(
                    record.id
                  )
              )

            if (
              index >= 0
            ) {
              updated[index] =
                record
            } else {
              updated.push(
                record
              )
            }
          }

          return updated
        }
      )


      /* =====================================================
         انتخاب اولین بخش Excel
      ===================================================== */

      const firstValidRecord =
        finalRecords.find(
          (record) =>
            normalizeId(
              record.sectionId
            )
        )

      if (
        firstValidRecord
      ) {
        setActiveSection(
          firstValidRecord.type
        )

        setSelectedSubSectionId(
          normalizeId(
            firstValidRecord.sectionId
          )
        )
      }


      alert(
        `${formatNumber(
          finalRecords.length
        )} تراکنش با موفقیت از Excel وارد شد.`
      )
    } catch (error) {
      console.error(
        "خطا در وارد کردن Excel:",
        error
      )

      alert(
        error?.message ||
          "وارد کردن فایل Excel انجام نشد."
      )
    } finally {
      setImportingExcel(false)

      event.target.value = ""
    }
  }


  /* =========================================================
     Add Record
  ========================================================= */

  function openAddRecordModal() {
    if (
      !selectedSubSection
    ) {
      alert(
        "ابتدا یک بخش مالی را انتخاب کنید."
      )

      return
    }

    setEditingRecord(null)

    setRecordForm({
      date: getToday(),
      description: "",
      amount: "",
      category: "",
      status: "تسویه شده",
      note: "",
    })

    setShowRecordModal(true)
  }


  /* =========================================================
     Edit Record
  ========================================================= */

  function openEditRecordModal(
    record
  ) {
    setEditingRecord(
      record
    )

    setRecordForm({
      date:
        normalizeDate(
          record.date
        ),

      description:
        record.description ||
        "",

      amount:
        record.amount ??
        "",

      category:
        record.category ||
        "",

      status:
        record.status ||
        "تسویه شده",

      note:
        record.note ||
        "",
    })

    setShowRecordModal(true)
  }


  /* =========================================================
     Save Record
  ========================================================= */

  async function handleSaveRecord() {
    if (
      !selectedSubSection
    ) {
      alert(
        "بخش مالی انتخاب نشده است."
      )

      return
    }

    if (
      !recordForm.description.trim()
    ) {
      alert(
        "شرح تراکنش را وارد کنید."
      )

      return
    }

    const amount =
      normalizeAmount(
        recordForm.amount
      )

    if (
      !amount ||
      amount <= 0
    ) {
      alert(
        "مبلغ معتبر وارد کنید."
      )

      return
    }

    setSaving(true)

    try {
      const record = {
        id:
          normalizeId(
            editingRecord?.id
          ) ||
          createId(
            activeSection
          ),

        date:
          normalizeDate(
            recordForm.date
          ),

        type:
          normalizeFinanceType(
            activeSection
          ),

        sectionId:
          normalizeId(
            selectedSubSection.id
          ),

        sectionName:
          selectedSubSection.name,

        section:
          selectedSubSection.name,

        description:
          recordForm.description.trim(),

        amount,

        category:
          recordForm.category.trim(),

        status:
          recordForm.status,

        note:
          recordForm.note.trim(),
      }

      const savedRecord =
        await saveFinanceRecord(
          record
        )

      setRecords(
        (previous) => {
          const exists =
            previous.some(
              (item) =>
                normalizeId(
                  item.id
                ) ===
                normalizeId(
                  savedRecord.id
                )
            )

          if (exists) {
            return previous.map(
              (item) =>
                normalizeId(
                  item.id
                ) ===
                normalizeId(
                  savedRecord.id
                )
                  ? savedRecord
                  : item
            )
          }

          return [
            ...previous,
            savedRecord,
          ]
        }
      )

      setShowRecordModal(
        false
      )

      setEditingRecord(null)
    } catch (error) {
      console.error(
        "خطا در ذخیره تراکنش:",
        error
      )

      alert(
        "ذخیره تراکنش انجام نشد."
      )
    } finally {
      setSaving(false)
    }
  }


  /* =========================================================
     Delete Record
  ========================================================= */

  async function handleDeleteRecord(
    record
  ) {
    const confirmed =
      window.confirm(
        `آیا از حذف این تراکنش مطمئن هستید؟\n\n${record.description}`
      )

    if (!confirmed) {
      return
    }

    try {
      await deleteFinanceRecord(
        record.id
      )

      setRecords(
        (previous) =>
          previous.filter(
            (item) =>
              normalizeId(
                item.id
              ) !==
              normalizeId(
                record.id
              )
          )
      )
    } catch (error) {
      console.error(
        "خطا در حذف تراکنش:",
        error
      )

      alert(
        "حذف تراکنش انجام نشد."
      )
    }
  }


  /* =========================================================
     Current Records
  ========================================================= */

  const currentRecords =
    useMemo(() => {
      if (
        activeSection ===
        "profit-loss"
      ) {
        return records
      }

      if (
        !selectedSubSection
      ) {
        return []
      }

      const selectedId =
        normalizeId(
          selectedSubSection.id
        )

      return records.filter(
        (record) =>
          normalizeFinanceType(
            record.type
          ) ===
            activeSection &&
          normalizeId(
            record.sectionId
          ) === selectedId
      )
    }, [
      records,
      activeSection,
      selectedSubSection,
    ])


  /* =========================================================
     Search
  ========================================================= */

  const filteredRecords =
    useMemo(() => {
      const text =
        normalizeName(
          searchText
        )

      if (!text) {
        return currentRecords
      }

      return currentRecords.filter(
        (record) =>
          normalizeName(
            record.description
          ).includes(text) ||

          normalizeName(
            record.category
          ).includes(text) ||

          normalizeName(
            record.sectionName ||
              record.section
          ).includes(text)
      )
    }, [
      currentRecords,
      searchText,
    ])


  /* =========================================================
     Current Summary
  ========================================================= */

  const currentTotal =
    useMemo(() => {
      return currentRecords.reduce(
        (
          total,
          record
        ) =>
          total +
          normalizeAmount(
            record.amount
          ),
        0
      )
    }, [
      currentRecords,
    ])


  const lastRecord =
    useMemo(() => {
      if (
        currentRecords.length ===
        0
      ) {
        return null
      }

      return [
        ...currentRecords,
      ].sort(
        (a, b) =>
          new Date(
            b.date
          ) -
          new Date(
            a.date
          )
      )[0]
    }, [
      currentRecords,
    ])


  /* =========================================================
     Global Statistics
  ========================================================= */

  const totalIncome =
    useMemo(() => {
      return records
        .filter(
          (record) =>
            normalizeFinanceType(
              record.type
            ) === "income"
        )
        .reduce(
          (
            total,
            record
          ) =>
            total +
            normalizeAmount(
              record.amount
            ),
          0
        )
    }, [records])


  const totalExpense =
    useMemo(() => {
      return records
        .filter(
          (record) =>
            normalizeFinanceType(
              record.type
            ) === "expense"
        )
        .reduce(
          (
            total,
            record
          ) =>
            total +
            normalizeAmount(
              record.amount
            ),
          0
        )
    }, [records])


  const profit =
    totalIncome -
    totalExpense


  /* =========================================================
     Loading
  ========================================================= */

  if (loading) {
    return (
      <div
        dir="rtl"
        className="flex min-h-screen items-center justify-center bg-[#1a1a1a] text-white"
      >
        <div className="text-gray-400">
          در حال بارگذاری اطلاعات مالی...
        </div>
      </div>
    )
  }


  /* =========================================================
     Main Render
  ========================================================= */

  return (
    <div
      dir="rtl"
      className="min-h-screen bg-[#1a1a1a] text-white"
    >

      {/* Header */}

      <header className="border-b border-[#d4a017]/30 bg-[#202020]">

        <div className="flex flex-col gap-4 px-4 py-5 md:flex-row md:items-center md:justify-between md:px-6">

          <div className="flex items-center gap-3 md:gap-4">

            <button
              onClick={onBack}
              className="flex h-10 w-10 items-center justify-center rounded-xl border border-white/10 bg-white/5 text-gray-300 transition hover:border-[#d4a017] hover:text-[#f0c040]"
              title="بازگشت"
            >
              <ArrowRight size={20} />
            </button>

            <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-[#d4a017]/15 text-[#f0c040]">
              <Wallet size={23} />
            </div>

            <div>
              <h1 className="text-xl font-bold">
                امور مالی
              </h1>

              <p className="mt-1 text-xs text-gray-500">
                مدیریت درآمد، هزینه و سود و زیان
              </p>
            </div>

          </div>


          <div className="flex flex-col gap-2 sm:flex-row">

            <input
              ref={excelInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={
                handleExcelImport
              }
              className="hidden"
            />

            <button
              onClick={
                openExcelFilePicker
              }
              disabled={
                importingExcel
              }
              className="flex items-center justify-center gap-2 rounded-xl border border-green-500/30 bg-green-500/10 px-4 py-2.5 text-sm font-bold text-green-400 transition hover:bg-green-500/20 disabled:cursor-not-allowed disabled:opacity-50"
            >
              <Upload size={17} />

              {importingExcel
                ? "در حال وارد کردن..."
                : "وارد کردن Excel"}
            </button>

            <button
              onClick={
                downloadFinanceExcelTemplate
              }
              className="flex items-center justify-center gap-2 rounded-xl border border-white/10 bg-white/5 px-4 py-2.5 text-sm text-gray-300 transition hover:border-[#d4a017]/50 hover:text-[#f0c040]"
            >
              <Download size={17} />

              قالب Excel
            </button>

          </div>

        </div>

      </header>


      {/* Main */}

      <main className="p-4 md:p-6">

        {/* Main Sections */}

        <div className="grid grid-cols-1 gap-4 md:grid-cols-3">

          {/* Income */}

          <button
            onClick={() =>
              handleMainSectionChange(
                "income"
              )
            }
            className={`rounded-2xl border p-5 text-right transition ${
              activeSection ===
              "income"
                ? "border-[#d4a017] bg-[#d4a017]/10"
                : "border-white/10 bg-[#202020] hover:border-[#d4a017]/50"
            }`}
          >

            <div className="mb-4 flex items-center justify-between">

              <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-green-500/10 text-green-400">
                <TrendingUp size={25} />
              </div>

              <TrendingUp
                size={20}
                className="text-gray-600"
              />

            </div>

            <h2 className="text-lg font-bold">
              درآمد
            </h2>

            <p className="mt-1 text-sm text-gray-500">
              مدیریت منابع درآمدی
            </p>

            <p className="mt-4 text-sm font-bold text-green-400">
              {formatMoney(
                totalIncome
              )}
            </p>

          </button>


          {/* Expense */}

          <button
            onClick={() =>
              handleMainSectionChange(
                "expense"
              )
            }
            className={`rounded-2xl border p-5 text-right transition ${
              activeSection ===
              "expense"
                ? "border-[#d4a017] bg-[#d4a017]/10"
                : "border-white/10 bg-[#202020] hover:border-[#d4a017]/50"
            }`}
          >

            <div className="mb-4 flex items-center justify-between">

              <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-red-500/10 text-red-400">
                <TrendingDown size={25} />
              </div>

              <TrendingDown
                size={20}
                className="text-gray-600"
              />

            </div>

            <h2 className="text-lg font-bold">
              هزینه
            </h2>

            <p className="mt-1 text-sm text-gray-500">
              مدیریت هزینه‌های سازمان
            </p>

            <p className="mt-4 text-sm font-bold text-red-400">
              {formatMoney(
                totalExpense
              )}
            </p>

          </button>


          {/* Profit / Loss */}

          <button
            onClick={() => {
              setActiveSection(
                "profit-loss"
              )

              setSelectedSubSectionId(
                ""
              )

              setSearchText("")
            }}
            className={`rounded-2xl border p-5 text-right transition ${
              activeSection ===
              "profit-loss"
                ? "border-[#d4a017] bg-[#d4a017]/10"
                : "border-white/10 bg-[#202020] hover:border-[#d4a017]/50"
            }`}
          >

            <div className="mb-4 flex items-center justify-between">

              <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-[#d4a017]/10 text-[#f0c040]">
                <DollarSign size={25} />
              </div>

              <DollarSign
                size={20}
                className="text-gray-600"
              />

            </div>

            <h2 className="text-lg font-bold">
              سود و زیان
            </h2>

            <p className="mt-1 text-sm text-gray-500">
              گزارش مالی و محاسبه سود و زیان
            </p>

            <p
              className={`mt-4 text-sm font-bold ${
                profit >= 0
                  ? "text-green-400"
                  : "text-red-400"
              }`}
            >
              {formatMoney(
                profit
              )}
            </p>

          </button>

        </div>


        {/* Content */}

        <div
          className={`mt-6 grid grid-cols-1 gap-6 ${
            activeSection ===
            "profit-loss"
              ? ""
              : "lg:grid-cols-[280px_1fr]"
          }`}
        >

          {/* Sidebar */}

          {activeSection !==
            "profit-loss" && (

            <aside className="rounded-2xl border border-white/10 bg-[#202020] p-4">

              <div className="mb-4 flex items-center justify-between">

                <div className="flex items-center gap-2">

                  {activeSection ===
                  "income" ? (
                    <TrendingUp
                      size={18}
                      className="text-green-400"
                    />
                  ) : (
                    <TrendingDown
                      size={18}
                      className="text-red-400"
                    />
                  )}

                  <h3 className="font-bold">
                    {activeSection ===
                    "income"
                      ? "بخش‌های درآمد"
                      : "بخش‌های هزینه"}
                  </h3>

                </div>

                <span className="rounded-lg bg-white/5 px-2 py-1 text-xs text-gray-500">
                  {currentSections.length}
                </span>

              </div>


              <div className="space-y-2">

                {currentSections.map(
                  (section) => {
                    const isSelected =
                      normalizeId(
                        selectedSubSectionId
                      ) ===
                      normalizeId(
                        section.id
                      )

                    return (
                      <div
                        key={
                          normalizeId(
                            section.id
                          )
                        }
                        className={`group flex items-center gap-2 rounded-xl border transition ${
                          isSelected
                            ? "border-[#d4a017] bg-[#d4a017]/10"
                            : "border-transparent hover:border-white/10 hover:bg-white/5"
                        }`}
                      >

                        <button
                          onClick={() => {
                            setSelectedSubSectionId(
                              normalizeId(
                                section.id
                              )
                            )

                            setSearchText("")
                          }}
                          className="flex flex-1 items-center gap-3 px-3 py-3 text-right"
                        >

                          <Building2
                            size={17}
                            className={
                              isSelected
                                ? "text-[#f0c040]"
                                : "text-gray-500"
                            }
                          />

                          <span className="text-sm">
                            {section.name}
                          </span>

                        </button>


                        <button
                          onClick={() =>
                            handleDeleteSection(
                              section
                            )
                          }
                          className="ml-2 flex h-8 w-8 items-center justify-center rounded-lg text-gray-600 transition hover:bg-red-500/10 hover:text-red-400"
                          title="حذف بخش"
                        >
                          <Trash2
                            size={16}
                          />
                        </button>

                      </div>
                    )
                  }
                )}


                <button
                  onClick={() => {
                    setNewSectionName(
                      ""
                    )

                    setShowAddModal(
                      true
                    )
                  }}
                  className="mt-3 flex w-full items-center justify-center gap-2 rounded-xl border border-dashed border-[#d4a017]/50 px-3 py-3 text-sm text-[#f0c040] transition hover:bg-[#d4a017]/10"
                >
                  <FolderPlus size={17} />

                  افزودن بخش{" "}

                  {activeSection ===
                  "income"
                    ? "درآمد"
                    : "هزینه"}

                </button>

              </div>

            </aside>
          )}


          {/* Main Panel */}

          <section
            className="rounded-2xl border border-white/10 bg-[#202020] p-4 md:p-6"
          >

            {activeSection ===
            "profit-loss" ? (

              <ProfitLossPanel
                totalIncome={
                  totalIncome
                }
                totalExpense={
                  totalExpense
                }
                profit={profit}
                records={records}
                config={config}
              />

            ) : selectedSubSection ? (

              <FinancialSectionPanel
                section={
                  selectedSubSection
                }
                type={
                  activeSection
                }
                records={
                  filteredRecords
                }
                total={
                  currentTotal
                }
                lastRecord={
                  lastRecord
                }
                searchText={
                  searchText
                }
                setSearchText={
                  setSearchText
                }
                onAddRecord={
                  openAddRecordModal
                }
                onEditRecord={
                  openEditRecordModal
                }
                onDeleteRecord={
                  handleDeleteRecord
                }
              />

            ) : (

              <EmptySectionPanel
                onAdd={() => {
                  setNewSectionName(
                    ""
                  )

                  setShowAddModal(
                    true
                  )
                }}
              />

            )}

          </section>

        </div>

      </main>


      {/* Add Section Modal */}

      {showAddModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4">

          <div className="w-full max-w-md rounded-2xl border border-[#d4a017]/30 bg-[#202020] p-6 shadow-2xl">

            <div className="mb-6 flex items-center justify-between">

              <div>

                <h2 className="text-lg font-bold">
                  افزودن بخش{" "}

                  {activeSection ===
                  "income"
                    ? "درآمد"
                    : "هزینه"}
                </h2>

                <p className="mt-1 text-xs leading-6 text-gray-500">
                  نام بخش مالی جدید را وارد کنید.
                  <br />
                  برای اتصال Excel، مقدار ستون
                  <span className="text-[#f0c040]">
                    {" "}section{" "}
                  </span>
                  یا
                  <span className="text-[#f0c040]">
                    {" "}sectionName
                  </span>
                  {" "}
                  باید با نام این بخش یکسان باشد.
                </p>

              </div>

              <button
                onClick={() =>
                  setShowAddModal(
                    false
                  )
                }
                className="rounded-lg p-2 text-gray-500 hover:bg-white/5 hover:text-white"
              >
                <X size={20} />
              </button>

            </div>


            <input
              value={
                newSectionName
              }
              onChange={(
                event
              ) =>
                setNewSectionName(
                  event.target.value
                )
              }
              onKeyDown={(
                event
              ) => {
                if (
                  event.key ===
                  "Enter"
                ) {
                  handleAddSection()
                }
              }}
              autoFocus
              placeholder={
                activeSection ===
                "income"
                  ? "مثلاً درآمد رستوران"
                  : "مثلاً هزینه تعمیرات"
              }
              className="w-full rounded-xl border border-white/10 bg-[#151515] px-4 py-3 text-sm text-white outline-none transition placeholder:text-gray-600 focus:border-[#d4a017]"
            />


            <div className="mt-5 flex gap-3">

              <button
                onClick={
                  handleAddSection
                }
                className="flex flex-1 items-center justify-center gap-2 rounded-xl bg-[#d4a017] px-4 py-3 text-sm font-bold text-black transition hover:bg-[#f0c040]"
              >
                <Plus size={18} />
                افزودن
              </button>

              <button
                onClick={() =>
                  setShowAddModal(
                    false
                  )
                }
                className="rounded-xl border border-white/10 px-5 py-3 text-sm text-gray-400 transition hover:bg-white/5 hover:text-white"
              >
                انصراف
              </button>

            </div>

          </div>

        </div>
      )}


      {/* Record Modal */}

      {showRecordModal && (
        <RecordModal
          form={recordForm}
          setForm={
            setRecordForm
          }
          editingRecord={
            editingRecord
          }
          section={
            selectedSubSection
          }
          type={
            activeSection
          }
          saving={saving}
          onSave={
            handleSaveRecord
          }
          onClose={() =>
            setShowRecordModal(
              false
            )
          }
        />
      )}

    </div>
  )
}


/* =========================================================
   Financial Section Panel
========================================================= */

function FinancialSectionPanel({
  section,
  type,
  records,
  total,
  lastRecord,
  searchText,
  setSearchText,
  onAddRecord,
  onEditRecord,
  onDeleteRecord,
}) {
  const isIncome =
    type === "income"

  return (
    <div>

      <div className="flex flex-col justify-between gap-4 border-b border-white/10 pb-5 md:flex-row md:items-center">

        <div className="flex items-center gap-4">

          <div
            className={`flex h-12 w-12 items-center justify-center rounded-xl ${
              isIncome
                ? "bg-green-500/10 text-green-400"
                : "bg-red-500/10 text-red-400"
            }`}
          >
            {isIncome ? (
              <TrendingUp size={24} />
            ) : (
              <TrendingDown size={24} />
            )}
          </div>

          <div>

            <h2 className="text-xl font-bold">
              {section.name}
            </h2>

            <p className="mt-1 text-xs text-gray-500">
              مدیریت تراکنش‌های این بخش
            </p>

          </div>

        </div>


        <button
          onClick={
            onAddRecord
          }
          className="flex items-center justify-center gap-2 rounded-xl bg-[#d4a017] px-4 py-2.5 text-sm font-bold text-black transition hover:bg-[#f0c040]"
        >
          <Plus size={17} />
          ثبت تراکنش
        </button>

      </div>


      <div className="mt-6 grid grid-cols-1 gap-4 md:grid-cols-3">

        <SummaryCard
          title={
            isIncome
              ? "مجموع درآمد"
              : "مجموع هزینه"
          }
          value={
            formatMoney(
              total
            )
          }
          icon={
            isIncome
              ? TrendingUp
              : TrendingDown
          }
        />

        <SummaryCard
          title="تعداد تراکنش"
          value={formatNumber(
            records.length
          )}
          icon={Receipt}
        />

        <SummaryCard
          title="آخرین ثبت"
          value={
            lastRecord?.date ||
            "—"
          }
          icon={
            CalendarDays
          }
        />

      </div>


      <div className="mt-6 flex flex-col gap-3 md:flex-row">

        <div className="relative flex-1">

          <Search
            size={18}
            className="absolute right-4 top-1/2 -translate-y-1/2 text-gray-600"
          />

          <input
            value={
              searchText
            }
            onChange={(event) =>
              setSearchText(
                event.target.value
              )
            }
            placeholder="جستجو در تراکنش‌ها..."
            className="w-full rounded-xl border border-white/10 bg-[#151515] py-3 pl-4 pr-11 text-sm text-white outline-none transition placeholder:text-gray-600 focus:border-[#d4a017]"
          />

        </div>

      </div>


      <div className="mt-6 overflow-hidden rounded-2xl border border-white/10">

        <div className="border-b border-white/10 bg-[#181818] px-5 py-4">

          <h3 className="font-bold">
            گزارش مالی
          </h3>

        </div>


        <div className="overflow-x-auto">

          <table className="w-full min-w-[900px]">

            <thead>

              <tr className="border-b border-white/10 bg-[#181818] text-right text-xs text-gray-500">

                <th className="px-5 py-4">
                  تاریخ
                </th>

                <th className="px-5 py-4">
                  شرح
                </th>

                <th className="px-5 py-4">
                  مبلغ
                </th>

                <th className="px-5 py-4">
                  دسته‌بندی
                </th>

                <th className="px-5 py-4">
                  وضعیت
                </th>

                <th className="px-5 py-4">
                  عملیات
                </th>

              </tr>

            </thead>


            <tbody>

              {records.length ===
              0 ? (

                <tr>

                  <td
                    colSpan="6"
                    className="px-5 py-16 text-center text-sm text-gray-600"
                  >

                    هنوز اطلاعات مالی برای این بخش
                    وارد نشده است.

                    <br />

                    <button
                      onClick={
                        onAddRecord
                      }
                      className="mt-4 inline-flex items-center gap-2 rounded-xl border border-[#d4a017]/40 px-4 py-2 text-xs text-[#f0c040] hover:bg-[#d4a017]/10"
                    >
                      <Plus size={15} />
                      ثبت اولین تراکنش
                    </button>

                  </td>

                </tr>

              ) : (

                records.map(
                  (record) => (

                    <tr
                      key={
                        normalizeId(
                          record.id
                        )
                      }
                      className="border-b border-white/5 transition hover:bg-white/[0.02]"
                    >

                      <td className="px-5 py-4 text-sm text-gray-400">
                        {record.date ||
                          "—"}
                      </td>

                      <td className="px-5 py-4 text-sm text-white">
                        {record.description ||
                          "—"}
                      </td>

                      <td
                        className={`px-5 py-4 text-sm font-bold ${
                          isIncome
                            ? "text-green-400"
                            : "text-red-400"
                        }`}
                      >
                        {formatMoney(
                          record.amount
                        )}
                      </td>

                      <td className="px-5 py-4 text-sm text-gray-400">
                        {record.category ||
                          "—"}
                      </td>

                      <td className="px-5 py-4">

                        <span className="rounded-lg bg-white/5 px-3 py-1 text-xs text-gray-400">
                          {record.status ||
                            "—"}
                        </span>

                      </td>

                      <td className="px-5 py-4">

                        <div className="flex items-center gap-2">

                          <button
                            onClick={() =>
                              onEditRecord(
                                record
                              )
                            }
                            className="flex h-8 w-8 items-center justify-center rounded-lg text-gray-500 transition hover:bg-[#d4a017]/10 hover:text-[#f0c040]"
                            title="ویرایش"
                          >
                            <Pencil
                              size={15}
                            />
                          </button>

                          <button
                            onClick={() =>
                              onDeleteRecord(
                                record
                              )
                            }
                            className="flex h-8 w-8 items-center justify-center rounded-lg text-gray-500 transition hover:bg-red-500/10 hover:text-red-400"
                            title="حذف"
                          >
                            <Trash2
                              size={15}
                            />
                          </button>

                        </div>

                      </td>

                    </tr>
                  )
                )
              )}

            </tbody>

          </table>

        </div>

      </div>


      <div className="mt-6 rounded-2xl border border-[#d4a017]/20 bg-[#d4a017]/5 p-5">

        <div className="mb-3 flex items-center gap-2 text-[#f0c040]">

          <FileText size={19} />

          <h3 className="font-bold">
            گزارش بخش
          </h3>

        </div>

        {records.length ===
        0 ? (

          <p className="text-sm leading-8 text-gray-400">
            در حال حاضر اطلاعاتی برای «
            {section.name}
            » ثبت نشده است.
            پس از ثبت تراکنش‌ها، مجموع مبلغ،
            تعداد تراکنش‌ها و آخرین ثبت در این
            بخش نمایش داده خواهد شد.
          </p>

        ) : (

          <p className="text-sm leading-8 text-gray-400">

            در بخش «
            {section.name}
            » تعداد{" "}

            <span className="font-bold text-white">
              {formatNumber(
                records.length
              )}
            </span>{" "}

            تراکنش ثبت شده است.

            مجموع مبلغ این بخش برابر با{" "}

            <span className="font-bold text-[#f0c040]">
              {formatMoney(
                total
              )}
            </span>{" "}

            است.

            {lastRecord && (
              <>
                آخرین تراکنش در تاریخ{" "}

                <span className="font-bold text-white">
                  {lastRecord.date}
                </span>{" "}

                با شرح «
                {lastRecord.description}
                » ثبت شده است.
              </>
            )}

          </p>
        )}

      </div>

    </div>
  )
}


/* =========================================================
   Profit / Loss Panel
========================================================= */

function ProfitLossPanel({
  totalIncome,
  totalExpense,
  profit,
  records,
  config,
}) {
  const incomeRecords =
    records.filter(
      (record) =>
        normalizeFinanceType(
          record.type
        ) === "income"
    )

  const expenseRecords =
    records.filter(
      (record) =>
        normalizeFinanceType(
          record.type
        ) === "expense"
    )


  const incomeBySection =
    (
      config.income || []
    ).map(
      (section) => ({
        ...section,

        total:
          incomeRecords
            .filter(
              (record) =>
                normalizeId(
                  record.sectionId
                ) ===
                normalizeId(
                  section.id
                )
            )
            .reduce(
              (
                total,
                record
              ) =>
                total +
                normalizeAmount(
                  record.amount
                ),
              0
            ),
      })
    )


  const expenseBySection =
    (
      config.expense || []
    ).map(
      (section) => ({
        ...section,

        total:
          expenseRecords
            .filter(
              (record) =>
                normalizeId(
                  record.sectionId
                ) ===
                normalizeId(
                  section.id
                )
            )
            .reduce(
              (
                total,
                record
              ) =>
                total +
                normalizeAmount(
                  record.amount
                ),
              0
            ),
      })
    )


  return (
    <div>

      <div className="border-b border-white/10 pb-5">

        <div className="flex items-center gap-4">

          <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-[#d4a017]/10 text-[#f0c040]">
            <DollarSign size={24} />
          </div>

          <div>

            <h2 className="text-xl font-bold">
              سود و زیان
            </h2>

            <p className="mt-1 text-xs text-gray-500">
              گزارش کلی وضعیت مالی
            </p>

          </div>

        </div>

      </div>


      <div className="mt-6 grid grid-cols-1 gap-4 md:grid-cols-3">

        <SummaryCard
          title="کل درآمد"
          value={formatMoney(
            totalIncome
          )}
          icon={TrendingUp}
        />

        <SummaryCard
          title="کل هزینه"
          value={formatMoney(
            totalExpense
          )}
          icon={TrendingDown}
        />

        <SummaryCard
          title={
            profit >= 0
              ? "سود"
              : "زیان"
          }
          value={formatMoney(
            Math.abs(profit)
          )}
          icon={DollarSign}
        />

      </div>


      <div
        className={`mt-6 rounded-2xl border p-6 ${
          profit >= 0
            ? "border-green-500/20 bg-green-500/5"
            : "border-red-500/20 bg-red-500/5"
        }`}
      >

        <div className="flex items-center gap-3">

          {profit >= 0 ? (
            <TrendingUp
              size={22}
              className="text-green-400"
            />
          ) : (
            <TrendingDown
              size={22}
              className="text-red-400"
            />
          )}

          <h3 className="font-bold">
            نتیجه مالی
          </h3>

        </div>

        <p className="mt-4 text-sm leading-8 text-gray-400">

          مجموع درآمد مجموعه{" "}

          <span className="font-bold text-green-400">
            {formatMoney(
              totalIncome
            )}
          </span>{" "}

          و مجموع هزینه‌ها{" "}

          <span className="font-bold text-red-400">
            {formatMoney(
              totalExpense
            )}
          </span>{" "}

          است.

          بنابراین نتیجه نهایی برابر با{" "}

          <span
            className={`font-bold ${
              profit >= 0
                ? "text-green-400"
                : "text-red-400"
            }`}
          >
            {formatMoney(
              Math.abs(profit)
            )}
          </span>{" "}

          {profit >= 0
            ? "سود"
            : "زیان"}{" "}

          است.

        </p>

      </div>


      <div className="mt-6 grid grid-cols-1 gap-6 lg:grid-cols-2">

        <div className="rounded-2xl border border-white/10 bg-[#181818] p-5">

          <div className="mb-4 flex items-center gap-2">

            <TrendingUp
              size={19}
              className="text-green-400"
            />

            <h3 className="font-bold">
              درآمد بر اساس بخش
            </h3>

          </div>


          <div className="space-y-3">

            {incomeBySection.map(
              (section) => (

                <div
                  key={
                    normalizeId(
                      section.id
                    )
                  }
                  className="flex items-center justify-between rounded-xl bg-white/[0.03] px-4 py-3"
                >

                  <span className="text-sm text-gray-400">
                    {section.name}
                  </span>

                  <span className="text-sm font-bold text-green-400">
                    {formatMoney(
                      section.total
                    )}
                  </span>

                </div>
              )
            )}

          </div>

        </div>


        <div className="rounded-2xl border border-white/10 bg-[#181818] p-5">

          <div className="mb-4 flex items-center gap-2">

            <TrendingDown
              size={19}
              className="text-red-400"
            />

            <h3 className="font-bold">
              هزینه بر اساس بخش
            </h3>

          </div>


          <div className="space-y-3">

            {expenseBySection.map(
              (section) => (

                <div
                  key={
                    normalizeId(
                      section.id
                    )
                  }
                  className="flex items-center justify-between rounded-xl bg-white/[0.03] px-4 py-3"
                >

                  <span className="text-sm text-gray-400">
                    {section.name}
                  </span>

                  <span className="text-sm font-bold text-red-400">
                    {formatMoney(
                      section.total
                    )}
                  </span>

                </div>
              )
            )}

          </div>

        </div>

      </div>


      <div className="mt-6 rounded-2xl border border-[#d4a017]/20 bg-[#d4a017]/5 p-5">

        <div className="mb-3 flex items-center gap-2 text-[#f0c040]">

          <FileText size={19} />

          <h3 className="font-bold">
            گزارش مالی
          </h3>

        </div>

        <p className="text-sm leading-8 text-gray-400">

          در حال حاضر{" "}

          <span className="font-bold text-green-400">
            {formatNumber(
              incomeRecords.length
            )}
          </span>{" "}

          تراکنش درآمدی و{" "}

          <span className="font-bold text-red-400">
            {formatNumber(
              expenseRecords.length
            )}
          </span>{" "}

          تراکنش هزینه‌ای ثبت شده است.

          مجموع درآمد برابر با{" "}

          <span className="font-bold text-green-400">
            {formatMoney(
              totalIncome
            )}
          </span>{" "}

          و مجموع هزینه برابر با{" "}

          <span className="font-bold text-red-400">
            {formatMoney(
              totalExpense
            )}
          </span>{" "}

          است.

        </p>

      </div>

    </div>
  )
}


/* =========================================================
   Empty State
========================================================= */

function EmptySectionPanel({
  onAdd,
}) {
  return (
    <div className="flex min-h-[450px] flex-col items-center justify-center text-center">

      <div className="flex h-16 w-16 items-center justify-center rounded-2xl bg-[#d4a017]/10 text-[#f0c040]">
        <FolderPlus size={30} />
      </div>

      <h2 className="mt-5 text-lg font-bold">
        هنوز بخشی ایجاد نشده است
      </h2>

      <p className="mt-2 max-w-md text-sm leading-7 text-gray-500">
        برای شروع یک بخش مالی جدید ایجاد کنید.
      </p>

      <button
        onClick={onAdd}
        className="mt-6 flex items-center gap-2 rounded-xl bg-[#d4a017] px-5 py-3 text-sm font-bold text-black transition hover:bg-[#f0c040]"
      >
        <Plus size={18} />
        افزودن بخش
      </button>

    </div>
  )
}


/* =========================================================
   Record Modal
========================================================= */

function RecordModal({
  form,
  setForm,
  editingRecord,
  section,
  type,
  saving,
  onSave,
  onClose,
}) {
  const isIncome =
    type === "income"

  function updateField(
    field,
    value
  ) {
    setForm(
      (previous) => ({
        ...previous,
        [field]: value,
      })
    )
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center overflow-y-auto bg-black/70 p-4">

      <div className="my-8 w-full max-w-2xl rounded-2xl border border-[#d4a017]/30 bg-[#202020] p-6 shadow-2xl">

        <div className="mb-6 flex items-start justify-between">

          <div>

            <h2 className="text-lg font-bold">
              {editingRecord
                ? "ویرایش تراکنش"
                : "ثبت تراکنش جدید"}
            </h2>

            <p className="mt-1 text-xs text-gray-500">
              {isIncome
                ? "ثبت درآمد"
                : "ثبت هزینه"}{" "}
              در بخش «
              {section?.name}
              »
            </p>

          </div>

          <button
            onClick={
              onClose
            }
            className="rounded-lg p-2 text-gray-500 hover:bg-white/5 hover:text-white"
          >
            <X size={20} />
          </button>

        </div>


        <div className="grid grid-cols-1 gap-4 md:grid-cols-2">

          <div>

            <label className="mb-2 block text-xs text-gray-500">
              تاریخ
            </label>

            <div className="relative">

              <CalendarDays
                size={17}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-600"
              />

              <input
                type="date"
                value={
                  form.date
                }
                onChange={(
                  event
                ) =>
                  updateField(
                    "date",
                    event.target.value
                  )
                }
                className="w-full rounded-xl border border-white/10 bg-[#151515] px-4 py-3 pr-10 text-sm text-white outline-none focus:border-[#d4a017]"
              />

            </div>

          </div>


          <div>

            <label className="mb-2 block text-xs text-gray-500">
              مبلغ
            </label>

            <input
              type="number"
              min="0"
              value={
                form.amount
              }
              onChange={(
                event
              ) =>
                updateField(
                  "amount",
                  event.target.value
                )
              }
              placeholder="مثلاً 15000000"
              className="w-full rounded-xl border border-white/10 bg-[#151515] px-4 py-3 text-sm text-white outline-none focus:border-[#d4a017]"
            />

          </div>


          <div className="md:col-span-2">

            <label className="mb-2 block text-xs text-gray-500">
              شرح تراکنش
            </label>

            <input
              value={
                form.description
              }
              onChange={(
                event
              ) =>
                updateField(
                  "description",
                  event.target.value
                )
              }
              placeholder="مثلاً فروش روزانه غذاخوری"
              className="w-full rounded-xl border border-white/10 bg-[#151515] px-4 py-3 text-sm text-white outline-none placeholder:text-gray-600 focus:border-[#d4a017]"
            />

          </div>


          <div>

            <label className="mb-2 block text-xs text-gray-500">
              دسته‌بندی
            </label>

            <input
              value={
                form.category
              }
              onChange={(
                event
              ) =>
                updateField(
                  "category",
                  event.target.value
                )
              }
              placeholder="مثلاً فروش، تعمیرات..."
              className="w-full rounded-xl border border-white/10 bg-[#151515] px-4 py-3 text-sm text-white outline-none placeholder:text-gray-600 focus:border-[#d4a017]"
            />

          </div>


          <div>

            <label className="mb-2 block text-xs text-gray-500">
              وضعیت
            </label>

            <select
              value={
                form.status
              }
              onChange={(
                event
              ) =>
                updateField(
                  "status",
                  event.target.value
                )
              }
              className="w-full rounded-xl border border-white/10 bg-[#151515] px-4 py-3 text-sm text-white outline-none focus:border-[#d4a017]"
            >

              <option value="تسویه شده">
                تسویه شده
              </option>

              <option value="در انتظار">
                در انتظار
              </option>

              <option value="لغو شده">
                لغو شده
              </option>

            </select>

          </div>


          <div className="md:col-span-2">

            <label className="mb-2 block text-xs text-gray-500">
              توضیحات
            </label>

            <textarea
              value={
                form.note
              }
              onChange={(
                event
              ) =>
                updateField(
                  "note",
                  event.target.value
                )
              }
              rows="3"
              placeholder="توضیحات تکمیلی..."
              className="w-full resize-none rounded-xl border border-white/10 bg-[#151515] px-4 py-3 text-sm text-white outline-none placeholder:text-gray-600 focus:border-[#d4a017]"
            />

          </div>

        </div>


        <div className="mt-6 flex gap-3">

          <button
            onClick={
              onSave
            }
            disabled={
              saving
            }
            className="flex flex-1 items-center justify-center gap-2 rounded-xl bg-[#d4a017] px-4 py-3 text-sm font-bold text-black transition hover:bg-[#f0c040] disabled:cursor-not-allowed disabled:opacity-50"
          >

            <Receipt size={17} />

            {saving
              ? "در حال ذخیره..."
              : editingRecord
              ? "ذخیره تغییرات"
              : "ثبت تراکنش"}

          </button>


          <button
            onClick={
              onClose
            }
            className="rounded-xl border border-white/10 px-5 py-3 text-sm text-gray-400 transition hover:bg-white/5 hover:text-white"
          >
            انصراف
          </button>

        </div>

      </div>

    </div>
  )
}


/* =========================================================
   Summary Card
========================================================= */

function SummaryCard({
  title,
  value,
  icon: Icon,
}) {
  return (
    <div className="rounded-2xl border border-white/10 bg-[#181818] p-5">

      <div className="flex items-center justify-between">

        <span className="text-sm text-gray-500">
          {title}
        </span>

        <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-[#d4a017]/10 text-[#f0c040]">
          <Icon size={18} />
        </div>

      </div>

      <div className="mt-4 text-xl font-bold">
        {value}
      </div>

    </div>
  )
}