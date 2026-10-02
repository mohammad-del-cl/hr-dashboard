import * as XLSX from "xlsx"

const DASHBOARD_SNAPSHOT_KEY = "__dashboardData"

/* =========================================================
   Helpers
========================================================= */

/*
  تبدیل اعداد فارسی و عربی به انگلیسی
*/
function normalizeDigits(value) {
  return String(value ?? "")
    .replace(/[۰-۹]/g, (digit) => {
      return "۰۱۲۳۴۵۶۷۸۹".indexOf(digit)
    })
    .replace(/[٠-٩]/g, (digit) => {
      return "٠١٢٣٤٥٦٧٨٩".indexOf(digit)
    })
}


/*
  تبدیل مبلغ به Number
*/
function normalizeAmount(value) {
  if (
    value === undefined ||
    value === null ||
    value === ""
  ) {
    return 0
  }

  const normalized =
    normalizeDigits(value)

  const cleaned = normalized
    .replace(/,/g, "")
    .replace(/٬/g, "")
    .replace(/،/g, "")
    .replace(/٫/g, ".")
    .replace(/\s/g, "")
    .replace(/[^\d.-]/g, "")

  return Number(cleaned) || 0
}


/*
  تبدیل نوع تراکنش به income / expense
*/
function normalizeType(value) {
  const type = String(value ?? "")
    .trim()
    .replace(/\u200c/g, " ")
    .replace(/ي/g, "ی")
    .replace(/ى/g, "ی")
    .replace(/ك/g, "ک")
    .replace(/\s+/g, " ")
    .toLowerCase()

  if (
    type === "expense" ||
    type === "expenses" ||
    type === "هزینه" ||
    type === "هزینه ها" ||
    type === "هزینه‌ها" ||
    type === "خرج" ||
    type === "مخارج" ||
    type === "cost"
  ) {
    return "expense"
  }

  if (
    type === "income" ||
    type === "incomes" ||
    type === "درآمد"
  ) {
    return "income"
  }

  /*
    اگر نوع ناشناخته باشد،
    مثل قبل income در نظر گرفته می‌شود.
  */
  return "income"
}


/*
  تبدیل تاریخ
*/
function normalizeDate(value) {
  if (!value) {
    return ""
  }


  /*
    اگر Excel تاریخ را به صورت Date
    تحویل داده باشد.
  */
  if (value instanceof Date) {
    if (
      Number.isNaN(
        value.getTime()
      )
    ) {
      return ""
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


  /*
    تبدیل اعداد فارسی
  */
  const text =
    normalizeDigits(value)
      .trim()
      .replace(/[\/.]/g, "-")


  /*
    اگر تاریخ به شکل:
    2026-09-21
    باشد.
  */
  const match =
    text.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})$/
    )

  if (match) {
    const year =
      match[1]

    const month =
      String(
        match[2]
      ).padStart(2, "0")

    const day =
      String(
        match[3]
      ).padStart(2, "0")

    return `${year}-${month}-${day}`
  }


  return text
}


/*
  گرفتن نام بخش از Excel

  هر دو ستون قابل قبول هستند:

  section
  sectionName
*/
function getSectionName(row) {
  return String(
    row?.section ??
      row?.sectionName ??
      row?.Section ??
      row?.SectionName ??
      ""
  )
    .trim()
    .replace(/\u200c/g, " ")
    .replace(/ي/g, "ی")
    .replace(/ى/g, "ی")
    .replace(/ك/g, "ک")
    .replace(/\s+/g, " ")
}


/* =========================================================
   Employees Excel
========================================================= */

export function readExcelFile(file) {
  return new Promise(
    (resolve, reject) => {
      const reader =
        new FileReader()


      reader.onload = (
        event
      ) => {
        try {
          const data =
            new Uint8Array(
              event.target.result
            )


          const workbook =
            XLSX.read(data, {
              type: "array",
              cellDates: true,
            })


          const sheet =
            workbook.Sheets[
              "employees"
            ]


          if (!sheet) {
            reject(
              new Error(
                "Sheet employees پیدا نشد"
              )
            )

            return
          }


          const employees =
            XLSX.utils.sheet_to_json(
              sheet,
              {
                defval: "",
              }
            )


          resolve(
            employees
          )
        } catch (error) {
          reject(
            new Error(
              error.message ||
                "خواندن فایل Excel انجام نشد"
            )
          )
        }
      }


      reader.onerror = () => {
        reject(
          new Error(
            "خواندن فایل Excel انجام نشد"
          )
        )
      }


      reader.readAsArrayBuffer(
        file
      )
    }
  )
}


/* =========================================================
   Finance Excel
========================================================= */

export function readFinanceExcelFile(
  file
) {
  return new Promise(
    (resolve, reject) => {
      const reader =
        new FileReader()


      reader.onload = (
        event
      ) => {
        try {
          const data =
            new Uint8Array(
              event.target.result
            )


          const workbook =
            XLSX.read(data, {
              type: "array",
              cellDates: true,
            })


          /*
            Sheet اصلی باید finance باشد.
          */
          let sheet =
            workbook.Sheets[
              "finance"
            ]


          /*
            برای اطمینان، Sheet فارسی را هم
            قبول می‌کنیم.
          */
          if (!sheet) {
            sheet =
              workbook.Sheets[
                "مالی"
              ]
          }


          if (!sheet) {
            reject(
              new Error(
                "Sheet finance پیدا نشد. نام Sheet باید finance باشد."
              )
            )

            return
          }


          const rows =
            XLSX.utils.sheet_to_json(
              sheet,
              {
                defval: "",
              }
            )


          if (!rows.length) {
            reject(
              new Error(
                "فایل Excel امور مالی خالی است."
              )
            )

            return
          }


          /* =================================================
             بررسی ستون‌های ضروری
          ================================================= */

          const requiredColumns = [
            "date",
            "type",
            "section",
            "description",
            "amount",
          ]


          const firstRow =
            rows[0]


          /*
            sectionName هم به جای section
            قابل قبول است.
          */
          const hasSection =
            Object.prototype.hasOwnProperty.call(
              firstRow,
              "section"
            ) ||
            Object.prototype.hasOwnProperty.call(
              firstRow,
              "sectionName"
            )


          const missingColumns =
            requiredColumns.filter(
              (column) => {
                if (
                  column ===
                  "section"
                ) {
                  return !hasSection
                }

                return !Object.prototype.hasOwnProperty.call(
                  firstRow,
                  column
                )
              }
            )


          if (
            missingColumns.length >
            0
          ) {
            reject(
              new Error(
                `ستون‌های زیر در فایل وجود ندارند: ${missingColumns.join(
                  ", "
                )}`
              )
            )

            return
          }


          /* =================================================
             تبدیل ردیف‌های Excel
          ================================================= */

          const financeRecords =
            rows.map(
              (
                row,
                index
              ) => {
                const sectionName =
                  getSectionName(
                    row
                  )


                const id =
                  String(
                    row?.id ??
                      ""
                  ).trim() ||
                  `excel-${Date.now()}-${index}`


                return {
                  /*
                    شناسه
                  */
                  id,


                  /*
                    تاریخ
                  */
                  date:
                    normalizeDate(
                      row?.date
                    ),


                  /*
                    income / expense
                  */
                  type:
                    normalizeType(
                      row?.type
                    ),


                  /*
                    نام بخش
                  */
                  section:
                    sectionName,

                  sectionName:
                    sectionName,


                  /*
                    FinancialAffairs
                    بعداً ID واقعی بخش را
                    پیدا می‌کند.
                  */
                  sectionId:
                    String(
                      row?.sectionId ??
                        ""
                    ).trim(),


                  /*
                    شرح
                  */
                  description:
                    String(
                      row?.description ??
                        ""
                    ).trim(),


                  /*
                    مبلغ
                  */
                  amount:
                    normalizeAmount(
                      row?.amount
                    ),


                  /*
                    دسته‌بندی
                  */
                  category:
                    String(
                      row?.category ??
                        ""
                    ).trim(),


                  /*
                    وضعیت
                  */
                  status:
                    String(
                      row?.status ??
                        "تسویه شده"
                    ).trim() ||
                    "تسویه شده",


                  /*
                    توضیحات
                  */
                  note:
                    String(
                      row?.note ??
                        ""
                    ).trim(),
                }
              }
            )


          resolve(
            financeRecords
          )
        } catch (error) {
          console.error(
            "Finance Excel Error:",
            error
          )

          reject(
            new Error(
              error.message ||
                "خواندن فایل Excel امور مالی انجام نشد."
            )
          )
        }
      }


      reader.onerror = () => {
        reject(
          new Error(
            "خواندن فایل Excel امور مالی انجام نشد."
          )
        )
      }


      reader.readAsArrayBuffer(
        file
      )
    }
  )
}


/* =========================================================
   Finance Excel Template
========================================================= */

export function downloadFinanceExcelTemplate() {
  /*
    ستون‌ها: id, date, type, section, description, amount,
             category, status, note

    - date: به شکل 2026-09-20 (یا سلول تاریخ Excel)
    - type: income یا expense
    - section: بخش یا مسیر زیربخش با جداکننده‌ی «/»
      مثال: درآمد غذاخوری / رستوران / ناهار
  */
  const rows = [
    ["FIN-001", "2026-09-20", "income", "درآمد غذاخوری / رستوران / ناهار", "فروش ناهار", 15000000, "فروش", "تسویه شده", "نمونه"],
    ["FIN-002", "2026-09-20", "income", "درآمد غذاخوری / رستوران / شام", "فروش شام", 12000000, "فروش", "تسویه شده", "نمونه"],
    ["FIN-003", "2026-09-20", "income", "درآمد غذاخوری / کافه", "فروش نوشیدنی", 4500000, "فروش", "تسویه شده", "نمونه"],
    ["FIN-004", "2026-09-20", "income", "درآمد استخر / بلیت", "فروش بلیت استخر", 8000000, "بلیت", "تسویه شده", "نمونه"],
    ["FIN-005", "2026-09-21", "income", "درآمد استخر / اشتراک", "فروش اشتراک ماهانه", 12000000, "اشتراک", "تسویه شده", "نمونه"],
    ["FIN-006", "2026-09-21", "income", "درآمد غذاخوری", "درآمد متفرقه", 700000, "متفرقه", "تسویه شده", ""],
    ["FIN-007", "2026-09-20", "expense", "هزینه غذاخوری / مواد اولیه", "خرید مواد اولیه", 3000000, "خرید", "تسویه شده", "نمونه"],
    ["FIN-008", "2026-09-21", "expense", "هزینه غذاخوری / حقوق", "حقوق پرسنل آشپزخانه", 18000000, "حقوق", "در انتظار", "نمونه"],
    ["FIN-009", "2026-09-20", "expense", "هزینه استخر / تعمیرات", "تعمیر تجهیزات", 5500000, "تعمیرات", "تسویه شده", "نمونه"],
    ["FIN-010", "2026-09-21", "expense", "هزینه استخر / مواد بهداشتی", "خرید کلر", 2100000, "مواد مصرفی", "تسویه شده", ""],
    ["FIN-011", "2026-09-21", "expense", "هزینه اداری", "قبض برق", 750000, "اداری", "تسویه شده", ""],
  ]

  const worksheet = XLSX.utils.aoa_to_sheet([
    [
      "id",
      "date",
      "type",
      "section",
      "description",
      "amount",
      "category",
      "status",
      "note",
    ],
    ...rows,
  ])

  worksheet["!cols"] = [
    { wch: 10 },
    { wch: 12 },
    { wch: 10 },
    { wch: 38 },
    { wch: 30 },
    { wch: 14 },
    { wch: 14 },
    { wch: 12 },
    { wch: 10 },
  ]

  const workbook = XLSX.utils.book_new()

  XLSX.utils.book_append_sheet(workbook, worksheet, "finance")

  XLSX.writeFile(workbook, "finance-template.xlsx")
}
/* =========================================================
   Final HR Excel Export / Import
   اضافه شده بدون حذف توابع نسخه اصلی
========================================================= */

export function downloadEmployeesExcel(
  employees = [],
  hrConfig = {}
) {
  const safeEmployees = Array.isArray(employees)
    ? employees
    : []

  const safeConfig =
    hrConfig && typeof hrConfig === "object"
      ? hrConfig
      : {}

  const customColumns = Array.isArray(
    safeConfig.customColumns
  )
    ? safeConfig.customColumns
    : []

  /* =========================================================
     Helpers
  ========================================================= */

  const isInternalKey = (key) => {
    const normalizedKey =
      String(key ?? "").trim()

    if (!normalizedKey) {
      return true
    }

    if (
      normalizedKey ===
      DASHBOARD_SNAPSHOT_KEY
    ) {
      return true
    }

    if (
      normalizedKey.startsWith("is_") &&
      normalizedKey.endsWith("_verified")
    ) {
      return true
    }

    return false
  }

  const isBlank = (value) => {
    return (
      value === undefined ||
      value === null ||
      String(value).trim() === ""
    )
  }

  const getSnapshot = (employee) => {
    const snapshot =
      employee?.[
        DASHBOARD_SNAPSHOT_KEY
      ]

    if (
      snapshot &&
      typeof snapshot === "object" &&
      !Array.isArray(snapshot)
    ) {
      return snapshot
    }

    /*
      سازگاری با رکوردهای قدیمی بدون Snapshot.
      فیلدهای داخلی و custom-* در Snapshot قرار نمی‌گیرند.
    */
    return Object.keys(employee || {})
      .reduce(
        (result, key) => {
          if (
            !isInternalKey(key) &&
            !key.startsWith("custom-")
          ) {
            result[key] =
              employee[key]
          }

          return result
        },
        {}
      )
  }

  const snapshots =
    safeEmployees.map(getSnapshot)

  /* =========================================================
     اطلاعات اصلی داشبورد

     این ستون‌ها فقط از Snapshot خوانده می‌شوند.
     بنابراین تغییرات HR روی آنها اثر نمی‌گذارد.
  ========================================================= */

  const dashboardHeaders = []
  const dashboardHeaderSet = new Set()

  for (
    const snapshot of snapshots
  ) {
    for (
      const key of Object.keys(
        snapshot || {}
      )
    ) {
      if (
        isInternalKey(key) ||
        key.startsWith("custom-")
      ) {
        continue
      }

      if (
        !dashboardHeaderSet.has(key)
      ) {
        dashboardHeaderSet.add(key)
        dashboardHeaders.push(key)
      }
    }
  }

  /* حذف ستون‌های کاملاً خالی داشبورد */
  const dashboardHeadersWithValue =
    dashboardHeaders.filter(
      (key) =>
        snapshots.some(
          (snapshot) =>
            !isBlank(
              snapshot?.[key]
            )
        )
    )

  /* =========================================================
     اطلاعات HR

     نکته مهم:
     department و sub_department دیگر همیشه تکرار نمی‌شوند.
     فقط وقتی مقدار HR با مقدار اصلی داشبورد فرق کرده باشد،
     hr_department یا hr_sub_department ساخته می‌شود.
  ========================================================= */

  const changedHrHeaders = []

  for (
    let index = 0;
    index < safeEmployees.length;
    index += 1
  ) {
    const employee =
      safeEmployees[index] || {}

    const snapshot =
      snapshots[index] || {}

    for (
      const key of Object.keys(
        employee
      )
    ) {
      if (
        isInternalKey(key) ||
        key === "id" ||
        key.startsWith("custom-")
      ) {
        continue
      }

      const currentValue =
        employee?.[key]

      const originalValue =
        snapshot?.[key]

      if (
        String(
          currentValue ?? ""
        ).trim() !==
        String(
          originalValue ?? ""
        ).trim()
      ) {
        const hrKey =
          `hr_${key}`

        if (
          !changedHrHeaders.includes(
            hrKey
          )
        ) {
          changedHrHeaders.push(
            hrKey
          )
        }
      }
    }
  }

  /*
    ستون‌های سفارشی HR فقط در صورت داشتن مقدار واقعی
    وارد فایل می‌شوند تا ستون خالی و بی‌مصرف ساخته نشود.
  */
  const activeCustomHeaders =
    customColumns
      .map(
        (column) =>
          String(
            column?.key ?? ""
          ).trim()
      )
      .filter(Boolean)
      .filter(
        (key) =>
          !isInternalKey(key) &&
          !key.startsWith("hr_")
      )
      .filter(
        (key) =>
          safeEmployees.some(
            (employee) =>
              !isBlank(
                employee?.[key]
              )
          )
      )

  const hrHeaders = [
    ...activeCustomHeaders,
    ...changedHrHeaders,
  ]

  /* =========================================================
     یک Sheet مشترک

     ترتیب:
     اطلاعات اصلی داشبورد
     + اطلاعات HR جدید/تغییریافته
  ========================================================= */

  const exportHeaders = [
    ...dashboardHeadersWithValue,
    ...hrHeaders,
  ].filter(
    (key, index, array) =>
      array.indexOf(key) === index
  )

  const employeeRows =
    safeEmployees.map(
      (employee, index) => {
        const snapshot =
          snapshots[index] || {}

        const row = {}

        for (
          const key of exportHeaders
        ) {
          if (
            dashboardHeadersWithValue.includes(
              key
            )
          ) {
            /*
              داده اصلی داشبورد:
              همیشه از Snapshot.
            */
            row[key] =
              snapshot?.[key] ??
              ""

            continue
          }

          if (
            key.startsWith("hr_")
          ) {
            const originalKey =
              key.slice(3)

            /*
              داده تغییرکرده HR:
              اگر خالی شده باشد نیز باید خالی صادر شود.
            */
            row[key] =
              employee?.[
                originalKey
              ] ??
              ""

            continue
          }

          /* اطلاعات سفارشی HR */
          row[key] =
            employee?.[key] ??
            ""
        }

        return row
      }
    )

  /* =========================================================
     Workbook
     فقط یک Sheet به نام employees
  ========================================================= */

  const workbook =
    XLSX.utils.book_new()

  const employeeSheet =
    XLSX.utils.json_to_sheet(
      employeeRows,
      {
        header: exportHeaders,
      }
    )

  employeeSheet["!cols"] =
    exportHeaders.map(
      (header) => ({
        wch: Math.min(
          Math.max(
            String(
              header ?? ""
            ).length + 3,
            12
          ),
          35
        ),
      })
    )

  XLSX.utils.book_append_sheet(
    workbook,
    employeeSheet,
    "employees"
  )

  const date =
    new Date()
      .toISOString()
      .slice(0, 10)

  XLSX.writeFile(
    workbook,
    `employees-final-${date}.xlsx`
  )
}

export function readHRExcelPackage(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()

    reader.onload = (event) => {
      try {
        const data =
          new Uint8Array(
            event.target.result
          )

        const workbook =
          XLSX.read(data, {
            type: "array",
            cellDates: true,
          })

        const employeeSheet =
          workbook.Sheets[
            "employees"
          ]

        if (!employeeSheet) {
          throw new Error(
            "Sheet employees پیدا نشد"
          )
        }

        const rawEmployees =
          XLSX.utils.sheet_to_json(
            employeeSheet,
            {
              defval: "",
            }
          )


        /*
          =====================================================
          فرمت جدید خروجی نهایی HR: فقط یک Sheet به نام employees
          =====================================================

          ساختار این فایل:
          - ستون‌های عادی = Snapshot اصلی داشبورد
          - ستون‌های hr_* = مقدار فعلی HR فقط در صورت تفاوت
          - ستون‌های custom-* = اطلاعات سفارشی HR

          در نتیجه می‌توان فایل نهایی را دوباره Import کرد و
          Snapshot داشبورد را بدون از دست رفتن بازیابی کرد.
        */
        const hasSeparateDashboardSheet =
          Boolean(
            workbook.Sheets["dashboard_data"]
          )

        const hasSeparateHrSheet =
          Boolean(
            workbook.Sheets["hr_data"]
          )

        if (
          !hasSeparateDashboardSheet &&
          !hasSeparateHrSheet
        ) {
          const employees =
            rawEmployees.map(
              (employee, index) => {
                const row =
                  employee || {}

                const rowId =
                  String(
                    row?.id ??
                    ""
                  ).trim()

                const id =
                  rowId ||
                  (
                    typeof crypto !==
                      "undefined" &&
                    typeof crypto.randomUUID ===
                      "function"
                      ? crypto.randomUUID()
                      : `employee-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 10)}`
                  )

                const snapshot =
                  Object.keys(row)
                    .reduce(
                      (result, key) => {
                        if (
                          key ===
                          DASHBOARD_SNAPSHOT_KEY
                        ) {
                          return result
                        }

                        if (
                          key.startsWith("hr_")
                        ) {
                          return result
                        }

                        if (
                          key.startsWith("is_") &&
                          key.endsWith("_verified")
                        ) {
                          return result
                        }

                        if (
                          key.startsWith("custom-")
                        ) {
                          return result
                        }

                        result[key] =
                          row[key]

                        return result
                      },
                      {}
                    )

                const currentEmployee = {
                  ...row,
                  id,
                }

                for (
                  const key of Object.keys(row)
                ) {
                  if (
                    !key.startsWith("hr_")
                  ) {
                    continue
                  }

                  const originalKey =
                    key.slice(3)

                  if (!originalKey) {
                    continue
                  }

                  currentEmployee[
                    originalKey
                  ] =
                    row[key]
                }

                return {
                  ...currentEmployee,
                  id,
                  [DASHBOARD_SNAPSHOT_KEY]:
                    snapshot,
                }
              }
            )

          resolve({
            employees,
            config: null,
          })

          return
        }

        /*
          =====================================================
          Snapshot داشبورد
          =====================================================
        */

        const dashboardSnapshots =
          new Map()

        const dashboardSnapshotsByIndex = []

        const dashboardSheet =
          workbook.Sheets[
            "dashboard_data"
          ]

        if (dashboardSheet) {
          const dashboardRows =
            XLSX.utils.sheet_to_json(
              dashboardSheet,
              {
                defval: "",
              }
            )

          dashboardRows.forEach(
            (row, index) => {
              const id =
                String(
                  row?.id ??
                  ""
                ).trim()

              let snapshot =
                null

              try {
                if (
                  typeof row?.data ===
                    "string" &&
                  row.data.trim()
                ) {
                  const parsed =
                    JSON.parse(
                      row.data
                    )

                  if (
                    parsed &&
                    typeof parsed ===
                      "object" &&
                    !Array.isArray(
                      parsed
                    )
                  ) {
                    snapshot =
                      parsed
                  }
                }
              } catch {
                snapshot =
                  null
              }

              dashboardSnapshotsByIndex[
                index
              ] = snapshot

              if (
                id &&
                snapshot
              ) {
                dashboardSnapshots.set(
                  id,
                  snapshot
                )
              }
            }
          )
        }

        /*
          =====================================================
          داده‌های فعلی HR
          =====================================================
        */

        const hrRowsById =
          new Map()

        const hrSheet =
          workbook.Sheets[
            "hr_data"
          ]

        if (hrSheet) {
          const hrRows =
            XLSX.utils.sheet_to_json(
              hrSheet,
              {
                defval: "",
              }
            )

          for (
            const row of hrRows
          ) {
            const id =
              String(
                row?.id ??
                ""
              ).trim()

            if (id) {
              hrRowsById.set(
                id,
                row
              )
            }
          }
        }

        const employees =
          rawEmployees.map(
            (employee, index) => {
              const id =
                String(
                  employee?.id ??
                  ""
                ).trim()

              const snapshotById =
                id
                  ? dashboardSnapshots.get(
                      id
                    )
                  : null

              const snapshot =
                snapshotById ||
                dashboardSnapshotsByIndex[
                  index
                ] ||
                Object.keys(
                  employee || {}
                ).reduce(
                  (result, key) => {
                    result[key] =
                      employee[key]

                    return result
                  },
                  {}
                )

              const hrRow =
                id
                  ? hrRowsById.get(
                      id
                    )
                  : null

              const mergedEmployee = {
                ...employee,
              }

              /*
                اطلاعات HR مستقیم:
                department / sub_department /
                custom columns
              */
              if (hrRow) {
                for (
                  const key of Object.keys(
                    hrRow
                  )
                ) {
                  if (key === "id") {
                    continue
                  }

                  if (
                    key.startsWith(
                      "hr_"
                    )
                  ) {
                    const originalKey =
                      key.slice(3)

                    if (
                      originalKey
                    ) {
                      mergedEmployee[
                        originalKey
                      ] =
                        hrRow[key]
                    }

                    continue
                  }

                  mergedEmployee[key] =
                    hrRow[key]
                }
              }

              return {
                ...mergedEmployee,
                id:
                  id ||
                  (
                    typeof crypto !==
                      "undefined" &&
                    typeof crypto.randomUUID ===
                      "function"
                      ? crypto.randomUUID()
                      : `employee-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
                  ),
                [DASHBOARD_SNAPSHOT_KEY]:
                  snapshot,
              }
            }
          )

        /*
          =====================================================
          HR Config
          =====================================================
        */

        let config = null

        const configSheet =
          workbook.Sheets[
            "hr_config"
          ]

        if (configSheet) {
          const rows =
            XLSX.utils.sheet_to_json(
              configSheet,
              {
                defval: "",
              }
            )

          const values = {}

          for (
            const row of rows
          ) {
            const key =
              String(
                row?.key ??
                ""
              ).trim()

            if (key) {
              values[key] =
                row?.value ??
                ""
            }
          }

          const parse = (
            value,
            fallback
          ) => {
            try {
              if (
                !value ||
                typeof value !==
                  "string"
              ) {
                return fallback
              }

              return JSON.parse(
                value
              )
            } catch {
              return fallback
            }
          }

          config = {
            departments:
              parse(
                values.departments,
                []
              ),
            customColumns:
              parse(
                values.customColumns,
                []
              ),
            hiddenColumns:
              parse(
                values.hiddenColumns,
                []
              ),
          }
        }

        /*
          اگر فایل قدیمی departments را به صورت Sheet جدا داشته باشد
          ولی hr_config نداشته باشد، همان ساختار را می‌سازیم.
        */

        if (
          !config &&
          workbook.Sheets[
            "departments"
          ]
        ) {
          const rows =
            XLSX.utils.sheet_to_json(
              workbook.Sheets[
                "departments"
              ],
              {
                defval: "",
              }
            )

          const departments = []

          for (
            const row of rows
          ) {
            const departmentName =
              String(
                row?.departmentName ??
                ""
              ).trim()

            const departmentId =
              String(
                row?.departmentId ??
                ""
              ).trim()

            const subName =
              String(
                row?.subDepartmentName ??
                ""
              ).trim()

            const subId =
              String(
                row?.subDepartmentId ??
                ""
              ).trim()

            if (!departmentName) {
              continue
            }

            let department =
              departments.find(
                (item) =>
                  item.name ===
                  departmentName
              )

            if (!department) {
              department = {
                id:
                  departmentId ||
                  `department-${departmentName}`,
                name:
                  departmentName,
                children: [],
              }

              departments.push(
                department
              )
            }

            if (
              subName &&
              !department.children.some(
                (item) =>
                  item.name ===
                  subName
              )
            ) {
              department.children.push({
                id:
                  subId ||
                  `subdepartment-${departmentName}-${subName}`,
                name:
                  subName,
              })
            }
          }

          config = {
            departments,
            customColumns: [],
            hiddenColumns: [],
          }
        }

        resolve({
          employees,
          config,
        })
      } catch (error) {
        reject(
          new Error(
            error.message ||
              "خواندن بسته Excel منابع انسانی انجام نشد."
          )
        )
      }
    }

    reader.onerror = () => {
      reject(
        new Error(
          "خواندن فایل Excel منابع انسانی انجام نشد."
        )
      )
    }

    reader.readAsArrayBuffer(
      file
    )
  })
}

/* =========================================================
   Finance Excel — خواندن و نوشتن (نسخه‌ی انعطاف‌پذیر)

   - نام Sheet مهم نیست: اگر finance / مالی پیدا نشد، اولین
     Sheet که ستون‌های تاریخ و مبلغ دارد خوانده می‌شود.
   - عنوان ستون‌ها فارسی یا انگلیسی، با هر حروف بزرگ/کوچک.
   - تاریخ: متن (2026-09-20) یا سلول تاریخ واقعی Excel.
   - بخش و زیربخش با جداکننده‌ی «/» نوشته می‌شود:
       درآمد غذاخوری / رستوران / ناهار
========================================================= */

const FINANCE_HEADER_ALIASES = {
  id: ["id", "شناسه", "کد"],
  date: ["date", "تاریخ"],
  type: ["type", "نوع", "نوعتراکنش"],
  section: [
    "section",
    "sectionname",
    "sectionpath",
    "بخش",
    "نامبخش",
    "مسیربخش",
    "زیربخش",
  ],
  description: ["description", "desc", "شرح", "شرحتراکنش"],
  amount: ["amount", "مبلغ", "مقدار"],
  category: ["category", "دسته", "دستهبندی"],
  status: ["status", "وضعیت"],
  note: ["note", "notes", "توضیحات", "یادداشت"],
  sectionId: ["sectionid", "شناسهبخش"],
}

function normalizeHeaderKey(value) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[\u200c\s_\-.]/g, "")
    .replace(/ي/g, "ی")
    .replace(/ى/g, "ی")
    .replace(/ك/g, "ک")
}

const FINANCE_HEADER_LOOKUP = (() => {
  const lookup = {}

  Object.entries(FINANCE_HEADER_ALIASES).forEach(
    ([field, aliases]) => {
      aliases.forEach((alias) => {
        lookup[normalizeHeaderKey(alias)] = field
      })
    }
  )

  return lookup
})()


/*
  مسیر بخش را استاندارد می‌کند:
  "غذاخوری>رستوران \ ناهار"  →  "غذاخوری / رستوران / ناهار"
*/
function normalizeSectionPathText(value) {
  return String(value ?? "")
    .replace(/\u200c/g, " ")
    .replace(/ي/g, "ی")
    .replace(/ى/g, "ی")
    .replace(/ك/g, "ک")
    .split(/[\/\\>»]/)
    .map((part) => part.replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .join(" / ")
}


/*
  تاریخ سلول Excel → YYYY-MM-DD
  (شماره‌ی سریال Excel، Date یا متن)
*/
function financeCellToDate(value) {
  if (value === undefined || value === null || value === "") {
    return ""
  }

  if (value instanceof Date) {
    return normalizeDate(value)
  }

  if (
    typeof value === "number" &&
    value > 20000 &&
    value < 80000
  ) {
    const parsed = XLSX.SSF.parse_date_code(value)

    if (parsed && parsed.y) {
      return `${parsed.y}-${String(parsed.m).padStart(
        2,
        "0"
      )}-${String(parsed.d).padStart(2, "0")}`
    }
  }

  return normalizeDate(value)
}


/*
  ردیف عنوان را پیدا می‌کند (در ۱۵ ردیف اول)
  و ستون‌ها را به نام استاندارد نگاشت می‌کند.
*/
function detectFinanceHeader(matrix) {
  const limit = Math.min(matrix.length, 15)

  for (let rowIndex = 0; rowIndex < limit; rowIndex++) {
    const columns = {}

    ;(matrix[rowIndex] || []).forEach((cell, colIndex) => {
      const field =
        FINANCE_HEADER_LOOKUP[normalizeHeaderKey(cell)]

      if (field && columns[field] === undefined) {
        columns[field] = colIndex
      }
    })

    if (
      columns.date !== undefined &&
      columns.amount !== undefined &&
      columns.section !== undefined
    ) {
      return { rowIndex, columns }
    }
  }

  return null
}


function pickFinanceSheet(workbook) {
  const names = workbook.SheetNames || []

  const preferred = names.find((name) =>
    ["finance", "مالی"].includes(
      String(name).trim().toLowerCase()
    )
  )

  const ordered = [
    ...(preferred ? [preferred] : []),
    ...names.filter(
      (name) =>
        name !== preferred &&
        String(name).trim().toLowerCase() !== "finance_config"
    ),
  ]

  for (const name of ordered) {
    const sheet = workbook.Sheets[name]

    if (!sheet) continue

    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: true,
    })

    const header = detectFinanceHeader(matrix)

    if (header) {
      return { name, matrix, header }
    }
  }

  return null
}


export function downloadFinalFinanceExcel(
  records = [],
  financeConfig = {}
) {
  const safeRecords = Array.isArray(records) ? records : []

  const safeConfig =
    financeConfig && typeof financeConfig === "object"
      ? financeConfig
      : {}

  const incomeSections = Array.isArray(safeConfig.income)
    ? safeConfig.income
    : []

  const expenseSections = Array.isArray(safeConfig.expense)
    ? safeConfig.expense
    : []

  /*
    مسیر کامل بخش (با زیربخش‌ها) از روی Config
  */
  function pathOf(sections, sectionId) {
    const names = []
    const seen = new Set()

    let current = sections.find(
      (item) => String(item?.id ?? "") === String(sectionId ?? "")
    )

    while (current && !seen.has(String(current.id))) {
      seen.add(String(current.id))
      names.unshift(current.name)

      const parentId = String(current.parentId ?? "")

      current = parentId
        ? sections.find(
            (item) => String(item?.id ?? "") === parentId
          )
        : null
    }

    return names.join(" / ")
  }

  const rows = safeRecords.map((record) => {
    const type = record?.type ?? ""

    const path = pathOf(
      type === "expense" ? expenseSections : incomeSections,
      record?.sectionId
    )

    return {
      id: record?.id ?? "",
      date: record?.date ?? "",
      type,
      section:
        path || record?.section || record?.sectionName || "",
      description: record?.description ?? "",
      amount: record?.amount ?? 0,
      category: record?.category ?? "",
      status: record?.status ?? "تسویه شده",
      note: record?.note ?? "",
      sectionId: record?.sectionId ?? "",
    }
  })

  const configRows = [
    ...incomeSections.map((item) => ({
      type: "income",
      id: item?.id ?? "",
      name: item?.name ?? "",
      parentId: item?.parentId ?? "",
    })),
    ...expenseSections.map((item) => ({
      type: "expense",
      id: item?.id ?? "",
      name: item?.name ?? "",
      parentId: item?.parentId ?? "",
    })),
  ]

  const workbook = XLSX.utils.book_new()

  const financeSheet = XLSX.utils.json_to_sheet(rows, {
    header: [
      "id",
      "date",
      "type",
      "section",
      "description",
      "amount",
      "category",
      "status",
      "note",
      "sectionId",
    ],
  })

  XLSX.utils.book_append_sheet(workbook, financeSheet, "finance")

  const configSheet = XLSX.utils.json_to_sheet(configRows, {
    header: ["type", "id", "name", "parentId"],
  })

  XLSX.utils.book_append_sheet(
    workbook,
    configSheet,
    "finance_config"
  )

  const date = new Date().toISOString().slice(0, 10)

  XLSX.writeFile(workbook, `finance-final-${date}.xlsx`)
}


export function readFinanceExcelPackage(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result)

        /*
          cellDates خاموش است تا تاریخ‌ها به‌صورت شماره‌ی سریال
          بیایند و بدون خطای منطقه‌ی زمانی تبدیل شوند.
        */
        const workbook = XLSX.read(data, {
          type: "array",
          cellDates: false,
        })

        const picked = pickFinanceSheet(workbook)

        if (!picked) {
          throw new Error(
            "ستون‌های لازم پیدا نشد. فایل باید حداقل ستون‌های date (تاریخ)، type (نوع)، section (بخش) و amount (مبلغ) داشته باشد. " +
              `Sheetهای فایل: ${(workbook.SheetNames || []).join(", ")}`
          )
        }

        const { matrix, header } = picked
        const { columns, rowIndex } = header

        if (columns.type === undefined) {
          throw new Error(
            "ستون type (نوع) پیدا نشد. مقدار آن باید income (درآمد) یا expense (هزینه) باشد."
          )
        }

        const get = (row, field) =>
          columns[field] === undefined
            ? ""
            : row[columns[field]] ?? ""

        const records = []

        for (let i = rowIndex + 1; i < matrix.length; i++) {
          const row = matrix[i] || []

          const description = String(get(row, "description")).trim()
          const amount = normalizeAmount(get(row, "amount"))
          const sectionPath = normalizeSectionPathText(
            get(row, "section")
          )

          // ردیف کاملاً خالی
          if (!description && !amount && !sectionPath) {
            continue
          }

          // ردیفی که نه مبلغ دارد نه شرح
          if (!amount && !description) {
            continue
          }

          const section = sectionPath || "بدون بخش"

          records.push({
            id:
              String(get(row, "id")).trim() ||
              `excel-${Date.now()}-${i}`,
            date: financeCellToDate(get(row, "date")),
            type: normalizeType(get(row, "type")),
            section,
            sectionName: section,
            sectionId: String(get(row, "sectionId")).trim(),
            description,
            amount,
            category: String(get(row, "category")).trim(),
            status:
              String(get(row, "status")).trim() || "تسویه شده",
            note: String(get(row, "note")).trim(),
          })
        }

        /*
          Sheet اختیاری finance_config (خروجی «دانلود اکسل نهایی»)
        */
        let config = null

        const configSheet = workbook.Sheets["finance_config"]

        if (configSheet) {
          const configRows = XLSX.utils.sheet_to_json(configSheet, {
            defval: "",
          })

          config = {
            id: "finance-config",
            income: [],
            expense: [],
          }

          for (const row of configRows) {
            const lower = {}

            Object.keys(row).forEach((key) => {
              lower[normalizeHeaderKey(key)] = row[key]
            })

            const type = normalizeType(lower.type ?? lower["نوع"])
            const name = String(
              lower.name ?? lower["نام"] ?? ""
            ).trim()
            const id = String(lower.id ?? "").trim()
            const parentId = String(lower.parentid ?? "").trim()

            if (!name) continue

            const target =
              type === "expense" ? config.expense : config.income

            const exists = target.some((item) =>
              id
                ? item.id === id
                : item.name === name &&
                  (item.parentId || "") === parentId
            )

            if (!exists) {
              target.push({
                id: id || `${type}-${name}`,
                name,
                parentId,
              })
            }
          }
        }

        resolve({ records, config })
      } catch (error) {
        reject(
          new Error(
            error.message ||
              "خواندن بسته Excel امور مالی انجام نشد."
          )
        )
      }
    }

    reader.onerror = () => {
      reject(new Error("خواندن فایل Excel امور مالی انجام نشد."))
    }

    reader.readAsArrayBuffer(file)
  })
}
