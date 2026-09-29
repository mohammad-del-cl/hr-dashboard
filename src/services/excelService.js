import * as XLSX from "xlsx"

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
  const rows = [

    {
      id: "FIN-001",
      date: "2026-09-20",
      type: "income",
      section: "درآمد غذاخوری",
      description:
        "فروش روزانه غذا",
      amount: 15000000,
      category: "فروش",
      status:
        "تسویه شده",
      note: "نمونه",
    },

    {
      id: "FIN-002",
      date: "2026-09-20",
      type: "income",
      section: "درآمد غذاخوری",
      description:
        "فروش نوشیدنی",
      amount: 4500000,
      category: "فروش",
      status:
        "تسویه شده",
      note: "نمونه",
    },

    {
      id: "FIN-003",
      date: "2026-09-20",
      type: "income",
      section: "درآمد استخر",
      description:
        "فروش بلیت استخر",
      amount: 8000000,
      category: "بلیت",
      status:
        "تسویه شده",
      note: "نمونه",
    },

    {
      id: "FIN-004",
      date: "2026-09-21",
      type: "income",
      section: "درآمد استخر",
      description:
        "فروش اشتراک",
      amount: 12000000,
      category: "اشتراک",
      status:
        "تسویه شده",
      note: "نمونه",
    },

    {
      id: "FIN-005",
      date: "2026-09-20",
      type: "expense",
      section: "هزینه غذاخوری",
      description:
        "خرید مواد اولیه",
      amount: 3000000,
      category: "خرید",
      status:
        "تسویه شده",
      note: "نمونه",
    },

    {
      id: "FIN-006",
      date: "2026-09-21",
      type: "expense",
      section: "هزینه غذاخوری",
      description:
        "خرید نوشیدنی",
      amount: 1800000,
      category: "خرید",
      status:
        "در انتظار",
      note: "نمونه",
    },

    {
      id: "FIN-007",
      date: "2026-09-20",
      type: "expense",
      section: "هزینه استخر",
      description:
        "تعمیر تجهیزات",
      amount: 5500000,
      category: "تعمیرات",
      status:
        "تسویه شده",
      note: "نمونه",
    },

    {
      id: "FIN-008",
      date: "2026-09-21",
      type: "expense",
      section: "هزینه استخر",
      description:
        "خرید مواد بهداشتی",
      amount: 2100000,
      category:
        "مواد مصرفی",
      status:
        "تسویه شده",
      note: "نمونه",
    },

    {
      id: "FIN-009",
      date: "2026-09-21",
      type: "income",
      section: "درآمد غذاخوری",
      description:
        "فروش ناهار",
      amount: 6700000,
      category: "فروش",
      status:
        "تسویه شده",
      note: "",
    },

    {
      id: "FIN-010",
      date: "2026-09-21",
      type: "expense",
      section: "هزینه غذاخوری",
      description:
        "هزینه حمل مواد اولیه",
      amount: 750000,
      category:
        "حمل و نقل",
      status:
        "تسویه شده",
      note: "",
    },

  ]


  const worksheet =
    XLSX.utils.json_to_sheet(
      rows
    )


  const workbook =
    XLSX.utils.book_new()


  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "finance"
  )


  XLSX.writeFile(
    workbook,
    "finance-template.xlsx"
  )
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

  const DASHBOARD_SNAPSHOT_KEY =
    "__dashboardData"

  const hiddenColumns = new Set(
    Array.isArray(safeConfig.hiddenColumns)
      ? safeConfig.hiddenColumns.map(String)
      : []
  )

  const customColumns = Array.isArray(
    safeConfig.customColumns
  )
    ? safeConfig.customColumns
    : []

  /*
    =========================================================
    لایه اول: اطلاعات اصلی داشبورد

    این اطلاعات همان داده‌ای هستند که با Excel وارد داشبورد
    شده‌اند و باید مستقل از تغییرات بخش HR باقی بمانند.
    =========================================================
  */

  const isInternalKey = (key) => {
    const normalizedKey =
      String(key ?? "").trim()

    if (!normalizedKey) {
      return true
    }

    if (
      normalizedKey === DASHBOARD_SNAPSHOT_KEY ||
      normalizedKey === "department" ||
      normalizedKey === "sub_department"
    ) {
      return true
    }

    /*
      ستون‌های سفارشی متعلق به HR هستند.
    */
    if (
      normalizedKey.startsWith("custom-")
    ) {
      return true
    }

    /*
      وضعیت تأیید، داده داخلی برنامه است و نباید
      به عنوان ستون اطلاعاتی خروجی نمایش داده شود.
    */
    if (
      normalizedKey.startsWith("is_") &&
      normalizedKey.endsWith("_verified")
    ) {
      return true
    }

    return false
  }

  /*
    ترتیب کلیدهای snapshot حفظ می‌شود تا ترتیب ستون‌های
    اطلاعات داشبورد همان ترتیب Excel اصلی باشد.
  */
  const dashboardHeaders = []

  const dashboardHeaderSet =
    new Set()

  for (const employee of safeEmployees) {
    const snapshot =
      employee?.[DASHBOARD_SNAPSHOT_KEY]

    if (
      !snapshot ||
      typeof snapshot !== "object" ||
      Array.isArray(snapshot)
    ) {
      continue
    }

    for (const key of Object.keys(snapshot)) {
      if (isInternalKey(key)) {
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

  /*
    =========================================================
    لایه دوم: اطلاعات منابع انسانی

    فقط اطلاعاتی که واقعاً متعلق به HR هستند به این بخش
    اضافه می‌شوند.
    =========================================================
  */

  const hrHeaders = []

  const addHrHeader = (key) => {
    const normalizedKey =
      String(key ?? "").trim()

    if (!normalizedKey) {
      return
    }

    if (
      hiddenColumns.has(normalizedKey) &&
      normalizedKey.startsWith("custom-")
    ) {
      return
    }

    if (
      !hrHeaders.includes(normalizedKey)
    ) {
      hrHeaders.push(normalizedKey)
    }
  }

  /*
    سازماندهی فعلی کارمند متعلق به HR است.
  */
  addHrHeader("department")
  addHrHeader("sub_department")

  /*
    ستون‌های سفارشی ساخته‌شده در HR.
  */
  for (
    const column of customColumns
  ) {
    addHrHeader(
      column?.key
    )
  }

  /*
    =========================================================
    حذف ستون‌های کاملاً خالی

    مثال:
      first_name
      last_name

    اگر هیچ مقداری در هیچ ردیفی نداشته باشند،
    اصلاً در Excel نهایی ایجاد نمی‌شوند.
    =========================================================
  */

  const hasAnyValue = (key) => {
    return safeEmployees.some(
      (employee) => {
        const snapshot =
          employee?.[
            DASHBOARD_SNAPSHOT_KEY
          ]

        let value = ""

        if (
          snapshot &&
          typeof snapshot === "object" &&
          !Array.isArray(snapshot) &&
          Object.prototype.hasOwnProperty.call(
            snapshot,
            key
          )
        ) {
          value = snapshot[key]
        } else {
          value = employee?.[key]
        }

        return !(
          value === undefined ||
          value === null ||
          String(value).trim() === ""
        )
      }
    )
  }

  const filteredDashboardHeaders =
    dashboardHeaders.filter(
      hasAnyValue
    )

  /*
    department / sub_department اگر هیچ مقداری نداشته باشند
    حذف می‌شوند.

    ستون سفارشی HR حتی اگر فعلاً خالی باشد باقی می‌ماند،
    چون کاربر آن را عمداً در منابع انسانی ساخته است.
  */
  const filteredHrHeaders =
    hrHeaders.filter(
      (key) => {
        const isCustom =
          customColumns.some(
            (column) =>
              String(
                column?.key ?? ""
              ).trim() === key
          )

        if (isCustom) {
          return true
        }

        return hasAnyValue(key)
      }
    )

  const exportHeaders = [
    ...filteredDashboardHeaders,
    ...filteredHrHeaders,
  ].filter(
    (key, index, array) =>
      array.indexOf(key) === index
  )

  /*
    =========================================================
    ساخت ردیف‌ها
    =========================================================
  */

  const employeeRows =
    safeEmployees.map(
      (employee) => {
        const row = {}

        const snapshot =
          employee?.[
            DASHBOARD_SNAPSHOT_KEY
          ]

        const hasSnapshot =
          snapshot &&
          typeof snapshot === "object" &&
          !Array.isArray(snapshot)

        for (
          const key of exportHeaders
        ) {
          let value = ""

          /*
            اطلاعات داشبورد:
            همیشه از snapshot خوانده می‌شوند.
          */
          if (
            filteredDashboardHeaders.includes(
              key
            ) &&
            hasSnapshot
          ) {
            value =
              snapshot[key]
          }
          /*
            اطلاعات HR:
            از مقدار فعلی بخش منابع انسانی خوانده می‌شوند.
          */
          else {
            value =
              employee?.[key]
          }

          row[key] =
            value === undefined ||
            value === null
              ? ""
              : value
        }

        return row
      }
    )

  /*
    =========================================================
    ساخت Workbook
    =========================================================
  */

  const workbook =
    XLSX.utils.book_new()

  /*
    Sheet اصلی خروجی:
    فقط اطلاعات واقعی داشبورد + HR
    */
  const employeeSheet =
    XLSX.utils.json_to_sheet(
      employeeRows,
      {
        header:
          exportHeaders,
      }
    )

  XLSX.utils.book_append_sheet(
    workbook,
    employeeSheet,
    "employees"
  )

  /*
    =========================================================
    Snapshot داشبورد برای import مجدد
    =========================================================

    فقط لایه داشبورد ذخیره می‌شود.
    اطلاعات HR مثل department و custom columns
    داخل snapshot قرار نمی‌گیرند.
  */

  const dashboardRows =
    safeEmployees.map(
      (employee, index) => {
        const rawSnapshot =
          employee?.[
            DASHBOARD_SNAPSHOT_KEY
          ]

        let source = {}

        if (
          rawSnapshot &&
          typeof rawSnapshot === "object" &&
          !Array.isArray(rawSnapshot)
        ) {
          source =
            Object.keys(
              rawSnapshot
            ).reduce(
              (result, key) => {
                if (
                  !isInternalKey(
                    key
                  )
                ) {
                  result[key] =
                    rawSnapshot[key]
                }

                return result
              },
              {}
            )
        } else {
          /*
            برای سازگاری با داده‌های قدیمی:
            اگر snapshot وجود نداشت،
            فقط فیلدهای غیر-HR از employee گرفته می‌شوند.
          */
          source =
            Object.keys(
              employee || {}
            ).reduce(
              (result, key) => {
                if (
                  !isInternalKey(
                    key
                  )
                ) {
                  result[key] =
                    employee[key]
                }

                return result
              },
              {}
            )
        }

        return {
          id:
            String(
              employee?.id ??
                `__row_${index}`
            ),

          data:
            JSON.stringify(
              source
            ),
        }
      }
    )

  const dashboardSheet =
    XLSX.utils.json_to_sheet(
      dashboardRows,
      {
        header: [
          "id",
          "data",
        ],
      }
    )

  XLSX.utils.book_append_sheet(
    workbook,
    dashboardSheet,
    "dashboard_data"
  )

  /*
    =========================================================
    تنظیمات منابع انسانی
    =========================================================
  */

  const configRows = [
    {
      key: "departments",
      value: JSON.stringify(
        Array.isArray(
          safeConfig.departments
        )
          ? safeConfig.departments
          : []
      ),
    },
    {
      key: "customColumns",
      value: JSON.stringify(
        customColumns
      ),
    },
    {
      key: "hiddenColumns",
      value: JSON.stringify(
        Array.from(
          hiddenColumns
        )
      ),
    },
  ]

  const configSheet =
    XLSX.utils.json_to_sheet(
      configRows,
      {
        header: [
          "key",
          "value",
        ],
      }
    )

  XLSX.utils.book_append_sheet(
    workbook,
    configSheet,
    "hr_config"
  )

  /*
    =========================================================
    واحدها و زیرواحدها
    =========================================================
  */

  const departmentRows = []

  const departments =
    Array.isArray(
      safeConfig.departments
    )
      ? safeConfig.departments
      : []

  for (
    const department
    of departments
  ) {
    const children =
      Array.isArray(
        department?.children
      )
        ? department.children
        : []

    if (!children.length) {
      departmentRows.push({
        departmentId:
          department?.id ?? "",

        departmentName:
          department?.name ?? "",

        subDepartmentId:
          "",

        subDepartmentName:
          "",
      })

      continue
    }

    for (
      const child
      of children
    ) {
      departmentRows.push({
        departmentId:
          department?.id ?? "",

        departmentName:
          department?.name ?? "",

        subDepartmentId:
          child?.id ?? "",

        subDepartmentName:
          child?.name ?? "",
      })
    }
  }

  const departmentSheet =
    XLSX.utils.json_to_sheet(
      departmentRows,
      {
        header: [
          "departmentId",
          "departmentName",
          "subDepartmentId",
          "subDepartmentName",
        ],
      }
    )

  XLSX.utils.book_append_sheet(
    workbook,
    departmentSheet,
    "departments"
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
        const data = new Uint8Array(event.target.result)
        const workbook = XLSX.read(data, {
          type: "array",
          cellDates: true,
        })

        const employeeSheet =
          workbook.Sheets["employees"]

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
          فایل نهایی پروژه یک Sheet مستقل برای snapshot داشبورد دارد.
        */
        const dashboardSnapshots = new Map()
        const dashboardSnapshotsByIndex = []
        const dashboardSheet =
          workbook.Sheets["dashboard_data"]

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
              const id = String(
                row?.id ?? ""
              ).trim()

              let snapshot = null

              try {
                if (
                  typeof row?.data === "string" &&
                  row.data.trim()
                ) {
                  const parsed =
                    JSON.parse(row.data)

                  if (
                    parsed &&
                    typeof parsed === "object" &&
                    !Array.isArray(parsed)
                  ) {
                    snapshot = parsed
                  }
                }
              } catch {
                snapshot = null
              }

              dashboardSnapshotsByIndex[index] =
                snapshot

              if (id && snapshot) {
                dashboardSnapshots.set(
                  id,
                  snapshot
                )
              }
            }
          )
        }

        const employees =
          rawEmployees.map(
            (employee, index) => {
              const id = String(
                employee?.id ?? ""
              ).trim()

              const snapshotById =
                id
                  ? dashboardSnapshots.get(id)
                  : null

              const snapshot =
                snapshotById ||
                dashboardSnapshotsByIndex[index] ||
                Object.keys(employee || {})
                  .reduce(
                    (result, key) => {
                      result[key] = employee[key]
                      return result
                    },
                    {}
                  )

              return {
                ...employee,
                id:
                  id || crypto.randomUUID(),
                [DASHBOARD_SNAPSHOT_KEY]:
                  snapshot,
              }
            }
          )

        let config = null
        const configSheet =
          workbook.Sheets["hr_config"]

        if (configSheet) {
          const rows =
            XLSX.utils.sheet_to_json(
              configSheet,
              {
                defval: "",
              }
            )

          const values = {}

          for (const row of rows) {
            const key = String(
              row?.key ?? ""
            ).trim()

            if (key) {
              values[key] =
                row?.value ?? ""
            }
          }

          const parse = (
            value,
            fallback
          ) => {
            try {
              if (
                !value ||
                typeof value !== "string"
              ) {
                return fallback
              }

              return JSON.parse(value)
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

        if (
          !config &&
          workbook.Sheets["departments"]
        ) {
          const rows =
            XLSX.utils.sheet_to_json(
              workbook.Sheets["departments"],
              {
                defval: "",
              }
            )

          const departments = []

          for (const row of rows) {
            const departmentName =
              String(
                row?.departmentName ?? ""
              ).trim()

            const departmentId =
              String(
                row?.departmentId ?? ""
              ).trim()

            const subName =
              String(
                row?.subDepartmentName ?? ""
              ).trim()

            const subId =
              String(
                row?.subDepartmentId ?? ""
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
                name: departmentName,
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
                  item.name === subName
              )
            ) {
              department.children.push({
                id:
                  subId ||
                  `subdepartment-${departmentName}-${subName}`,
                name: subName,
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

    reader.readAsArrayBuffer(file)
  })
}


export function downloadFinalFinanceExcel(
  records = [],
  financeConfig = {}
) {
  const safeRecords = Array.isArray(records) ? records : []
  const safeConfig = financeConfig && typeof financeConfig === "object"
    ? financeConfig
    : {}

  const rows = safeRecords.map((record) => ({
    id: record?.id ?? "",
    date: record?.date ?? "",
    type: record?.type ?? "",
    section: record?.section ?? record?.sectionName ?? "",
    description: record?.description ?? "",
    amount: record?.amount ?? 0,
    category: record?.category ?? "",
    status: record?.status ?? "تسویه شده",
    note: record?.note ?? "",
    sectionId: record?.sectionId ?? "",
  }))

  const configRows = [
    ...(Array.isArray(safeConfig.income) ? safeConfig.income : []).map((item) => ({
      type: "income",
      id: item?.id ?? "",
      name: item?.name ?? "",
    })),
    ...(Array.isArray(safeConfig.expense) ? safeConfig.expense : []).map((item) => ({
      type: "expense",
      id: item?.id ?? "",
      name: item?.name ?? "",
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
    header: ["type", "id", "name"],
  })
  XLSX.utils.book_append_sheet(workbook, configSheet, "finance_config")

  const date = new Date().toISOString().slice(0, 10)
  XLSX.writeFile(workbook, `finance-final-${date}.xlsx`)
}


export function readFinanceExcelPackage(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result)
        const workbook = XLSX.read(data, {
          type: "array",
          cellDates: true,
        })

        const sheet =
          workbook.Sheets["finance"] ||
          workbook.Sheets["مالی"]

        if (!sheet) {
          throw new Error(
            "Sheet finance پیدا نشد. نام Sheet باید finance باشد."
          )
        }

        const rows = XLSX.utils.sheet_to_json(sheet, {
          defval: "",
        })

        const records = rows.map((row, index) => {
          const sectionName = getSectionName(row)

          return {
            id:
              String(row?.id ?? "").trim() ||
              `excel-${Date.now()}-${index}`,
            date: normalizeDate(row?.date),
            type: normalizeType(row?.type),
            section: sectionName,
            sectionName,
            sectionId: String(row?.sectionId ?? "").trim(),
            description: String(row?.description ?? "").trim(),
            amount: normalizeAmount(row?.amount),
            category: String(row?.category ?? "").trim(),
            status:
              String(row?.status ?? "تسویه شده").trim() ||
              "تسویه شده",
            note: String(row?.note ?? "").trim(),
          }
        })

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
            const type = normalizeType(row?.type)
            const name = String(row?.name ?? "").trim()
            const id = String(row?.id ?? "").trim()

            if (!name) continue

            const target = type === "expense" ? config.expense : config.income
            if (!target.some((item) => item.name === name)) {
              target.push({
                id: id || `${type}-${name}`,
                name,
              })
            }
          }
        }

        resolve({ records, config })
      } catch (error) {
        reject(
          new Error(
            error.message || "خواندن بسته Excel امور مالی انجام نشد."
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







