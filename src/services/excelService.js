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