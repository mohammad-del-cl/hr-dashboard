import {
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react"

import Chart from "react-apexcharts"

import HumanResources from "./components/HumanResources"
import FinancialAffairs from "./components/FinancialAffairs"
import Login from "./components/Login"

import {
  getEmployees,
  saveEmployees,
} from "./services/dataService"

import {
  LayoutDashboard,
  Users,
  Building2,
  BarChart3,
  Settings,
  Upload,
  Bell,
  Search,
  Check,
  Wallet,
  ChevronDown,
  LogOut,
} from "lucide-react"

import { readExcelFile } from "./services/excelService"
import { AUTH_STORAGE_KEY } from "./config/authConfig"

/* =========================================================
   Default Departments
========================================================= */

const DEFAULT_DEPARTMENTS = [
  "پشتیبانی",
  "مالی",
  "منابع انسانی",
  "فناوری اطلاعات",
  "بازاریابی",
  "فروش",
  "اداری",
  "تولید",
]

/* =========================================================
   Chart Font
========================================================= */

const CHART_FONT =
  "Vazirmatn, Tahoma, Arial, sans-serif"

const CHART_LABEL_STYLE = {
  fontFamily: CHART_FONT,
  fontSize: "12px",
  fontWeight: 500,
}

const CHART_DATA_LABEL_STYLE = {
  fontFamily: CHART_FONT,
  fontSize: "12px",
  fontWeight: 600,
}

/* =========================================================
   Text Normalization
   برای جلوگیری از مشکل ي/ی و ك/ک و فاصله‌ها
========================================================= */

function normalizeText(value) {
  return String(value ?? "")
    .replace(/ي/g, "ی")
    .replace(/ى/g, "ی")
    .replace(/ك/g, "ک")
    .replace(/\u200c/g, "")
    .replace(/\s+/g, " ")
    .trim()
}

/* =========================================================
   Create Department List
   واحدهای موجود در Excel هم به لیست اضافه می‌شوند.
========================================================= */

function buildDepartments(employees) {
  const result = [...DEFAULT_DEPARTMENTS]

  for (const employee of employees) {
    const department = String(
      employee?.department ?? ""
    ).trim()

    if (!department) {
      continue
    }

    const exists = result.some(
      (item) =>
        normalizeText(item) ===
        normalizeText(department)
    )

    if (!exists) {
      result.push(department)
    }
  }

  return result
}

/* =========================================================
   App
========================================================= */

function App() {
  const fileInputRef = useRef(null)

  /* =======================================================
     Authentication
  ======================================================= */

  const [isAuthenticated, setIsAuthenticated] = useState(
    () =>
      typeof window !== "undefined" &&
      window.localStorage.getItem(AUTH_STORAGE_KEY) === "true"
  )

  function handleLoginSuccess() {
    window.localStorage.setItem(AUTH_STORAGE_KEY, "true")
    setIsAuthenticated(true)
  }

  function handleLogout() {
    window.localStorage.removeItem(AUTH_STORAGE_KEY)
    setIsAuthenticated(false)
    setCurrentPage("dashboard")
  }

  /* =======================================================
     Employees
  ======================================================= */

  const [employees, setEmployees] = useState([])

  const [fileName, setFileName] = useState("")

  const [searchText, setSearchText] = useState("")

  const [selectedCompany, setSelectedCompany] =
    useState("شرکت نمونه")

  const [selectedDepartments, setSelectedDepartments] =
    useState(DEFAULT_DEPARTMENTS)

  const [loadingData, setLoadingData] = useState(true)

  const [isJobTitleChartOpen, setIsJobTitleChartOpen] =
    useState(true)

  /* =======================================================
     Current Page
  ======================================================= */

  const [currentPage, setCurrentPage] =
    useState("dashboard")

  /* =======================================================
     Dynamic Departments
  ======================================================= */

  const departments = useMemo(() => {
    return buildDepartments(employees)
  }, [employees])

  /* =======================================================
     Load Saved Data
  ======================================================= */

  useEffect(() => {
    let mounted = true

    async function loadSavedData() {
      try {
        const savedEmployees =
          await getEmployees()

        if (!mounted) {
          return
        }

        const safeEmployees =
          Array.isArray(savedEmployees)
            ? savedEmployees
            : []

        setEmployees(safeEmployees)

        setSelectedDepartments(
          buildDepartments(
            safeEmployees
          )
        )
      } catch (error) {
        console.error(
          "خطا در خواندن اطلاعات ذخیره شده:",
          error
        )
      } finally {
        if (mounted) {
          setLoadingData(false)
        }
      }
    }

    loadSavedData()

    return () => {
      mounted = false
    }
  }, [])

  /* =======================================================
     Excel Upload
  ======================================================= */

  const handleFileChange = async (event) => {
    const file =
      event.target.files?.[0]

    if (!file) {
      return
    }

    try {
      const data =
        await readExcelFile(file)

      const safeData =
        Array.isArray(data)
          ? data
          : []

      setEmployees(safeData)

      setFileName(
        file.name
      )

      setSearchText("")

      setSelectedDepartments(
        buildDepartments(
          safeData
        )
      )

      await saveEmployees(
        safeData
      )
    } catch (error) {
      alert(
        error.message ||
          "خطا در خواندن فایل Excel"
      )
    }

    event.target.value = ""
  }

  /* =======================================================
     Upload Button
  ======================================================= */

  const handleUploadClick = () => {
    fileInputRef.current?.click()
  }

  /* =======================================================
     Employees Changed From HR
  ======================================================= */

  const handleEmployeesChange =
    async (newEmployees) => {
      const safeEmployees =
        Array.isArray(
          newEmployees
        )
          ? newEmployees
          : []

      setEmployees(
        safeEmployees
      )

      setSelectedDepartments(
        (current) => {
          const allDepartments =
            buildDepartments(
              safeEmployees
            )

          const hadAll =
            current.length >=
            departments.length

          if (hadAll) {
            return allDepartments
          }

          return current.filter(
            (item) =>
              allDepartments.some(
                (department) =>
                  normalizeText(
                    department
                  ) ===
                  normalizeText(
                    item
                  )
              )
          )
        }
      )

      try {
        await saveEmployees(
          safeEmployees
        )
      } catch (error) {
        console.error(
          "خطا در ذخیره تغییرات کارکنان:",
          error
        )
      }
    }

  /* =======================================================
     Department Filters
  ======================================================= */

  const toggleDepartment = (
    department
  ) => {
    setSelectedDepartments(
      (current) => {
        const exists =
          current.some(
            (item) =>
              normalizeText(
                item
              ) ===
              normalizeText(
                department
              )
          )

        if (exists) {
          return current.filter(
            (item) =>
              normalizeText(
                item
              ) !==
              normalizeText(
                department
              )
          )
        }

        return [
          ...current,
          department,
        ]
      }
    )
  }

  /* =======================================================
     Select All Departments
  ======================================================= */

  const selectAllDepartments = () => {
    setSelectedDepartments(
      departments
    )
  }

  /* =======================================================
     Clear Departments
  ======================================================= */

  const clearDepartments = () => {
    setSelectedDepartments([])
  }

  /* =======================================================
     Filtered Employees
  ======================================================= */

  const filteredEmployees =
    useMemo(() => {
      const text =
        normalizeText(
          searchText
        ).toLowerCase()

      return employees.filter(
        (employee) => {
          const department =
            String(
              employee?.department ||
                ""
            ).trim()

          const departmentMatch =
            selectedDepartments.length >
              0 &&
            selectedDepartments.some(
              (item) =>
                normalizeText(
                  item
                ) ===
                normalizeText(
                  department
                )
            )

          if (!departmentMatch) {
            return false
          }

          if (!text) {
            return true
          }

          return [
            employee?.name,
            employee?.id,
            employee?.job_title,
            employee?.department,
            employee?.sub_department,
          ].some(
            (value) =>
              normalizeText(
                value
              )
                .toLowerCase()
                .includes(text)
          )
        }
      )
    }, [
      employees,
      searchText,
      selectedDepartments,
    ])

  /* =======================================================
     Total Employees
     
     تعداد کل رکوردهای Excel
======================================================= */

  const totalEmployees =
    employees.length

  /* =======================================================
     Dashboard Employees
     
     نمودارها و جستجو از این لیست استفاده می‌کنند.
======================================================= */

  const dashboardEmployees =
    filteredEmployees

  /* =======================================================
     Total Salary
======================================================= */

  const totalSalary =
    employees.reduce(
      (total, employee) => {
        return (
          total +
          Number(
            employee?.salary || 0
          )
        )
      },
      0
    )

  /* =======================================================
     Average Age
======================================================= */

  const averageAge =
    employees.length > 0
      ? employees.reduce(
          (total, employee) => {
            return (
              total +
              Number(
                employee?.age || 0
              )
            )
          },
          0
        ) / employees.length
      : 0

  /* =======================================================
     Gender
======================================================= */

  const maleCount =
    dashboardEmployees.filter(
      (employee) =>
        normalizeText(
          employee?.gender
        ) === "مرد"
    ).length

  const femaleCount =
    dashboardEmployees.filter(
      (employee) =>
        normalizeText(
          employee?.gender
        ) === "زن"
    ).length

  const genderChartOptions = {
    chart: {
      type: "donut",
      background: "transparent",
      fontFamily: CHART_FONT,
    },

    labels: [
      "مرد",
      "زن",
    ],

    colors: [
      "#f0c040",
      "#fff0cf",
    ],

    theme: {
      mode: "dark",
    },

    legend: {
      position: "bottom",

      labels: {
        colors: "#ffffff",
      },

      fontFamily: CHART_FONT,
      fontSize: "12px",
    },

    dataLabels: {
      enabled: true,

      style: {
        ...CHART_DATA_LABEL_STYLE,
        colors: [
          "#ffffff",
        ],
      },
    },

    stroke: {
      colors: ["#202020"],
      width: 2,
    },

    plotOptions: {
      pie: {
        donut: {
          size: "58%",
        },
      },
    },

    tooltip: {
      theme: "dark",
      style: {
        fontFamily: CHART_FONT,
      },
    },
  }

  const genderChartSeries = [
    maleCount,
    femaleCount,
  ]

  /* =======================================================
     Experience
======================================================= */

  const newEmployees =
    dashboardEmployees.filter(
      (employee) =>
        Number(
          employee?.experience_years ||
            0
        ) < 2
    ).length

  const normalEmployees =
    dashboardEmployees.filter(
      (employee) => {
        const experience =
          Number(
            employee?.experience_years ||
              0
          )

        return (
          experience >= 2 &&
          experience <= 5
        )
      }
    ).length

  const experiencedEmployees =
    dashboardEmployees.filter(
      (employee) => {
        const experience =
          Number(
            employee?.experience_years ||
              0
          )

        return (
          experience >= 6 &&
          experience <= 10
        )
      }
    ).length

  const seniorEmployees =
    dashboardEmployees.filter(
      (employee) =>
        Number(
          employee?.experience_years ||
            0
        ) > 10
    ).length

  const experienceChartOptions = {
    chart: {
      type: "donut",
      background: "transparent",
      fontFamily: CHART_FONT,
    },

    labels: [
      "تازه‌کار",
      "معمولی",
      "باتجربه",
      "ارشد",
    ],

    colors: [
      "#a9822e",
      "#c99108",
      "#f4cf72",
      "#fff3d2",
    ],

    theme: {
      mode: "dark",
    },

    legend: {
      position: "bottom",

      labels: {
        colors: "#ffffff",
      },

      fontFamily: CHART_FONT,
      fontSize: "12px",
    },

    dataLabels: {
      enabled: true,

      style: {
        ...CHART_DATA_LABEL_STYLE,
        colors: [
          "#ffffff",
        ],
      },
    },

    stroke: {
      colors: ["#202020"],
      width: 2,
    },

    plotOptions: {
      pie: {
        donut: {
          size: "55%",
        },
      },
    },

    tooltip: {
      theme: "dark",
      style: {
        fontFamily: CHART_FONT,
      },
    },
  }

  const experienceChartSeries = [
    newEmployees,
    normalEmployees,
    experiencedEmployees,
    seniorEmployees,
  ]

  /* =======================================================
     Contract
======================================================= */

  const officialCount =
    dashboardEmployees.filter(
      (employee) =>
        normalizeText(
          employee?.contract_type
        ) === "رسمی"
    ).length

  const probationaryCount =
    dashboardEmployees.filter(
      (employee) =>
        normalizeText(
          employee?.contract_type
        ) === "آزمایشی"
    ).length

  const contractorCount =
    dashboardEmployees.filter(
      (employee) =>
        normalizeText(
          employee?.contract_type
        ) === "پیمانکاری"
    ).length

  const contractualCount =
    dashboardEmployees.filter(
      (employee) =>
        normalizeText(
          employee?.contract_type
        ) === "قراردادی"
    ).length

  const contractChartOptions = {
    chart: {
      type: "bar",
      background: "transparent",

      toolbar: {
        show: false,
      },

      fontFamily: CHART_FONT,
    },

    plotOptions: {
      bar: {
        borderRadius: 4,
        columnWidth: "48%",
        distributed: true,
      },
    },

    xaxis: {
      categories: [
        "رسمی",
        "آزمایشی",
        "پیمانکاری",
        "قراردادی",
      ],

      labels: {
        style: {
          ...CHART_LABEL_STYLE,
          colors: [
            "#ffffff",
            "#ffffff",
            "#ffffff",
            "#ffffff",
          ],
        },
      },

      axisBorder: {
        color: "#ffffff44",
      },

      axisTicks: {
        color: "#ffffff44",
      },
    },

    yaxis: {
      labels: {
        style: {
          ...CHART_LABEL_STYLE,
          colors: "#ffffff",
        },
      },
    },

    theme: {
      mode: "dark",
    },

    colors: [
      "#c99108",
      "#f0c040",
      "#d4a017",
      "#e8c66a",
    ],

    dataLabels: {
      enabled: true,

      style: {
        ...CHART_DATA_LABEL_STYLE,
        colors: [
          "#ffffff",
        ],
      },
    },

    tooltip: {
      theme: "dark",

      style: {
        fontFamily: CHART_FONT,
      },
    },

    grid: {
      borderColor:
        "#ffffff10",
    },
  }

  const contractChartSeries = [
    {
      name: "تعداد کارکنان",

      data: [
        officialCount,
        probationaryCount,
        contractorCount,
        contractualCount,
      ],
    },
  ]

  /* =======================================================
     Marital
======================================================= */

  const singleCount =
    dashboardEmployees.filter(
      (employee) =>
        normalizeText(
          employee?.marital_status
        ) === "مجرد"
    ).length

  const marriedCount =
    dashboardEmployees.filter(
      (employee) =>
        normalizeText(
          employee?.marital_status
        ) === "متأهل"
    ).length

  const maritalChartOptions = {
    chart: {
      type: "donut",
      background: "transparent",
      fontFamily: CHART_FONT,
    },

    labels: [
      "متأهل",
      "مجرد",
    ],

    colors: [
      "#fff0cf",
      "#f0c040",
    ],

    theme: {
      mode: "dark",
    },

    legend: {
      show: false,
    },

    dataLabels: {
      enabled: true,

      style: {
        ...CHART_DATA_LABEL_STYLE,
        colors: [
          "#ffffff",
        ],
      },
    },

    stroke: {
      colors: ["#202020"],
      width: 2,
    },

    plotOptions: {
      pie: {
        donut: {
          size: "58%",
        },
      },
    },

    tooltip: {
      theme: "dark",

      style: {
        fontFamily: CHART_FONT,
      },
    },
  }

  const maritalChartSeries = [
    marriedCount,
    singleCount,
  ]

  /* =======================================================
     Insurance
======================================================= */

  const insuredCount =
    dashboardEmployees.filter(
      (employee) =>
        normalizeText(
          employee?.has_insurance
        ) === "بله"
    ).length

  const uninsuredCount =
    dashboardEmployees.filter(
      (employee) =>
        normalizeText(
          employee?.has_insurance
        ) === "خیر"
    ).length

  const insuranceChartOptions = {
    chart: {
      type: "donut",
      background: "transparent",
      fontFamily: CHART_FONT,
    },

    labels: [
      "بیمه دارد",
      "بیمه ندارد",
    ],

    colors: [
      "#fff0cf",
      "#f0c040",
    ],

    theme: {
      mode: "dark",
    },

    legend: {
      show: false,
    },

    dataLabels: {
      enabled: true,

      style: {
        ...CHART_DATA_LABEL_STYLE,
        colors: [
          "#ffffff",
        ],
      },
    },

    stroke: {
      colors: ["#202020"],
      width: 2,
    },

    plotOptions: {
      pie: {
        donut: {
          size: "58%",
        },
      },
    },

    tooltip: {
      theme: "dark",

      style: {
        fontFamily: CHART_FONT,
      },
    },
  }

  const insuranceChartSeries = [
    insuredCount,
    uninsuredCount,
  ]

  /* =======================================================
     Education
     
     اصلاح مهم:
     چون نمودار horizontal است،
     categories باید داخل yaxis باشد.
======================================================= */

  const degreeStats =
    useMemo(() => {
      const counts = {}

      for (const employee of dashboardEmployees) {
        let degree = normalizeText(
          employee?.degree
        )

        if (!degree) {
          degree = "نامشخص"
        }

        if (!counts[degree]) {
          counts[degree] = 0
        }

        counts[degree] += 1
      }

      return Object.entries(counts).sort(
        (a, b) => b[1] - a[1]
      )
    }, [dashboardEmployees])

  const degreeNames = degreeStats.map(
    ([degree]) => degree
  )

  const degreeCounts = degreeStats.map(
    ([, value]) => value
  )

  const educationChartOptions =
    useMemo(() => {
      return {
        chart: {
          type: "bar",
          background: "transparent",

          toolbar: {
            show: false,
          },

          fontFamily: CHART_FONT,
        },

        plotOptions: {
          bar: {
            horizontal: true,
            borderRadius: 4,
            barHeight: "58%",
            distributed: true,
          },
        },

        xaxis: {
          categories: degreeNames,

          labels: {
            style: {
              ...CHART_LABEL_STYLE,
              colors: "#ffffff",
            },
          },

          axisBorder: {
            color: "#ffffff44",
          },

          axisTicks: {
            color: "#ffffff44",
          },
        },

        yaxis: {
          labels: {
            style: {
              ...CHART_LABEL_STYLE,
              colors: "#ffffff",
            },
          },
        },

        theme: {
          mode: "dark",
        },

        colors: [
          "#c99108",
          "#d4a017",
          "#e0b84e",
          "#f0c040",
          "#f4d987",
          "#b88724",
          "#e8ca75",
          "#d9ae42",
        ],

        dataLabels: {
          enabled: true,

          formatter: (value) => `${value} نفر`,

          style: {
            ...CHART_DATA_LABEL_STYLE,
            colors: [
              "#ffffff",
            ],
          },
        },

        tooltip: {
          theme: "dark",

          x: {
            show: true,
          },

          y: {
            formatter: (value) => `${value} نفر`,
          },

          style: {
            fontFamily: CHART_FONT,
          },
        },

        grid: {
          borderColor: "#ffffff10",
        },
      }
    }, [degreeNames])

  const educationChartSeries = [
    {
      name: "تعداد کارکنان",
      data: degreeCounts,
    },
  ]

  /* =======================================================
     Age
======================================================= */

  const ageGroups = [
    {
      title: "18-20",
      min: 18,
      max: 20,
    },
    {
      title: "21-23",
      min: 21,
      max: 23,
    },
    {
      title: "24-26",
      min: 24,
      max: 26,
    },
    {
      title: "27-29",
      min: 27,
      max: 29,
    },
    {
      title: "30-32",
      min: 30,
      max: 32,
    },
    {
      title: "33-35",
      min: 33,
      max: 35,
    },
    {
      title: "36-38",
      min: 36,
      max: 38,
    },
    {
      title: "39-41",
      min: 39,
      max: 41,
    },
    {
      title: "42-44",
      min: 42,
      max: 44,
    },
    {
      title: "45-47",
      min: 45,
      max: 47,
    },
    {
      title: "48-50",
      min: 48,
      max: 50,
    },
    {
      title: "51+",
      min: 51,
      max: Infinity,
    },
  ]

  const ageCounts =
    ageGroups.map(
      (group) => {
        return dashboardEmployees.filter(
          (employee) => {
            const age =
              Number(
                employee?.age || 0
              )

            return (
              age >= group.min &&
              age <= group.max
            )
          }
        ).length
      }
    )

  const ageChartOptions = {
    chart: {
      type: "treemap",
      background: "transparent",

      toolbar: {
        show: false,
      },

      fontFamily: CHART_FONT,
    },

    legend: {
      show: false,
    },

    plotOptions: {
      treemap: {
        distributed: true,
        enableShades: false,
      },
    },

    colors: [
      "#b9913d",
      "#c9a24c",
      "#d8b76b",
      "#e6ca91",
      "#f0d8aa",
      "#bd984b",
      "#d0ad64",
      "#dfc68e",
      "#eeddbb",
      "#a98236",
      "#d7b977",
      "#f4e6c5",
    ],

    dataLabels: {
      enabled: true,

      style: {
        fontFamily: CHART_FONT,
        fontSize: "11px",
        fontWeight: 700,
        colors: [
          "#1a1a1a",
        ],
      },

      formatter: (
        text,
        opts
      ) => {
        const data =
          opts.w.config
            .series[0]
            .data[
              opts.dataPointIndex
            ]

        return `${text}\n${data.y}`
      },
    },

    tooltip: {
      theme: "dark",

      style: {
        fontFamily: CHART_FONT,
      },
    },
  }

  const ageChartSeries = [
    {
      data: ageGroups.map(
        (group, index) => ({
          x: group.title,
          y: ageCounts[index],
        })
      ),
    },
  ]

  /* =======================================================
     JOB TITLE
     
     اصلاح مهم:
     1. همه عنوان‌های شغلی نمایش داده می‌شوند.
     2. از dashboardEmployees استفاده می‌شود
        تا فیلتر واحد و جستجو اعمال شود.
======================================================= */

  const jobTitleStats =
    useMemo(() => {
      const counts = {}

      for (const employee of dashboardEmployees) {
        let title =
          normalizeText(
            employee?.job_title
          )

        if (!title) {
          title = "نامشخص"
        }

        if (!counts[title]) {
          counts[title] = 0
        }

        counts[title] += 1
      }

      return Object.entries(
        counts
      ).sort(
        (a, b) =>
          b[1] - a[1]
      )
    }, [dashboardEmployees])

  /* =======================================================
     All Job Titles
======================================================= */

  const jobTitles =
    jobTitleStats.map(
      ([title]) => title
    )

  /* =======================================================
     Job Title Values
======================================================= */

  const jobTitleValues =
    jobTitleStats.map(
      ([, value]) => value
    )

  /* =======================================================
     Job Title Display Rules
  ======================================================= */

  const JOB_TITLE_VISIBLE_LIMIT = 5

  const jobTitleNeedsCollapse =
    jobTitles.length > JOB_TITLE_VISIBLE_LIMIT

  const jobTitleChartHeight = Math.max(
    300,
    jobTitles.length * 48
  )

  /* =======================================================
     Job Title Chart
     
     اصلاح مهم:
     چون horizontal=true است،
     categories باید داخل yaxis باشد.
======================================================= */

  const jobTitleChartOptions =
    useMemo(() => {
      const dynamicHeight =
        Math.max(
          310,
          jobTitles.length * 48
        )

      return {
        chart: {
          type: "bar",
          background:
            "transparent",

          toolbar: {
            show: false,
          },

          fontFamily: CHART_FONT,
        },

        plotOptions: {
          bar: {
            horizontal: true,
            borderRadius: 4,
            barHeight: "58%",
            distributed: true,
          },
        },

        xaxis: {
          categories: jobTitles,

          labels: {
            style: {
              ...CHART_LABEL_STYLE,
              colors: "#ffffff",
            },
          },

          axisBorder: {
            color: "#ffffff44",
          },

          axisTicks: {
            color: "#ffffff44",
          },
        },

        yaxis: {
          labels: {
            style: {
              ...CHART_LABEL_STYLE,
              colors: "#ffffff",
            },
          },
        },

        theme: {
          mode: "dark",
        },

        legend: {
          show: false,
        },

        colors: [
          "#f0c040",
          "#e4bd54",
          "#d4a017",
          "#c99108",
          "#e8ca75",
          "#b88724",
          "#f4d987",
          "#c6a23a",
          "#e7c765",
          "#ad821e",
          "#f2d68e",
          "#d9ae42",
        ],

        dataLabels: {
          enabled: true,

          formatter: (value) => `${value} نفر`,

          style: {
            ...CHART_DATA_LABEL_STYLE,
            colors: [
              "#ffffff",
            ],
          },
        },

        tooltip: {
          theme: "dark",

          x: {
            show: true,
          },

          y: {
            formatter: (value) => `${value} نفر`,
          },

          style: {
            fontFamily: CHART_FONT,
          },
        },

        grid: {
          borderColor:
            "#ffffff10",
        },

        _dynamicHeight:
          dynamicHeight,
      }
    }, [jobTitles])

  const jobTitleChartSeries = [
    {
      name: "تعداد کارکنان",
      data:
        jobTitleValues,
    },
  ]

  /* =======================================================
     Login Page
     تا زمانی که کاربر وارد نشده، هیچ صفحه‌ای از داشبورد
     (نه HR، نه Finance، نه Dashboard) نمایش داده نمی‌شود.
======================================================= */

  if (!isAuthenticated) {
    return <Login onLoginSuccess={handleLoginSuccess} />
  }

  /* =======================================================
     HR Page
======================================================= */

  if (
    currentPage === "hr"
  ) {
    return (
      <HumanResources
        employees={employees}
        onEmployeesChange={
          handleEmployeesChange
        }
        onBack={() =>
          setCurrentPage(
            "dashboard"
          )
        }
      />
    )
  }

  /* =======================================================
     Finance Page
======================================================= */

  if (
    currentPage === "finance"
  ) {
    return (
      <FinancialAffairs
        onBack={() =>
          setCurrentPage(
            "dashboard"
          )
        }
      />
    )
  }

  /* =======================================================
     Dashboard
======================================================= */

  return (
    <div
      dir="rtl"
      className="min-h-screen bg-[#1a1a1a] text-white"
    >
      {/* =====================================================
          Header
      ===================================================== */}

      <header className="flex h-20 items-center justify-between border-b border-[#d4a017]/40 bg-black px-4 md:px-6">
        <div className="flex items-center gap-3">
          <div className="hidden h-12 w-12 items-center justify-center rounded-xl bg-[#d4a017] text-black sm:flex">
            <LayoutDashboard
              size={26}
            />
          </div>

          <div>
            <h1 className="text-xl font-black text-white md:text-3xl">
              داشبورد منابع انسانی
            </h1>

            <p className="mt-1 text-xs text-[#d4a017] md:text-sm">
              سیستم تحلیل و مدیریت منابع انسانی
            </p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <button className="rounded-lg p-2 transition hover:bg-white/5">
            <Bell size={20} />
          </button>

          <div className="hidden h-8 w-px bg-white/10 md:block" />

          <div className="hidden text-left md:block">
            <p className="text-sm font-medium">
              مدیر سیستم
            </p>

            <p className="text-xs text-gray-500">
              Administrator
            </p>
          </div>

          <div className="flex h-9 w-9 items-center justify-center rounded-full bg-[#d4a017] font-bold text-black">
            م
          </div>

          <button
            type="button"
            onClick={handleLogout}
            title="خروج"
            className="flex items-center gap-2 rounded-lg border border-white/10 px-3 py-2 text-xs font-bold text-gray-300 transition hover:border-red-500/40 hover:bg-red-500/10 hover:text-red-400"
          >
            <LogOut size={16} />
            <span className="hidden sm:inline">خروج</span>
          </button>
        </div>
      </header>

      {/* =====================================================
          Layout
      ===================================================== */}

      <div className="flex min-h-[calc(100vh-80px)]">
        {/* ===================================================
            Sidebar
        =================================================== */}

        <aside className="hidden w-64 shrink-0 overflow-y-auto border-l border-[#d4a017]/30 bg-[#151515] p-4 lg:block">
          {/* Company */}

          <div className="mb-5 rounded-xl border border-[#d4a017]/40 bg-[#1c1c1c] p-4">
            <p className="mb-2 text-xs text-gray-500">
              شرکت
            </p>

            <select
              value={
                selectedCompany
              }
              onChange={(event) =>
                setSelectedCompany(
                  event.target.value
                )
              }
              className="w-full bg-transparent text-sm text-white outline-none"
            >
              <option className="bg-[#1a1a1a]">
                شرکت نمونه
              </option>

              <option className="bg-[#1a1a1a]">
                شرکت دوم
              </option>
            </select>
          </div>

          {/* Navigation */}

          <nav className="mb-6 space-y-2">
            <SidebarItem
              icon={
                <LayoutDashboard
                  size={19}
                />
              }
              title="داشبورد"
              active={
                currentPage ===
                "dashboard"
              }
              onClick={() =>
                setCurrentPage(
                  "dashboard"
                )
              }
            />

            <SidebarItem
              icon={
                <Users size={19} />
              }
              title="کارکنان"
              active={
                currentPage ===
                "hr"
              }
              onClick={() =>
                setCurrentPage("hr")
              }
            />

            <SidebarItem
              icon={
                <Building2
                  size={19}
                />
              }
              title="واحدها و دپارتمان‌ها"
              onClick={() =>
                setCurrentPage("hr")
              }
            />

            <SidebarItem
              icon={
                <Wallet size={19} />
              }
              title="امور مالی"
              active={
                currentPage ===
                "finance"
              }
              onClick={() =>
                setCurrentPage(
                  "finance"
                )
              }
            />

            <SidebarItem
              icon={
                <BarChart3
                  size={19}
                />
              }
              title="گزارش‌ها"
            />

            <SidebarItem
              icon={
                <Settings
                  size={19}
                />
              }
              title="تنظیمات"
            />
          </nav>

          {/* Departments */}

          <div className="rounded-xl border border-[#d4a017]/30 bg-[#1c1c1c] p-3">
            <div className="mb-3 flex items-center justify-between">
              <p className="text-sm font-bold text-white">
                واحدها
              </p>

              <div className="flex gap-1">
                <button
                  onClick={
                    selectAllDepartments
                  }
                  className="text-[10px] text-[#f0c040] hover:text-white"
                >
                  همه
                </button>

                <span className="text-gray-700">
                  /
                </span>

                <button
                  onClick={
                    clearDepartments
                  }
                  className="text-[10px] text-gray-400 hover:text-white"
                >
                  هیچ‌کدام
                </button>
              </div>
            </div>

            <div className="space-y-2">
              {departments.map(
                (department) => {
                  const checked =
                    selectedDepartments.some(
                      (item) =>
                        normalizeText(
                          item
                        ) ===
                        normalizeText(
                          department
                        )
                    )

                  return (
                    <button
                      key={
                        department
                      }
                      onClick={() =>
                        toggleDepartment(
                          department
                        )
                      }
                      className="flex w-full items-center justify-between gap-2 rounded-lg border border-white/5 bg-[#202020] px-3 py-2.5 transition hover:border-[#d4a017]/50"
                    >
                      <span className="text-sm text-gray-200">
                        {
                          department
                        }
                      </span>

                      <span
                        className={`flex h-5 w-5 items-center justify-center rounded-md border ${
                          checked
                            ? "border-[#d4a017] bg-[#d4a017] text-black"
                            : "border-white/20 text-transparent"
                        }`}
                      >
                        <Check
                          size={14}
                          strokeWidth={
                            3
                          }
                        />
                      </span>
                    </button>
                  )
                }
              )}
            </div>
          </div>

          {/* Upload */}

          <div className="mt-5">
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={
                handleFileChange
              }
              className="hidden"
            />

            <button
              onClick={
                handleUploadClick
              }
              className="flex w-full items-center justify-center gap-2 rounded-xl bg-[#d4a017] px-4 py-3 text-sm font-bold text-black transition hover:bg-[#f0c040]"
            >
              <Upload size={18} />
              آپلود فایل Excel
            </button>
          </div>
        </aside>

        {/* ===================================================
            Main
        =================================================== */}

        <main className="min-w-0 flex-1 overflow-auto p-3 md:p-5 lg:p-6">
          {/* Mobile */}

          <div className="mb-4 flex flex-col gap-2 lg:hidden">
            <input
              id="mobile-excel"
              type="file"
              accept=".xlsx,.xls"
              onChange={
                handleFileChange
              }
              className="hidden"
            />

            <label
              htmlFor="mobile-excel"
              className="flex items-center justify-center gap-2 rounded-xl bg-[#d4a017] px-4 py-3 text-sm font-bold text-black"
            >
              <Upload size={18} />
              آپلود فایل Excel
            </label>

            <button
              onClick={() =>
                setCurrentPage(
                  "hr"
                )
              }
              className="flex items-center justify-center gap-2 rounded-xl border border-[#d4a017]/50 bg-[#202020] px-4 py-3 text-sm font-bold text-[#f0c040]"
            >
              <Users size={18} />
              مدیریت منابع انسانی
            </button>

            <button
              onClick={() =>
                setCurrentPage(
                  "finance"
                )
              }
              className="flex items-center justify-center gap-2 rounded-xl border border-[#d4a017]/50 bg-[#202020] px-4 py-3 text-sm font-bold text-[#f0c040]"
            >
              <Wallet size={18} />
              امور مالی
            </button>
          </div>

          {/* Title */}

          <div className="mb-5 flex flex-col justify-between gap-4 md:flex-row md:items-center">
            <div>
              <h2 className="text-xl font-bold md:text-2xl">
                داشبورد منابع انسانی
              </h2>

              <p className="mt-1 text-sm text-gray-500">
                نمای کلی اطلاعات کارکنان
              </p>
            </div>

            <div className="relative">
              <Search
                size={18}
                className="absolute right-3 top-3 text-gray-500"
              />

              <input
                type="text"
                value={
                  searchText
                }
                onChange={(
                  event
                ) =>
                  setSearchText(
                    event.target.value
                  )
                }
                placeholder="جستجوی کارکنان..."
                className="w-full rounded-xl border border-white/10 bg-[#202020] py-2.5 pl-4 pr-10 text-sm outline-none focus:border-[#d4a017] md:w-64"
              />
            </div>
          </div>

          {/* Loading */}

          {loadingData && (
            <div className="mb-5 rounded-xl border border-[#d4a017]/30 bg-[#202020] p-4 text-sm text-gray-400">
              در حال خواندن اطلاعات ذخیره‌شده...
            </div>
          )}

          {/* File */}

          {fileName && (
            <div className="mb-5 rounded-xl border border-[#d4a017]/30 bg-[#202020] p-4">
              <div className="flex flex-col justify-between gap-2 sm:flex-row sm:items-center">
                <div>
                  <p className="text-xs text-gray-500">
                    فایل فعلی
                  </p>

                  <p className="mt-1 font-bold text-[#f0c040]">
                    {
                      fileName
                    }
                  </p>
                </div>

                <p className="text-sm text-gray-400">
                  نمایش{" "}
                  {
                    dashboardEmployees.length
                  }{" "}
                  از{" "}
                  {
                    employees.length
                  }{" "}
                  رکورد
                </p>
              </div>
            </div>
          )}

          {/* =================================================
              KPI
          ================================================= */}

          <div className="grid grid-cols-2 gap-3 md:grid-cols-3 xl:grid-cols-5">
            <KpiCard
              title="تعداد نفرات"
              value={
                totalEmployees
              }
            />

            <KpiCard
              title="مجموع حقوق"
              value={formatNumber(
                totalSalary
              )}
            />

            <KpiCard
              title="میانگین سنی"
              value={averageAge.toFixed(
                0
              )}
            />

            <KpiCard
              title="نسبت به کل"
              value={
                employees.length >
                0
                  ? "100%"
                  : "0%"
              }
            />

            <KpiCard
              title="تعداد فعال"
              value={
                totalEmployees
              }
            />
          </div>

          {/* =================================================
              Gender
          ================================================= */}

          <div className="mt-4 grid grid-cols-1 gap-3 xl:grid-cols-12">
            <div className="xl:col-span-4">
              <DashboardCard title="تفکیک کارکنان بر اساس جنسیت">
                {
                  dashboardEmployees.length >
                  0 ? (
                    <Chart
                      options={
                        genderChartOptions
                      }
                      series={
                        genderChartSeries
                      }
                      type="donut"
                      height={
                        245
                      }
                    />
                  ) : (
                    <EmptyChart />
                  )
                }

                <div className="mt-1 flex justify-center gap-10">
                  <SmallStat
                    label="مرد"
                    value={
                      maleCount
                    }
                  />

                  <SmallStat
                    label="زن"
                    value={
                      femaleCount
                    }
                  />
                </div>
              </DashboardCard>
            </div>

            <div className="xl:col-span-8">
              <DashboardCard title="وضعیت داده‌های داشبورد">
                <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
                  <KpiMini
                    title="کل کارکنان"
                    value={
                      totalEmployees
                    }
                  />

                  <KpiMini
                    title="رکوردهای قابل نمایش"
                    value={
                      dashboardEmployees.length
                    }
                  />

                  <KpiMini
                    title="عنوان‌های شغلی"
                    value={
                      jobTitles.length
                    }
                  />

                  <KpiMini
                    title="واحدها"
                    value={
                      departments.length
                    }
                  />
                </div>
              </DashboardCard>
            </div>
          </div>

          {/* =================================================
              Second row
          ================================================= */}

          <div className="mt-3 grid grid-cols-1 gap-3 xl:grid-cols-12">
            <div className="xl:col-span-5">
              <DashboardCard title="نفرات به تفکیک سابقه خدمت">
                {
                  dashboardEmployees.length >
                  0 ? (
                    <Chart
                      options={
                        experienceChartOptions
                      }
                      series={
                        experienceChartSeries
                      }
                      type="donut"
                      height={
                        290
                      }
                    />
                  ) : (
                    <EmptyChart />
                  )
                }
              </DashboardCard>
            </div>

            <div className="xl:col-span-4">
              <DashboardCard title="نفرات به تفکیک وضعیت قرارداد">
                {
                  dashboardEmployees.length >
                  0 ? (
                    <Chart
                      options={
                        contractChartOptions
                      }
                      series={
                        contractChartSeries
                      }
                      type="bar"
                      height={
                        290
                      }
                    />
                  ) : (
                    <EmptyChart />
                  )
                }
              </DashboardCard>
            </div>

            <div className="xl:col-span-3">
              <DashboardCard title="وضعیت تأهل">
                <MiniStats
                  firstLabel="متأهل"
                  firstValue={
                    marriedCount
                  }
                  secondLabel="مجرد"
                  secondValue={
                    singleCount
                  }
                />

                {
                  dashboardEmployees.length >
                  0 ? (
                    <Chart
                      options={
                        maritalChartOptions
                      }
                      series={
                        maritalChartSeries
                      }
                      type="donut"
                      height={
                        210
                      }
                    />
                  ) : (
                    <EmptyChart />
                  )
                }
              </DashboardCard>
            </div>
          </div>

          {/* =================================================
              Third row
              دو نمودار با برچسب‌های متنی طولانی کنار هم
          ================================================= */}

          <div className="mt-3 grid grid-cols-1 gap-3 xl:grid-cols-12">
            <div className="xl:col-span-7">
              <DashboardCard>
                {jobTitleNeedsCollapse ? (
                  <>
                    <button
                      type="button"
                      onClick={() =>
                        setIsJobTitleChartOpen(
                          (previous) => !previous
                        )
                      }
                      className="flex w-full items-center justify-between gap-3 text-right"
                      aria-expanded={isJobTitleChartOpen}
                    >
                      <span className="text-sm font-bold text-white md:text-base">
                        نفرات به تفکیک عنوان شغلی
                        <span className="mr-2 text-xs font-normal text-gray-500">
                          ({jobTitles.length} عنوان)
                        </span>
                      </span>
                      <span
                        className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-[#a97808] bg-[#242424] text-[#f0b323] transition-transform duration-200 ${
                          isJobTitleChartOpen ? "rotate-180" : "rotate-0"
                        }`}
                      >
                        <ChevronDown size={18} />
                      </span>
                    </button>
                    {isJobTitleChartOpen && (
                      <div className="mt-3 max-h-[330px] overflow-y-auto overflow-x-hidden pr-1">
                        {jobTitles.length > 0 ? (
                          <Chart
                            options={jobTitleChartOptions}
                            series={jobTitleChartSeries}
                            type="bar"
                            height={jobTitleChartHeight}
                          />
                        ) : (
                          <EmptyChart />
                        )}
                      </div>
                    )}
                  </>
                ) : (
                  <>
                    <h3 className="mb-2 text-sm font-bold text-white md:text-base">
                      نفرات به تفکیک عنوان شغلی
                    </h3>
                    {jobTitles.length > 0 ? (
                      <div className="overflow-x-hidden">
                        <Chart
                          options={jobTitleChartOptions}
                          series={jobTitleChartSeries}
                          type="bar"
                          height={jobTitleChartHeight}
                        />
                      </div>
                    ) : (
                      <EmptyChart />
                    )}
                  </>
                )}
              </DashboardCard>
            </div>

            <div className="xl:col-span-5">
              <DashboardCard title="نفرات به تفکیک مدرک تحصیلی">
                {
                  degreeNames.length >
                  0 ? (
                    <div className="overflow-x-hidden overflow-y-auto">
                      <Chart
                        options={
                          educationChartOptions
                        }
                        series={
                          educationChartSeries
                        }
                        type="bar"
                        height={300}
                      />
                    </div>
                  ) : (
                    <EmptyChart />
                  )
                }
              </DashboardCard>
            </div>
          </div>

          {/* =================================================
              Fourth row
              بیمه کنار نمودار سن تا فضای صفحه بهتر استفاده شود
          ================================================= */}

          <div className="mt-3 grid grid-cols-1 gap-3 xl:grid-cols-12">
            <div className="xl:col-span-4">
              <DashboardCard title="وضعیت بیمه">
                <MiniStats
                  firstLabel="بیمه دارد"
                  firstValue={
                    insuredCount
                  }
                  secondLabel="بیمه ندارد"
                  secondValue={
                    uninsuredCount
                  }
                />

                {
                  dashboardEmployees.length >
                  0 ? (
                    <Chart
                      options={
                        insuranceChartOptions
                      }
                      series={
                        insuranceChartSeries
                      }
                      type="donut"
                      height={
                        245
                      }
                    />
                  ) : (
                    <EmptyChart />
                  )
                }
              </DashboardCard>
            </div>

            <div className="xl:col-span-8">
              <DashboardCard title="نفرات به تفکیک بازه سنی">
                {
                  dashboardEmployees.length >
                  0 ? (
                    <Chart
                      options={
                        ageChartOptions
                      }
                      series={
                        ageChartSeries
                      }
                      type="treemap"
                      height={
                        360
                      }
                    />
                  ) : (
                    <EmptyChart />
                  )
                }
              </DashboardCard>
            </div>
          </div>

        </main>
      </div>
    </div>
  )
}

/* =========================================================
   Sidebar Item
========================================================= */

function SidebarItem({
  icon,
  title,
  active = false,
  onClick,
}) {
  return (
    <button
      onClick={onClick}
      className={`
        flex w-full items-center gap-3 rounded-xl px-4 py-3
        text-sm transition
        ${
          active
            ? "bg-[#d4a017] font-bold text-black"
            : "text-gray-400 hover:bg-white/5 hover:text-white"
        }
      `}
    >
      {icon}

      <span>
        {title}
      </span>
    </button>
  )
}

/* =========================================================
   KPI Card
========================================================= */

function KpiCard({
  title,
  value,
}) {
  return (
    <div className="min-h-[105px] rounded-xl border border-[#d4a017]/40 bg-[#202020] p-4 md:p-5">
      <p className="text-xs text-gray-400 md:text-sm">
        {title}
      </p>

      <p className="mt-3 truncate text-xl font-bold text-[#f0c040] md:text-2xl">
        {value}
      </p>
    </div>
  )
}

/* =========================================================
   Mini KPI
========================================================= */

function KpiMini({
  title,
  value,
}) {
  return (
    <div className="flex flex-col justify-center rounded-xl border border-[#d4a017]/35 bg-[#202020] p-4">
      <p className="text-xs text-gray-400">
        {title}
      </p>

      <p className="mt-3 truncate text-lg font-black text-[#f0c040] md:text-xl">
        {value}
      </p>
    </div>
  )
}

/* =========================================================
   Dashboard Card
========================================================= */

function DashboardCard({
  title,
  children,
}) {
  return (
    <div className="h-full overflow-hidden rounded-xl border border-[#d4a017]/40 bg-[#202020] p-3 md:p-4">
      <h3 className="mb-2 text-sm font-bold text-white md:text-base">
        {title}
      </h3>

      {children}
    </div>
  )
}

/* =========================================================
   Small Stat
========================================================= */

function SmallStat({
  label,
  value,
}) {
  return (
    <div className="text-center">
      <p className="text-xs text-gray-400">
        {label}
      </p>

      <p className="text-lg font-bold text-[#f0c040]">
        {formatNumber(
          value
        )}
      </p>
    </div>
  )
}

/* =========================================================
   Mini Stats
========================================================= */

function MiniStats({
  firstLabel,
  firstValue,
  secondLabel,
  secondValue,
}) {
  return (
    <div className="mb-1 flex items-center justify-between text-xs md:text-sm">
      <div>
        <span className="text-gray-400">
          {
            firstLabel
          }
        </span>

        <strong className="mr-2 text-[#f0c040]">
          {formatNumber(
            firstValue
          )}
        </strong>
      </div>

      <div>
        <span className="text-gray-400">
          {
            secondLabel
          }
        </span>

        <strong className="mr-2 text-[#f0c040]">
          {formatNumber(
            secondValue
          )}
        </strong>
      </div>
    </div>
  )
}

/* =========================================================
   Empty Chart
========================================================= */

function EmptyChart() {
  return (
    <div className="flex h-56 items-center justify-center rounded-lg border border-white/5 bg-[#191919]">
      <p className="text-sm text-gray-600">
        هنوز داده‌ای وارد نشده است
      </p>
    </div>
  )
}

/* =========================================================
   Number Format
========================================================= */

function formatNumber(number) {
  return new Intl.NumberFormat(
    "fa-IR"
  ).format(
    Number(
      number || 0
    )
  )
}

export default App