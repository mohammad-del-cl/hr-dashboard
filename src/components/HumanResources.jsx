import {
  useEffect,
  useMemo,
  useState,
} from "react"

import {
  ArrowRight,
  Building2,
  Check,
  ChevronDown,
  ChevronLeft,
  FolderPlus,
  Plus,
  Search,
  Trash2,
  UserPlus,
  UserRound,
  Users,
  X,
  Download,
  Pencil,
} from "lucide-react"

import {
  getHRConfig,
  saveHRConfig,
} from "../services/dataService"

import {
  downloadEmployeesExcel,
} from "../services/excelService"

import HRChildTree from "./HRChildTree"

import {
  splitSubPath,
  joinSubPath,
  findNodeById,
  flattenChildren,
  countNodeEmployees,
  employeeInBranch,
  addChildNode,
  removeChildNode,
  mergeEmployeesIntoTreeDeep,
} from "./hrTree"

/* =========================================================
   Default departments
========================================================= */

const DEFAULT_DEPARTMENTS = [
  {
    id: "hr",
    name: "منابع انسانی",
    children: [
      {
        id: "hr-admin",
        name: "اداری",
      },
      {
        id: "hr-cafeteria",
        name: "غذاخوری",
      },
    ],
  },
  {
    id: "finance",
    name: "مالی",
    children: [],
  },
  {
    id: "support",
    name: "پشتیبانی",
    children: [],
  },
  {
    id: "it",
    name: "فناوری اطلاعات",
    children: [],
  },
  {
    id: "marketing",
    name: "بازاریابی",
    children: [],
  },
  {
    id: "sales",
    name: "فروش",
    children: [],
  },
  {
    id: "administration",
    name: "اداری",
    children: [],
  },
  {
    id: "production",
    name: "تولید",
    children: [],
  },
]

/* =========================================================
   Default columns
========================================================= */

const DEFAULT_COLUMNS = [
  {
    key: "marital_status",
    label: "وضعیت تأهل",
  },
  {
    key: "age",
    label: "سن",
  },
  {
    key: "degree",
    label: "مدرک تحصیلی",
  },
  {
    key: "job_title",
    label: "عنوان شغلی",
  },
  {
    key: "contract_type",
    label: "نوع قرارداد",
  },
  {
    key: "has_insurance",
    label: "بیمه",
  },
  {
    key: "salary",
    label: "حقوق",
  },
]

/* =========================================================
   Dashboard fields (سوال‌های ثابت)

   این فیلدها مستقیم در آمار داشبورد استفاده می‌شوند.
   مقدارهای گزینه‌ها باید دقیقاً با شرط‌های App.jsx یکی باشند.
========================================================= */

// کلید snapshot داشبورد (همان مقدار App.jsx و excelService.js)
const DASHBOARD_SNAPSHOT_KEY = "__dashboardData"

const DASHBOARD_FIELDS = [
  {
    key: "gender",
    label: "جنسیت",
    type: "select",
    options: ["مرد", "زن"],
  },
  { key: "age", label: "سن", type: "number" },
  {
    key: "marital_status",
    label: "وضعیت تأهل",
    type: "select",
    options: ["مجرد", "متأهل"],
  },
  { key: "degree", label: "مدرک تحصیلی", type: "text" },
  { key: "job_title", label: "عنوان شغلی", type: "text" },
  { key: "experience_years", label: "سابقه خدمت (سال)", type: "number" },
  {
    key: "contract_type",
    label: "نوع قرارداد",
    type: "select",
    options: ["رسمی", "آزمایشی", "پیمانکاری", "قراردادی"],
  },
  {
    key: "has_insurance",
    label: "بیمه",
    type: "select",
    options: ["بله", "خیر"],
  },
  { key: "salary", label: "حقوق (تومان)", type: "number" },
]

/* =========================================================
   Text helpers
========================================================= */

function normalizeText(value) {
  return String(value ?? "")
    .trim()
    .replace(/\u200c/g, " ")
    .replace(/\u200f/g, "")
    .replace(/\u200e/g, "")
    .replace(/ي/g, "ی")
    .replace(/ى/g, "ی")
    .replace(/ك/g, "ک")
    .replace(/\s+/g, " ")
    .toLowerCase()
}

function createId(prefix, name) {
  const normalized = normalizeText(name)

  const safeName = normalized
    .replace(/[^\p{L}\p{N}]+/gu, "-")
    .replace(/^-+|-+$/g, "")

  return `${prefix}-${safeName || Date.now()}`
}

function getEmployeeId(employee) {
  if (
    employee?.id !== undefined &&
    employee?.id !== null &&
    employee?.id !== ""
  ) {
    return String(employee.id)
  }

  // اگر Excel ستون id نداشت، از چند فیلد ثابت یک شناسه
  // قابل تکرار می‌سازیم تا با فیلتر شدن جدول تغییر نکند.
  const fingerprint = [
    employee?.name,
    employee?.email,
    employee?.phone,
    employee?.age,
    employee?.department,
    employee?.sub_department,
  ]
    .map((value) => normalizeText(value))
    .join("|")

  return `employee-${fingerprint || "unknown"}`
}

/* =========================================================
   Department helpers
========================================================= */

function normalizeChildNodes(children, parentName) {
  if (!Array.isArray(children)) {
    return []
  }

  return children.map((child) => {
    const name =
      typeof child === "string"
        ? child
        : child?.name

    return {
      id:
        child?.id ||
        createId(
          "subdepartment",
          `${parentName}-${name}`
        ),

      name: name || "زیرواحد بدون نام",

      children: normalizeChildNodes(
        child?.children,
        `${parentName}-${name}`
      ),
    }
  })
}

function normalizeDepartmentTree(tree) {
  if (Array.isArray(tree)) {
    return tree.map((department) => ({
      id:
        department.id ||
        createId(
          "department",
          department.name
        ),

      name:
        department.name ||
        "واحد بدون نام",

      children: normalizeChildNodes(
        department.children,
        department.name
      ),
    }))
  }

  if (
    tree &&
    typeof tree === "object"
  ) {
    return Object.entries(tree).map(
      ([departmentName, children]) => ({
        id: createId(
          "department",
          departmentName
        ),

        name: departmentName,

        children: normalizeChildNodes(
          children,
          departmentName
        ),
      })
    )
  }

  return []
}

/* =========================================================
   Merge departments from Excel
========================================================= */

function mergeEmployeesIntoTree(
  currentTree,
  employees
) {
  return mergeEmployeesIntoTreeDeep(
    currentTree,
    employees,
    createId
  )
}

/* =========================================================
   Main component
========================================================= */

function HumanResources({
  employees = [],
  onEmployeesChange,
  onBack,
}) {
  const [departments, setDepartments] =
    useState(DEFAULT_DEPARTMENTS)

  const [columns, setColumns] =
    useState(DEFAULT_COLUMNS)

  // ستون‌هایی که کاربر از جدول منابع انسانی مخفی کرده است.
  // اطلاعات اصلی کارمند حذف نمی‌شود؛ فقط ستون از این صفحه پنهان می‌شود.
  const [hiddenColumns, setHiddenColumns] =
    useState([])

  const [
    selectedDepartmentId,
    setSelectedDepartmentId,
  ] = useState("hr")

  const [
    selectedChildId,
    setSelectedChildId,
  ] = useState(null)

  // در شروع، کل کارکنان نمایش داده می‌شوند تا هیچ شغلی
  // به خاطر فیلتر واحد از جدول پنهان نشود.
  const [
    showAllEmployees,
    setShowAllEmployees,
  ] = useState(true)

  const [searchText, setSearchText] =
    useState("")

  const [
    showAddDepartment,
    setShowAddDepartment,
  ] = useState(false)

  const [
    showAddSubdepartment,
    setShowAddSubdepartment,
  ] = useState(false)

  const [
    showAddColumn,
    setShowAddColumn,
  ] = useState(false)

  const [
    newDepartmentName,
    setNewDepartmentName,
  ] = useState("")

  const [
    newSubdepartmentName,
    setNewSubdepartmentName,
  ] = useState("")

  const [
    newColumnName,
    setNewColumnName,
  ] = useState("")

  const [
    savingConfig,
    setSavingConfig,
  ] = useState(false)

  const [
    loadingConfig,
    setLoadingConfig,
  ] = useState(true)

  const [
    expandedDepartments,
    setExpandedDepartments,
  ] = useState({})

  // افزودن دستی کارمند
  const [
    showAddEmployee,
    setShowAddEmployee,
  ] = useState(false)

  const [
    newEmployeeForm,
    setNewEmployeeForm,
  ] = useState({})

  const [
    newEmployeeDepartmentId,
    setNewEmployeeDepartmentId,
  ] = useState("")

  const [
    newEmployeeSubPath,
    setNewEmployeeSubPath,
  ] = useState("")

  // ویرایش کارمند: اگر مقدار داشته باشد، مودال فرم در حالت ویرایش است
  const [
    editingEmployeeId,
    setEditingEmployeeId,
  ] = useState(null)

  const [
    applyEditToDashboard,
    setApplyEditToDashboard,
  ] = useState(false)

  /* =========================================================
     Load HR config
  ========================================================= */

  useEffect(() => {
    let mounted = true

    async function loadConfig() {
      try {
        const config =
          await getHRConfig()

        if (!mounted) {
          return
        }

        const savedDepartments =
          normalizeDepartmentTree(
            config?.departments
          )

        const mergedDepartments =
          mergeEmployeesIntoTree(
            savedDepartments.length
              ? savedDepartments
              : DEFAULT_DEPARTMENTS,
            employees
          )

        setDepartments(
          mergedDepartments
        )

        const savedHiddenColumns =
          Array.isArray(config?.hiddenColumns)
            ? config.hiddenColumns.map(String)
            : []

        setHiddenColumns(savedHiddenColumns)

        const savedColumns =
          Array.isArray(
            config?.customColumns
          )
            ? config.customColumns
            : []

        const mergedColumns =
          DEFAULT_COLUMNS
            .filter(
              (column) =>
                !savedHiddenColumns.includes(column.key)
            )
            .map((column) => ({
              ...column,
              custom: false,
            }))

        for (const column of savedColumns) {
          if (
            !column?.key ||
            !column?.label
          ) {
            continue
          }

          if (
            mergedColumns.some(
              (item) =>
                item.key === column.key
            )
          ) {
            continue
          }

          mergedColumns.push({
            key: column.key,
            label: column.label,
            custom: true,
          })
        }

        setColumns(
          mergedColumns
        )

        setExpandedDepartments({
          hr: true,
        })
      } catch (error) {
        console.error(
          "خطا در خواندن تنظیمات منابع انسانی:",
          error
        )

        const merged =
          mergeEmployeesIntoTree(
            DEFAULT_DEPARTMENTS,
            employees
          )

        setDepartments(merged)
      } finally {
        if (mounted) {
          setLoadingConfig(false)
        }
      }
    }

    loadConfig()

    return () => {
      mounted = false
    }
  }, [])

  /* =========================================================
     Merge new Excel departments
  ========================================================= */

  useEffect(() => {
    if (!employees.length) {
      return
    }

    setDepartments((current) =>
      mergeEmployeesIntoTree(
        current,
        employees
      )
    )
  }, [employees])

  /* =========================================================
     Selected department
  ========================================================= */

  const selectedDepartment =
    useMemo(() => {
      return (
        departments.find(
          (department) =>
            department.id ===
            selectedDepartmentId
        ) ||
        departments[0] ||
        null
      )
    }, [
      departments,
      selectedDepartmentId,
    ])

  const selectedChildInfo =
    useMemo(() => {
      if (
        !selectedDepartment ||
        !selectedChildId
      ) {
        return null
      }

      return findNodeById(
        selectedDepartment.children,
        selectedChildId
      )
    }, [
      selectedDepartment,
      selectedChildId,
    ])

  const selectedChild =
    selectedChildInfo?.node || null

  const selectedPathNames =
    selectedChildInfo?.pathNames || []

  /* =========================================================
     Current employees
  ========================================================= */

  const departmentEmployees =
    useMemo(() => {
      // حالت «کل سازمان»: همه رکوردهای Excel را نشان بده.
      // این باعث می‌شود کارمندانی که واحدشان در ساختار فعلی
      // وجود ندارد هم از جدول حذف نشوند.
      if (showAllEmployees) {
        return employees
      }

      if (!selectedDepartment) {
        return []
      }

      // شاخه‌ی انتخاب‌شده + همه‌ی زیرشاخه‌های آن
      return employees.filter(
        (employee) =>
          employeeInBranch(
            employee,
            selectedDepartment.name,
            selectedPathNames
          )
      )
    }, [
      employees,
      selectedDepartment,
      selectedPathNames,
      showAllEmployees,
    ])

  const filteredEmployees =
    useMemo(() => {
      const search =
        normalizeText(searchText)

      if (!search) {
        return departmentEmployees
      }

      return departmentEmployees.filter(
        (employee) =>
          Object.values(
            employee
          ).some((value) =>
            normalizeText(
              value
            ).includes(search)
          )
      )
    }, [
      departmentEmployees,
      searchText,
    ])

  /* =========================================================
     Save HR configuration
  ========================================================= */

  const saveConfig = async (
    nextDepartments,
    nextColumns
  ) => {
    setSavingConfig(true)

    try {
      const nextHiddenColumns =
        DEFAULT_COLUMNS
          .map((column) => column.key)
          .filter(
            (key) =>
              !nextColumns.some(
                (column) => column.key === key
              )
          )

      setHiddenColumns(nextHiddenColumns)

      await saveHRConfig({
        departments: nextDepartments,

        customColumns: nextColumns.filter(
          (column) => column.custom
        ),

        hiddenColumns: nextHiddenColumns,

        verifiedCells: {},
      })
    } catch (error) {
      console.error(
        "خطا در ذخیره تنظیمات:",
        error
      )

      alert(
        "ذخیره تنظیمات انجام نشد."
      )
    } finally {
      setSavingConfig(false)
    }
  }

  /* =========================================================
     Add root department
  ========================================================= */

  const handleAddDepartment =
    async () => {
      const name =
        newDepartmentName.trim()

      if (!name) {
        return
      }

      const exists =
        departments.some(
          (department) =>
            normalizeText(
              department.name
            ) ===
            normalizeText(name)
        )

      if (exists) {
        alert(
          "این واحد قبلاً وجود دارد."
        )
        return
      }

      const newDepartment = {
        id: createId(
          "department",
          name
        ),
        name,
        children: [],
      }

      const nextDepartments = [
        ...departments,
        newDepartment,
      ]

      setDepartments(
        nextDepartments
      )

      setSelectedDepartmentId(
        newDepartment.id
      )

      setSelectedChildId(null)

      setExpandedDepartments(
        (current) => ({
          ...current,
          [newDepartment.id]:
            true,
        })
      )

      setNewDepartmentName("")

      setShowAddDepartment(false)

      await saveConfig(
        nextDepartments,
        columns
      )
    }

  /* =========================================================
     Add subdepartment
  ========================================================= */

  const handleAddSubdepartment =
    async () => {
      if (!selectedDepartment) {
        return
      }

      const name =
        newSubdepartmentName.trim()

      if (!name) {
        return
      }

      // اگر زیرواحدی انتخاب شده باشد، زیر همان ساخته می‌شود؛
      // وگرنه مستقیم زیر واحد اصلی.
      const siblings = selectedChild
        ? selectedChild.children || []
        : selectedDepartment.children || []

      const exists = siblings.some(
        (child) =>
          normalizeText(child.name) ===
          normalizeText(name)
      )

      if (exists) {
        alert(
          "این زیرواحد قبلاً وجود دارد."
        )
        return
      }

      const newChild = {
        id: createId(
          "subdepartment",
          [
            selectedDepartment.name,
            ...selectedPathNames,
            name,
          ].join("-")
        ),
        name,
        children: [],
      }

      const nextDepartments =
        addChildNode(
          departments,
          selectedDepartment.id,
          selectedChild?.id || null,
          newChild
        )

      setDepartments(
        nextDepartments
      )

      setSelectedChildId(
        newChild.id
      )

      setExpandedDepartments(
        (current) => ({
          ...current,
          [selectedDepartment.id]:
            true,
        })
      )

      setNewSubdepartmentName("")

      setShowAddSubdepartment(
        false
      )

      await saveConfig(
        nextDepartments,
        columns
      )
    }

  /* =========================================================
     Delete root department
  ========================================================= */

  const handleDeleteDepartment =
    async (department) => {
      const employeeCount =
        getDepartmentCount(
          department.name
        )

      if (employeeCount > 0) {
        alert(
          `امکان حذف واحد «${department.name}» وجود ندارد.\n\n${employeeCount} کارمند در این واحد قرار دارند.\nابتدا کارکنان این واحد را به واحد دیگری منتقل کن.`
        )

        return
      }

      const confirmed =
        window.confirm(
          `آیا مطمئنی می‌خواهی واحد «${department.name}» حذف شود؟\n\nتمام زیرواحدهای این واحد نیز حذف خواهند شد.`
        )

      if (!confirmed) {
        return
      }

      const nextDepartments =
        departments.filter(
          (item) =>
            item.id !==
            department.id
        )

      setDepartments(
        nextDepartments
      )

      if (
        selectedDepartmentId ===
        department.id
      ) {
        const nextDepartment =
          nextDepartments[0]

        setSelectedDepartmentId(
          nextDepartment?.id || null
        )

        setSelectedChildId(null)
        setSearchText("")
      }

      setExpandedDepartments(
        (current) => {
          const next = {
            ...current,
          }

          delete next[department.id]

          return next
        }
      )

      await saveConfig(
        nextDepartments,
        columns
      )
    }

  /* =========================================================
     Delete subdepartment
  ========================================================= */

  const handleDeleteSubdepartment =
    async (
      department,
      child,
      pathNames
    ) => {
      const employeeCount =
        countNodeEmployees(
          employees,
          department.name,
          pathNames || [child.name]
        )

      if (employeeCount > 0) {
        alert(
          `امکان حذف زیرواحد «${child.name}» وجود ندارد.\n\n${employeeCount} کارمند در این زیرواحد (یا زیرشاخه‌های آن) قرار دارند.\nابتدا کارکنان را به زیرواحد دیگری منتقل کن یا زیرواحد آنها را خالی کن.`
        )

        return
      }

      if (
        (child.children || []).length > 0
      ) {
        alert(
          `زیرواحد «${child.name}» دارای زیرشاخه است.\nابتدا زیرشاخه‌ها را حذف کن.`
        )

        return
      }

      const confirmed =
        window.confirm(
          `آیا مطمئنی می‌خواهی زیرواحد «${child.name}» از «${department.name}» حذف شود؟`
        )

      if (!confirmed) {
        return
      }

      const nextDepartments =
        removeChildNode(
          departments,
          department.id,
          child.id
        )

      setDepartments(
        nextDepartments
      )

      if (
        selectedDepartmentId ===
          department.id &&
        selectedChildId ===
          child.id
      ) {
        setSelectedChildId(null)
        setSearchText("")
      }

      await saveConfig(
        nextDepartments,
        columns
      )
    }

  /* =========================================================
     Add custom column
  ========================================================= */

  const handleAddColumn =
    async () => {
      const label =
        newColumnName.trim()

      if (!label) {
        return
      }

      const key = createId(
        "custom",
        label
      )

      const exists =
        columns.some(
          (column) =>
            normalizeText(
              column.label
            ) ===
            normalizeText(label)
        )

      if (exists) {
        alert(
          "این ستون قبلاً وجود دارد."
        )

        return
      }

      const newColumn = {
        key,
        label,
        custom: true,
      }

      const nextColumns = [
        ...columns,
        newColumn,
      ]

      setColumns(
        nextColumns
      )

      setNewColumnName("")

      setShowAddColumn(false)

      await saveConfig(
        departments,
        nextColumns
      )
    }

  /* =========================================================
     Delete ANY column
  ========================================================= */

  const handleDeleteColumn =
    async (column) => {
      const confirmed =
        window.confirm(
          `آیا ستون «${column.label}» از جدول منابع انسانی حذف شود؟`
        )

      if (!confirmed) {
        return
      }

      const nextColumns =
        columns.filter(
          (item) => item.key !== column.key
        )

      setColumns(nextColumns)

      // برای ستون‌های اصلی، اطلاعات Excel حذف نمی‌شود؛
      // فقط ستون از صفحه منابع انسانی مخفی می‌شود.
      //
      // برای ستون سفارشی، چون داده فقط مخصوص همین ستون است،
      // مقدارهای آن ستون و وضعیت تأییدش هم پاک می‌شوند.
      if (column.custom && onEmployeesChange) {
        const nextEmployees =
          employees.map((employee) => {
            const copy = { ...employee }

            delete copy[column.key]
            delete copy[`is_${column.key}_verified`]

            return copy
          })

        await onEmployeesChange(nextEmployees)
      }

      // saveConfig خودش hiddenColumns را از روی nextColumns
      // محاسبه و ذخیره می‌کند؛ بنابراین ستون‌های پیش‌فرض هم
      // بعد از Refresh حذف‌شده باقی می‌مانند.
      await saveConfig(
        departments,
        nextColumns
      )
    }

  /* =========================================================
     Update employee field
  ========================================================= */

  const updateEmployeeField =
    async (
      employeeId,
      key,
      value
    ) => {
      const nextEmployees =
        employees.map(
          (employee, index) => {
            const currentId =
              getEmployeeId(employee)

            if (
              currentId !==
              String(employeeId)
            ) {
              return employee
            }

            return {
              ...employee,
              [key]: value,
            }
          }
        )

      if (onEmployeesChange) {
        await onEmployeesChange(
          nextEmployees
        )
      }
    }

  /* =========================================================
     Move employee
  ========================================================= */

  const updateEmployeeDepartment =
    async (
      employeeId,
      departmentName,
      subDepartmentName
    ) => {
      const nextEmployees =
        employees.map(
          (employee, index) => {
            const currentId =
              getEmployeeId(employee)

            if (
              currentId !==
              String(employeeId)
            ) {
              return employee
            }

            return {
              ...employee,

              department:
                departmentName,

              sub_department:
                subDepartmentName ||
                "",
            }
          }
        )

      if (onEmployeesChange) {
        await onEmployeesChange(
          nextEmployees
        )
      }
    }

  /* =========================================================
     Add employee manually
  ========================================================= */

  const openAddEmployee = (
    departmentId,
    subPath
  ) => {
    const fallbackDepartment =
      departments.find(
        (department) =>
          department.id === departmentId
      ) ||
      selectedDepartment ||
      departments[0] ||
      null

    setEditingEmployeeId(null)
    setApplyEditToDashboard(false)
    setNewEmployeeForm({})
    setNewEmployeeDepartmentId(
      fallbackDepartment?.id || ""
    )
    setNewEmployeeSubPath(subPath || "")
    setShowAddEmployee(true)
  }

  // باز کردن فرم در حالت ویرایش با اطلاعات فعلی کارمند
  const openEditEmployee = (employeeId) => {
    const target = employees.find(
      (item) => getEmployeeId(item) === String(employeeId)
    )

    if (!target) {
      return
    }

    const department = departments.find(
      (item) =>
        normalizeText(item.name) ===
        normalizeText(target.department)
    )

    const form = { name: target.name ?? "" }

    for (const field of [
      ...DASHBOARD_FIELDS,
      ...columns.filter((column) => column.custom),
    ]) {
      form[field.key] = target[field.key] ?? ""
    }

    setEditingEmployeeId(String(employeeId))
    setApplyEditToDashboard(false)
    setNewEmployeeForm(form)
    setNewEmployeeDepartmentId(department?.id || "")
    setNewEmployeeSubPath(
      joinSubPath(splitSubPath(target.sub_department))
    )
    setShowAddEmployee(true)
  }

  // دکمه‌ی «کارمند جدید»: واحد/زیرواحد فعلی را پیش‌فرض می‌گذارد
  const openAddEmployeeFromSelection = () => {
    openAddEmployee(
      selectedDepartment?.id,
      !showAllEmployees && selectedChild
        ? joinSubPath(selectedPathNames)
        : ""
    )
  }

  const handleAddEmployee = async () => {
    const name = String(
      newEmployeeForm.name || ""
    ).trim()

    if (!name) {
      alert("نام کارمند را وارد کن.")
      return
    }

    const department = departments.find(
      (item) =>
        item.id === newEmployeeDepartmentId
    )

    if (!department) {
      alert("واحد کارمند را انتخاب کن.")
      return
    }

    const isEditing = Boolean(editingEmployeeId)

    const baseEmployee = isEditing
      ? employees.find(
          (item) =>
            getEmployeeId(item) === editingEmployeeId
        )
      : null

    if (isEditing && !baseEmployee) {
      alert("کارمند پیدا نشد.")
      return
    }

    const newEmployee = isEditing
      ? {
          ...baseEmployee,
          name,
          department: department.name,
          sub_department: newEmployeeSubPath || "",
        }
      : {
          id: crypto.randomUUID(),
          name,
          department: department.name,
          sub_department: newEmployeeSubPath || "",
        }

    // سوال‌های ثابت داشبورد + سوال‌های سفارشی منابع انسانی
    const questionFields = [
      ...DASHBOARD_FIELDS,
      ...columns.filter((column) => column.custom),
    ]

    for (const column of questionFields) {
      const raw = String(
        newEmployeeForm[column.key] ?? ""
      ).trim()

      if (raw === "") {
        if (isEditing) {
          delete newEmployee[column.key]
        }

        continue
      }

      if (
        column.type === "number" ||
        column.key === "salary" ||
        column.key === "age"
      ) {
        const numeric = Number(
          raw
            .replace(/[۰-۹]/g, (digit) =>
              "۰۱۲۳۴۵۶۷۸۹".indexOf(digit)
            )
            .replace(/[,٬،\s]/g, "")
        )

        newEmployee[column.key] =
          Number.isFinite(numeric) ? numeric : raw
      } else {
        newEmployee[column.key] = raw
      }
    }

    if (isEditing) {
      /*
        به‌صورت پیش‌فرض ویرایش فقط اطلاعات صفحه منابع انسانی را
        عوض می‌کند و snapshot داشبورد دست‌نخورده می‌ماند.
        اگر کاربر گزینه «اعمال روی داشبورد» را بزند، snapshot هم
        با مقدارهای جدید به‌روز می‌شود.
      */
      if (applyEditToDashboard) {
        const snapshot = {
          ...(baseEmployee[DASHBOARD_SNAPSHOT_KEY] || {}),
        }

        for (const key of [
          "name",
          "department",
          "sub_department",
          ...DASHBOARD_FIELDS.map((field) => field.key),
        ]) {
          const value = newEmployee[key]

          if (
            value === undefined ||
            value === null ||
            value === ""
          ) {
            delete snapshot[key]
          } else {
            snapshot[key] = value
          }
        }

        newEmployee[DASHBOARD_SNAPSHOT_KEY] = snapshot
      }

      if (onEmployeesChange) {
        await onEmployeesChange(
          employees.map((item) =>
            getEmployeeId(item) === editingEmployeeId
              ? newEmployee
              : item
          )
        )
      }

      setShowAddEmployee(false)
      setEditingEmployeeId(null)
      setNewEmployeeForm({})
      return
    }

    if (onEmployeesChange) {
      await onEmployeesChange([
        ...employees,
        newEmployee,
      ])
    }

    // نمایش همان واحد/زیرواحدی که کارمند در آن ثبت شد
    const parts = splitSubPath(newEmployeeSubPath)

    let targetChildId = null
    let siblings = department.children || []

    for (const part of parts) {
      const node = siblings.find(
        (child) =>
          normalizeText(child.name) ===
          normalizeText(part)
      )

      if (!node) {
        break
      }

      targetChildId = node.id
      siblings = node.children || []
    }

    setShowAllEmployees(false)
    setSelectedDepartmentId(department.id)
    setSelectedChildId(targetChildId)
    setSearchText("")

    setExpandedDepartments((current) => ({
      ...current,
      [department.id]: true,
    }))

    setShowAddEmployee(false)
    setNewEmployeeForm({})
  }

  /* =========================================================
     Delete employee
  ========================================================= */

  const handleDeleteEmployee = async (employeeId) => {
    const target = employees.find(
      (item) => getEmployeeId(item) === String(employeeId)
    )

    if (!target) {
      return
    }

    const confirmed = window.confirm(
      `کارمند «${target.name || "بدون نام"}» حذف شود؟\nاین کار قابل بازگشت نیست.`
    )

    if (!confirmed) {
      return
    }

    if (onEmployeesChange) {
      await onEmployeesChange(
        employees.filter(
          (item) => getEmployeeId(item) !== String(employeeId)
        )
      )
    }
  }

  /* =========================================================
     Toggle verification
  ========================================================= */

  const toggleVerification =
    async (employeeId, columnKey) => {
      if (!employeeId) {
        return
      }

      const verificationKey =
        `is_${columnKey}_verified`

      const targetEmployee =
        employees.find(
          (item) =>
            getEmployeeId(item) ===
            String(employeeId)
        )

      if (!targetEmployee) {
        return
      }

      const nextValue =
        !Boolean(targetEmployee[verificationKey])

      // تأیید بر اساس شناسه ثابت همان کارمند انجام می‌شود.
      // بنابراین فیلتر، جستجو یا تغییر ترتیب جدول روی فرد دیگری اثر نمی‌گذارد.
      const nextEmployees =
        employees.map((item) =>
          getEmployeeId(item) ===
          String(employeeId)
            ? {
                ...item,
                [verificationKey]: nextValue,
              }
            : item
        )

      if (onEmployeesChange) {
        await onEmployeesChange(nextEmployees)
      }
    }

  /* =========================================================
     Department selection
  ========================================================= */

  const selectAllEmployees = () => {
    setShowAllEmployees(true)
    setSelectedChildId(null)
    setSearchText("")
  }

  const selectDepartment = (
    department
  ) => {
    setShowAllEmployees(false)

    setSelectedDepartmentId(
      department.id
    )

    setSelectedChildId(null)

    setSearchText("")

    setExpandedDepartments(
      (current) => ({
        ...current,
        [department.id]:
          true,
      })
    )
  }

  const selectChild = (
    department,
    child
  ) => {
    setShowAllEmployees(false)

    setSelectedDepartmentId(
      department.id
    )

    setSelectedChildId(
      child.id
    )

    setSearchText("")

    setExpandedDepartments(
      (current) => ({
        ...current,
        [department.id]:
          true,
      })
    )
  }

  const openAddChild = (
    department,
    child
  ) => {
    selectChild(department, child)
    setShowAddSubdepartment(true)
  }

  const toggleDepartment =
    (departmentId) => {
      setExpandedDepartments(
        (current) => ({
          ...current,

          [departmentId]:
            !current[departmentId],
        })
      )
    }

  /* =========================================================
     Department statistics
  ========================================================= */

  const getDepartmentCount =
    (departmentName) => {
      return employees.filter(
        (employee) =>
          normalizeText(
            employee.department
          ) ===
          normalizeText(
            departmentName
          )
      ).length
    }

  const getChildCount =
    (
      departmentName,
      childName
    ) => {
      return employees.filter(
        (employee) =>
          normalizeText(
            employee.department
          ) ===
            normalizeText(
              departmentName
            ) &&
          normalizeText(
            employee.sub_department
          ) ===
            normalizeText(
              childName
            )
      ).length
    }

  /* =========================================================
     Render
  ========================================================= */

  return (
    <div
      dir="rtl"
      className="min-h-screen bg-[#1a1a1a] text-white"
    >
      {/* =====================================================
          Header
      ===================================================== */}

      <header className="sticky top-0 z-30 flex min-h-20 items-center justify-between border-b border-[#d4a017]/40 bg-[#111111]/95 px-4 py-4 backdrop-blur md:px-6">
        <div className="flex items-center gap-3">
          <button
            onClick={onBack}
            className="flex h-10 w-10 items-center justify-center rounded-xl border border-white/10 bg-[#202020] text-gray-300 transition hover:border-[#d4a017] hover:text-[#f0c040]"
            title="بازگشت به داشبورد"
          >
            <ArrowRight size={19} />
          </button>

          <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-[#d4a017] text-black">
            <Users size={23} />
          </div>

          <div>
            <h1 className="text-lg font-black md:text-2xl">
              مدیریت منابع انسانی
            </h1>

            <p className="mt-1 text-xs text-[#d4a017] md:text-sm">
              مدیریت کارکنان، واحدها و اطلاعات سازمانی
            </p>
          </div>
        </div>

        <div className="hidden items-center gap-3 sm:flex">
          <div className="rounded-lg border border-white/10 bg-[#202020] px-4 py-2 text-xs text-gray-400">
            {savingConfig
              ? "در حال ذخیره..."
              : "ذخیره شده"}
          </div>

          <div className="flex h-9 w-9 items-center justify-center rounded-full bg-[#d4a017] font-bold text-black">
            م
          </div>
        </div>
      </header>

      {/* =====================================================
          Main layout
      ===================================================== */}

      <div className="flex min-h-[calc(100vh-80px)] flex-col lg:flex-row">

        {/* ===================================================
            Department sidebar
        =================================================== */}

        <aside className="w-full shrink-0 border-b border-[#d4a017]/30 bg-[#151515] p-4 lg:w-80 lg:border-b-0 lg:border-l">

          {/* Sidebar header */}

          <div className="mb-4 flex items-center justify-between">
            <div>
              <h2 className="text-base font-black">
                ساختار سازمانی
              </h2>

              <p className="mt-1 text-xs text-gray-500">
                واحدها و زیرواحدها
              </p>
            </div>

            <button
              onClick={() =>
                setShowAddDepartment(true)
              }
              className="flex items-center gap-1.5 rounded-lg bg-[#d4a017] px-3 py-2 text-xs font-bold text-black transition hover:bg-[#f0c040]"
            >
              <Plus size={15} />
              واحد جدید
            </button>
          </div>

          {/* All employees */}

          <button
            onClick={selectAllEmployees}
            className={`mb-3 flex w-full items-center gap-2 rounded-xl border px-3 py-3 text-right transition ${
              showAllEmployees
                ? "border-[#d4a017] bg-[#d4a017]/10 text-[#f0c040]"
                : "border-white/5 bg-[#1d1d1d] text-gray-300 hover:border-[#d4a017]/50"
            }`}
          >
            <Users size={17} />
            <span className="flex-1 text-sm font-bold">
              کل کارکنان
            </span>
            <span className="rounded-md bg-white/5 px-2 py-1 text-[10px] text-gray-500">
              {employees.length}
            </span>
          </button>

          {/* Department tree */}

          <div className="max-h-[55vh] space-y-2 overflow-y-auto pr-1 lg:max-h-[calc(100vh-190px)]">

            {departments.map(
              (department) => {
                const expanded =
                  Boolean(
                    expandedDepartments[
                      department.id
                    ]
                  )

                const active =
                  !showAllEmployees &&
                  selectedDepartmentId ===
                    department.id &&
                  !selectedChildId

                return (
                  <div
                    key={
                      department.id
                    }
                  >
                    {/* Root department */}

                    <div
                      className={`flex items-center gap-1 rounded-xl border transition ${
                        active
                          ? "border-[#d4a017] bg-[#d4a017]/10"
                          : "border-white/5 bg-[#1d1d1d]"
                      }`}
                    >
                      <button
                        onClick={() =>
                          toggleDepartment(
                            department.id
                          )
                        }
                        className="flex h-10 w-8 items-center justify-center text-gray-500 hover:text-white"
                      >
                        {department
                          .children
                          ?.length > 0 ? (
                          expanded ? (
                            <ChevronDown
                              size={16}
                            />
                          ) : (
                            <ChevronLeft
                              size={16}
                            />
                          )
                        ) : null}
                      </button>

                      <button
                        onClick={() =>
                          selectDepartment(
                            department
                          )
                        }
                        className="flex min-w-0 flex-1 items-center gap-2 py-2.5 text-right"
                      >
                        <Building2
                          size={17}
                          className={
                            active
                              ? "text-[#f0c040]"
                              : "text-gray-500"
                          }
                        />

                        <span
                          className={`truncate text-sm ${
                            active
                              ? "font-bold text-[#f0c040]"
                              : "text-gray-200"
                          }`}
                        >
                          {
                            department.name
                          }
                        </span>
                      </button>

                      <span className="rounded-md bg-white/5 px-2 py-1 text-[10px] text-gray-500">
                        {getDepartmentCount(
                          department.name
                        )}
                      </span>

                      {/* Add employee to root department */}

                      <button
                        onClick={() =>
                          openAddEmployee(
                            department.id,
                            ""
                          )
                        }
                        className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg text-gray-600 transition hover:bg-[#d4a017]/10 hover:text-[#f0c040]"
                        title="افزودن کارمند به این واحد"
                      >
                        <UserPlus
                          size={15}
                        />
                      </button>

                      {/* Delete root department */}

                      <button
                        onClick={() =>
                          handleDeleteDepartment(
                            department
                          )
                        }
                        className="ml-1 flex h-8 w-8 shrink-0 items-center justify-center rounded-lg text-gray-600 transition hover:bg-red-500/10 hover:text-red-400"
                        title="حذف واحد"
                      >
                        <Trash2
                          size={15}
                        />
                      </button>
                    </div>

                    {/* Children (recursive tree) */}

                    {expanded &&
                      department
                        .children
                        ?.length > 0 && (
                        <div className="mr-5 mt-1 border-r border-white/10 pr-2">
                          <HRChildTree
                            department={department}
                            nodes={department.children}
                            employees={employees}
                            activeChildId={
                              !showAllEmployees &&
                              selectedDepartmentId ===
                                department.id
                                ? selectedChildId
                                : null
                            }
                            onSelect={selectChild}
                            onAddChild={openAddChild}
                            onDelete={
                              handleDeleteSubdepartment
                            }
                            onAddEmployee={(
                              dept,
                              node,
                              pathNames
                            ) =>
                              openAddEmployee(
                                dept.id,
                                joinSubPath(
                                  pathNames
                                )
                              )
                            }
                          />
                        </div>
                      )}
                  </div>
                )
              }
            )}

          </div>
        </aside>

        {/* ===================================================
            Content
        =================================================== */}

        <main className="min-w-0 flex-1 overflow-auto p-4 md:p-6">

          {/* Page title */}

          <div className="mb-5 flex flex-col justify-between gap-4 xl:flex-row xl:items-center">
            <div>
              <div className="flex flex-wrap items-center gap-2">

                <h2 className="text-xl font-black md:text-2xl">
                  {selectedChild
                    ? selectedChild.name
                    : selectedDepartment?.name ||
                      "منابع انسانی"}
                </h2>

                {selectedChild && (
                  <>
                    <ChevronLeft
                      size={18}
                      className="text-gray-600"
                    />

                    <span className="rounded-lg border border-[#d4a017]/30 bg-[#d4a017]/10 px-2.5 py-1 text-xs text-[#f0c040]">
                      {[
                        selectedDepartment?.name,
                        ...selectedPathNames.slice(0, -1),
                      ].join(" / ")}
                    </span>
                  </>
                )}

              </div>

              <p className="mt-1 text-sm text-gray-500">
                {filteredEmployees.length} نفر
                در این بخش
              </p>
            </div>

            <div className="flex flex-col gap-2 sm:flex-row">

              <div className="relative">
                <Search
                  size={17}
                  className="absolute right-3 top-3 text-gray-500"
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
                  placeholder="جستجو در کارکنان..."
                  className="w-full rounded-xl border border-white/10 bg-[#202020] py-2.5 pl-4 pr-10 text-sm outline-none transition focus:border-[#d4a017] sm:w-64"
                />
              </div>

              <button
                onClick={() =>
                  setShowAddSubdepartment(
                    true
                  )
                }
                disabled={
                  !selectedDepartment
                }
                className="flex items-center justify-center gap-2 rounded-xl border border-[#d4a017]/50 bg-[#202020] px-4 py-2.5 text-sm font-bold text-[#f0c040] transition hover:bg-[#d4a017]/10 disabled:cursor-not-allowed disabled:opacity-40"
              >
                <FolderPlus
                  size={17}
                />
                زیرواحد جدید
              </button>

              <button
                onClick={() =>
                  downloadEmployeesExcel(
                    employees,
                    {
                      departments,
                      customColumns: columns.filter(
                        (column) => column.custom
                      ),
                      hiddenColumns,
                    }
                  )
                }
                disabled={
                  employees.length === 0
                }
                className="flex items-center justify-center gap-2 rounded-xl border border-[#d4a017]/50 bg-[#d4a017]/10 px-4 py-2.5 text-sm font-bold text-[#f0c040] transition hover:bg-[#d4a017]/20 disabled:cursor-not-allowed disabled:opacity-40"
                title="دانلود آخرین وضعیت منابع انسانی"
              >
                <Download size={17} />
                دانلود اکسل نهایی
              </button>

              <button
                onClick={
                  openAddEmployeeFromSelection
                }
                disabled={
                  departments.length === 0
                }
                className="flex items-center justify-center gap-2 rounded-xl border border-[#d4a017]/50 bg-[#202020] px-4 py-2.5 text-sm font-bold text-[#f0c040] transition hover:bg-[#d4a017]/10 disabled:cursor-not-allowed disabled:opacity-40"
                title="افزودن دستی کارمند به واحد/زیرواحد انتخاب‌شده"
              >
                <UserPlus size={17} />
                کارمند جدید
              </button>

              <button
                onClick={() =>
                  setShowAddColumn(true)
                }
                className="flex items-center justify-center gap-2 rounded-xl bg-[#d4a017] px-4 py-2.5 text-sm font-bold text-black transition hover:bg-[#f0c040]"
              >
                <Plus size={17} />
                ستون جدید
              </button>

            </div>
          </div>

          {/* Loading */}

          {loadingConfig && (
            <div className="mb-4 rounded-xl border border-[#d4a017]/30 bg-[#202020] p-4 text-sm text-gray-400">
              در حال خواندن تنظیمات منابع انسانی...
            </div>
          )}

          {/* Organization notice */}

          <div className="mb-4 rounded-xl border border-[#d4a017]/30 bg-[#202020] p-4">

            <div className="flex flex-col justify-between gap-3 md:flex-row md:items-center">

              <div>
                <p className="text-sm font-bold text-[#f0c040]">
                  مدیریت ساختار سازمانی
                </p>

                <p className="mt-1 text-xs leading-6 text-gray-500">
                  واحدها و زیرواحدهایی که داخل Excel
                  وجود داشته باشند به‌صورت خودکار
                  شناسایی می‌شوند. همچنین می‌توانی
                  واحد جدید بسازی و کارکنان را به آن
                  منتقل کنی.
                </p>
              </div>

              <div className="shrink-0 rounded-lg border border-white/10 bg-[#181818] px-3 py-2 text-xs text-gray-400">
                {showAllEmployees
                  ? "کل سازمان"
                  : selectedChild
                    ? [
                        selectedDepartment?.name,
                        ...selectedPathNames,
                      ].join(" / ")
                    : selectedDepartment?.name ||
                      "انتخاب نشده"}
              </div>

            </div>
          </div>

          {/* Employee table */}

          <div className="overflow-hidden rounded-2xl border border-[#d4a017]/40 bg-[#202020]">

            {/* Table header */}

            <div className="flex flex-col justify-between gap-3 border-b border-white/10 p-4 md:flex-row md:items-center">

              <div>
                <h3 className="text-base font-black">
                  لیست کارکنان
                </h3>

                <p className="mt-1 text-xs text-gray-500">
                  در حالت «کل کارکنان» همه رکوردهای Excel نمایش داده می‌شوند؛
                  برای فیلتر کردن بر اساس واحد، از ساختار سازمانی استفاده کن.
                </p>
              </div>

              <span className="rounded-lg border border-[#d4a017]/30 bg-[#d4a017]/10 px-3 py-1.5 text-xs text-[#f0c040]">
                {filteredEmployees.length} رکورد
              </span>

            </div>

            {filteredEmployees.length ===
            0 ? (
              <EmptyEmployees
                selectedDepartment={
                  selectedDepartment
                }
                selectedChild={
                  selectedChild
                }
                showAllEmployees={
                  showAllEmployees
                }
                onClearChild={() =>
                  setSelectedChildId(
                    null
                  )
                }
                onAddEmployee={
                  openAddEmployeeFromSelection
                }
              />
            ) : (
              <div className="overflow-x-auto">

                <table className="min-w-[1250px] w-full border-collapse">

                  <thead>
                    <tr className="border-b border-white/10 bg-[#181818]">

                      <th className="sticky right-0 z-10 border-l border-white/5 bg-[#181818] px-4 py-3 text-right text-xs font-bold text-gray-400">
                        کارمند
                      </th>

                      <th className="px-4 py-3 text-right text-xs font-bold text-gray-400">
                        سازماندهی
                      </th>

                      {columns.map(
                        (column) => (
                          <th
                            key={
                              column.key
                            }
                            className="px-4 py-3 text-right text-xs font-bold text-gray-400"
                          >
                            <div className="flex min-w-[120px] items-center justify-between gap-2">

                              <span>
                                {
                                  column.label
                                }
                              </span>

                              <button
                                onClick={() =>
                                  handleDeleteColumn(
                                    column
                                  )
                                }
                                className="text-gray-600 transition hover:text-red-400"
                                title="حذف ستون"
                              >
                                <Trash2
                                  size={14}
                                />
                              </button>

                            </div>
                          </th>
                        )
                      )}

                      <th className="px-4 py-3 text-right text-xs font-bold text-gray-400">
                        عملیات
                      </th>

                    </tr>
                  </thead>

                  <tbody>
                    {filteredEmployees.map(
                      (employee) => {
                        const employeeId =
                          getEmployeeId(employee)

                        return (
                          <EmployeeRow
                            key={
                              employeeId
                            }
                            employee={
                              employee
                            }
                            employeeId={
                              employeeId
                            }
                            departments={
                              departments
                            }
                            columns={
                              columns
                            }
                            onUpdateField={
                              updateEmployeeField
                            }
                            onUpdateDepartment={
                              updateEmployeeDepartment
                            }
                            onToggleVerification={
                              toggleVerification
                            }
                            onDelete={
                              handleDeleteEmployee
                            }
                            onEdit={
                              openEditEmployee
                            }
                          />
                        )
                      }
                    )}
                  </tbody>

                </table>
              </div>
            )}
          </div>
        </main>
      </div>

      {/* =====================================================
          Add Department Modal
      ===================================================== */}

      {showAddDepartment && (
        <Modal
          title="افزودن واحد جدید"
          onClose={() =>
            setShowAddDepartment(
              false
            )
          }
        >
          <p className="mb-4 text-sm leading-6 text-gray-400">
            نام واحد اصلی سازمان را وارد کن.
          </p>

          <input
            autoFocus
            value={
              newDepartmentName
            }
            onChange={(event) =>
              setNewDepartmentName(
                event.target.value
              )
            }
            onKeyDown={(event) => {
              if (
                event.key ===
                "Enter"
              ) {
                handleAddDepartment()
              }
            }}
            placeholder="مثلاً منابع فنی"
            className="w-full rounded-xl border border-white/10 bg-[#181818] px-4 py-3 text-sm outline-none focus:border-[#d4a017]"
          />

          <div className="mt-4 flex justify-end gap-2">

            <button
              onClick={() =>
                setShowAddDepartment(
                  false
                )
              }
              className="rounded-xl border border-white/10 px-4 py-2.5 text-sm text-gray-400 hover:text-white"
            >
              لغو
            </button>

            <button
              onClick={
                handleAddDepartment
              }
              className="rounded-xl bg-[#d4a017] px-5 py-2.5 text-sm font-bold text-black hover:bg-[#f0c040]"
            >
              افزودن
            </button>

          </div>
        </Modal>
      )}

      {/* =====================================================
          Add Subdepartment Modal
      ===================================================== */}

      {showAddSubdepartment && (
        <Modal
          title="افزودن زیرواحد"
          onClose={() =>
            setShowAddSubdepartment(
              false
            )
          }
        >
          <p className="mb-4 text-sm leading-6 text-gray-400">
            زیرواحد جدید برای «
            <span className="font-bold text-[#f0c040]">
              {[
                selectedDepartment?.name,
                ...selectedPathNames,
              ].join(" / ")}
            </span>
            » ایجاد می‌شود.
          </p>

          <input
            autoFocus
            value={
              newSubdepartmentName
            }
            onChange={(event) =>
              setNewSubdepartmentName(
                event.target.value
              )
            }
            onKeyDown={(event) => {
              if (
                event.key ===
                "Enter"
              ) {
                handleAddSubdepartment()
              }
            }}
            placeholder="مثلاً مدیریت"
            className="w-full rounded-xl border border-white/10 bg-[#181818] px-4 py-3 text-sm outline-none focus:border-[#d4a017]"
          />

          <div className="mt-4 flex justify-end gap-2">

            <button
              onClick={() =>
                setShowAddSubdepartment(
                  false
                )
              }
              className="rounded-xl border border-white/10 px-4 py-2.5 text-sm text-gray-400 hover:text-white"
            >
              لغو
            </button>

            <button
              onClick={
                handleAddSubdepartment
              }
              className="rounded-xl bg-[#d4a017] px-5 py-2.5 text-sm font-bold text-black hover:bg-[#f0c040]"
            >
              افزودن
            </button>

          </div>
        </Modal>
      )}

      {/* =====================================================
          Add Employee Modal
      ===================================================== */}

      {showAddEmployee && (
        <Modal
          title={
            editingEmployeeId
              ? "ویرایش اطلاعات کارمند"
              : "افزودن کارمند جدید"
          }
          onClose={() =>
            setShowAddEmployee(false)
          }
        >
          <div className="max-h-[65vh] space-y-3 overflow-y-auto pr-1">

            <div>
              <label className="mb-1 block text-xs text-gray-400">
                نام و نام خانوادگی *
              </label>

              <input
                autoFocus
                value={newEmployeeForm.name || ""}
                onChange={(event) =>
                  setNewEmployeeForm(
                    (current) => ({
                      ...current,
                      name: event.target.value,
                    })
                  )
                }
                placeholder="مثلاً علی رضایی"
                className="w-full rounded-xl border border-white/10 bg-[#181818] px-4 py-2.5 text-sm outline-none focus:border-[#d4a017]"
              />
            </div>

            <div>
              <label className="mb-1 block text-xs text-gray-400">
                واحد
              </label>

              <select
                value={newEmployeeDepartmentId}
                onChange={(event) => {
                  setNewEmployeeDepartmentId(
                    event.target.value
                  )
                  setNewEmployeeSubPath("")
                }}
                className="w-full rounded-xl border border-white/10 bg-[#181818] px-4 py-2.5 text-sm text-white outline-none focus:border-[#d4a017]"
              >
                {departments.map((department) => (
                  <option
                    key={department.id}
                    value={department.id}
                    className="bg-[#181818]"
                  >
                    {department.name}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="mb-1 block text-xs text-gray-400">
                زیرواحد
              </label>

              <select
                value={newEmployeeSubPath}
                disabled={
                  flattenChildren(
                    departments.find(
                      (department) =>
                        department.id ===
                        newEmployeeDepartmentId
                    )?.children || []
                  ).length === 0
                }
                onChange={(event) =>
                  setNewEmployeeSubPath(
                    event.target.value
                  )
                }
                className="w-full rounded-xl border border-white/10 bg-[#181818] px-4 py-2.5 text-sm text-white outline-none focus:border-[#d4a017] disabled:cursor-not-allowed disabled:opacity-40"
              >
                <option
                  value=""
                  className="bg-[#181818]"
                >
                  بدون زیرواحد (مستقیم زیر واحد اصلی)
                </option>

                {flattenChildren(
                  departments.find(
                    (department) =>
                      department.id ===
                      newEmployeeDepartmentId
                  )?.children || []
                ).map((option) => (
                  <option
                    key={option.id}
                    value={option.path}
                    className="bg-[#181818]"
                  >
                    {option.path}
                  </option>
                ))}
              </select>
            </div>

            {/* دسته اول: سوال‌های ثابت داشبورد */}

            <div className="rounded-xl border border-[#d4a017]/30 bg-[#d4a017]/5 p-3">
              <p className="mb-1 text-sm font-bold text-[#f0c040]">
                اطلاعات ثابت (مورد نیاز داشبورد)
              </p>

              <p className="mb-3 text-[11px] leading-5 text-gray-500">
                این اطلاعات در آمار و نمودارهای داشبورد استفاده می‌شوند.
              </p>

              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                {DASHBOARD_FIELDS.map((field) => (
                  <div key={field.key}>
                    <label className="mb-1 block text-xs text-gray-400">
                      {field.label}
                    </label>

                    {field.type === "select" ? (
                      <select
                        value={newEmployeeForm[field.key] ?? ""}
                        onChange={(event) =>
                          setNewEmployeeForm((current) => ({
                            ...current,
                            [field.key]: event.target.value,
                          }))
                        }
                        className="w-full rounded-xl border border-white/10 bg-[#181818] px-4 py-2.5 text-sm text-white outline-none focus:border-[#d4a017]"
                      >
                        <option value="" className="bg-[#181818]">
                          انتخاب کنید
                        </option>

                        {field.options.map((option) => (
                          <option
                            key={option}
                            value={option}
                            className="bg-[#181818]"
                          >
                            {option}
                          </option>
                        ))}
                      </select>
                    ) : (
                      <input
                        type={field.type === "number" ? "number" : "text"}
                        value={newEmployeeForm[field.key] ?? ""}
                        onChange={(event) =>
                          setNewEmployeeForm((current) => ({
                            ...current,
                            [field.key]: event.target.value,
                          }))
                        }
                        placeholder="اختیاری"
                        className="w-full rounded-xl border border-white/10 bg-[#181818] px-4 py-2.5 text-sm outline-none placeholder:text-gray-700 focus:border-[#d4a017]"
                      />
                    )}
                  </div>
                ))}
              </div>
            </div>

            {/* دسته دوم: سوال‌های منابع انسانی (ستون‌های سفارشی) */}

            <div className="rounded-xl border border-white/10 bg-[#181818] p-3">
              <p className="mb-1 text-sm font-bold text-white">
                اطلاعات منابع انسانی
              </p>

              <p className="mb-3 text-[11px] leading-5 text-gray-500">
                ستون‌هایی که خودت در بخش منابع انسانی اضافه کرده‌ای.
              </p>

              {columns.filter((column) => column.custom).length === 0 ? (
                <p className="text-xs text-gray-600">
                  هنوز ستون سفارشی‌ای تعریف نشده است. با دکمه «ستون جدید»
                  می‌توانی سوال جدید اضافه کنی.
                </p>
              ) : (
                <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                  {columns
                    .filter((column) => column.custom)
                    .map((column) => (
                      <div key={column.key}>
                        <label className="mb-1 block text-xs text-gray-400">
                          {column.label}
                        </label>

                        <input
                          value={newEmployeeForm[column.key] ?? ""}
                          onChange={(event) =>
                            setNewEmployeeForm((current) => ({
                              ...current,
                              [column.key]: event.target.value,
                            }))
                          }
                          placeholder="اختیاری"
                          className="w-full rounded-xl border border-white/10 bg-[#202020] px-4 py-2.5 text-sm outline-none placeholder:text-gray-700 focus:border-[#d4a017]"
                        />
                      </div>
                    ))}
                </div>
              )}
            </div>

            {editingEmployeeId && (
              <label className="flex cursor-pointer items-start gap-2 rounded-xl border border-white/10 bg-[#181818] p-3 text-xs leading-5 text-gray-400">
                <input
                  type="checkbox"
                  checked={applyEditToDashboard}
                  onChange={(event) =>
                    setApplyEditToDashboard(
                      event.target.checked
                    )
                  }
                  className="mt-1 accent-[#d4a017]"
                />
                <span>
                  تغییرات اطلاعات ثابت روی آمار داشبورد هم اعمال شود
                  (در حالت عادی فقط در صفحه منابع انسانی تغییر می‌کند)
                </span>
              </label>
            )}

          </div>

          <div className="mt-4 flex justify-end gap-2">

            <button
              onClick={() =>
                setShowAddEmployee(false)
              }
              className="rounded-xl border border-white/10 px-4 py-2.5 text-sm text-gray-400 hover:text-white"
            >
              لغو
            </button>

            <button
              onClick={handleAddEmployee}
              className="rounded-xl bg-[#d4a017] px-5 py-2.5 text-sm font-bold text-black hover:bg-[#f0c040]"
            >
              {editingEmployeeId
                ? "ذخیره تغییرات"
                : "افزودن کارمند"}
            </button>

          </div>
        </Modal>
      )}

      {/* =====================================================
          Add Column Modal
      ===================================================== */}

      {showAddColumn && (
        <Modal
          title="افزودن ستون جدید"
          onClose={() =>
            setShowAddColumn(
              false
            )
          }
        >
          <p className="mb-4 text-sm leading-6 text-gray-400">
            یک ستون سفارشی برای اطلاعات کارکنان
            ایجاد کن. مقدار آن برای هر کارمند
            قابل ویرایش خواهد بود.
          </p>

          <input
            autoFocus
            value={
              newColumnName
            }
            onChange={(event) =>
              setNewColumnName(
                event.target.value
              )
            }
            onKeyDown={(event) => {
              if (
                event.key ===
                "Enter"
              ) {
                handleAddColumn()
              }
            }}
            placeholder="مثلاً شماره پرسنلی"
            className="w-full rounded-xl border border-white/10 bg-[#181818] px-4 py-3 text-sm outline-none focus:border-[#d4a017]"
          />

          <div className="mt-4 flex justify-end gap-2">

            <button
              onClick={() =>
                setShowAddColumn(
                  false
                )
              }
              className="rounded-xl border border-white/10 px-4 py-2.5 text-sm text-gray-400 hover:text-white"
            >
              لغو
            </button>

            <button
              onClick={
                handleAddColumn
              }
              className="rounded-xl bg-[#d4a017] px-5 py-2.5 text-sm font-bold text-black hover:bg-[#f0c040]"
            >
              افزودن ستون
            </button>

          </div>
        </Modal>
      )}

    </div>
  )
}

/* =========================================================
   Employee row
========================================================= */

function EmployeeRow({
  employee,
  employeeId,
  departments,
  columns,
  onUpdateField,
  onUpdateDepartment,
  onToggleVerification,
  onDelete,
  onEdit,
}) {
  const employeeDepartment =
    String(
      employee.department || ""
    ).trim()

  const employeeSubDepartment =
    String(
      employee.sub_department ||
        ""
    ).trim()

  const selectedDepartment =
    departments.find(
      (department) =>
        normalizeText(
          department.name
        ) ===
        normalizeText(
          employeeDepartment
        )
    )

  const subOptions = flattenChildren(
    selectedDepartment?.children || []
  )

  const currentSubPath = joinSubPath(
    splitSubPath(employeeSubDepartment)
  )

  return (
    <tr className="border-b border-white/5 transition hover:bg-white/[0.025]">

      {/* Employee */}

      <td className="sticky right-0 z-10 border-l border-white/5 bg-[#202020] px-4 py-3">
        <div className="flex min-w-[180px] items-center gap-3">

          <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-[#d4a017]/15 text-[#f0c040]">
            <UserRound size={17} />
          </div>

          <div className="min-w-0">

            <p className="truncate text-sm font-bold text-white">
              {employee.name ||
                "بدون نام"}
            </p>

            <p className="mt-0.5 truncate text-[10px] text-gray-600">
              #{employeeId}
            </p>

          </div>
        </div>
      </td>

      {/* Organization */}

      <td className="px-4 py-3">
        <div className="flex min-w-[260px] flex-col gap-2">

          <select
            value={
              selectedDepartment?.id ||
              ""
            }
            onChange={(event) => {
              const nextDepartment =
                departments.find(
                  (department) =>
                    department.id ===
                    event.target.value
                )

              if (!nextDepartment) {
                return
              }

              onUpdateDepartment(
                employeeId,
                nextDepartment.name,
                ""
              )
            }}
            className="rounded-lg border border-white/10 bg-[#181818] px-3 py-2 text-xs text-white outline-none focus:border-[#d4a017]"
          >
            <option
              value=""
              className="bg-[#181818]"
            >
              انتخاب واحد
            </option>

            {departments.map(
              (department) => (
                <option
                  key={
                    department.id
                  }
                  value={
                    department.id
                  }
                  className="bg-[#181818]"
                >
                  {
                    department.name
                  }
                </option>
              )
            )}
          </select>

          <select
            value={
              subOptions.find(
                (option) =>
                  normalizeText(
                    option.path
                  ) ===
                  normalizeText(
                    currentSubPath
                  )
              )?.path || ""
            }
            disabled={
              !selectedDepartment ||
              subOptions.length === 0
            }
            onChange={(event) =>
              onUpdateDepartment(
                employeeId,
                selectedDepartment?.name ||
                  employeeDepartment,
                event.target.value
              )
            }
            className="rounded-lg border border-white/10 bg-[#181818] px-3 py-2 text-xs text-white outline-none focus:border-[#d4a017] disabled:cursor-not-allowed disabled:opacity-40"
          >
            <option
              value=""
              className="bg-[#181818]"
            >
              {subOptions.length
                ? "انتخاب زیرواحد"
                : "بدون زیرواحد"}
            </option>

            {subOptions.map(
              (option) => (
                <option
                  key={option.id}
                  value={option.path}
                  className="bg-[#181818]"
                >
                  {option.path}
                </option>
              )
            )}
          </select>

        </div>
      </td>

      {/* Standard / custom columns */}

      {columns.map(
        (column) => {
          const value =
            employee[
              column.key
            ] ?? ""

          const verified =
            Boolean(
              employee[
                `is_${column.key}_verified`
              ]
            )

          return (
            <td
              key={
                column.key
              }
              className="px-4 py-3"
            >
              <div className="flex min-w-[140px] items-center gap-2">

                {column.custom ? (
                  <input
                    value={
                      value
                    }
                    onChange={(
                      event
                    ) =>
                      onUpdateField(
                        employeeId,
                        column.key,
                        event.target
                          .value
                      )
                    }
                    placeholder="وارد کنید..."
                    className="min-w-0 flex-1 rounded-lg border border-white/10 bg-[#181818] px-3 py-2 text-xs text-white outline-none placeholder:text-gray-700 focus:border-[#d4a017]"
                  />
                ) : (
                  <span className="min-w-0 flex-1 truncate rounded-lg bg-[#181818] px-3 py-2 text-xs text-gray-300">
                    {formatCellValue(
                      column.key,
                      value
                    )}
                  </span>
                )}

                <button
                  onClick={() =>
                    onToggleVerification(
                      employeeId,
                      column.key
                    )
                  }
                  title={
                    verified
                      ? "تأیید شده"
                      : "تأیید نشده"
                  }
                  className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border transition ${
                    verified
                      ? "border-[#d4a017] bg-[#d4a017] text-black"
                      : "border-white/10 bg-[#181818] text-gray-700 hover:border-[#d4a017]/50 hover:text-[#d4a017]"
                  }`}
                >
                  <Check
                    size={15}
                    strokeWidth={3}
                  />
                </button>

              </div>
            </td>
          )
        }
      )}

      {/* Delete employee */}

      <td className="px-4 py-3">
        <div className="flex items-center gap-1">
          <button
            onClick={() => onEdit(employeeId)}
            title="ویرایش کارمند"
            className="flex h-8 w-8 items-center justify-center rounded-lg text-gray-600 transition hover:bg-[#d4a017]/10 hover:text-[#f0c040]"
          >
            <Pencil size={15} />
          </button>

          <button
            onClick={() => onDelete(employeeId)}
            title="حذف کارمند"
            className="flex h-8 w-8 items-center justify-center rounded-lg text-gray-600 transition hover:bg-red-500/10 hover:text-red-400"
          >
            <Trash2 size={15} />
          </button>
        </div>
      </td>

    </tr>
  )
}

/* =========================================================
   Format cell
========================================================= */

function formatCellValue(
  key,
  value
) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "—"
  }

  if (key === "salary") {
    return `${new Intl.NumberFormat(
      "fa-IR"
    ).format(
      Number(value) || 0
    )}`
  }

  return String(value)
}

/* =========================================================
   Empty employees
========================================================= */

function EmptyEmployees({
  selectedDepartment,
  selectedChild,
  showAllEmployees,
  onClearChild,
  onAddEmployee,
}) {
  return (
    <div className="flex min-h-[350px] flex-col items-center justify-center p-6 text-center">

      <div className="flex h-16 w-16 items-center justify-center rounded-2xl bg-[#d4a017]/10 text-[#d4a017]">
        <Users size={28} />
      </div>

      <h3 className="mt-4 text-base font-bold">
        کارمندی در این بخش پیدا نشد
      </h3>

      <p className="mt-2 max-w-md text-sm leading-7 text-gray-500">
        {showAllEmployees
          ? "در فایل Excel هیچ رکورد کارمندی پیدا نشد."
          : selectedChild
            ? `در زیرواحد «${selectedChild.name}» از واحد «${selectedDepartment?.name}» کارمندی وجود ندارد. اگر می‌خواهی کارمندی را به این بخش منتقل کنی، از جدول کارکنان واحد اصلی استفاده کن و واحد/زیرواحد او را تغییر بده.`
            : `در واحد «${selectedDepartment?.name}» کارمندی وجود ندارد.`}
      </p>

      {onAddEmployee && (
        <button
          onClick={onAddEmployee}
          className="mt-4 flex items-center gap-2 rounded-xl bg-[#d4a017] px-4 py-2.5 text-xs font-bold text-black hover:bg-[#f0c040]"
        >
          <UserPlus size={15} />
          افزودن کارمند
        </button>
      )}

      {selectedChild && (
        <button
          onClick={
            onClearChild
          }
          className="mt-4 rounded-xl border border-[#d4a017]/50 bg-[#202020] px-4 py-2.5 text-xs font-bold text-[#f0c040] hover:bg-[#d4a017]/10"
        >
          نمایش کل واحد
        </button>
      )}

    </div>
  )
}

/* =========================================================
   Modal
========================================================= */

function Modal({
  title,
  children,
  onClose,
}) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4 backdrop-blur-sm">

      <div
        className="w-full max-w-md rounded-2xl border border-[#d4a017]/40 bg-[#202020] p-5 shadow-2xl"
        onClick={(event) =>
          event.stopPropagation()
        }
      >

        <div className="mb-5 flex items-center justify-between">

          <h3 className="text-base font-black">
            {title}
          </h3>

          <button
            onClick={onClose}
            className="flex h-8 w-8 items-center justify-center rounded-lg text-gray-500 transition hover:bg-white/5 hover:text-white"
          >
            <X size={18} />
          </button>

        </div>

        {children}

      </div>
    </div>
  )
}

export default HumanResources