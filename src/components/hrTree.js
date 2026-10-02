/* =========================================================
   hrTree.js
   توابع کمکی درخت واحدها در منابع انسانی (عمق نامحدود)

   ساختار ذخیره‌سازی (در hrConfig.departments):
   [{ id, name, children: [{ id, name, children: [...] }] }]

   ذخیره‌ی واحد کارمند (بدون تغییر در فیلدهای قبلی):
   - department      = نام واحد اصلی
   - sub_department  = مسیر زیرواحد با جداکننده « / »
                       مثال: "غذاخوری / رستوران / صبحانه"
   اگر فقط یک نام باشد (مثل قبل) همان زیرواحد سطح اول است،
   پس داده‌ها و اکسل‌های قبلی بدون تغییر کار می‌کنند.
========================================================= */

export const SUB_PATH_SEPARATOR = " / "

export function normText(value) {
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

/* "الف / ب > ج" → ["الف","ب","ج"] */
export function splitSubPath(value) {
  return String(value ?? "")
    .replace(/\u200c/g, " ")
    .split(/[\/\\>»]/)
    .map((part) => part.replace(/\s+/g, " ").trim())
    .filter(Boolean)
}

export function joinSubPath(parts) {
  return (parts || []).join(SUB_PATH_SEPARATOR)
}

/* پیدا کردن گره با id در هر عمق → { node, pathNames } */
export function findNodeById(nodes, id, parentNames = []) {
  const list = Array.isArray(nodes) ? nodes : []

  for (const node of list) {
    const pathNames = [...parentNames, node.name]

    if (node.id === id) {
      return { node, pathNames }
    }

    const found = findNodeById(node.children, id, pathNames)

    if (found) {
      return found
    }
  }

  return null
}

/* لیست تخت همه‌ی زیرواحدها با مسیر کامل (برای select کارمند) */
export function flattenChildren(nodes, parentNames = []) {
  const result = []

  for (const node of Array.isArray(nodes) ? nodes : []) {
    const pathNames = [...parentNames, node.name]

    result.push({
      id: node.id,
      name: node.name,
      pathNames,
      path: joinSubPath(pathNames),
      depth: pathNames.length,
    })

    result.push(...flattenChildren(node.children, pathNames))
  }

  return result
}

/*
  آیا کارمند داخل این شاخه (یا هر زیرشاخه‌ی آن) است؟
  pathNames خالی = کل واحد اصلی
*/
export function employeeInBranch(employee, departmentName, pathNames = []) {
  if (normText(employee?.department) !== normText(departmentName)) {
    return false
  }

  if (!pathNames.length) {
    return true
  }

  const employeePath = splitSubPath(employee?.sub_department)

  if (employeePath.length < pathNames.length) {
    return false
  }

  return pathNames.every(
    (name, index) => normText(employeePath[index]) === normText(name)
  )
}

export function countNodeEmployees(employees, departmentName, pathNames = []) {
  return (employees || []).filter((employee) =>
    employeeInBranch(employee, departmentName, pathNames)
  ).length
}

/* افزودن گره به والد (parentNodeId = null یعنی مستقیم زیر واحد اصلی) */
export function addChildNode(departments, departmentId, parentNodeId, newNode) {
  const insert = (nodes) =>
    (nodes || []).map((node) =>
      node.id === parentNodeId
        ? { ...node, children: [...(node.children || []), newNode] }
        : { ...node, children: insert(node.children) }
    )

  return departments.map((department) => {
    if (department.id !== departmentId) {
      return department
    }

    if (!parentNodeId) {
      return {
        ...department,
        children: [...(department.children || []), newNode],
      }
    }

    return { ...department, children: insert(department.children) }
  })
}

/* حذف گره از هر عمق */
export function removeChildNode(departments, departmentId, nodeId) {
  const remove = (nodes) =>
    (nodes || [])
      .filter((node) => node.id !== nodeId)
      .map((node) => ({ ...node, children: remove(node.children) }))

  return departments.map((department) =>
    department.id === departmentId
      ? { ...department, children: remove(department.children) }
      : department
  )
}

/*
  ساختن شاخه‌های ناموجود از روی مسیر کارمندان (ورود از Excel)
  createId همان createId داخل HumanResources.jsx است.
*/
export function mergeEmployeesIntoTreeDeep(currentTree, employees, createId) {
  const cloneNodes = (nodes) =>
    (Array.isArray(nodes) ? nodes : []).map((node) => ({
      ...node,
      children: cloneNodes(node.children),
    }))

  const result = (currentTree || []).map((department) => ({
    ...department,
    children: cloneNodes(department.children),
  }))

  for (const employee of employees || []) {
    const departmentName = String(employee?.department || "").trim()

    if (!departmentName) {
      continue
    }

    let department = result.find(
      (item) => normText(item.name) === normText(departmentName)
    )

    if (!department) {
      department = {
        id: createId("department", departmentName),
        name: departmentName,
        children: [],
      }

      result.push(department)
    }

    const parts = splitSubPath(employee?.sub_department)

    let siblings = department.children

    parts.forEach((part, index) => {
      let node = siblings.find(
        (child) => normText(child.name) === normText(part)
      )

      if (!node) {
        node = {
          id: createId(
            "subdepartment",
            `${departmentName}-${parts.slice(0, index + 1).join("-")}`
          ),
          name: part,
          children: [],
        }

        siblings.push(node)
      }

      siblings = node.children
    })
  }

  return result
}
