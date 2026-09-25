import { openDB } from "idb"

const DB_NAME = "hr-dashboard-db"
const DB_VERSION = 2

const EMPLOYEES_STORE = "employees"
const HR_CONFIG_STORE = "hrConfig"
const FINANCE_CONFIG_STORE = "financeConfig"
const FINANCE_RECORDS_STORE = "financeRecords"

let dbPromise = null

function getDB() {
  if (!dbPromise) {
    dbPromise = openDB(DB_NAME, DB_VERSION, {
      upgrade(db) {
        /* =====================================================
           Employees
        ===================================================== */

        if (!db.objectStoreNames.contains(EMPLOYEES_STORE)) {
          db.createObjectStore(EMPLOYEES_STORE, {
            keyPath: "id",
          })
        }


        /* =====================================================
           HR Configuration
        ===================================================== */

        if (!db.objectStoreNames.contains(HR_CONFIG_STORE)) {
          db.createObjectStore(HR_CONFIG_STORE, {
            keyPath: "id",
          })
        }


        /* =====================================================
           Finance Configuration
        ===================================================== */

        if (
          !db.objectStoreNames.contains(
            FINANCE_CONFIG_STORE
          )
        ) {
          db.createObjectStore(
            FINANCE_CONFIG_STORE,
            {
              keyPath: "id",
            }
          )
        }


        /* =====================================================
           Finance Records
        ===================================================== */

        if (
          !db.objectStoreNames.contains(
            FINANCE_RECORDS_STORE
          )
        ) {
          db.createObjectStore(
            FINANCE_RECORDS_STORE,
            {
              keyPath: "id",
            }
          )
        }
      },
    })
  }

  return dbPromise
}


/* =========================================================
   Employees
========================================================= */

export async function saveEmployees(employees) {
  const db = await getDB()

  const tx = db.transaction(
    EMPLOYEES_STORE,
    "readwrite"
  )

  const store =
    tx.objectStore(EMPLOYEES_STORE)

  await store.clear()

  for (const employee of employees) {
    const normalizedEmployee = {
      ...employee,

      id:
        employee.id !== undefined &&
        employee.id !== null &&
        employee.id !== ""
          ? String(employee.id)
          : crypto.randomUUID(),
    }

    await store.put(normalizedEmployee)
  }

  await tx.done
}


export async function getEmployees() {
  const db = await getDB()

  const employees =
    await db.getAll(
      EMPLOYEES_STORE
    )

  return employees || []
}


export async function clearEmployees() {
  const db = await getDB()

  await db.clear(
    EMPLOYEES_STORE
  )
}


/* =========================================================
   HR Configuration
========================================================= */

const DEFAULT_HR_CONFIG = {
  id: "hr-config",

  departments: [
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
  ],

  customColumns: [],

  verifiedCells: {},
}


export async function getHRConfig() {
  const db = await getDB()

  const config =
    await db.get(
      HR_CONFIG_STORE,
      "hr-config"
    )

  if (!config) {
    await saveHRConfig(
      DEFAULT_HR_CONFIG
    )

    return DEFAULT_HR_CONFIG
  }

  return {
    ...DEFAULT_HR_CONFIG,
    ...config,
  }
}


export async function saveHRConfig(
  config
) {
  const db = await getDB()

  await db.put(
    HR_CONFIG_STORE,
    {
      ...config,
      id: "hr-config",
    }
  )
}


export async function resetHRConfig() {
  await saveHRConfig(
    DEFAULT_HR_CONFIG
  )

  return DEFAULT_HR_CONFIG
}


/* =========================================================
   Finance Configuration
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
   Get Finance Configuration
========================================================= */

export async function getFinanceConfig() {
  const db = await getDB()

  let config =
    await db.get(
      FINANCE_CONFIG_STORE,
      "finance-config"
    )


  /*
    اگر نسخه قبلی FinancialAffairs
    اطلاعات مالی را داخل hrConfig ذخیره
    کرده باشد، آنها را به ساختار جدید منتقل می‌کنیم.
  */

  if (!config) {
    const oldHRConfig =
      await db.get(
        HR_CONFIG_STORE,
        "hr-config"
      )

    if (
      oldHRConfig &&
      oldHRConfig.financialSections
    ) {
      config = {
        id: "finance-config",

        income:
          oldHRConfig
            .financialSections
            .income ||
          DEFAULT_FINANCE_CONFIG.income,

        expense:
          oldHRConfig
            .financialSections
            .expense ||
          DEFAULT_FINANCE_CONFIG.expense,
      }

      await db.put(
        FINANCE_CONFIG_STORE,
        config
      )

      return config
    }


    await saveFinanceConfig(
      DEFAULT_FINANCE_CONFIG
    )

    return DEFAULT_FINANCE_CONFIG
  }


  return {
    ...DEFAULT_FINANCE_CONFIG,
    ...config,

    income:
      config.income ||
      DEFAULT_FINANCE_CONFIG.income,

    expense:
      config.expense ||
      DEFAULT_FINANCE_CONFIG.expense,
  }
}


/* =========================================================
   Save Finance Configuration
========================================================= */

export async function saveFinanceConfig(
  config
) {
  const db = await getDB()

  await db.put(
    FINANCE_CONFIG_STORE,
    {
      ...config,
      id: "finance-config",
    }
  )
}


/* =========================================================
   Reset Finance Configuration
========================================================= */

export async function resetFinanceConfig() {
  await saveFinanceConfig(
    DEFAULT_FINANCE_CONFIG
  )

  return DEFAULT_FINANCE_CONFIG
}


/* =========================================================
   Finance Records
========================================================= */

/*
  ساختار هر رکورد مالی:

  {
    id,
    date,
    type,
    sectionId,
    sectionName,
    description,
    amount,
    category,
    status,
    note
  }

  type:

  income
  expense
*/


export async function saveFinanceRecords(
  records
) {
  const db = await getDB()

  const tx = db.transaction(
    FINANCE_RECORDS_STORE,
    "readwrite"
  )

  const store =
    tx.objectStore(
      FINANCE_RECORDS_STORE
    )

  await store.clear()

  for (const record of records) {
    const normalizedRecord = {
      ...record,

      id:
        record.id !== undefined &&
        record.id !== null &&
        record.id !== ""
          ? String(record.id)
          : crypto.randomUUID(),

      amount:
        Number(record.amount) || 0,
    }

    await store.put(
      normalizedRecord
    )
  }

  await tx.done
}


/* =========================================================
   Get Finance Records
========================================================= */

export async function getFinanceRecords() {
  const db = await getDB()

  const records =
    await db.getAll(
      FINANCE_RECORDS_STORE
    )

  return records || []
}


/* =========================================================
   Add / Update One Finance Record
========================================================= */

export async function saveFinanceRecord(
  record
) {
  const db = await getDB()

  const normalizedRecord = {
    ...record,

    id:
      record.id !== undefined &&
      record.id !== null &&
      record.id !== ""
        ? String(record.id)
        : crypto.randomUUID(),

    amount:
      Number(record.amount) || 0,
  }

  await db.put(
    FINANCE_RECORDS_STORE,
    normalizedRecord
  )

  return normalizedRecord
}


/* =========================================================
   Delete One Finance Record
========================================================= */

export async function deleteFinanceRecord(
  recordId
) {
  const db = await getDB()

  await db.delete(
    FINANCE_RECORDS_STORE,
    String(recordId)
  )
}


/* =========================================================
   Clear All Finance Records
========================================================= */

export async function clearFinanceRecords() {
  const db = await getDB()

  await db.clear(
    FINANCE_RECORDS_STORE
  )
}


/* =========================================================
   Get Finance Records By Type
========================================================= */

export async function getFinanceRecordsByType(
  type
) {
  const records =
    await getFinanceRecords()

  return records.filter(
    (record) =>
      record.type === type
  )
}


/* =========================================================
   Get Finance Records By Section
========================================================= */

export async function getFinanceRecordsBySection(
  sectionId
) {
  const records =
    await getFinanceRecords()

  return records.filter(
    (record) =>
      record.sectionId === sectionId
  )
}