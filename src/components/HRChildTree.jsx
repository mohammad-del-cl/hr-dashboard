import { useState } from "react"

import {
  ChevronDown,
  ChevronLeft,
  FolderPlus,
  Trash2,
  UserPlus,
} from "lucide-react"

import { countNodeEmployees } from "./hrTree"

/* =========================================================
   HRChildTree
   نمایش بازگشتی زیرواحدها (عمق نامحدود).
   عدد کنار هر شاخه = تعداد کارکنان همان شاخه + همه‌ی زیرشاخه‌ها
========================================================= */

export default function HRChildTree({
  department,
  nodes,
  employees,
  activeChildId,
  parentNames = [],
  onSelect,
  onAddChild,
  onDelete,
  onAddEmployee,
}) {
  const [collapsed, setCollapsed] = useState({})

  return (
    <div className="space-y-1">
      {(nodes || []).map((node) => {
        const pathNames = [...parentNames, node.name]
        const hasChildren = (node.children || []).length > 0
        const isCollapsed = Boolean(collapsed[node.id])
        const active = activeChildId === node.id

        return (
          <div key={node.id}>
            <div
              className={`flex items-center gap-1 rounded-lg transition ${
                active ? "bg-[#d4a017]/15" : "hover:bg-white/5"
              }`}
            >
              {hasChildren ? (
                <button
                  onClick={() =>
                    setCollapsed((current) => ({
                      ...current,
                      [node.id]: !current[node.id],
                    }))
                  }
                  className="flex h-8 w-6 shrink-0 items-center justify-center text-gray-500 hover:text-[#f0c040]"
                  title={isCollapsed ? "باز کردن" : "بستن"}
                >
                  {isCollapsed ? (
                    <ChevronLeft size={14} />
                  ) : (
                    <ChevronDown size={14} />
                  )}
                </button>
              ) : (
                <span className="inline-block h-8 w-6 shrink-0" />
              )}

              <button
                onClick={() => onSelect(department, node, pathNames)}
                className={`flex min-w-0 flex-1 items-center justify-between rounded-lg px-2 py-2 text-right text-xs ${
                  active
                    ? "font-bold text-[#f0c040]"
                    : "text-gray-400 hover:text-white"
                }`}
              >
                <span className="flex min-w-0 items-center gap-2">
                  <span className="h-1.5 w-1.5 shrink-0 rounded-full bg-current" />
                  <span className="truncate">{node.name}</span>
                </span>

                <span className="mr-2 rounded-md bg-white/5 px-1.5 py-0.5 text-[10px] text-gray-600">
                  {countNodeEmployees(
                    employees,
                    department.name,
                    pathNames
                  )}
                </span>
              </button>

              {onAddEmployee && (
                <button
                  onClick={() =>
                    onAddEmployee(department, node, pathNames)
                  }
                  className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg text-gray-600 transition hover:bg-[#d4a017]/10 hover:text-[#f0c040]"
                  title="افزودن کارمند به این زیرواحد"
                >
                  <UserPlus size={14} />
                </button>
              )}

              <button
                onClick={() => onAddChild(department, node, pathNames)}
                className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg text-gray-600 transition hover:bg-[#d4a017]/10 hover:text-[#f0c040]"
                title="افزودن زیرواحد"
              >
                <FolderPlus size={14} />
              </button>

              <button
                onClick={() => onDelete(department, node, pathNames)}
                className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg text-gray-600 transition hover:bg-red-500/10 hover:text-red-400"
                title="حذف زیرواحد"
              >
                <Trash2 size={14} />
              </button>
            </div>

            {hasChildren && !isCollapsed && (
              <div className="mr-4 mt-1 border-r border-white/10 pr-2">
                <HRChildTree
                  department={department}
                  nodes={node.children}
                  employees={employees}
                  activeChildId={activeChildId}
                  parentNames={pathNames}
                  onSelect={onSelect}
                  onAddChild={onAddChild}
                  onDelete={onDelete}
                  onAddEmployee={onAddEmployee}
                />
              </div>
            )}
          </div>
        )
      })}
    </div>
  )
}
