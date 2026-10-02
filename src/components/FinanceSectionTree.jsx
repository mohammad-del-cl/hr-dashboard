import { useEffect, useMemo, useState } from "react"

import {
  Building2,
  ChevronDown,
  ChevronLeft,
  FileText,
  FolderPlus,
  Trash2,
} from "lucide-react"

import { buildSectionTree, normalizeId } from "./financeTree"


/* =========================================================
   Section Tree (نمایش درختی بخش‌ها و زیربخش‌ها)
========================================================= */

export default function SectionTree({
  sections,
  selectedId,
  onSelect,
  onAddChild,
  onDelete,
  allLabel,
}) {
  const [collapsed, setCollapsed] = useState({})

  const tree = useMemo(
    () => buildSectionTree(sections),
    [sections]
  )

  /*
    وقتی بخشی انتخاب می‌شود (مثلاً زیربخش تازه ساخته‌شده)،
    والدهای آن باز می‌شوند تا انتخاب دیده شود.
  */
  useEffect(() => {
    const id = normalizeId(selectedId)

    if (!id) {
      return
    }

    const list = Array.isArray(sections) ? sections : []
    const ancestors = []
    const seen = new Set([id])

    let current = list.find(
      (section) => normalizeId(section.id) === id
    )

    while (current) {
      const parentId = normalizeId(current.parentId)

      if (!parentId || seen.has(parentId)) {
        break
      }

      seen.add(parentId)
      ancestors.push(parentId)

      current = list.find(
        (section) => normalizeId(section.id) === parentId
      )
    }

    if (ancestors.length === 0) {
      return
    }

    setCollapsed((previous) => {
      if (!ancestors.some((item) => previous[item])) {
        return previous
      }

      const next = { ...previous }

      ancestors.forEach((item) => {
        delete next[item]
      })

      return next
    })
  }, [selectedId, sections])

  function toggle(id) {
    setCollapsed((previous) => ({
      ...previous,
      [id]: !previous[id],
    }))
  }

  function renderNode(node, depth) {
    const id = normalizeId(node.id)
    const hasChildren = node.children.length > 0
    const isCollapsed = Boolean(collapsed[id])
    const isSelected = normalizeId(selectedId) === id

    return (
      <div key={id}>
        <div
          style={{ marginRight: depth * 16 }}
          className={`group flex items-center gap-1 rounded-xl border transition ${
            isSelected
              ? "border-[#d4a017] bg-[#d4a017]/10"
              : "border-transparent hover:border-white/10 hover:bg-white/5"
          }`}
        >
          {hasChildren ? (
            <button
              onClick={() => toggle(id)}
              className="mr-1 flex h-7 w-7 shrink-0 items-center justify-center rounded-lg text-gray-500 transition hover:text-[#f0c040]"
              title={isCollapsed ? "باز کردن" : "بستن"}
            >
              {isCollapsed ? (
                <ChevronLeft size={16} />
              ) : (
                <ChevronDown size={16} />
              )}
            </button>
          ) : (
            <span className="mr-1 inline-block h-7 w-7 shrink-0" />
          )}

          <button
            onClick={() => onSelect(id)}
            className="flex min-w-0 flex-1 items-center gap-2 px-1 py-2.5 text-right"
          >
            <Building2
              size={16}
              className={
                isSelected
                  ? "shrink-0 text-[#f0c040]"
                  : "shrink-0 text-gray-500"
              }
            />

            <span className="truncate text-sm">
              {node.name}
            </span>
          </button>

          {onAddChild && (
            <button
              onClick={() => onAddChild(node)}
              className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg text-gray-600 transition hover:bg-[#d4a017]/10 hover:text-[#f0c040]"
              title="افزودن زیربخش"
            >
              <FolderPlus size={15} />
            </button>
          )}

          {onDelete && (
            <button
              onClick={() => onDelete(node)}
              className="ml-1 flex h-8 w-8 shrink-0 items-center justify-center rounded-lg text-gray-600 transition hover:bg-red-500/10 hover:text-red-400"
              title="حذف بخش"
            >
              <Trash2 size={15} />
            </button>
          )}
        </div>

        {hasChildren && !isCollapsed && (
          <div className="mt-1 space-y-1">
            {node.children.map((child) =>
              renderNode(child, depth + 1)
            )}
          </div>
        )}
      </div>
    )
  }

  return (
    <div className="space-y-1">
      {allLabel && (
        <button
          onClick={() => onSelect("")}
          className={`flex w-full items-center gap-2 rounded-xl border px-3 py-2.5 text-right text-sm transition ${
            !normalizeId(selectedId)
              ? "border-[#d4a017] bg-[#d4a017]/10"
              : "border-transparent hover:border-white/10 hover:bg-white/5"
          }`}
        >
          <FileText
            size={16}
            className={
              !normalizeId(selectedId)
                ? "text-[#f0c040]"
                : "text-gray-500"
            }
          />
          {allLabel}
        </button>
      )}

      {tree.map((node) => renderNode(node, 0))}

      {tree.length === 0 && (
        <p className="px-2 py-4 text-center text-xs text-gray-600">
          هنوز بخشی ثبت نشده است.
        </p>
      )}
    </div>
  )
}


