# Loading Plan View — Proposed Clean Design

## Core Principle

**Every row in the table = one independent Loading Plan record in Dataverse.**
No parent/child. No grouping. No capping. What you see is what gets saved and exported.

---

## Current Problems

| # | Problem | Root Cause |
|---|---------|-----------|
| 1 | 7,743 lines, hard to maintain | Old assignment table code still lingering, duplicate functions, dead code |
| 2 | Two `saveAllChanges()` functions (lines 6855 & 7484) | Copy-paste during refactoring |
| 3 | `syncAssignmentToOrderItems()`, `refreshOrderItemsDisplay()` are orphaned | Were for the old 2-table design |
| 4 | `attachAssignmentEvents()` is a no-op | Old assignment table removed but stub left |
| 5 | FG dims inline form code gone but modal functions orphaned | Partially restored, but messy |
| 6 | Container cards rendering is ~400 lines with complex volume logic | Over-engineered for the actual need |
| 7 | `splitIntoContainers()`, `mergeContainerItems()`, `validateContainerItemQuantities()` — rarely used | Built for a workflow nobody uses |
| 8 | Lock mode selector strings reference dead elements | `.ci-quantity`, old assignment table elements |

---

## Proposed Architecture: 5 Clean Modules

### Module 1: **State & Config** (~100 lines)

```
Constants:
  - API endpoints (LP, Containers, Container Items, Orders, FG Master)
  - Container type definitions (type → maxWeight, maxVolume)
  - Density constants (COOLANT=1.07, default=0.9)
  - Pallet weight constant (19.38 kg)

State:
  - CURRENT_DCL_ID        — from URL param
  - DCL_CONTAINERS         — array of {id, dataverseId, type, maxWeight}
  - DCL_CONTAINER_ITEMS    — array of {id, lpId, quantity, containerGuid, isSplitItem}
  - FG_MASTER              — array of product specs (cached on load)
  - LOADING_DATE           — for shipment window validation
  - CURRENCY_CODE          — per-DCL

No window exports except what's strictly needed for the HTML onclick handlers.
```

### Module 2: **API Layer** (~200 lines)

All Dataverse CRUD in one place. Every function is async, returns data or throws.

```
LP Row CRUD:
  - createLpRow(dclGuid, payload) → serverId
  - updateLpRow(serverId, payload)
  - deleteLpRow(serverId)
  - fetchLpRows(dclGuid) → rows[]

Container CRUD:
  - createContainer(dclGuid, type, maxWeight) → containerId
  - deleteContainer(containerId)
  - fetchContainers(dclGuid) → containers[]

Container Item CRUD:
  - createContainerItem(lpId, qty, containerGuid, isSplit) → ciId
  - updateContainerItem(ciId, fields)
  - deleteContainerItem(ciId)
  - fetchContainerItems(dclGuid) → items[]

Oracle:
  - fetchOrderNumbers(dclGuid) → orderNos[]
  - fetchOrderLines(orderNo) → lines[]
  - fetchShippedHistory(orderNo) → shipped[]

FG Master:
  - fetchFgMaster() → specs[]
  - createFgRecord(fields) → id
  - updateFgRecord(id, fields)

DCL Master:
  - patchDclTotals(dclGuid, totals)
  - patchDclComments(dclGuid, comments)
  - fetchDclStatus(dclGuid) → status
```

### Module 3: **Table Engine** (~500 lines)

The Order Items table — rendering, editing, calculations, totals.

```
Rendering:
  - renderRow(item, index) → <tr>
  - renderTable(items[]) — clear tbody, render all rows
  - renumberRows()

Row Calculation:
  - recalcRow(tr)
      UOM = parse from packaging (e.g. "24x500ml" → 12L per case)
      Pending = Order Qty − Loading Qty
      Total Liters = Loading Qty × UOM
      Net Weight = Total Liters × density
      Pallet Weight = palletized ? (pallets × 19.38) : 0
      Gross Weight = Net Weight + Pallet Weight + Loading Qty
      (skip if cell has manualOverride flag)
  - recalcAllRows()
  - recomputeTotals()
      Simple loop: count rows, sum orderQty, loadingQty, net, gross
      Per-row pending = orderQty − loadingQty
      Update DOM: #totalItems, #totalOrderQty, etc.
      Patch DCL master totals

Events (delegated on tbody):
  - Click: Split, Remove, Add Dims
  - Change: Loading Qty, Container dropdown, Palletized, Pallets, Release Status
  - Input: contentEditable cells (order no, item code, desc, packaging, etc.)
  - Blur: recalc on focus-out

Persistence:
  - saveAllChanges() — ONE function, loops dirty rows, PATCH each
  - discardChanges() — reload from server
  - createRowOnServer(tr)
  - updateRowOnServer(tr)
  - deleteRowOnServer(serverId)
```

### Module 4: **Container Manager** (~400 lines)

Everything about containers — add, remove, assign, auto-assign, render cards, render summaries.

```
Container CRUD:
  - addContainers(type, qty)
      Create N containers on server
      Add to DCL_CONTAINERS
      Re-render cards

  - removeContainer(id)
      Unassign all items from this container first
      Delete from server
      Remove from DCL_CONTAINERS
      Re-render

Container Assignment:
  - assignItemToContainer(ciId, containerGuid)
      PATCH container item
      Update local state
      Re-render cards + summaries

  - autoAssignAll()
      Get unassigned items, sort by weight descending
      FFD: for each item, find first container that fits
      Batch PATCH assignments
      Re-render

Container UI:
  - renderContainerCards()
      For each container:
        - List assigned items (item code, qty, weight)
        - Show fill bar (weight used / max weight)
        - Show utilization %
        - "Remove" button per container

  - renderContainerSummaries()
      Per-container: weight used, weight remaining, item count
      Total: all containers combined

  - rebuildDropdowns()
      Refresh <select class="assign-container"> options in all rows
      Set selected value from container item state
```

### Module 5: **Import & Split** (~250 lines)

```
Import from Oracle:
  - importFromOracle()
      Fetch order numbers for DCL
      For each order: fetch lines + shipped history
      Build items[] with computed fields
      Deduplicate against existing table rows (by orderNo+itemCode)
      appendAndPersist(newItems)
      Ensure container items exist
      Show result: "Imported X new, skipped Y existing"

Split:
  - splitRow(tr, quantities[])
      quantities = array of new quantities (e.g. [60, 40] for 2-way)
      First qty updates the original row
      For each remaining qty: clone row, set qty, insert after original
      Create server records + container items
      recalcAllRows() + recomputeTotals()
      No special styling — just normal rows

  - showSplitDialog(tr)
      Modal with 3 modes:
        Simple: "Split X units off" → [remaining, splitQty]
        Equal: "Split into N parts" → [qty/N, qty/N, ...]
        By qty: "X units per container" → [x, x, x, remainder]
      Confirm → splitRow(tr, quantities)

Load Existing:
  - loadExistingRows(dclGuid)
      Fetch saved LP rows from Dataverse
      Render into table
      Fetch container items → populate state
      Sync container dropdowns
```

---

## Simplified Data Flow

```
IMPORT FROM ORACLE
  Oracle API → computeItemData() → renderRow() → createLpRow() → ensureContainerItems()

EDIT ROW
  User edits cell → recalcRow() → mark dirty → [Save All] → updateLpRow()

SPLIT ROW
  User clicks Split → showSplitDialog() → splitRow()
    → update original (updateLpRow + updateContainerItem)
    → create new rows (createLpRow + createContainerItem) × N
    → recalcAllRows() + recomputeTotals()

ADD CONTAINER
  User selects type + qty → createContainer() × N → renderContainerCards()

ASSIGN TO CONTAINER
  User changes dropdown → assignItemToContainer() → renderCards + renderSummaries

AUTO-ASSIGN
  Click Auto-Assign → autoAssignAll() → batch PATCH → renderCards + renderSummaries

SAVE ALL
  Click Save → for each dirty row: updateLpRow() → patchDclTotals() → clear dirty flags
```

---

## What Gets Removed (Dead Code)

| Item | Lines Saved |
|------|------------|
| Duplicate `saveAllChanges()` (line 7484) | ~80 |
| `syncAssignmentToOrderItems()` | ~30 |
| `refreshOrderItemsDisplay()` | ~50 |
| `attachAssignmentEvents()` no-op | ~10 |
| `rebuildAssignmentTableWithSync()` wrapper | ~5 |
| `splitIntoContainers()` algorithm (unused) | ~150 |
| `mergeContainerItems()` (unused) | ~50 |
| `validateContainerItemQuantities()` (unused) | ~30 |
| `saveContainerAllocationsToDataverse()` stub | ~5 |
| `debugValidateAssignmentsDeep()` | ~80 |
| Dead CSS (.assignment-table, .container-dropdown, .split-indicator, etc.) | ~200 |
| Dead lock-mode selectors | ~5 |
| **Total** | **~695 lines** |

---

## Summary Table After Cleanup

| What | Current | Proposed |
|------|---------|----------|
| Total JS lines | ~7,743 | ~5,500 (estimate after dead code removal) |
| `saveAllChanges()` functions | 2 | 1 |
| recomputeTotals logic | 100 lines with grouping/capping | 30 lines straight sum |
| Split row treatment | Parent/child with badges/arrows | Normal independent row |
| Import duplicate handling | None (added recently) | Built-in dedup by orderNo+itemCode |
| Container assignment | Works | Works (with case-insensitive GUID fix) |
| FG Dimensions | Orphaned modal | Connected via Dims button |

---

## What NOT to Change

These work correctly and should stay as-is:

- `recalcRow()` formula chain (UOM → liters → net → gross)
- `allocateItemsToContainers()` FFD algorithm
- `safeAjax()` wrapper with token handling
- `showSimpleSplitPrompt()` modal UI
- Container type definitions and capacity map
- `matchFgForOutstanding()` fuzzy FG matching
- `lockLoadingPlanIfSubmitted()` read-only mode
- LP Comments auto-save with debounce
- Keyboard shortcuts (Ctrl+S, Ctrl+Z)
- Fullscreen mode
- Unsaved changes warning
