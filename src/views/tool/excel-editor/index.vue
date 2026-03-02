<template>
  <div class="excel-editor-page">
    <section class="hero-card">
      <div>
        <p class="eyebrow">Format Preserving Editor</p>
        <h1>在线 Excel 编辑</h1>
        <p class="hero-desc">
          当前版本会基于后端原始工作簿直接修改单元格内容，并尽量保留合并单元格、底色、字体、列宽、行高等原始格式。
        </p>
      </div>
      <div class="hero-actions">
        <el-button @click="goToFileCenter">返回文件管理</el-button>
        <el-button :loading="loadingServerFile" :disabled="!currentFileName" @click="reloadWorkbook">重新加载</el-button>
        <el-button type="success" :loading="saving" :disabled="!hasWorkbook" @click="saveToServer">在线保存</el-button>
        <el-button :disabled="!currentFileName" @click="downloadServerWorkbook">下载当前文件</el-button>
      </div>
    </section>

    <section class="summary-grid">
      <article class="summary-card">
        <span class="summary-label">当前文件</span>
        <strong>{{ currentFileName || '未选择' }}</strong>
        <span>{{ workbookName }}</span>
      </article>
      <article class="summary-card">
        <span class="summary-label">工作表</span>
        <strong>{{ currentSheet ? currentSheet.name : '未加载' }}</strong>
        <span>{{ sheets.length }} 个工作表</span>
      </article>
      <article class="summary-card">
        <span class="summary-label">保存状态</span>
        <strong>{{ isDirty ? `有 ${dirtyCount} 处未保存修改` : '已同步' }}</strong>
        <span>{{ isDirty ? '保存时会提交改动单元格和合并布局' : '当前内容与服务器文件一致' }}</span>
      </article>
    </section>

    <section class="editor-shell" v-loading="pageLoading">
      <template v-if="hasWorkbook">
        <div class="toolbar">
          <div class="toolbar-left">
            <el-tag type="info" effect="light">编辑文件：{{ currentFileName }}</el-tag>
            <el-tag :type="isDirty ? 'warning' : 'success'" effect="light">{{ isDirty ? '未保存' : '已保存' }}</el-tag>
          </div>
          <div class="toolbar-right">
            <span class="selection-tip">{{ selectionLabel }}</span>
            <el-button size="small" :disabled="!canMergeSelection" @click="mergeSelectedCells">合并单元格</el-button>
            <el-button size="small" :disabled="!canUnmergeSelection" @click="unmergeSelectedCells">取消合并</el-button>
          </div>
        </div>

        <el-tabs v-model="activeSheetName" class="sheet-tabs">
          <el-tab-pane
            v-for="sheet in sheets"
            :key="sheet.name"
            :label="sheet.name"
            :name="sheet.name"
          />
        </el-tabs>

        <div v-if="currentSheet" class="grid-wrap">
          <div class="sheet-stage" :style="buildSheetStageStyle(currentSheet)">
            <div v-if="hasBackgroundImages(currentSheet)" class="sheet-image-layer sheet-image-layer-background">
              <div
                v-for="(image, imageIndex) in listBackgroundImages(currentSheet)"
                :key="buildImageKey(image, imageIndex)"
                class="sheet-image-item background"
                :style="buildImageStyle(image)"
              >
                <img class="sheet-image" :src="image.src" :alt="image.description || 'background image'" draggable="false" />
              </div>
            </div>

            <table class="sheet-grid">
              <thead>
                <tr>
                  <th class="corner-header"></th>
                  <th
                    v-for="columnIndex in currentSheet.maxColumnCount"
                    :key="`header-col-${columnIndex}`"
                    class="column-header"
                    :style="buildColumnHeaderStyle(currentSheet, columnIndex - 1)"
                  >
                    <span class="column-label">{{ toColumnLabel(columnIndex - 1) }}</span>
                    <span
                      class="resize-handle col-resize-handle"
                      @mousedown="startColumnResize(columnIndex - 1, $event)"
                    ></span>
                  </th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="row in currentSheet.rows" :key="`row-${row.rowIndex}`">
                  <td class="row-index" :style="buildRowIndexStyle(row)">
                    {{ row.rowIndex + 1 }}
                    <span
                      class="resize-handle row-resize-handle"
                      @mousedown="startRowResize(row.rowIndex, $event)"
                    ></span>
                  </td>
                  <td
                    v-for="cell in row.cells"
                    :key="`cell-${row.rowIndex}-${cell.colIndex}`"
                    class="sheet-cell"
                    :class="buildCellClass(cell)"
                    :colspan="cell.colSpan"
                    :rowspan="cell.rowSpan"
                    :style="buildCellTdStyle(cell)"
                    @click="handleCellSelection(cell, $event)"
                  >
                    <template v-if="hasValidationOptions(cell)">
                      <select
                        v-model="cell.value"
                        class="cell-input cell-select"
                        :class="{ dirty: cell.dirty, formula: cell.formula }"
                        :style="buildCellInputStyle(cell)"
                        @change="handleCellInput(cell)"
                      >
                        <option v-if="cell.validationAllowBlank" value=""></option>
                        <option
                          v-if="shouldKeepCurrentValidationValue(cell)"
                          :value="cell.value"
                        >
                          {{ cell.value }}
                        </option>
                        <option
                          v-for="option in cell.validationOptions"
                          :key="`list-${cell.rowIndex}-${cell.colIndex}-${option}`"
                          :value="option"
                        >
                          {{ option }}
                        </option>
                      </select>
                    </template>
                    <textarea
                      v-else
                      v-model="cell.value"
                      class="cell-input"
                      :class="{ dirty: cell.dirty, formula: cell.formula }"
                      :style="buildCellInputStyle(cell)"
                      :maxlength="MAX_CELL_TEXT_LENGTH"
                      :wrap="cell.style?.whiteSpace === 'pre-wrap' ? 'soft' : 'off'"
                      spellcheck="false"
                      @input="handleCellInput(cell)"
                    />
                  </td>
                </tr>
              </tbody>
            </table>

            <div v-if="hasAnchoredImages(currentSheet)" class="sheet-image-layer sheet-image-layer-foreground">
              <div
                v-for="(image, imageIndex) in listAnchoredImages(currentSheet)"
                :key="buildImageKey(image, imageIndex)"
                class="sheet-image-item anchored"
                :style="buildImageStyle(image)"
              >
                <img class="sheet-image" :src="image.src" :alt="image.description || 'embedded image'" draggable="false" />
              </div>
            </div>
          </div>
        </div>
      </template>

      <el-empty v-else description="还没有加载可编辑的文件，请先去文件管理页选择一份已上传 Excel。" />
    </section>
  </div>
</template>

<script setup>
import { computed, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import { saveAs } from 'file-saver'
import { applyExcelWorkbookChanges, getExcelEditorContent, getExcelEditorInfo, getExcelWorkbookView } from '@/api/tool/excelEditor'

const DEFAULT_ROW_HEIGHT = 28
const MIN_ROW_HEIGHT = 24
const DEFAULT_COLUMN_WIDTH = 96
const MIN_COLUMN_WIDTH = 72
const DEFAULT_CELL_HEIGHT = 32
const COLUMN_HEADER_HEIGHT = 34
const ROW_INDEX_WIDTH = 56
const MAX_CELL_TEXT_LENGTH = 32767

const route = useRoute()
const router = useRouter()

const workbookName = ref('未加载')
const sheets = ref([])
const activeSheetName = ref('')
const currentFileName = ref('')
const pageLoading = ref(false)
const loadingServerFile = ref(false)
const saving = ref(false)
const isDirty = ref(false)
const dirtyCount = ref(0)
const revertingRoute = ref(false)
const selection = ref(null)
const showedCellLimitHint = ref(false)
const columnResizeState = ref(null)
const rowResizeState = ref(null)

const currentSheet = computed(() => sheets.value.find((sheet) => sheet.name === activeSheetName.value) || null)
const hasWorkbook = computed(() => sheets.value.length > 0)
const activeSelection = computed(() => {
  if (!selection.value || selection.value.sheetName !== activeSheetName.value) {
    return null
  }
  return selection.value
})
const selectionLabel = computed(() => {
  if (!activeSelection.value) {
    return '点击单元格选择，Shift + 点击扩展选区'
  }
  return `选区 ${formatSelectionRange(activeSelection.value)}`
})
const canMergeSelection = computed(() => {
  if (!currentSheet.value || !activeSelection.value) {
    return false
  }
  return getSelectionArea(activeSelection.value) > 1
})
const canUnmergeSelection = computed(() => {
  if (!currentSheet.value || !activeSelection.value) {
    return false
  }
  return currentSheet.value.cells.some((cell) => isMergedCell(cell) && rangesIntersect(activeSelection.value, getCellRange(cell)))
})

onMounted(async () => {
  await initializeFromRoute()
})

onBeforeUnmount(() => {
  clearResizeState()
})

watch(
  () => route.query.fileName,
  async (nextValue, previousValue) => {
    if (revertingRoute.value) {
      revertingRoute.value = false
      return
    }
    if (normalizeQueryFileName(nextValue) === normalizeQueryFileName(previousValue)) {
      return
    }
    const confirmed = await confirmReplaceWorkbook()
    if (!confirmed) {
      revertingRoute.value = true
      router.replace({
        path: '/excel-editor/index',
        query: currentFileName.value ? { fileName: currentFileName.value } : {}
      })
      return
    }
    await initializeFromRoute()
  }
)

watch(activeSheetName, () => {
  clearResizeState()
  clearSelection()
})

async function initializeFromRoute() {
  const fileName = normalizeQueryFileName(route.query.fileName)
  if (fileName) {
    await loadWorkbookByFileName(fileName, true)
    return
  }
  try {
    const response = await getExcelEditorInfo()
    const fallbackFileName = response?.data?.currentFileName || ''
    if (!fallbackFileName) {
      clearWorkbook()
      return
    }
    await loadWorkbookByFileName(fallbackFileName, true)
  } catch (error) {
    clearWorkbook()
    ElMessage.error(error?.message || '读取文件信息失败')
  }
}

async function confirmReplaceWorkbook() {
  if (!hasWorkbook.value || !isDirty.value) {
    return true
  }
  try {
    await ElMessageBox.confirm('当前有未保存修改，继续切换会丢失这些改动，是否继续？', '确认切换', {
      type: 'warning',
      confirmButtonText: '继续',
      cancelButtonText: '取消'
    })
    return true
  } catch {
    return false
  }
}

async function loadWorkbookByFileName(fileName, skipConfirm = false) {
  if (!fileName) {
    clearWorkbook()
    return
  }
  if (!skipConfirm) {
    const confirmed = await confirmReplaceWorkbook()
    if (!confirmed) {
      return
    }
  }
  loadingServerFile.value = true
  pageLoading.value = true
  try {
    const response = await getExcelWorkbookView(fileName)
    applyWorkbookView(response?.data || {})
    ElMessage.success('文件已加载')
  } catch (error) {
    ElMessage.error(error?.message || '加载文件失败')
  } finally {
    loadingServerFile.value = false
    pageLoading.value = false
  }
}

function applyWorkbookView(data) {
  const nextSheets = Array.isArray(data?.sheets) ? data.sheets.map(normalizeSheet) : []
  sheets.value = nextSheets
  activeSheetName.value = nextSheets[0]?.name || ''
  currentFileName.value = data?.currentFileName || data?.fileName || ''
  workbookName.value = data?.fileName || currentFileName.value || '未加载'
  clearSelection()
  refreshDirtyState()
}

function normalizeSheet(sheet) {
  const name = sheet?.name || `Sheet${Math.random().toString(36).slice(2, 6)}`
  const rowHeights = []
  const columnWidths = []
  const cells = []
  const images = normalizeImages(sheet?.images)
  const validations = normalizeValidations(sheet?.validations)
  const freezePane = normalizeFreezePane(sheet?.freezePane)
  let rowCount = Math.max(Number(sheet?.rowCount ?? 0), 1)
  let maxColumnCount = Math.max(Number(sheet?.maxColumnCount ?? 0), 1)
  const sourceRows = Array.isArray(sheet?.rows) ? sheet.rows : []

  for (const sourceRow of sourceRows) {
    const rowIndex = Math.max(Number(sourceRow?.rowIndex ?? 0), 0)
    rowCount = Math.max(rowCount, rowIndex + 1)
    ensureArraySize(rowHeights, rowIndex + 1, DEFAULT_ROW_HEIGHT)
    rowHeights[rowIndex] = normalizeRowHeight(sourceRow?.heightPx)

    const sourceCells = Array.isArray(sourceRow?.cells) ? sourceRow.cells : []
    for (const sourceCell of sourceCells) {
      const cell = normalizeCell(sourceCell)
      cells.push(cell)
      rowCount = Math.max(rowCount, cell.rowIndex + cell.rowSpan)
      maxColumnCount = Math.max(maxColumnCount, cell.colIndex + cell.colSpan)
      ensureArraySize(columnWidths, cell.colIndex + cell.colSpan, null)
      applyColumnWidthHint(columnWidths, cell)
    }
  }

  ensureArraySize(rowHeights, rowCount, DEFAULT_ROW_HEIGHT)
  ensureArraySize(columnWidths, maxColumnCount, null)
  for (let index = 0; index < columnWidths.length; index += 1) {
    if (columnWidths[index] == null) {
      columnWidths[index] = DEFAULT_COLUMN_WIDTH
    }
  }

  const normalized = {
    name,
    rowCount,
    maxColumnCount,
    rowHeights,
    columnWidths,
    freezePane,
    validations,
    images,
    cells: dedupeCells(cells),
    rows: [],
    originalMergeKeys: [],
    originalRowHeights: [],
    originalColumnWidths: []
  }

  rebuildSheetRows(normalized)
  normalized.originalMergeKeys = collectSheetMergeKeys(normalized)
  normalized.originalRowHeights = normalized.rowHeights.map((height) => Math.round(Number(height)))
  normalized.originalColumnWidths = normalized.columnWidths.map((width) => Math.round(Number(width)))
  return normalized
}

function normalizeCell(cell) {
  const value = cell?.value == null ? '' : String(cell.value)
  return {
    rowIndex: Math.max(Number(cell?.rowIndex ?? 0), 0),
    colIndex: Math.max(Number(cell?.colIndex ?? 0), 0),
    rowSpan: Math.max(Number(cell?.rowSpan ?? 1), 1),
    colSpan: Math.max(Number(cell?.colSpan ?? 1), 1),
    widthPx: Math.max(Number(cell?.widthPx ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH),
    heightPx: Math.max(Number(cell?.heightPx ?? DEFAULT_CELL_HEIGHT), MIN_ROW_HEIGHT),
    value,
    originalValue: value,
    displayValue: cell?.displayValue == null ? '' : String(cell.displayValue),
    formula: Boolean(cell?.formula),
    dirty: false,
    forceDirty: false,
    validationOptions: [],
    validationAllowBlank: true,
    style: typeof cell?.style === 'object' && cell.style ? { ...cell.style } : {}
  }
}

function normalizeFreezePane(freezePane) {
  return {
    xSplit: Math.max(Number(freezePane?.xSplit ?? 0), 0),
    ySplit: Math.max(Number(freezePane?.ySplit ?? 0), 0),
    leftColumn: Math.max(Number(freezePane?.leftColumn ?? 0), 0),
    topRow: Math.max(Number(freezePane?.topRow ?? 0), 0)
  }
}

function normalizeValidations(validations) {
  const source = Array.isArray(validations) ? validations : []
  return source
    .map((item) => ({
      firstRow: Math.max(Number(item?.firstRow ?? 0), 0),
      lastRow: Math.max(Number(item?.lastRow ?? 0), 0),
      firstColumn: Math.max(Number(item?.firstColumn ?? 0), 0),
      lastColumn: Math.max(Number(item?.lastColumn ?? 0), 0),
      allowBlank: item?.allowBlank !== false,
      options: Array.isArray(item?.options) ? item.options.map((option) => String(option)) : []
    }))
    .filter((item) => item.options.length > 0)
}

function normalizeImages(images) {
  const source = Array.isArray(images) ? images : []
  return source
    .map((item) => ({
      kind: item?.kind === 'background' ? 'background' : 'anchored',
      row1: Math.max(Number(item?.row1 ?? 0), 0),
      col1: Math.max(Number(item?.col1 ?? 0), 0),
      row2: Math.max(Number(item?.row2 ?? 0), 0),
      col2: Math.max(Number(item?.col2 ?? 0), 0),
      dx1Px: Math.max(Number(item?.dx1Px ?? 0), 0),
      dx2Px: Math.max(Number(item?.dx2Px ?? 0), 0),
      dy1Px: Math.max(Number(item?.dy1Px ?? 0), 0),
      dy2Px: Math.max(Number(item?.dy2Px ?? 0), 0),
      widthPx: Math.max(Number(item?.widthPx ?? 0), 1),
      heightPx: Math.max(Number(item?.heightPx ?? 0), 1),
      src: item?.src ? String(item.src) : '',
      mimeType: item?.mimeType ? String(item.mimeType) : '',
      extension: item?.extension ? String(item.extension) : '',
      description: item?.description ? String(item.description) : ''
    }))
    .filter((item) => item.src)
}

function normalizeRowHeight(value) {
  return Math.max(Number(value ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT)
}

function applyColumnWidthHint(columnWidths, cell) {
  if (cell.colSpan === 1) {
    columnWidths[cell.colIndex] = cell.widthPx
    return
  }
  const fallbackWidth = Math.max(Math.round(cell.widthPx / cell.colSpan), MIN_COLUMN_WIDTH)
  for (let offset = 0; offset < cell.colSpan; offset += 1) {
    const columnIndex = cell.colIndex + offset
    if (columnWidths[columnIndex] == null) {
      columnWidths[columnIndex] = fallbackWidth
    }
  }
}

function clearWorkbook() {
  workbookName.value = '未加载'
  sheets.value = []
  activeSheetName.value = ''
  currentFileName.value = ''
  clearSelection()
  isDirty.value = false
  dirtyCount.value = 0
}

function clearSelection() {
  selection.value = null
}

function handleCellInput(cell) {
  const normalized = normalizeValue(cell.value)
  if (normalized.length > MAX_CELL_TEXT_LENGTH) {
    cell.value = normalized.slice(0, MAX_CELL_TEXT_LENGTH)
    if (!showedCellLimitHint.value) {
      ElMessage.warning(`单元格内容已超过 Excel 限制，已截断到 ${MAX_CELL_TEXT_LENGTH} 字符`)
      showedCellLimitHint.value = true
    }
  }
  cell.dirty = cell.forceDirty || normalizeValue(cell.value) !== normalizeValue(cell.originalValue)
  refreshDirtyState()
}

function handleCellSelection(cell, event) {
  if (!currentSheet.value) {
    return
  }
  const cellRange = getCellRange(cell)
  if (event?.shiftKey && activeSelection.value) {
    const anchor = activeSelection.value
    selection.value = {
      sheetName: currentSheet.value.name,
      anchorStartRow: anchor.anchorStartRow,
      anchorEndRow: anchor.anchorEndRow,
      anchorStartCol: anchor.anchorStartCol,
      anchorEndCol: anchor.anchorEndCol,
      startRow: Math.min(anchor.anchorStartRow, cellRange.startRow),
      endRow: Math.max(anchor.anchorEndRow, cellRange.endRow),
      startCol: Math.min(anchor.anchorStartCol, cellRange.startCol),
      endCol: Math.max(anchor.anchorEndCol, cellRange.endCol)
    }
    return
  }
  selection.value = {
    sheetName: currentSheet.value.name,
    anchorStartRow: cellRange.startRow,
    anchorEndRow: cellRange.endRow,
    anchorStartCol: cellRange.startCol,
    anchorEndCol: cellRange.endCol,
    startRow: cellRange.startRow,
    endRow: cellRange.endRow,
    startCol: cellRange.startCol,
    endCol: cellRange.endCol
  }
}

function isCellSelected(cell) {
  if (!activeSelection.value) {
    return false
  }
  return rangesIntersect(activeSelection.value, getCellRange(cell))
}

async function mergeSelectedCells() {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    ElMessage.warning('请先选择要合并的单元格')
    return
  }
  if (getSelectionArea(range) < 2) {
    ElMessage.info('至少选择两个单元格后才能合并')
    return
  }

  const cellsToClear = sheet.cells.filter((cell) => {
    if (!isCellTopLeftInsideRange(cell, range)) {
      return false
    }
    if (cell.rowIndex === range.startRow && cell.colIndex === range.startCol) {
      return false
    }
    return normalizeValue(cell.value) !== ''
  })
  if (cellsToClear.length) {
    try {
      await ElMessageBox.confirm('合并后仅保留左上角内容，其余内容会被清空，是否继续？', '确认合并', {
        type: 'warning',
        confirmButtonText: '继续合并',
        cancelButtonText: '取消'
      })
    } catch {
      return
    }
  }

  expandMergedCellsToSingles(sheet, range)
  const topLeftCell = findCellAt(sheet, range.startRow, range.startCol)
  if (!topLeftCell) {
    ElMessage.warning('选区无效，无法执行合并')
    return
  }

  const nextCells = []
  for (const cell of sheet.cells) {
    if (isCellTopLeftInsideRange(cell, range)) {
      if (cell === topLeftCell) {
        nextCells.push(cell)
      }
      continue
    }
    nextCells.push(cell)
  }

  topLeftCell.rowSpan = range.endRow - range.startRow + 1
  topLeftCell.colSpan = range.endCol - range.startCol + 1
  topLeftCell.forceDirty = false
  topLeftCell.dirty = normalizeValue(topLeftCell.value) !== normalizeValue(topLeftCell.originalValue)

  sheet.cells = dedupeCells(nextCells)
  rebuildSheetRows(sheet)
  selection.value = {
    sheetName: sheet.name,
    anchorStartRow: range.startRow,
    anchorEndRow: range.endRow,
    anchorStartCol: range.startCol,
    anchorEndCol: range.endCol,
    startRow: range.startRow,
    endRow: range.endRow,
    startCol: range.startCol,
    endCol: range.endCol
  }
  refreshDirtyState()
  ElMessage.success('已更新合并区域，保存后会写回文件')
}

function unmergeSelectedCells() {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    ElMessage.warning('请先选择已合并的单元格')
    return
  }

  const targetKeys = new Set(
    sheet.cells
      .filter((cell) => isMergedCell(cell) && rangesIntersect(range, getCellRange(cell)))
      .map((cell) => buildCellKey(cell.rowIndex, cell.colIndex))
  )

  if (!targetKeys.size) {
    ElMessage.info('当前选区内没有已合并单元格')
    return
  }

  const nextCells = []
  for (const cell of sheet.cells) {
    const cellKey = buildCellKey(cell.rowIndex, cell.colIndex)
    if (!targetKeys.has(cellKey)) {
      nextCells.push(cell)
      continue
    }
    nextCells.push(...splitMergedCell(sheet, cell, true))
  }

  sheet.cells = dedupeCells(nextCells)
  rebuildSheetRows(sheet)
  refreshDirtyState()
  ElMessage.success('已拆分合并区域，保存后会写回文件')
}

function expandMergedCellsToSingles(sheet, range) {
  const targetKeys = new Set(
    sheet.cells
      .filter((cell) => isMergedCell(cell) && rangesIntersect(range, getCellRange(cell)))
      .map((cell) => buildCellKey(cell.rowIndex, cell.colIndex))
  )

  if (!targetKeys.size) {
    return
  }

  const nextCells = []
  for (const cell of sheet.cells) {
    const cellKey = buildCellKey(cell.rowIndex, cell.colIndex)
    if (!targetKeys.has(cellKey)) {
      nextCells.push(cell)
      continue
    }
    nextCells.push(...splitMergedCell(sheet, cell, true))
  }

  sheet.cells = dedupeCells(nextCells)
}

function splitMergedCell(sheet, cell, forceDirty = false) {
  if (!isMergedCell(cell)) {
    return [cell]
  }

  const items = []
  for (let rowOffset = 0; rowOffset < cell.rowSpan; rowOffset += 1) {
    for (let colOffset = 0; colOffset < cell.colSpan; colOffset += 1) {
      const rowIndex = cell.rowIndex + rowOffset
      const colIndex = cell.colIndex + colOffset
      if (rowOffset === 0 && colOffset === 0) {
        cell.rowSpan = 1
        cell.colSpan = 1
        syncCellDimensions(sheet, cell)
        items.push(cell)
        continue
      }
      items.push(createSyntheticCell(sheet, rowIndex, colIndex, cell.style, forceDirty))
    }
  }
  return items
}

function createSyntheticCell(sheet, rowIndex, colIndex, style, forceDirty) {
  const value = ''
  return {
    rowIndex,
    colIndex,
    rowSpan: 1,
    colSpan: 1,
    widthPx: resolveRangeWidth(sheet, colIndex, 1),
    heightPx: resolveRangeHeight(sheet, rowIndex, 1),
    value,
    originalValue: value,
    displayValue: value,
    formula: false,
    dirty: Boolean(forceDirty),
    forceDirty: Boolean(forceDirty),
    style: style ? { ...style } : {}
  }
}

function findCellAt(sheet, rowIndex, colIndex) {
  return sheet.cells.find((cell) => cell.rowIndex === rowIndex && cell.colIndex === colIndex) || null
}

function isCellTopLeftInsideRange(cell, range) {
  return cell.rowIndex >= range.startRow
    && cell.rowIndex <= range.endRow
    && cell.colIndex >= range.startCol
    && cell.colIndex <= range.endCol
}

function refreshDirtyState() {
  let cellChanges = 0
  for (const sheet of sheets.value) {
    for (const cell of sheet.cells) {
      if (cell.dirty) {
        cellChanges += 1
      }
    }
  }
  const mergeChanges = countMergeChanges()
  const dimensionChanges = countDimensionChanges()
  dirtyCount.value = cellChanges + mergeChanges + dimensionChanges
  isDirty.value = dirtyCount.value > 0
}

function countMergeChanges() {
  let count = 0
  for (const sheet of sheets.value) {
    const current = new Set(collectSheetMergeKeys(sheet))
    const original = new Set(sheet.originalMergeKeys || [])
    for (const key of current) {
      if (!original.has(key)) {
        count += 1
      }
    }
    for (const key of original) {
      if (!current.has(key)) {
        count += 1
      }
    }
  }
  return count
}

function collectSheetMergeKeys(sheet) {
  return sheet.cells
    .filter((cell) => isMergedCell(cell))
    .map((cell) => `${cell.rowIndex}:${cell.colIndex}:${cell.rowSpan}:${cell.colSpan}`)
    .sort()
}

function collectChanges() {
  const changes = []
  for (const sheet of sheets.value) {
    for (const cell of sheet.cells) {
      if (!cell.dirty) {
        continue
      }
      changes.push({
        sheetName: sheet.name,
        rowIndex: cell.rowIndex,
        colIndex: cell.colIndex,
        value: normalizeValue(cell.value)
      })
    }
  }
  return changes
}

function collectMergeRegions() {
  const regions = []
  for (const sheet of sheets.value) {
    for (const cell of sheet.cells) {
      if (!isMergedCell(cell)) {
        continue
      }
      regions.push({
        sheetName: sheet.name,
        firstRow: cell.rowIndex,
        lastRow: cell.rowIndex + cell.rowSpan - 1,
        firstColumn: cell.colIndex,
        lastColumn: cell.colIndex + cell.colSpan - 1
      })
    }
  }
  return regions
}

function countDimensionChanges() {
  let count = 0
  for (const sheet of sheets.value) {
    const rowLength = Math.max(sheet.rowHeights.length, (sheet.originalRowHeights || []).length)
    for (let index = 0; index < rowLength; index += 1) {
      const current = Math.round(Number(sheet.rowHeights[index] ?? DEFAULT_ROW_HEIGHT))
      const original = Math.round(Number(sheet.originalRowHeights?.[index] ?? DEFAULT_ROW_HEIGHT))
      if (current !== original) {
        count += 1
      }
    }
    const columnLength = Math.max(sheet.columnWidths.length, (sheet.originalColumnWidths || []).length)
    for (let index = 0; index < columnLength; index += 1) {
      const current = Math.round(Number(sheet.columnWidths[index] ?? DEFAULT_COLUMN_WIDTH))
      const original = Math.round(Number(sheet.originalColumnWidths?.[index] ?? DEFAULT_COLUMN_WIDTH))
      if (current !== original) {
        count += 1
      }
    }
  }
  return count
}

function collectRowHeightChanges() {
  const changes = []
  for (const sheet of sheets.value) {
    const rowLength = Math.max(sheet.rowHeights.length, (sheet.originalRowHeights || []).length)
    for (let index = 0; index < rowLength; index += 1) {
      const current = Math.round(Number(sheet.rowHeights[index] ?? DEFAULT_ROW_HEIGHT))
      const original = Math.round(Number(sheet.originalRowHeights?.[index] ?? DEFAULT_ROW_HEIGHT))
      if (current !== original) {
        changes.push({
          sheetName: sheet.name,
          rowIndex: index,
          heightPx: current
        })
      }
    }
  }
  return changes
}

function collectColumnWidthChanges() {
  const changes = []
  for (const sheet of sheets.value) {
    const columnLength = Math.max(sheet.columnWidths.length, (sheet.originalColumnWidths || []).length)
    for (let index = 0; index < columnLength; index += 1) {
      const current = Math.round(Number(sheet.columnWidths[index] ?? DEFAULT_COLUMN_WIDTH))
      const original = Math.round(Number(sheet.originalColumnWidths?.[index] ?? DEFAULT_COLUMN_WIDTH))
      if (current !== original) {
        changes.push({
          sheetName: sheet.name,
          colIndex: index,
          widthPx: current
        })
      }
    }
  }
  return changes
}

async function saveToServer() {
  if (!currentFileName.value || !hasWorkbook.value) {
    ElMessage.warning('请先选择并加载一份 Excel 文件')
    return
  }

  const changes = collectChanges()
  const mergeChanges = countMergeChanges()
  const rowHeightChanges = collectRowHeightChanges()
  const columnWidthChanges = collectColumnWidthChanges()
  if (!changes.length && !mergeChanges && !rowHeightChanges.length && !columnWidthChanges.length) {
    ElMessage.info('当前没有需要保存的修改')
    return
  }

  saving.value = true
  try {
    const response = await applyExcelWorkbookChanges({
      fileName: currentFileName.value,
      changes,
      mergeRegions: mergeChanges ? collectMergeRegions() : undefined,
      rowHeights: rowHeightChanges.length ? rowHeightChanges : undefined,
      columnWidths: columnWidthChanges.length ? columnWidthChanges : undefined
    })
    const nextFileName = response?.data?.currentFileName || currentFileName.value
    await loadWorkbookByFileName(nextFileName, true)
    ElMessage.success(response?.msg || 'Excel 已在线保存')
  } catch (error) {
    ElMessage.error(error?.message || '在线保存失败')
  } finally {
    saving.value = false
  }
}

async function reloadWorkbook() {
  if (!currentFileName.value) {
    ElMessage.warning('当前没有选中的文件')
    return
  }
  await loadWorkbookByFileName(currentFileName.value)
}

async function downloadServerWorkbook() {
  if (!currentFileName.value) {
    ElMessage.warning('当前没有可下载的文件')
    return
  }
  try {
    const buffer = await getExcelEditorContent(currentFileName.value)
    saveAs(new Blob([buffer]), currentFileName.value)
    ElMessage.success('文件已下载')
  } catch (error) {
    ElMessage.error(error?.message || '下载失败')
  }
}

function hasBackgroundImages(sheet) {
  return sheetHasBackgroundImage(sheet)
}

function hasAnchoredImages(sheet) {
  const images = Array.isArray(sheet?.images) ? sheet.images : []
  return images.some((image) => image.kind === 'anchored')
}

function listBackgroundImages(sheet) {
  const images = Array.isArray(sheet?.images) ? sheet.images : []
  return images.filter((image) => image.kind === 'background')
}

function listAnchoredImages(sheet) {
  const images = Array.isArray(sheet?.images) ? sheet.images : []
  return images.filter((image) => image.kind === 'anchored')
}

function sheetHasBackgroundImage(sheet) {
  const images = Array.isArray(sheet?.images) ? sheet.images : []
  return images.some((image) => image.kind === 'background')
}

function buildImageKey(image, index) {
  return `${image.kind}:${image.row1}:${image.col1}:${image.widthPx}:${image.heightPx}:${index}`
}

function buildSheetStageStyle(sheet) {
  return {
    width: `${resolveSheetPixelWidth(sheet)}px`,
    minWidth: '100%',
    minHeight: `${resolveSheetPixelHeight(sheet)}px`
  }
}

function buildImageStyle(image) {
  const sheet = currentSheet.value
  if (!sheet) {
    return {}
  }

  const dx = Math.max(Number(image?.dx1Px ?? 0), 0)
  const dy = Math.max(Number(image?.dy1Px ?? 0), 0)
  const top = COLUMN_HEADER_HEIGHT + resolveFrozenTopOffset(sheet, image.row1) + dy
  const left = ROW_INDEX_WIDTH + resolveFrozenLeftOffset(sheet, image.col1) + dx

  const style = {
    top: `${top}px`,
    left: `${left}px`,
    width: `${Math.max(Number(image?.widthPx ?? 0), 1)}px`,
    height: `${Math.max(Number(image?.heightPx ?? 0), 1)}px`
  }

  const freezePane = sheet.freezePane || { xSplit: 0, ySplit: 0 }
  const freezeRows = Math.max(Number(freezePane.ySplit ?? 0), 0)
  const freezeColumns = Math.max(Number(freezePane.xSplit ?? 0), 0)
  const inFrozenRow = image.kind !== 'background' && image.row1 < freezeRows
  const inFrozenColumn = image.kind !== 'background' && image.col1 < freezeColumns
  if (inFrozenRow || inFrozenColumn) {
    style.position = 'sticky'
    if (inFrozenRow) {
      style.top = `${top}px`
    }
    if (inFrozenColumn) {
      style.left = `${left}px`
    }
    style.zIndex = inFrozenRow && inFrozenColumn ? 8 : 7
  }
  return style
}

function buildCellTdStyle(cell) {
  const sheet = currentSheet.value
  const defaultBackground = sheet && sheetHasBackgroundImage(sheet) ? 'transparent' : '#ffffff'
  const style = {
    width: `${cell.widthPx}px`,
    minWidth: `${cell.widthPx}px`,
    height: `${cell.heightPx}px`,
    backgroundColor: cell.style?.backgroundColor || defaultBackground,
    textAlign: cell.style?.textAlign || 'left',
    verticalAlign: cell.style?.verticalAlign || 'top',
    borderTop: cell.style?.borderTop || '',
    borderRight: cell.style?.borderRight || '',
    borderBottom: cell.style?.borderBottom || '',
    borderLeft: cell.style?.borderLeft || ''
  }

  if (!sheet) {
    return style
  }

  const freezePane = sheet.freezePane || { xSplit: 0, ySplit: 0 }
  const freezeRows = Math.max(Number(freezePane.ySplit ?? 0), 0)
  const freezeColumns = Math.max(Number(freezePane.xSplit ?? 0), 0)
  const inFrozenRow = cell.rowIndex < freezeRows
  const inFrozenColumn = cell.colIndex < freezeColumns

  if (inFrozenRow || inFrozenColumn) {
    style.position = 'sticky'
    if (inFrozenRow) {
      style.top = `${COLUMN_HEADER_HEIGHT + resolveFrozenTopOffset(sheet, cell.rowIndex)}px`
    }
    if (inFrozenColumn) {
      style.left = `${ROW_INDEX_WIDTH + resolveFrozenLeftOffset(sheet, cell.colIndex)}px`
    }
    style.zIndex = inFrozenRow && inFrozenColumn ? 5 : 4
    style.boxShadow = 'inset 0 0 0 1px rgba(15, 122, 90, 0.08)'
  }

  return style
}

function buildCellClass(cell) {
  return {
    selected: isCellSelected(cell),
    'has-validation': hasValidationOptions(cell)
  }
}

function buildRowIndexStyle(row) {
  const style = {
    height: `${row.heightPx}px`
  }

  const sheet = currentSheet.value
  if (!sheet) {
    return style
  }
  const freezeRows = Math.max(Number(sheet.freezePane?.ySplit ?? 0), 0)
  if (row.rowIndex < freezeRows) {
    style.top = `${COLUMN_HEADER_HEIGHT + resolveFrozenTopOffset(sheet, row.rowIndex)}px`
    style.zIndex = 7
  }
  return style
}

function buildColumnHeaderStyle(sheet, columnIndex) {
  const width = Math.max(Number(sheet?.columnWidths?.[columnIndex] ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH)
  const style = {
    width: `${width}px`,
    minWidth: `${width}px`,
    height: `${COLUMN_HEADER_HEIGHT}px`
  }
  const freezeColumns = Math.max(Number(sheet?.freezePane?.xSplit ?? 0), 0)
  if (columnIndex < freezeColumns) {
    style.left = `${ROW_INDEX_WIDTH + resolveFrozenLeftOffset(sheet, columnIndex)}px`
    style.zIndex = 10
  }
  return style
}

function buildCellInputStyle(cell) {
  return {
    minHeight: `${Math.max(cell.heightPx - 8, 28)}px`,
    height: `${Math.max(cell.heightPx - 8, 28)}px`,
    textAlign: cell.style?.textAlign || 'left',
    color: cell.style?.color || '',
    fontWeight: cell.style?.fontWeight || '',
    fontStyle: cell.style?.fontStyle || '',
    fontSize: cell.style?.fontSize || '',
    fontFamily: cell.style?.fontFamily || '',
    textDecoration: cell.style?.textDecoration || '',
    whiteSpace: cell.style?.whiteSpace || 'nowrap'
  }
}

function resolveFrozenTopOffset(sheet, rowIndex) {
  let total = 0
  for (let index = 0; index < rowIndex; index += 1) {
    total += Math.max(Number(sheet.rowHeights[index] ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT)
  }
  return total
}

function resolveFrozenLeftOffset(sheet, columnIndex) {
  let total = 0
  for (let index = 0; index < columnIndex; index += 1) {
    total += Math.max(Number(sheet.columnWidths[index] ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH)
  }
  return total
}

function resolveSheetPixelWidth(sheet) {
  if (!sheet) {
    return ROW_INDEX_WIDTH + DEFAULT_COLUMN_WIDTH
  }
  let total = ROW_INDEX_WIDTH
  const maxColumnCount = Math.max(Number(sheet.maxColumnCount ?? 0), 1)
  for (let index = 0; index < maxColumnCount; index += 1) {
    total += Math.max(Number(sheet.columnWidths[index] ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH)
  }
  return total
}

function resolveSheetPixelHeight(sheet) {
  if (!sheet) {
    return COLUMN_HEADER_HEIGHT + DEFAULT_ROW_HEIGHT
  }
  let total = COLUMN_HEADER_HEIGHT
  const rowCount = Math.max(Number(sheet.rowCount ?? 0), 1)
  for (let index = 0; index < rowCount; index += 1) {
    total += Math.max(Number(sheet.rowHeights[index] ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT)
  }
  return total
}

function hasValidationOptions(cell) {
  return Array.isArray(cell?.validationOptions) && cell.validationOptions.length > 0
}

function buildCellDatalistId(cell) {
  const sheet = activeSheetName.value || 'sheet'
  const safeSheetName = String(sheet).replace(/[^a-zA-Z0-9_-]/g, '_')
  return `excel-list-${safeSheetName}-${cell.rowIndex}-${cell.colIndex}`
}

function startColumnResize(colIndex, event) {
  if (!currentSheet.value) {
    return
  }
  event.preventDefault()
  event.stopPropagation()
  const sheet = currentSheet.value
  columnResizeState.value = {
    sheetName: sheet.name,
    colIndex,
    startX: Number(event.clientX || 0),
    startWidth: Math.max(Number(sheet.columnWidths[colIndex] ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH)
  }
  rowResizeState.value = null
  document.body.style.cursor = 'col-resize'
  window.addEventListener('mousemove', handleResizePointerMove)
  window.addEventListener('mouseup', stopResize)
}

function startRowResize(rowIndex, event) {
  if (!currentSheet.value) {
    return
  }
  event.preventDefault()
  event.stopPropagation()
  const sheet = currentSheet.value
  rowResizeState.value = {
    sheetName: sheet.name,
    rowIndex,
    startY: Number(event.clientY || 0),
    startHeight: Math.max(Number(sheet.rowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT)
  }
  columnResizeState.value = null
  document.body.style.cursor = 'row-resize'
  window.addEventListener('mousemove', handleResizePointerMove)
  window.addEventListener('mouseup', stopResize)
}

function handleResizePointerMove(event) {
  const sheet = currentSheet.value
  if (!sheet) {
    clearResizeState()
    return
  }

  if (columnResizeState.value && columnResizeState.value.sheetName === sheet.name) {
    const state = columnResizeState.value
    const delta = Number(event.clientX || 0) - state.startX
    const nextWidth = Math.max(MIN_COLUMN_WIDTH, Math.round(state.startWidth + delta))
    if (sheet.columnWidths[state.colIndex] !== nextWidth) {
      sheet.columnWidths[state.colIndex] = nextWidth
      rebuildSheetRows(sheet)
      refreshDirtyState()
    }
    return
  }

  if (rowResizeState.value && rowResizeState.value.sheetName === sheet.name) {
    const state = rowResizeState.value
    const delta = Number(event.clientY || 0) - state.startY
    const nextHeight = Math.max(MIN_ROW_HEIGHT, Math.round(state.startHeight + delta))
    if (sheet.rowHeights[state.rowIndex] !== nextHeight) {
      sheet.rowHeights[state.rowIndex] = nextHeight
      rebuildSheetRows(sheet)
      refreshDirtyState()
    }
  }
}

function stopResize() {
  clearResizeState()
}

function clearResizeState() {
  columnResizeState.value = null
  rowResizeState.value = null
  document.body.style.cursor = ''
  window.removeEventListener('mousemove', handleResizePointerMove)
  window.removeEventListener('mouseup', stopResize)
}

function normalizeValue(value) {
  return value == null ? '' : String(value)
}

function normalizeQueryFileName(value) {
  if (Array.isArray(value)) {
    return value[0] || ''
  }
  return value || ''
}

function getCellRange(cell) {
  return {
    startRow: cell.rowIndex,
    endRow: cell.rowIndex + cell.rowSpan - 1,
    startCol: cell.colIndex,
    endCol: cell.colIndex + cell.colSpan - 1
  }
}

function getSelectionArea(range) {
  return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1)
}

function isMergedCell(cell) {
  return cell.rowSpan > 1 || cell.colSpan > 1
}

function rangesIntersect(left, right) {
  return left.startRow <= right.endRow
    && left.endRow >= right.startRow
    && left.startCol <= right.endCol
    && left.endCol >= right.startCol
}

function formatSelectionRange(range) {
  const startLabel = `${toColumnLabel(range.startCol)}${range.startRow + 1}`
  const endLabel = `${toColumnLabel(range.endCol)}${range.endRow + 1}`
  return startLabel === endLabel ? startLabel : `${startLabel}:${endLabel}`
}

function toColumnLabel(columnIndex) {
  let label = ''
  let current = Math.max(Number(columnIndex ?? 0), 0) + 1
  while (current > 0) {
    const remainder = (current - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    current = Math.floor((current - 1) / 26)
  }
  return label || 'A'
}

function ensureArraySize(target, size, defaultValue) {
  while (target.length < size) {
    target.push(defaultValue)
  }
}

function ensureSheetBounds(sheet, requiredRowCount, requiredColumnCount) {
  sheet.rowCount = Math.max(sheet.rowCount, requiredRowCount, 1)
  sheet.maxColumnCount = Math.max(sheet.maxColumnCount, requiredColumnCount, 1)
  ensureArraySize(sheet.rowHeights, sheet.rowCount, DEFAULT_ROW_HEIGHT)
  ensureArraySize(sheet.columnWidths, sheet.maxColumnCount, DEFAULT_COLUMN_WIDTH)
  if (!sheet.freezePane) {
    sheet.freezePane = normalizeFreezePane(null)
  }
  sheet.freezePane.xSplit = Math.min(Math.max(Number(sheet.freezePane.xSplit ?? 0), 0), sheet.maxColumnCount)
  sheet.freezePane.ySplit = Math.min(Math.max(Number(sheet.freezePane.ySplit ?? 0), 0), sheet.rowCount)
}

function resolveRangeWidth(sheet, startColumnIndex, colSpan) {
  ensureSheetBounds(sheet, sheet.rowCount, startColumnIndex + colSpan)
  let total = 0
  for (let offset = 0; offset < colSpan; offset += 1) {
    total += Math.max(Number(sheet.columnWidths[startColumnIndex + offset] ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH)
  }
  return total
}

function resolveRangeHeight(sheet, startRowIndex, rowSpan) {
  ensureSheetBounds(sheet, startRowIndex + rowSpan, sheet.maxColumnCount)
  let total = 0
  for (let offset = 0; offset < rowSpan; offset += 1) {
    total += Math.max(Number(sheet.rowHeights[startRowIndex + offset] ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT)
  }
  return total
}

function syncCellDimensions(sheet, cell) {
  ensureSheetBounds(sheet, cell.rowIndex + cell.rowSpan, cell.colIndex + cell.colSpan)
  cell.widthPx = resolveRangeWidth(sheet, cell.colIndex, cell.colSpan)
  cell.heightPx = resolveRangeHeight(sheet, cell.rowIndex, cell.rowSpan)
  const validation = resolveCellValidationOptions(sheet, cell.rowIndex, cell.colIndex)
  cell.validationOptions = validation.options
  cell.validationAllowBlank = validation.allowBlank
}

function resolveCellValidationOptions(sheet, rowIndex, colIndex) {
  const validations = Array.isArray(sheet?.validations) ? sheet.validations : []
  for (const validation of validations) {
    if (rowIndex < validation.firstRow || rowIndex > validation.lastRow) {
      continue
    }
    if (colIndex < validation.firstColumn || colIndex > validation.lastColumn) {
      continue
    }
    return {
      options: Array.isArray(validation.options) ? validation.options : [],
      allowBlank: validation.allowBlank !== false
    }
  }
  return {
    options: [],
    allowBlank: true
  }
}

function shouldKeepCurrentValidationValue(cell) {
  if (!cell || !normalizeValue(cell.value)) {
    return false
  }
  return !cell.validationOptions.includes(cell.value)
}

function dedupeCells(cells) {
  const cellMap = new Map()
  for (const cell of cells) {
    cellMap.set(buildCellKey(cell.rowIndex, cell.colIndex), cell)
  }
  return Array.from(cellMap.values()).sort((left, right) => {
    if (left.rowIndex !== right.rowIndex) {
      return left.rowIndex - right.rowIndex
    }
    return left.colIndex - right.colIndex
  })
}

function rebuildSheetRows(sheet) {
  let rowCount = Math.max(Number(sheet.rowCount ?? 0), 1)
  let maxColumnCount = Math.max(Number(sheet.maxColumnCount ?? 0), 1)

  for (const cell of sheet.cells) {
    rowCount = Math.max(rowCount, cell.rowIndex + cell.rowSpan)
    maxColumnCount = Math.max(maxColumnCount, cell.colIndex + cell.colSpan)
  }

  ensureSheetBounds(sheet, rowCount, maxColumnCount)
  sheet.cells = dedupeCells(sheet.cells)

  const rows = Array.from({ length: sheet.rowCount }, (_, rowIndex) => ({
    rowIndex,
    heightPx: Math.max(Number(sheet.rowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT),
    cells: []
  }))

  for (const cell of sheet.cells) {
    syncCellDimensions(sheet, cell)
    rows[cell.rowIndex].cells.push(cell)
  }

  sheet.rows = rows
}

function buildCellKey(rowIndex, colIndex) {
  return `${rowIndex}:${colIndex}`
}

function goToFileCenter() {
  router.push('/excel-editor/files')
}
</script>

<style lang="scss" scoped>
.excel-editor-page {
  min-height: calc(100vh - 84px);
  padding: 16px;
  background: #f6f8fb;
}

.hero-card,
.summary-card,
.editor-shell {
  border: 1px solid #e4e8ef;
  border-radius: 8px;
  background: #fff;
  box-shadow: none;
}

.hero-card {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 12px;
  padding: 12px 14px;
}

.eyebrow {
  margin: 0 0 4px;
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: #6b7280;
}

h1 {
  margin: 0;
  font-size: 20px;
  line-height: 1.3;
  color: #111827;
}

.hero-desc {
  margin: 6px 0 0;
  color: #6b7280;
  line-height: 1.4;
  font-size: 13px;
}

.hero-actions,
.toolbar-left,
.toolbar-right {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}

.hero-actions {
  align-content: flex-start;
  justify-content: flex-end;
}

.summary-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 10px;
  margin-bottom: 12px;
}

.summary-card {
  display: flex;
  flex-direction: column;
  gap: 4px;
  min-height: 0;
  padding: 10px 12px;
}

.summary-label {
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #6b7280;
}

.summary-card strong {
  font-size: 14px;
  line-height: 1.35;
  color: #111827;
}

.summary-card span:last-child {
  color: #6b7280;
  line-height: 1.35;
  font-size: 12px;
}

.editor-shell {
  padding: 12px;
}

.toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 14px;
}

.toolbar-right {
  align-items: center;
  justify-content: flex-end;
}

.selection-tip {
  font-size: 12px;
  color: #4b5563;
}

.sheet-tabs {
  margin-bottom: 12px;
}

.grid-wrap {
  overflow: auto;
  max-height: calc(100vh - 390px);
  border: 1px solid #e5e7eb;
  border-radius: 6px;
  background: #fff;
}

.sheet-stage {
  position: relative;
  width: max-content;
  min-width: 100%;
}

.sheet-grid {
  position: relative;
  z-index: 3;
  min-width: 100%;
  border-collapse: collapse;
}

.sheet-grid td,
.sheet-grid th {
  border: 1px solid #e5e7eb;
}

.sheet-grid thead th {
  position: sticky;
  top: 0;
}

.corner-header {
  position: sticky;
  top: 0;
  left: 0;
  z-index: 11;
  min-width: 56px;
  width: 56px;
  height: 34px;
  background: #f3f4f6;
}

.column-header {
  position: sticky;
  top: 0;
  z-index: 9;
  padding: 0 18px 0 10px;
  text-align: center;
  font-weight: 700;
  color: #374151;
  background: #f3f4f6;
}

.column-label {
  display: inline-block;
  line-height: 1;
}

.sheet-image-layer {
  position: absolute;
  inset: 0;
  pointer-events: none;
}

.sheet-image-layer-background {
  z-index: 1;
}

.sheet-image-layer-foreground {
  z-index: 6;
}

.sheet-image-item {
  position: absolute;
  overflow: hidden;
}

.sheet-image-item.background {
  opacity: 1;
}

.sheet-image {
  width: 100%;
  height: 100%;
  display: block;
  object-fit: fill;
  user-select: none;
}

.sheet-cell {
  position: relative;
  background: #fff;
  transition: box-shadow 0.18s ease, background-color 0.18s ease;
}

.sheet-cell.selected {
  box-shadow: inset 0 0 0 2px rgba(37, 99, 235, 0.65);
  background: rgba(37, 99, 235, 0.06);
}

.sheet-cell.selected .cell-input {
  background: rgba(37, 99, 235, 0.04);
}

.row-index {
  position: sticky;
  left: 0;
  z-index: 3;
  min-width: 56px;
  width: 56px;
  padding: 0 8px;
  text-align: center;
  font-weight: 700;
  color: #374151;
  background: #f3f4f6;
}

.resize-handle {
  position: absolute;
  user-select: none;
}

.col-resize-handle {
  top: 0;
  right: -4px;
  width: 8px;
  height: 100%;
  cursor: col-resize;
}

.row-resize-handle {
  left: 0;
  bottom: -4px;
  width: 100%;
  height: 8px;
  cursor: row-resize;
}

.cell-input {
  display: block;
  width: 100%;
  min-width: 100%;
  padding: 8px 10px;
  border: none;
  outline: none;
  resize: none;
  background: transparent;
  line-height: 1.45;
  overflow: hidden;
}

.cell-select {
  padding-right: 26px;
}

.sheet-cell.has-validation::after {
  content: "▼";
  position: absolute;
  right: 8px;
  bottom: 6px;
  font-size: 10px;
  line-height: 1;
  color: rgba(55, 65, 81, 0.6);
  pointer-events: none;
}

.cell-input.dirty {
  background: rgba(255, 245, 214, 0.85);
}

.cell-input.formula {
  box-shadow: inset 0 0 0 1px rgba(47, 119, 208, 0.14);
}

@media (max-width: 1200px) {
  .summary-grid {
    grid-template-columns: 1fr;
  }
}

@media (max-width: 960px) {
  .excel-editor-page {
    padding: 16px;
  }

  .hero-card,
  .toolbar,
  .toolbar-right {
    flex-direction: column;
    align-items: flex-start;
  }
}
</style>
