<template>
  <div class="excel-editor-page">
    <section class="hero-card">
      <div>
        <p class="eyebrow">Format Preserving Editor</p>
        <h1>在线Excel编辑</h1>
        <p class="hero-desc">
          当前版本会基于后端原始工作簿直接改单元格值，尽量保留合并单元格、底色、字体、列宽、行高等原始格式。
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
        <span>{{ isDirty ? '保存时会只提交改动过的单元格' : '当前内容与后端文件一致' }}</span>
      </article>
    </section>

    <section class="editor-shell" v-loading="pageLoading">
      <template v-if="hasWorkbook">
        <div class="toolbar">
          <div class="toolbar-left">
            <el-tag type="info" effect="light">编辑文件：{{ currentFileName }}</el-tag>
            <el-tag :type="isDirty ? 'warning' : 'success'" effect="light">{{ isDirty ? '未保存' : '已保存' }}</el-tag>
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

        <div class="grid-wrap">
          <table class="sheet-grid">
            <tbody>
              <tr v-for="row in currentSheet.rows" :key="`row-${row.rowIndex}`">
                <td class="row-index" :style="{ height: `${row.heightPx}px` }">{{ row.rowIndex + 1 }}</td>
                <td
                  v-for="cell in row.cells"
                  :key="`cell-${row.rowIndex}-${cell.colIndex}`"
                  :colspan="cell.colSpan"
                  :rowspan="cell.rowSpan"
                  :style="buildCellTdStyle(cell)"
                >
                  <textarea
                    v-model="cell.value"
                    class="cell-input"
                    :class="{ dirty: cell.dirty, formula: cell.formula }"
                    :style="buildCellInputStyle(cell)"
                    spellcheck="false"
                    @input="handleCellInput(cell)"
                  />
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </template>

      <el-empty v-else description="还没有加载可编辑的文件，请先去文件管理页选择一份已上传 Excel。" />
    </section>
  </div>
</template>

<script setup>
import { computed, onMounted, ref, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import { saveAs } from 'file-saver'
import { applyExcelWorkbookChanges, getExcelEditorContent, getExcelEditorInfo, getExcelWorkbookView } from '@/api/tool/excelEditor'

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

const currentSheet = computed(() => sheets.value.find((sheet) => sheet.name === activeSheetName.value) || null)
const hasWorkbook = computed(() => sheets.value.length > 0)

onMounted(async () => {
  await initializeFromRoute()
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
    await ElMessageBox.confirm('当前有未保存修改，继续操作会丢失这些改动，是否继续？', '确认切换', {
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
  isDirty.value = false
  dirtyCount.value = 0
}

function normalizeSheet(sheet) {
  return {
    name: sheet?.name || `Sheet${Math.random().toString(36).slice(2, 6)}`,
    rows: Array.isArray(sheet?.rows) ? sheet.rows.map(normalizeRow) : []
  }
}

function normalizeRow(row) {
  return {
    rowIndex: Number(row?.rowIndex ?? 0),
    heightPx: Number(row?.heightPx ?? 28),
    cells: Array.isArray(row?.cells) ? row.cells.map(normalizeCell) : []
  }
}

function normalizeCell(cell) {
  const value = cell?.value == null ? '' : String(cell.value)
  return {
    rowIndex: Number(cell?.rowIndex ?? 0),
    colIndex: Number(cell?.colIndex ?? 0),
    rowSpan: Math.max(Number(cell?.rowSpan ?? 1), 1),
    colSpan: Math.max(Number(cell?.colSpan ?? 1), 1),
    widthPx: Math.max(Number(cell?.widthPx ?? 96), 72),
    heightPx: Math.max(Number(cell?.heightPx ?? 32), 24),
    value,
    originalValue: value,
    displayValue: cell?.displayValue == null ? '' : String(cell.displayValue),
    formula: Boolean(cell?.formula),
    dirty: false,
    style: typeof cell?.style === 'object' && cell.style ? { ...cell.style } : {}
  }
}

function clearWorkbook() {
  workbookName.value = '未加载'
  sheets.value = []
  activeSheetName.value = ''
  currentFileName.value = ''
  isDirty.value = false
  dirtyCount.value = 0
}

function handleCellInput(cell) {
  cell.dirty = normalizeValue(cell.value) !== normalizeValue(cell.originalValue)
  refreshDirtyState()
}

function refreshDirtyState() {
  let count = 0
  for (const sheet of sheets.value) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (cell.dirty) {
          count += 1
        }
      }
    }
  }
  dirtyCount.value = count
  isDirty.value = count > 0
}

function collectChanges() {
  const changes = []
  for (const sheet of sheets.value) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
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
  }
  return changes
}

async function saveToServer() {
  if (!currentFileName.value || !hasWorkbook.value) {
    ElMessage.warning('请先选择并加载一份 Excel 文件')
    return
  }
  const changes = collectChanges()
  if (!changes.length) {
    ElMessage.info('当前没有需要保存的修改')
    return
  }

  saving.value = true
  try {
    const response = await applyExcelWorkbookChanges({
      fileName: currentFileName.value,
      changes
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

function buildCellTdStyle(cell) {
  return {
    width: `${cell.widthPx}px`,
    minWidth: `${cell.widthPx}px`,
    height: `${cell.heightPx}px`,
    backgroundColor: cell.style?.backgroundColor || '',
    textAlign: cell.style?.textAlign || 'left',
    verticalAlign: cell.style?.verticalAlign || 'top'
  }
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
    whiteSpace: cell.style?.whiteSpace || 'pre-wrap'
  }
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

function goToFileCenter() {
  router.push('/excel-editor/files')
}
</script>

<style lang="scss" scoped>
.excel-editor-page {
  min-height: calc(100vh - 84px);
  padding: 24px;
  background:
    radial-gradient(circle at top right, rgba(16, 109, 79, 0.12), transparent 28%),
    linear-gradient(180deg, #f4f7f2 0%, #edf2ed 100%);
}

.hero-card,
.summary-card,
.editor-shell {
  border: 1px solid rgba(16, 88, 68, 0.1);
  border-radius: 18px;
  background: rgba(255, 255, 255, 0.95);
  box-shadow: 0 16px 38px rgba(31, 44, 37, 0.08);
}

.hero-card {
  display: flex;
  justify-content: space-between;
  gap: 20px;
  margin-bottom: 16px;
  padding: 24px;
}

.eyebrow {
  margin: 0 0 8px;
  font-size: 12px;
  font-weight: 700;
  letter-spacing: 0.14em;
  text-transform: uppercase;
  color: #0f7a5a;
}

h1 {
  margin: 0;
  font-size: 30px;
  line-height: 1.1;
  color: #17352a;
}

.hero-desc {
  margin: 10px 0 0;
  color: #4f6559;
  line-height: 1.6;
}

.hero-actions,
.toolbar-left {
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
  gap: 16px;
  margin-bottom: 16px;
}

.summary-card {
  display: flex;
  flex-direction: column;
  gap: 6px;
  min-height: 110px;
  padding: 18px 20px;
}

.summary-label {
  font-size: 12px;
  font-weight: 700;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: #728579;
}

.summary-card strong {
  font-size: 18px;
  line-height: 1.35;
  color: #20392d;
}

.summary-card span:last-child {
  color: #627669;
  line-height: 1.5;
}

.editor-shell {
  padding: 18px;
}

.toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 14px;
}

.sheet-tabs {
  margin-bottom: 12px;
}

.grid-wrap {
  overflow: auto;
  max-height: calc(100vh - 390px);
  border: 1px solid #dce7de;
  border-radius: 14px;
  background: #fff;
}

.sheet-grid {
  min-width: 100%;
  border-collapse: collapse;
}

.sheet-grid td {
  border: 1px solid #dfe6e1;
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
  color: #537063;
  background: #edf7f1;
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
  .toolbar {
    flex-direction: column;
    align-items: flex-start;
  }
}
</style>
