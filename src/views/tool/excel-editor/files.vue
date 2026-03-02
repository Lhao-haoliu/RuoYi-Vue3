<template>
  <div class="excel-files-page">
    <input
      ref="uploadInputRef"
      accept=".xlsx,.xls"
      class="upload-input"
      type="file"
      @change="handleFileSelect"
    >

    <section class="hero-card">
      <div>
        <p class="eyebrow">Excel Files</p>
        <h1>在线Excel文件管理</h1>
        <p class="hero-desc">
          这里专门查看所有已上传文件。上传后会进入列表，你可以明确选择要编辑的那一份，再进入在线 Excel 编辑页。
        </p>
      </div>
      <div class="hero-actions">
        <el-button :loading="loading" @click="refreshStorageInfo">刷新列表</el-button>
        <el-button type="primary" :loading="uploading" @click="triggerUpload">上传Excel</el-button>
      </div>
    </section>

    <section class="summary-grid">
      <article class="summary-card">
        <span class="summary-label">已上传文件</span>
        <strong>{{ uploadedFiles.length }} 个</strong>
        <span>按最近修改时间排序</span>
      </article>
      <article class="summary-card">
        <span class="summary-label">当前后端文件</span>
        <strong>{{ currentFileName || '未指定' }}</strong>
        <span>打开编辑时默认读取这一份</span>
      </article>
    </section>

    <section class="table-card" v-loading="loading">
      <el-table :data="uploadedFiles" stripe>
        <el-table-column label="文件名" min-width="320" prop="fileName" />
        <el-table-column label="大小" min-width="140">
          <template #default="{ row }">
            {{ formatFileSize(row.size) }}
          </template>
        </el-table-column>
        <el-table-column label="更新时间" min-width="200">
          <template #default="{ row }">
            {{ formatTime(row.lastModified) }}
          </template>
        </el-table-column>
        <el-table-column label="状态" width="140">
          <template #default="{ row }">
            <el-tag :type="row.fileName === currentFileName ? 'success' : 'info'" effect="light">
              {{ row.fileName === currentFileName ? '当前文件' : '可编辑' }}
            </el-tag>
          </template>
        </el-table-column>
        <el-table-column label="操作" width="180" fixed="right">
          <template #default="{ row }">
            <el-button link type="primary" @click="openEditor(row.fileName)">编辑此文件</el-button>
          </template>
        </el-table-column>
      </el-table>

      <el-empty
        v-if="!uploadedFiles.length && !loading"
        description="还没有已上传文件，请先上传一份 Excel。"
      />
    </section>
  </div>
</template>

<script setup>
import { onMounted, ref } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import { getExcelEditorInfo, uploadExcelEditorFile } from '@/api/tool/excelEditor'

const router = useRouter()

const uploadInputRef = ref(null)
const loading = ref(false)
const uploading = ref(false)
const uploadedFiles = ref([])
const currentFileName = ref('')

onMounted(async () => {
  await refreshStorageInfo()
})

async function refreshStorageInfo() {
  loading.value = true
  try {
    const response = await getExcelEditorInfo()
    uploadedFiles.value = Array.isArray(response?.data?.files) ? response.data.files : []
    currentFileName.value = response?.data?.currentFileName || ''
  } catch (error) {
    ElMessage.error(error?.message || '读取文件列表失败')
  } finally {
    loading.value = false
  }
}

function triggerUpload() {
  uploadInputRef.value?.click()
}

async function handleFileSelect(event) {
  const file = event.target.files?.[0]
  event.target.value = ''
  if (!file) {
    return
  }
  uploading.value = true
  try {
    const formData = new FormData()
    formData.append('file', file)
    const response = await uploadExcelEditorFile(formData)
    ElMessage.success(response.msg || 'Excel 已上传')
    await refreshStorageInfo()
    const nextFileName = response?.data?.currentFileName || file.name
    openEditor(nextFileName)
  } catch (error) {
    ElMessage.error(error?.message || '上传 Excel 失败')
  } finally {
    uploading.value = false
  }
}

function openEditor(fileName) {
  if (!fileName) {
    ElMessage.warning('没有可编辑的文件')
    return
  }
  router.push({
    path: '/excel-editor/index',
    query: { fileName }
  })
}

function formatFileSize(size) {
  const value = Number(size || 0)
  if (value < 1024) {
    return `${value} B`
  }
  if (value < 1024 * 1024) {
    return `${(value / 1024).toFixed(1)} KB`
  }
  return `${(value / (1024 * 1024)).toFixed(2)} MB`
}

function formatTime(timestamp) {
  if (!timestamp) {
    return '-'
  }
  const date = new Date(timestamp)
  const pad = (value) => String(value).padStart(2, '0')
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`
}
</script>

<style lang="scss" scoped>
.excel-files-page {
  min-height: calc(100vh - 84px);
  padding: 24px;
  background:
    radial-gradient(circle at top left, rgba(19, 101, 161, 0.12), transparent 24%),
    linear-gradient(180deg, #f4f8fb 0%, #edf3f7 100%);
}

.upload-input {
  display: none;
}

.hero-card,
.summary-card,
.table-card {
  border: 1px solid rgba(24, 78, 121, 0.1);
  border-radius: 18px;
  background: rgba(255, 255, 255, 0.96);
  box-shadow: 0 16px 38px rgba(33, 53, 73, 0.08);
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
  color: #1f6da2;
}

h1 {
  margin: 0;
  font-size: 30px;
  line-height: 1.1;
  color: #1d3551;
}

.hero-desc {
  margin: 10px 0 0;
  color: #52677c;
  line-height: 1.6;
}

.hero-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  align-content: flex-start;
}

.summary-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 16px;
  margin-bottom: 16px;
}

.summary-card {
  display: flex;
  flex-direction: column;
  gap: 6px;
  min-height: 106px;
  padding: 18px 20px;
}

.summary-label {
  font-size: 12px;
  font-weight: 700;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: #74879a;
}

.summary-card strong {
  font-size: 18px;
  line-height: 1.35;
  color: #22384f;
}

.summary-card span:last-child {
  color: #607487;
}

.table-card {
  padding: 18px;
}

@media (max-width: 960px) {
  .excel-files-page {
    padding: 16px;
  }

  .hero-card {
    flex-direction: column;
    align-items: flex-start;
  }

  .summary-grid {
    grid-template-columns: 1fr;
  }
}
</style>
