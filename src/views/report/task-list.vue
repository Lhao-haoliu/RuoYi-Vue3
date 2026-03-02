<template>
  <div class="app-container">
    <h2>任务明细</h2>
    <el-form :inline="true" class="mb8">
      <el-form-item label="任务名称">
        <el-input v-model="keyword" placeholder="请输入任务名称" clearable style="width: 220px" />
      </el-form-item>
      <el-form-item>
        <el-button type="primary" v-hasPermi="['report:view']">查询</el-button>
      </el-form-item>
    </el-form>
    <el-table :data="filteredTasks" border>
      <el-table-column prop="id" label="ID" width="100" />
      <el-table-column prop="name" label="任务名称" />
      <el-table-column prop="status" label="状态" width="160" />
      <el-table-column label="操作" width="120">
        <template #default="scope">
          <el-button link type="danger" v-hasPermi="['report:export']" @click="handleDelete(scope.row.id)">
            删除
          </el-button>
        </template>
      </el-table-column>
    </el-table>
  </div>
</template>

<script setup>
import { computed, ref } from "vue"
import { ElMessage } from "element-plus"
import portalData from "@/mock/portal-data.json"

const keyword = ref("")
const mockTasks = ref([...(portalData.tasks || [])])

const filteredTasks = computed(() => {
  const key = keyword.value.trim().toLowerCase()
  if (!key) return mockTasks.value
  return mockTasks.value.filter((item) => String(item.name).toLowerCase().includes(key))
})

function handleDelete(id) {
  mockTasks.value = mockTasks.value.filter((item) => item.id !== id)
  ElMessage.success("已删除（前端模拟）")
}
</script>
