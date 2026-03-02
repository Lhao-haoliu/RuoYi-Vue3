<template>
  <div class="app-container">
    <h2>统计报表</h2>
    <el-form :inline="true" class="mb8">
      <el-form-item label="指标名称">
        <el-input v-model="keyword" placeholder="请输入指标名称" clearable style="width: 220px" />
      </el-form-item>
      <el-form-item>
        <el-button type="primary" v-hasPermi="['report:view']">查询</el-button>
      </el-form-item>
    </el-form>
    <el-table :data="filteredStatistics" border>
      <el-table-column prop="name" label="指标" />
      <el-table-column prop="value" label="数值" width="180" />
      <el-table-column label="操作" width="120">
        <template #default="scope">
          <el-button link type="danger" v-hasPermi="['report:export']" @click="handleDelete(scope.row.name)">
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
const mockStatistics = ref([...(portalData.statistics || [])])

const filteredStatistics = computed(() => {
  const key = keyword.value.trim().toLowerCase()
  if (!key) return mockStatistics.value
  return mockStatistics.value.filter((item) => String(item.name).toLowerCase().includes(key))
})

function handleDelete(name) {
  mockStatistics.value = mockStatistics.value.filter((item) => item.name !== name)
  ElMessage.success("已删除（前端模拟）")
}
</script>
