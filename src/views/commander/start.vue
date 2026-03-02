<template>
  <div class="app-container">
    <h2>任务启动</h2>
    <div class="mb8">
      <el-button type="primary" v-hasPermi="['commander:start:mockRun']" @click="mockStart">模拟启动任务</el-button>
      <el-button v-hasPermi="['commander:start:mockStop']" @click="mockStop">模拟停止任务</el-button>
    </div>
    <el-table :data="mockStartItems" border>
      <el-table-column prop="id" label="ID" width="100" />
      <el-table-column prop="name" label="任务名称" />
      <el-table-column prop="owner" label="指挥员" width="180" />
      <el-table-column prop="status" label="状态" width="140" />
    </el-table>
  </div>
</template>

<script setup>
import { ElMessage } from "element-plus"
import portalData from "@/mock/portal-data.json"

const mockStartItems = portalData.startItems || []

function mockStart() {
  ElMessage.success("模拟启动成功（仅前端）")
}

function mockStop() {
  ElMessage.warning("模拟停止成功（仅前端）")
}
</script>
