<template>
  <div class="app-container">
    <h2>任务启动</h2>
    <el-table :data="mockStartItems" border>
      <el-table-column prop="id" label="ID" width="100" />
      <el-table-column prop="name" label="任务名称" />
      <el-table-column prop="owner" label="指挥员" width="180" />
      <el-table-column prop="status" label="状态" width="140" />
    </el-table>
  </div>
</template>

<script setup>
import portalData from '@/mock/portal-data.json'

const mockStartItems = portalData.startItems
</script>
