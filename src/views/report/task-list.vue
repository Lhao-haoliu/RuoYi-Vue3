<template>
  <div class="app-container">
    <h2>任务明细</h2>
    <el-table :data="mockTasks" border>
      <el-table-column prop="id" label="ID" width="100" />
      <el-table-column prop="name" label="任务名称" />
      <el-table-column prop="status" label="状态" width="160" />
    </el-table>
  </div>
</template>

<script setup>
import portalData from '@/mock/portal-data.json'

const mockTasks = portalData.tasks
</script>
