<template>
  <div class="app-container">
    <h2>统计报表</h2>
    <el-table :data="mockStatistics" border>
      <el-table-column prop="name" label="指标" />
      <el-table-column prop="value" label="数值" width="180" />
    </el-table>
  </div>
</template>

<script setup>
import portalData from '@/mock/portal-data.json'

const mockStatistics = portalData.statistics
</script>
