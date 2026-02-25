<template>
  <div class="app-container">
    <h2>公版图管理</h2>
    <el-table :data="mockPublicMaps" border>
      <el-table-column prop="id" label="ID" width="100" />
      <el-table-column prop="name" label="名称" />
      <el-table-column prop="version" label="版本" width="180" />
      <el-table-column prop="status" label="状态" width="140" />
    </el-table>
  </div>
</template>

<script setup>
import portalData from '@/mock/portal-data.json'

const mockPublicMaps = portalData.publicMaps
</script>
