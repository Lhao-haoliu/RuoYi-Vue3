<template>
  <div class="portal-dashboard">
    <section class="portal-banner portal-card">
      <div>
        <h2>{{ dashboard.banner.greeting }}，{{ userName }}</h2>
        <p>{{ dashboard.banner.subTitle }}</p>
      </div>
      <el-button class="portal-primary-btn">进入 Commander</el-button>
    </section>

    <section class="portal-grid kpi-grid">
      <article class="portal-card kpi-card" v-for="item in dashboard.overview" :key="item.label">
        <header class="kpi-label">{{ item.label }}</header>
        <div class="kpi-number">{{ item.value }}</div>
        <div class="kpi-trend" :class="trendClass(item.trend)">{{ item.trend }}</div>
      </article>
    </section>

    <section class="portal-grid content-grid">
      <article class="portal-card">
        <header class="card-header">
          <el-icon><PieChart /></el-icon>
          <span>任务状态分布</span>
          <el-button class="portal-secondary-btn" size="small">筛选</el-button>
        </header>
        <div class="status-list">
          <div class="status-row" v-for="row in dashboard.statusDistribution" :key="row.name">
            <label>{{ row.name }}</label>
            <div class="status-bar" :style="{ background: statusTrackColor(row.name) }">
              <span :style="{ width: `${row.value}%`, background: statusColor(row.name) }"></span>
            </div>
            <strong>{{ row.value }}%</strong>
          </div>
        </div>
      </article>

      <article class="portal-card">
        <header class="card-header">
          <el-icon><Warning /></el-icon>
          <span>最近异常 / Important</span>
          <el-button class="portal-secondary-btn" size="small">导出</el-button>
        </header>
        <div class="event-list">
          <div class="event-item" v-for="event in dashboard.recentImportant" :key="event.time + event.title">
            <span class="event-time">{{ event.time }}</span>
            <span class="event-title">{{ event.title }}</span>
            <el-tag :type="tagType(event.level)" effect="light">{{ event.level }}</el-tag>
          </div>
        </div>
      </article>

      <article class="portal-card">
        <header class="card-header">
          <el-icon><Grid /></el-icon>
          <span>快捷入口</span>
        </header>
        <div class="quick-actions">
          <el-button
            v-for="quick in quickMenu.quickEntry"
            :key="quick.path"
            class="portal-primary-btn quick-btn"
            @click="router.push(quick.path)"
          >
            {{ quick.title }}
          </el-button>
        </div>
      </article>
    </section>
  </div>
</template>

<script setup>
import { PieChart, Warning, Grid } from '@element-plus/icons-vue'
import useUserStore from '@/store/modules/user'
import dashboard from '@/mock/dashboard.json'
import quickMenu from '@/mock/menu.json'

const router = useRouter()
const userStore = useUserStore()

const userName = computed(() => userStore.nickName || userStore.name || '管理员')

function tagType(level) {
  if (level === 'Success') return 'success'
  if (level === 'Warning') return 'warning'
  if (level === 'Error') return 'danger'
  return 'info'
}

function trendClass(trend) {
  const value = String(trend || '').trim()
  if (value.startsWith('-')) {
    return 'is-negative'
  }
  return 'is-positive'
}

function statusColor(status) {
  if (status === 'Success') return '#10B981'
  if (status === 'Warning') return '#F59E0B'
  if (status === 'Error') return '#EF4444'
  return '#2563EB'
}

function statusTrackColor(status) {
  if (status === 'Success') return 'rgb(16 185 129 / 0.1)'
  if (status === 'Warning') return 'rgb(245 158 11 / 0.1)'
  if (status === 'Error') return 'rgb(239 68 68 / 0.1)'
  return 'rgb(37 99 235 / 0.1)'
}
</script>

<style scoped lang="scss">
.portal-dashboard {
  display: grid;
  gap: 18px;
}

.portal-banner {
  min-height: 140px;
  display: flex;
  align-items: center;
  justify-content: space-between;

  h2 {
    margin: 0 0 8px;
    font-size: 30px;
    color: #111827;
  }

  p {
    margin: 0;
    color: #4b5563;
    font-size: 14px;
  }
}

.portal-grid {
  display: grid;
  gap: 16px;
}

.kpi-grid {
  grid-template-columns: repeat(4, minmax(0, 1fr));
}

.content-grid {
  grid-template-columns: 1.3fr 1fr 1fr;
}

.card-header {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 14px;
  color: #111827;
  font-weight: 500;
  font-size: 14px;

  .el-button {
    margin-left: auto;
  }
}

.kpi-label {
  margin-bottom: 10px;
  color: #4b5563;
  font-size: 13px;
  font-weight: 600;
}

.kpi-number {
  font-size: 24px;
  font-weight: 700;
  line-height: 1.2;
  color: #111827;
}

.kpi-trend {
  margin-top: 8px;
  font-size: 13px;
  font-weight: 600;
}

.kpi-trend.is-positive {
  color: #10b981;
}

.kpi-trend.is-negative {
  color: #ef4444;
}

.status-list {
  display: grid;
  gap: 12px;
}

.status-row {
  display: grid;
  grid-template-columns: 72px 1fr 46px;
  align-items: center;
  gap: 8px;

  label {
    color: #4b5563;
    font-size: 13px;
  }

  strong {
    color: #111827;
    font-size: 13px;
  }
}

.status-bar {
  width: 100%;
  height: 8px;
  border-radius: 999px;
  overflow: hidden;

  span {
    display: block;
    height: 100%;
  }
}

.event-list {
  display: grid;
  gap: 10px;
}

.event-item {
  display: grid;
  grid-template-columns: 58px 1fr auto;
  align-items: center;
  gap: 8px;
  padding: 10px 12px;
  border: 1px solid #e5e7eb;
  border-radius: 6px;
}

.event-time {
  color: #9ca3af;
  font-size: 13px;
}

.event-title {
  font-size: 13px;
  color: #111827;
}

.quick-actions {
  display: grid;
  gap: 10px;
  align-content: flex-start;
}

.quick-btn {
  width: 100%;
}

@media (max-width: 1200px) {
  .kpi-grid {
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }

  .content-grid {
    grid-template-columns: 1fr;
  }
}
</style>
