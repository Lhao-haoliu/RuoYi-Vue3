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
        <header class="card-header">
          <el-icon><DataLine /></el-icon>
          <span>{{ item.label }}</span>
        </header>
        <div class="kpi-number">{{ item.value }}</div>
        <div class="kpi-trend">{{ item.trend }}</div>
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
            <div class="status-bar">
              <span :style="{ width: `${row.value}%`, background: row.color }"></span>
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
import { DataLine, PieChart, Warning, Grid } from '@element-plus/icons-vue'
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
</script>

<style scoped lang="scss">
.portal-dashboard {
  display: grid;
  gap: 16px;
}

.portal-banner {
  min-height: 140px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  h2 {
    margin: 0 0 8px;
    font-size: 28px;
    color: #0f1f3d;
  }
  p {
    margin: 0;
    color: #54617d;
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
  color: #2e3a59;
  font-weight: 500;
  font-size: 14px;
  .el-button {
    margin-left: auto;
  }
}

.kpi-number {
  font-size: 32px;
  font-weight: 700;
  line-height: 1.2;
  color: #2e3a59;
}

.kpi-trend {
  margin-top: 4px;
  font-size: 13px;
  color: #4a90e2;
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
    color: #55627f;
    font-size: 13px;
  }
  strong {
    color: #2e3a59;
    font-size: 13px;
  }
}

.status-bar {
  width: 100%;
  height: 10px;
  border-radius: 999px;
  background: #edf2fa;
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
  border: 1px solid #ecf0f7;
  border-radius: 10px;
}

.event-time {
  color: #6f7c97;
  font-size: 13px;
}

.event-title {
  font-size: 13px;
  color: #2e3a59;
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
