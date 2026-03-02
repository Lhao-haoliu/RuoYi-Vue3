<template>
  <div class="portal-shell">
    <aside class="portal-sidebar">
      <div class="portal-brand">
        <img :src="portalLogo" alt="SEMI Logo" />
        <div class="portal-brand-text">
          <h1>eRPA</h1>
          <p>SEMI控制台</p>
        </div>
      </div>
      <el-scrollbar class="portal-menu-scroll">
        <el-menu
          :default-active="activePath"
          :collapse="false"
          :unique-opened="true"
          :collapse-transition="false"
          background-color="transparent"
          text-color="#b9cae7"
          active-text-color="#ffffff"
          router
        >
          <portal-nav-item
            v-for="item in menuRoutes"
            :key="item.path"
            :item="item"
            :base-path="''"
          />
        </el-menu>
      </el-scrollbar>
    </aside>

    <section class="portal-main">
      <header class="portal-topbar">
        <div class="topbar-left">
          <el-input v-model="searchKey" class="portal-search" placeholder="全局搜索模块..." clearable />
        </div>
        <div class="topbar-right">
          <el-badge :value="3" class="portal-notify">
            <el-button circle class="portal-icon-btn">
              <el-icon><Bell /></el-icon>
            </el-button>
          </el-badge>
          <div class="portal-time">{{ nowText }}</div>
          <el-dropdown>
            <div class="portal-user">
              <el-avatar :src="userStore.avatar || undefined" :size="34" />
              <span>{{ userStore.nickName || userStore.name || "用户" }}</span>
            </div>
            <template #dropdown>
              <el-dropdown-menu>
                <el-dropdown-item @click="router.push('/user/profile')">个人中心</el-dropdown-item>
                <el-dropdown-item divided @click="handleLogout">退出登录</el-dropdown-item>
              </el-dropdown-menu>
            </template>
          </el-dropdown>
        </div>
      </header>
      <main class="portal-content">
        <router-view />
      </main>
    </section>
  </div>
</template>

<script setup>
import { Bell } from '@element-plus/icons-vue'
import useUserStore from '@/store/modules/user'
import usePermissionStore from '@/store/modules/permission'
import PortalNavItem from '@/layout/components/portal/PortalNavItem.vue'
import portalLogo from '@/assets/logo/company-logo.svg'

const router = useRouter()
const route = useRoute()
const userStore = useUserStore()
const permissionStore = usePermissionStore()

const searchKey = ref('')
const nowText = ref('')
let timer = null

const activePath = computed(() => route.path)
const menuRoutes = computed(() =>
  (permissionStore.sidebarRouters || []).filter((r) => !r.hidden && r.meta && r.meta.title)
)

function updateNow() {
  const now = new Date()
  nowText.value = now.toLocaleString('zh-CN', {
    hour12: false,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  })
}

function handleLogout() {
  userStore.logOut().then(() => {
    router.push('/login')
  })
}

onMounted(() => {
  updateNow()
  timer = setInterval(updateNow, 1000)
})

onBeforeUnmount(() => {
  if (timer) clearInterval(timer)
})
</script>

<style scoped lang="scss">
.portal-shell {
  min-height: 100vh;
  display: grid;
  grid-template-columns: 264px 1fr;
  background: #f7f9fc;
  transition: grid-template-columns 0.2s ease;
}

.portal-sidebar {
  background: linear-gradient(180deg, #0f1f3d 0%, #132b55 100%);
  border-right: 1px solid rgba(255, 255, 255, 0.08);
  display: flex;
  flex-direction: column;
}

.portal-brand {
  height: 76px;
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 0 18px;
  border-bottom: 1px solid rgba(255, 255, 255, 0.09);
  img {
    width: 104px;
    height: 30px;
    object-fit: contain;
  }
  .portal-brand-text {
    overflow: hidden;
  }
  h1 {
    margin: 0;
    font-size: 18px;
    font-weight: 700;
    color: #fff;
    letter-spacing: 0.6px;
  }
  p {
    margin: 0;
    font-size: 12px;
    color: #8ea6cf;
  }
}

.portal-menu-scroll {
  height: calc(100vh - 76px);
  :deep(.el-menu) {
    border-right: none;
    padding: 10px 8px 20px;
  }
  :deep(.el-menu-item),
  :deep(.el-sub-menu__title) {
    border-radius: 10px;
    margin-bottom: 6px;
  }
  :deep(.el-menu-item.is-active) {
    background: linear-gradient(90deg, #4a90e2 0%, #6c5dd3 100%);
  }
  :deep(.el-menu-item:hover),
  :deep(.el-sub-menu__title:hover) {
    background: rgba(255, 255, 255, 0.12);
  }
}

.portal-main {
  display: flex;
  flex-direction: column;
  min-width: 0;
}

.portal-topbar {
  height: 76px;
  padding: 0 22px;
  background: #ffffff;
  border-bottom: 1px solid #e6ecf5;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.topbar-left,
.topbar-right {
  display: flex;
  align-items: center;
  gap: 12px;
}

.portal-search {
  width: 320px;
}

.portal-notify {
  margin-right: 6px;
}

.portal-icon-btn {
  border-color: #d8e2f3;
}

.portal-time {
  font-size: 13px;
  color: #5f6b81;
  min-width: 170px;
}

.portal-user {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  cursor: pointer;
  color: #2e3a59;
  font-size: 14px;
}

.portal-content {
  padding: 22px;
}

@media (max-width: 980px) {
  .portal-shell {
    grid-template-columns: 1fr;
  }
  .portal-sidebar {
    display: none;
  }
  .portal-search {
    width: 180px;
  }
}
</style>

