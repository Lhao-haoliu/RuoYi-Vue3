<template>
  <template v-if="item && !item.hidden && item.meta">
    <el-sub-menu v-if="hasChildren" :index="menuIndex">
      <template #title>
        <el-icon><component :is="resolveIcon(item.meta.icon)" /></el-icon>
        <span>{{ item.meta.title }}</span>
      </template>
      <portal-nav-item
        v-for="child in visibleChildren"
        :key="`${menuIndex}-${child.path}`"
        :item="child"
        :base-path="menuIndex"
      />
    </el-sub-menu>
    <el-menu-item v-else :index="menuIndex">
      <el-icon><component :is="resolveIcon(item.meta.icon)" /></el-icon>
      <span>{{ item.meta.title }}</span>
    </el-menu-item>
  </template>
</template>

<script setup>
import * as ElIcons from '@element-plus/icons-vue'
defineOptions({ name: 'PortalNavItem' })

const props = defineProps({
  item: {
    type: Object,
    required: true
  },
  basePath: {
    type: String,
    default: ''
  }
})

const iconMap = {
  dashboard: 'DataLine',
  user: 'User',
  peoples: 'UserFilled',
  tree: 'Share',
  'tree-table': 'Grid',
  edit: 'EditPen',
  message: 'Bell',
  log: 'DocumentCopy',
  monitor: 'Monitor',
  tool: 'Tools',
  chart: 'PieChart',
  form: 'Document',
  list: 'Tickets',
  job: 'Timer',
  guide: 'Compass',
  system: 'Setting',
  code: 'Cpu',
  build: 'Operation',
  swagger: 'Connection',
  menu: 'Menu',
  post: 'Suitcase',
  dict: 'Collection',
  redis: 'Coin',
  online: 'View'
}

const visibleChildren = computed(() => (props.item.children || []).filter((i) => !i.hidden && i.meta))
const hasChildren = computed(() => visibleChildren.value.length > 0)
const menuIndex = computed(() => normalizePath(props.basePath, props.item.path || ''))

function normalizePath(basePath, routePath) {
  if (/^https?:\/\//.test(routePath)) {
    return routePath
  }
  const raw = `${basePath || ''}/${routePath || ''}`.replace(/\/+/g, '/')
  return raw.startsWith('/') ? raw : `/${raw}`
}

function resolveIcon(name) {
  const key = iconMap[name] || 'Menu'
  return ElIcons[key] || ElIcons.Menu
}
</script>
