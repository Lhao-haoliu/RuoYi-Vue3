import { createApp } from 'vue'

import Cookies from 'js-cookie'

import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import locale from 'element-plus/es/locale/lang/zh-cn'

import '@/assets/styles/index.scss' // global css

import App from './App'
import store from './store'
import router from './router'
import directive from './directive' // directive

// 娉ㄥ唽鎸囦护
import plugins from './plugins' // plugins
import { download } from '@/utils/request'

// svg鍥炬爣
import 'virtual:svg-icons-register'
import SvgIcon from '@/components/SvgIcon'
import elementIcons from '@/components/SvgIcon/svgicon'

import './permission' // permission control

import { useDict } from '@/utils/dict'
import { getConfigKey } from "@/api/system/config"
import { parseTime, resetForm, addDateRange, handleTree, selectDictLabel, selectDictLabels } from '@/utils/ruoyi'

// 鍒嗛〉缁勪欢
import Pagination from '@/components/Pagination'
// 鑷畾涔夎〃鏍煎伐鍏风粍浠?
import RightToolbar from '@/components/RightToolbar'
// 瀵屾枃鏈粍浠?
import Editor from "@/components/Editor"
// 鏂囦欢涓婁紶缁勪欢
import FileUpload from "@/components/FileUpload"
// 鍥剧墖涓婁紶缁勪欢
import ImageUpload from "@/components/ImageUpload"
// 鍥剧墖棰勮缁勪欢
import ImagePreview from "@/components/ImagePreview"
// 瀛楀吀鏍囩缁勪欢
import DictTag from '@/components/DictTag'

const app = createApp(App)

// 鍏ㄥ眬鏂规硶鎸傝浇
app.config.globalProperties.useDict = useDict
app.config.globalProperties.download = download
app.config.globalProperties.parseTime = parseTime
app.config.globalProperties.resetForm = resetForm
app.config.globalProperties.handleTree = handleTree
app.config.globalProperties.addDateRange = addDateRange
app.config.globalProperties.getConfigKey = getConfigKey
app.config.globalProperties.selectDictLabel = selectDictLabel
app.config.globalProperties.selectDictLabels = selectDictLabels

// 鍏ㄥ眬缁勪欢鎸傝浇
app.component('DictTag', DictTag)
app.component('Pagination', Pagination)
app.component('FileUpload', FileUpload)
app.component('ImageUpload', ImageUpload)
app.component('ImagePreview', ImagePreview)
app.component('RightToolbar', RightToolbar)
app.component('Editor', Editor)

app.use(router)
app.use(store)
app.use(plugins)
app.use(elementIcons)
app.component('svg-icon', SvgIcon)

directive(app)

// 浣跨敤element-plus 骞朵笖璁剧疆鍏ㄥ眬鐨勫ぇ灏?
app.use(ElementPlus, {
  locale: locale,
  // 鏀寔 large銆乨efault銆乻mall
  size: Cookies.get('size') || 'default'
})

app.mount('#app')

