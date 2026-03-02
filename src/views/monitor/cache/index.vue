<template>
  <div class="app-container monitor-industrial">
    <el-row :gutter="10">
      <el-col :span="24" class="card-box">
        <el-card>
          <template #header>
            <Monitor style="width: 1em; height: 1em; vertical-align: middle;" />
            <span style="vertical-align: middle;">基本信息</span>
          </template>
          <div class="el-table el-table--enable-row-hover el-table--medium">
            <table cellspacing="0" style="width: 100%">
              <tbody>
                <tr>
                  <td class="el-table__cell is-leaf"><div class="cell">Redis版本</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.info">{{ cache.info.redis_version }}</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell">运行模式</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.info">{{ cache.info.redis_mode == 'standalone' ? '单机' : '集群' }}</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell">端口</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.info">{{ cache.info.tcp_port }}</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell">客户端数</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.info">{{ cache.info.connected_clients }}</div></td>
                </tr>
                <tr>
                  <td class="el-table__cell is-leaf"><div class="cell">运行时间(天)</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.info">{{ cache.info.uptime_in_days }}</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell">使用内存</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.info">{{ cache.info.used_memory_human }}</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell">使用CPU</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.info">{{ parseFloat(cache.info.used_cpu_user_children).toFixed(2) }}</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell">内存配置</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.info">{{ cache.info.maxmemory_human }}</div></td>
                </tr>
                <tr>
                  <td class="el-table__cell is-leaf"><div class="cell">AOF是否开启</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.info">{{ cache.info.aof_enabled == '0' ? '否' : '是' }}</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell">RDB是否成功</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.info">{{ cache.info.rdb_last_bgsave_status }}</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell">Key数量</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.dbSize">{{ cache.dbSize }}</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell">网络入口/出口</div></td>
                  <td class="el-table__cell is-leaf"><div class="cell" v-if="cache.info">{{ cache.info.instantaneous_input_kbps }}kps/{{ cache.info.instantaneous_output_kbps }}kps</div></td>
                </tr>
              </tbody>
            </table>
          </div>
        </el-card>
      </el-col>

      <el-col :span="12" class="card-box">
        <el-card>
          <template #header>
            <PieChart style="width: 1em; height: 1em; vertical-align: middle;" />
            <span style="vertical-align: middle;">命令统计</span>
          </template>
          <div class="el-table el-table--enable-row-hover el-table--medium">
            <div ref="commandstats" style="height: 420px" />
          </div>
        </el-card>
      </el-col>

      <el-col :span="12" class="card-box">
        <el-card>
          <template #header>
            <Odometer style="width: 1em; height: 1em; vertical-align: middle;" />
            <span style="vertical-align: middle;">内存信息</span>
          </template>
          <div class="el-table el-table--enable-row-hover el-table--medium">
            <div ref="usedmemory" style="height: 420px" />
          </div>
        </el-card>
      </el-col>
    </el-row>
  </div>
</template>

<script setup name="Cache">
import { getCache } from '@/api/monitor/cache'
import * as echarts from 'echarts'

const cache = ref([])
const commandstats = ref(null)
const usedmemory = ref(null)
const { proxy } = getCurrentInstance()

function buildGaugeOption(usedValue, usedText) {
  const maxValue = 1000
  const safeValue = Number.isFinite(usedValue) ? usedValue : 0
  const ratio = Math.max(Math.min(safeValue / maxValue, 1), 0)

  return {
    tooltip: {
      formatter: `{b}<br/>{a}: ${usedText}`
    },
    series: [
      {
        name: '峰值',
        type: 'gauge',
        min: 0,
        max: maxValue,
        startAngle: 210,
        endAngle: -30,
        axisLine: {
          lineStyle: {
            width: 8,
            color: [
              [ratio, '#2563EB'],
              [1, 'rgba(37, 99, 235, 0.12)']
            ]
          }
        },
        progress: {
          show: false
        },
        axisTick: { show: false },
        splitLine: { show: false },
        axisLabel: { show: false },
        pointer: {
          show: false
        },
        anchor: { show: false },
        title: {
          offsetCenter: [0, '35%'],
          color: '#4B5563',
          fontSize: 13
        },
        detail: {
          valueAnimation: true,
          offsetCenter: [0, '0%'],
          formatter: usedText,
          color: '#111827',
          fontSize: 24,
          fontWeight: 700
        },
        data: [
          {
            value: safeValue,
            name: '内存消耗'
          }
        ]
      }
    ]
  }
}

function getList() {
  proxy.$modal.loading('正在加载缓存监控数据，请稍候。')
  getCache().then(response => {
    proxy.$modal.closeLoading()
    cache.value = response.data

    const commandstatsIntance = echarts.init(commandstats.value)
    commandstatsIntance.setOption({
      color: ['#2563EB', '#10B981', '#F59E0B', '#EF4444', '#60A5FA', '#A7F3D0'],
      tooltip: {
        trigger: 'item',
        formatter: '{a} <br/>{b} : {c} ({d}%)'
      },
      series: [
        {
          name: '命令',
          type: 'pie',
          radius: ['35%', '65%'],
          center: ['50%', '46%'],
          itemStyle: {
            borderColor: '#ffffff',
            borderWidth: 1
          },
          data: response.data.commandStats,
          animationEasing: 'cubicInOut',
          animationDuration: 800,
          label: {
            color: '#4B5563'
          }
        }
      ]
    })

    const usedmemoryInstance = echarts.init(usedmemory.value)
    const usedMemoryValue = parseFloat(cache.value.info.used_memory_human)
    usedmemoryInstance.setOption(buildGaugeOption(usedMemoryValue, cache.value.info.used_memory_human))

    window.addEventListener('resize', () => {
      commandstatsIntance.resize()
      usedmemoryInstance.resize()
    })
  })
}

getList()
</script>

<style scoped lang="scss">
.monitor-industrial {
  .el-card {
    border: 1px solid #bfdbfe;
    border-radius: 6px;
    box-shadow: none;
  }
}
</style>
