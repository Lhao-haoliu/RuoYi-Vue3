<template>
  <div>
    <template v-for="(item, index) in options" :key="item.value + '-' + index">
      <el-tag
        v-if="isValueMatch(item.value)"
        :disable-transitions="true"
        :type="item.elTagType || 'info'"
        :class="['dict-tag-pill', item.elTagClass]"
        effect="light"
      >
        {{ item.label }}
      </el-tag>
    </template>
    <template v-if="unmatch && showValue">
      {{ unmatchArray | handleArray }}
    </template>
  </div>
</template>

<script setup>
const unmatchArray = ref([])

const props = defineProps({
  options: {
    type: Array,
    default: null,
  },
  value: [Number, String, Array],
  showValue: {
    type: Boolean,
    default: true,
  },
  separator: {
    type: String,
    default: ',',
  }
})

const values = computed(() => {
  if (props.value === null || typeof props.value === 'undefined' || props.value === '') return []
  if (typeof props.value === 'number' || typeof props.value === 'boolean') return [props.value]
  return Array.isArray(props.value) ? props.value.map(item => '' + item) : String(props.value).split(props.separator)
})

const unmatch = computed(() => {
  unmatchArray.value = []
  if (props.value === null || typeof props.value === 'undefined' || props.value === '' || !Array.isArray(props.options) || props.options.length === 0) return false
  let hasUnmatched = false
  values.value.forEach(item => {
    if (!props.options.some(v => v.value == item)) {
      unmatchArray.value.push(item)
      hasUnmatched = true
    }
  })
  return hasUnmatched
})

function handleArray(array) {
  if (array.length === 0) return ''
  return array.reduce((pre, cur) => {
    return pre + ' ' + cur
  })
}

function isValueMatch(itemValue) {
  return values.value.some(val => val == itemValue)
}
</script>

<style scoped>
.dict-tag-pill + .dict-tag-pill {
  margin-left: 8px;
}
</style>
