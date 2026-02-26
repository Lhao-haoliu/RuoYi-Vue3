<template>
  <section class="page-container">
    <header class="page-container__header">
      <div class="page-container__title-wrap">
        <h1 class="page-container__title">{{ title }}</h1>
        <p v-if="description" class="page-container__desc">{{ description }}</p>
      </div>
      <div class="page-container__actions">
        <slot name="actions" />
      </div>
    </header>

    <section v-if="$slots.search" class="page-container__search-card">
      <slot name="search" />
    </section>

    <section class="page-container__table-card">
      <slot v-if="hasTableSlot" name="table" />
      <slot v-else />
    </section>

    <slot v-if="hasTableSlot" />
  </section>
</template>

<script setup>
import { computed, useSlots } from "vue"

const slots = useSlots()
const hasTableSlot = computed(() => Boolean(slots.table))

defineProps({
  title: {
    type: String,
    default: ""
  },
  description: {
    type: String,
    default: ""
  }
})
</script>
