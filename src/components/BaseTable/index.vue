<template>
  <div class="base-table">
    <div v-if="$slots.toolbar" class="base-table__toolbar">
      <slot name="toolbar" />
    </div>
    <div class="base-table__content">
      <slot />
    </div>
    <div v-if="$slots.pagination" class="base-table__pagination">
      <slot name="pagination" />
    </div>
  </div>
</template>
