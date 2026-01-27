<template>
  <RouterView />
</template>

<style scoped></style>

<script>
import { ref } from 'vue'
import ribbon from './components/ribbon.js'

// 立即注册ribbon对象到window对象，确保WPS加载插件时就能访问到
window.ribbon = ribbon

export default {
  setup() {
    const message = ref('你好，HND插件')

    return {
      message
    }
  }
}
</script>
