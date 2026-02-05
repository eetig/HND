<template>
  <div class="moyu-container">
    <iframe
      src="https://sudoku.com/zh/killer/zhuanjia/"
      frameborder="0"
      width="100%"
      height="750"
      scrolling="yes"
    ></iframe>
  </div>
</template>

<script>
export default {
  name: 'MoyuTaskPane'
}
</script>

<style scoped>
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

.moyu-container {
  width: 100%;
  height: 100%;
  padding: 0;
  margin: 0;
  overflow: hidden;
}

iframe {
  width: 100%;
  height: 750px;
  border: none;
  margin: 0;
  padding: 0;
  display: block;
}
</style>