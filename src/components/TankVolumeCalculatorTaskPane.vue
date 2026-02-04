<template>
  <div class="tank-volume-calculator-taskpane">
    <iframe
      :src="calculatorUrl"
      frameborder="0"
      width="100%"
      height="750"
      scrolling="yes"
    ></iframe>
  </div>
</template>

<script>
export default {
  name: 'TankVolumeCalculatorTaskPane',
  data() {
    return {
      calculatorUrl: 'https://www.23bei.com/tool/521.html'
    }
  }
}
</script>

<style scoped>
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

.tank-volume-calculator-taskpane {
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