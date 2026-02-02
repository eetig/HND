<template>
  <div class="tank-volume-calculator-taskpane">
    <h2>椭圆封头卧式储罐液位容积计算</h2>
    <div class="iframe-container">
      <iframe
        :src="calculatorUrl"
        frameborder="0"
        width="100%"
        height="600"
        scrolling="yes"
      ></iframe>
    </div>
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
.tank-volume-calculator-taskpane {
  padding: 16px;
  font-family: Arial, sans-serif;
}

h2 {
  color: #333;
  margin-bottom: 20px;
  font-size: 18px;
}

.iframe-container {
  border: 1px solid #ddd;
  border-radius: 4px;
  overflow: hidden;
}

iframe {
  border: none;
}
</style>