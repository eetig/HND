<template>
  <div class="qrcode-taskpane">
    <h2>二维码生成器</h2>
    <div class="iframe-container">
      <iframe
        :src="qrcodeUrl"
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
  name: 'QRCodeTaskPane',
  data() {
    return {
      qrcodeUrl: 'https://zh.qr-code.net/'
    }
  }
}
</script>

<style scoped>
.qrcode-taskpane {
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