<template>
  <div class="qrcode-taskpane">
    <iframe
      :src="qrcodeUrl"
      frameborder="0"
      width="100%"
      height="750"
      scrolling="yes"
    ></iframe>
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
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

.qrcode-taskpane {
  width: 100%;
  height: 100%;
  padding: 0;
  margin: 0;
  font-family: Arial, sans-serif;
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