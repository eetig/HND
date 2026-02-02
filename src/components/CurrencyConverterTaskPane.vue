<template>
  <div class="currency-container">
    <iframe 
      src="https://webfem.com/tools/huilv/" 
      class="currency-iframe"
      frameborder="0"
      allowfullscreen
    ></iframe>
  </div>
</template>

<script>
export default {
  name: 'CurrencyConverterTaskPane'
}
</script>

<style scoped>
.currency-container {
  width: 100%;
  height: 100%;
  overflow: hidden;
  display: flex;
  justify-content: center;
  align-items: center;
  padding: 20px;
}

.currency-iframe {
  width: 100%;
  max-width: 800px;
  height: 100%;
  min-height: 600px;
  border: none;
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}
</style>