<script setup>
import HelloWorld from './components/HelloWorld.vue'
import panel from './components/panel.vue'
import upload from './components/upload.vue'
</script>

<template>
  <div id="upload">
    <upload ></upload>
  </div>
  <div id="panel">
    <panel></panel>
  </div>
</template>

<style scoped>
#upload {
  position: absolute;
  top: 60%;
  left: 50%;
}
#panel {
  position: absolute;
  top: 10%;
  left: 40%;
}
</style>
