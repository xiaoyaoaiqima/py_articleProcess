<template>
  <a-collapse v-model:activeKey="activeKey" ghost style="width: 50rem">
    <a-collapse-panel key="1" header="学校" :show-arrow="false" style="display: inline-block;">
      <div>
        <a-button v-for="school in schools" :key="school" type="text"
                  @click="handleButtonClick('1', school)">{{ school }}
        </a-button>
      </div>
    </a-collapse-panel>
    <a-collapse-panel key="2" header="学院" :show-arrow="false">
      <div>
        <a-button v-for="college in BJcolleges" :key="college" type="text"
                  @click="handleButtonClick('2', college)">{{ college }}
        </a-button>
      </div>
    </a-collapse-panel>
    <a-collapse-panel key="3" header="专业" collapsible="disabled">
      <p>{{ text }}</p>
    </a-collapse-panel>
    <a-collapse-panel key="4" header="教师" collapsible="disabled">
      <p>{{ text }}</p>
    </a-collapse-panel>
    <a-collapse-panel key="4" header="专业课程" collapsible="disabled">
      <p>{{ text }}</p>
    </a-collapse-panel>
  </a-collapse>
</template>
<script setup>
import {BJcolleges, BJlessons, BJmajors, BJteachers, schools} from '../assets/static.js'
import {ref, watch} from 'vue';

const activeKey = ref(['1']);
const results = ref({});

const handleButtonClick = (panelKey, btnKey) => {
  console.log(panelKey)
  activeKey.value = [panelKey]; // 设置激活的 panel key
  // if (panelKey===)
  results.value[panelKey] = btnKey; // 存储点击结果
  // console.log('Button clicked:', btnKey); // 处理按钮点击事件
};

watch(activeKey, val => {
  // console.log(val);
});


const Clicked_btn = ref({
  btn1:0,
  btn2:0,
  btn3:0,
  btn4:0,
  btn5:0
});


watch(results, val => {
  Clicked_btn.value = val['1']; // 获取学校按钮的点击结果
});
</script>
