<script setup lang="ts">
import { ref } from "vue";
import VLookupPage from './components/VLookupPage.vue';

interface FeatureItem {
  id: string;
  name: string;
  icon?: string;
}

interface FeatureGroup {
  id: string;
  name: string;
  expanded: boolean;
  items: FeatureItem[];
}

// 应用标题
const appTitle = "办公软件";

// 搜索框内容
const searchQuery = ref("");

// 功能分组数据
const featureGroups = ref<FeatureGroup[]>([
  {
    id: "office",
    name: "Office",
    expanded: true,
    items: [
      {
        id: "vlookup",
        name: "VLOOKUP 助手",
        icon: "📊"
      }
    ]
  }
]);

// 切换分组展开/折叠状态
function toggleGroup(groupId: string) {
  const group = featureGroups.value.find(g => g.id === groupId);
  if (group) {
    group.expanded = !group.expanded;
  }
}

// 当前页面状态
const currentPage = ref<'main' | 'vlookup'>('main');

// 点击功能项
function handleFeatureClick(featureId: string) {
  console.log(`点击了功能: ${featureId}`);
  if (featureId === 'vlookup') {
    currentPage.value = 'vlookup';
  }
}

// 从VLookup页面返回
function handleBackFromVLookup() {
  currentPage.value = 'main';
}
</script>

<template>
  <div class="app-container">
    <!-- 根据当前页面状态显示不同内容 -->
    <template v-if="currentPage === 'main'">
    <!-- 顶部标题栏 -->
    <header class="app-header">
      <div class="header-left">
        <button class="menu-button">☰</button>
        <h1 class="app-title">{{ appTitle }}</h1>
      </div>
      <div class="header-right">
        <button class="search-button">🔍</button>
      </div>
    </header>
    
    <!-- 搜索框 -->
    <div class="search-container">
      <input 
        v-model="searchQuery"
        type="text" 
        placeholder="搜索功能..."
        class="search-input"
      />
    </div>
    
    <!-- 功能分组列表 -->
    <div class="feature-list">
      <div 
        v-for="group in featureGroups" 
        :key="group.id"
        :class="['feature-group', { 'office-group': group.id === 'office' }]"
      >
        <!-- 分组标题 -->
        <div 
          class="group-header"
          @click="toggleGroup(group.id)"
        >
          <span class="group-name">{{ group.name }}</span>
          <span class="group-toggle">
            {{ group.expanded ? '▼' : '▶' }}
          </span>
        </div>
        
        <!-- 功能项列表 -->
        <div 
          v-if="group.expanded"
          class="group-items"
        >
          <div 
            v-for="item in group.items" 
            :key="item.id"
            class="feature-item"
            @click="handleFeatureClick(item.id)"
          >
            <span class="feature-icon">{{ item.icon || '🔧' }}</span>
            <span class="feature-name">{{ item.name }}</span>
          </div>
        </div>
      </div>
    </div>
    </template>
    
    <!-- VLOOKUP页面 -->
    <VLookupPage v-else-if="currentPage === 'vlookup'" @back="handleBackFromVLookup" />
  </div>
</template>

<style scoped>
/* 应用容器样式 */
.app-container {
  width: 100%;
  height: 100vh;
  background-color: var(--color-background);
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

/* 头部样式 */
.app-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 12px 16px;
  background-color: var(--color-background);
  border-bottom: 1px solid var(--color-border);
}

.header-left {
  display: flex;
  align-items: center;
  gap: 12px;
}

.menu-button,
.search-button {
  width: 32px;
  height: 32px;
  border: none;
  background: none;
  font-size: 18px;
  cursor: pointer;
  border-radius: 4px;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: background-color 0.2s;
}

.menu-button:hover,
.search-button:hover {
  background-color: rgba(0, 0, 0, 0.05);
}

.app-title {
  font-size: 18px;
  font-weight: 600;
  color: #333;
  margin: 0;
}

/* 搜索框样式 */
.search-container {
  padding: 12px 16px;
}

.search-input {
  width: 100%;
  padding: 10px 16px;
  border: 1px solid var(--color-primary-light);
  border-radius: 20px;
  background-color: var(--color-background);
  font-size: 14px;
  outline: none;
  transition: border-color 0.2s;
  box-sizing: border-box;
}

.search-input:focus {
  border-color: var(--color-primary);
  box-shadow: 0 0 0 2px rgba(20, 184, 166, 0.2);
}

/* 功能列表样式 */
.feature-list {
  flex: 1;
  overflow-y: auto;
  padding: 0 16px 16px;
}

.feature-group {
  margin-bottom: 16px;
}

/* 分组标题样式 */
  .group-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 10px 12px;
    background-color: #f8f8f8;
    border-radius: 8px;
    cursor: pointer;
    margin-bottom: 8px;
    border-left: 3px solid var(--color-text-secondary);
    transition: background-color 0.2s;
  }
  
  /* Office分组特定样式 */
  .office-group .group-header {
    background-color: var(--color-primary-lighter);
    border-left-color: var(--color-primary);
  }
  
  .office-group .group-header:hover {
    background-color: var(--color-primary-light);
  }

.group-header:hover {
  background-color: var(--color-border-light);
}

.group-name {
  font-size: 14px;
  font-weight: 600;
  color: #333;
}

.group-toggle {
  font-size: 12px;
  color: #666;
}

/* 功能项网格样式 */
.group-items {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(100px, 1fr));
  gap: 12px;
}

.feature-item {
  display: flex;
  flex-direction: column;
  align-items: center;
  padding: 12px 8px;
  background-color: var(--color-background);
  border-radius: 8px;
  cursor: pointer;
  transition: all 0.2s;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
  text-align: center;
}

.feature-item:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.feature-icon {
  font-size: 24px;
  margin-bottom: 6px;
}

.feature-name {
  font-size: 13px;
  color: var(--color-text-primary);
  line-height: 1.2;
}

/* 滚动条样式 */
.feature-list::-webkit-scrollbar {
  width: 6px;
}

.feature-list::-webkit-scrollbar-track {
  background: #f1f1f1;
  border-radius: 3px;
}

.feature-list::-webkit-scrollbar-thumb {
  background: var(--color-border);
  border-radius: 3px;
}

.feature-list::-webkit-scrollbar-thumb:hover {
  background: var(--color-text-secondary);
}

/* 响应式调整 */
@media (max-width: 768px) {
  .group-items {
    grid-template-columns: repeat(auto-fill, minmax(80px, 1fr));
  }
  
  .feature-name {
    font-size: 12px;
  }
}
</style>

<style>
/* 全局颜色变量定义 */
:root {
  /* 主色调 - 绿色系列 */
  --color-primary: #14b8a6;
  --color-primary-dark: #0d9488;
  --color-primary-light: #e6f7f4;
  --color-primary-lighter: #f0fdfa;
  
  /* 中性色调 */
  --color-background: #ffffff;
  --color-border: #e0e0e0;
  --color-border-light: #f0f0f0;
  --color-text-primary: #333333;
  --color-text-secondary: #666666;
}

/* 全局样式重置 */
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

html, body {
  height: 100%;
  overflow: hidden;
}

#app {
  height: 100vh;
}

/* 移除默认按钮样式 */
button {
  background: none;
  border: none;
  font-family: inherit;
}
</style>