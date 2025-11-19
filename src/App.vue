<script setup lang="ts">
import { ref } from "vue";
import VLookupPage from './components/VLookupPage.vue';

interface FeatureItem {
  id: string;
  name: string;
  icon?: string;
  enabled?: boolean; // 是否已实现
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
        name: "Office 工具",
        expanded: true,
        items: [
          {
            id: "vlookup",
            name: "VLOOKUP 助手",
            icon: "📊",
            enabled: true
          },
          {
            id: "excel-formatter",
            name: "Excel 格式化器",
            icon: "📈",
            enabled: false
          },
          {
            id: "ppt-templates",
            name: "PPT 模板库",
            icon: "📑",
            enabled: false
          },
          {
            id: "word-tools",
            name: "Word 辅助工具",
            icon: "📝",
            enabled: false
          }
        ]
      },
  {
        id: "data-analysis",
        name: "数据分析",
        expanded: true,
        items: [
          {
            id: "data-visualization",
            name: "数据可视化",
            icon: "📊",
            enabled: false
          },
          {
            id: "statistical-analysis",
            name: "统计分析助手",
            icon: "📈",
            enabled: false
          },
          {
            id: "data-cleaner",
            name: "数据清洗工具",
            icon: "🧹",
            enabled: false
          }
        ]
      },
  {
        id: "productivity",
        name: "效率工具",
        expanded: true,
        items: [
          {
            id: "batch-rename",
            name: "批量重命名",
            icon: "📋",
            enabled: false
          },
          {
            id: "pdf-tools",
            name: "PDF 处理工具",
            icon: "📄",
            enabled: false
          },
          {
            id: "screenshot-manager",
            name: "截图管理",
            icon: "🖼️",
            enabled: false
          }
        ]
      },
  {
        id: "ai-assistant",
        name: "AI 助手",
        expanded: true,
        items: [
          {
            id: "text-summarizer",
            name: "文本摘要",
            icon: "📝",
            enabled: false
          },
          {
            id: "translation-tool",
            name: "翻译工具",
            icon: "🌐",
            enabled: false
          },
          {
            id: "content-generator",
            name: "内容生成器",
            icon: "🤖",
            enabled: false
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

// 获取功能项描述
function getDescription(featureId: string): string {
  const descriptions: Record<string, string> = {
    'vlookup': '快速生成Excel VLOOKUP公式',
    'excel-formatter': '自动格式化Excel表格样式',
    'ppt-templates': '提供专业PPT模板和素材',
    'word-tools': 'Word文档编辑和格式转换',
    'data-visualization': '直观展示数据图表',
    'statistical-analysis': '基础统计计算和分析',
    'data-cleaner': '快速清洗和整理数据',
    'batch-rename': '批量修改文件名和格式',
    'pdf-tools': 'PDF转换、合并和编辑',
    'screenshot-manager': '智能管理和标注截图',
    'text-summarizer': '快速生成文本摘要',
    'translation-tool': '多语言翻译和校对',
    'content-generator': 'AI辅助内容创作'
  };
  return descriptions[featureId] || '办公助手功能';
}
</script>

<template>
  <div class="app-container">
    <!-- 根据当前页面状态显示不同内容 -->
    <div v-if="currentPage === 'main'">
      <!-- 顶部标题栏 -->
      <header class="app-header">
        <div class="logo-container">
          <span class="app-icon">💼</span>
          <h1 class="app-title">{{ appTitle }}</h1>
        </div>
      </header>
      
      <!-- 主内容区域 -->
      <main class="main-content">
        <!-- 功能卡片网格 -->
        <div class="features-grid">
          <div 
            v-for="group in featureGroups" 
            :key="group.id"
            class="features-section"
          >
            <!-- 分组标题 -->
            <div 
              class="group-header"
              @click="toggleGroup(group.id)"
            >
              <h2 class="group-title">{{ group.name }}</h2>
              <span class="group-toggle" :class="{ 'expanded': group.expanded }">
                {{ group.expanded ? '−' : '+' }}
              </span>
            </div>
            
            <!-- 功能项卡片列表 - 仅在展开状态显示 -->
            <div v-if="group.expanded" class="cards-container">
              <div 
                v-for="item in group.items" 
                :key="item.id"
                :class="['feature-card', { 'feature-card-disabled': !item.enabled }]"
                @click="item.enabled !== false && handleFeatureClick(item.id)"
                :title="item.name"
              >
                <div class="card-icon">{{ item.icon || '🔧' }}</div>
                <h3 class="card-title">{{ item.name }}</h3>
                <div class="card-description">
                  {{ getDescription(item.id) }}
                </div>
                <span class="card-arrow">→</span>
              </div>
            </div>
          </div>
        </div>
      </main>
    </div>
    
    <!-- VLOOKUP页面 -->
    <VLookupPage v-else-if="currentPage === 'vlookup'" @back="handleBackFromVLookup" />
  </div>
</template>

<style scoped>
/* 应用容器样式 */
.app-container {
  width: 100%;
  min-height: 100vh;
  background-color: var(--color-background);
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
  display: flex;
  flex-direction: column;
  overflow: visible;
}

/* 头部样式 - 大气简洁版 */
.app-header {
  display: flex;
  justify-content: center;
  align-items: center;
  padding: 24px 20px;
  background-color: var(--color-background);
  border-bottom: 1px solid var(--color-border);
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.03);
}

.logo-container {
  display: flex;
  align-items: center;
  gap: 16px;
}

.app-icon {
  font-size: 36px;
  transition: transform 0.3s ease;
}

.app-icon:hover {
  transform: scale(1.1);
}

.app-title {
  font-size: 28px;
  font-weight: 700;
  color: var(--color-text-primary);
  margin: 0;
  letter-spacing: -0.5px;
}

/* 主内容区域 */
.main-content {
  flex: 1;
  display: flex;
  padding: 40px 20px;
  overflow-y: auto;
}

/* 功能卡片网格 */
.features-grid {
  width: 100%;
  max-width: 1200px;
  display: flex;
  flex-direction: column;
  gap: 40px;
}

.features-section {
  width: 100%;
  background: rgba(255, 255, 255, 0.8);
  border-radius: 12px;
  border: 1px solid var(--color-border);
  overflow: hidden;
  transition: all 0.3s ease;
}

.features-section:hover {
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.04);
}

/* 分组标题样式 */
.group-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 20px 24px;
  background: linear-gradient(135deg, var(--color-background) 0%, #f9fafb 100%);
  border-bottom: 1px solid var(--color-border);
  cursor: pointer;
  transition: all 0.3s ease;
}

.group-header:hover {
  background: var(--color-primary-light);
}

.group-title {
  font-size: 20px;
  font-weight: 600;
  color: var(--color-text-primary);
  margin: 0;
}

.group-toggle {
  font-size: 24px;
  font-weight: 300;
  color: var(--color-primary);
  transition: transform 0.3s ease;
}

.group-toggle.expanded {
  transform: rotate(0deg);
}

/* 功能卡片容器 */
.cards-container {
  display: flex;
  flex-wrap: wrap;
  gap: 24px;
  padding: 24px;
  justify-content: center;
  animation: slideDown 0.3s ease-out;
}

/* 展开动画 */
@keyframes slideDown {
  from {
    opacity: 0;
    max-height: 0;
    padding-top: 0;
    padding-bottom: 0;
  }
  to {
    opacity: 1;
    max-height: 1000px;
    padding-top: 24px;
    padding-bottom: 24px;
  }
}

/* 功能卡片样式 */
.feature-card {
  position: relative;
  width: 320px;
  padding: 40px 30px;
  background: linear-gradient(135deg, var(--color-background) 0%, #f9fafb 100%);
  border: 1px solid var(--color-border);
  border-radius: 12px;
  cursor: pointer;
  transition: all 0.3s ease;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.04);
  display: flex;
  flex-direction: column;
  align-items: center;
  text-align: center;
  overflow: hidden;
}

.feature-card:hover {
  transform: translateY(-8px);
  box-shadow: 0 16px 48px rgba(0, 0, 0, 0.08);
  border-color: var(--color-primary);
}

.feature-card::before {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  height: 4px;
  background: linear-gradient(90deg, var(--color-primary), var(--color-primary-dark));
  transform: scaleX(0);
  transform-origin: left;
  transition: transform 0.3s ease;
}

.feature-card:hover::before {
  transform: scaleX(1);
}

.card-icon {
  font-size: 64px;
  margin-bottom: 24px;
  transition: transform 0.3s ease;
}

.feature-card:hover .card-icon {
  transform: scale(1.1);
}

.card-title {
  font-size: 22px;
  font-weight: 600;
  color: var(--color-text-primary);
  margin: 0 0 12px;
  transition: color 0.3s ease;
}

.feature-card:hover .card-title {
  color: var(--color-primary);
}

.card-description {
  font-size: 15px;
  color: var(--color-text-secondary);
  line-height: 1.6;
  margin: 0 0 20px;
}

.card-arrow {
  position: absolute;
  bottom: 24px;
  right: 24px;
  font-size: 18px;
  color: var(--color-primary);
  opacity: 0;
  transform: translateX(10px);
  transition: all 0.3s ease;
}

.feature-card:hover .card-arrow {
  opacity: 1;
  transform: translateX(0);
}

/* 未实现功能的置灰样式 */
.feature-card-disabled {
  opacity: 0.6;
  cursor: not-allowed;
  filter: grayscale(70%);
}

.feature-card-disabled:hover {
  transform: none;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.04);
  border-color: var(--color-border);
}

.feature-card-disabled:hover::before {
  transform: scaleX(0);
}

.feature-card-disabled:hover .card-icon {
  transform: none;
}

.feature-card-disabled:hover .card-title {
  color: var(--color-text-primary);
}

.feature-card-disabled:hover .card-arrow {
  opacity: 0;
  transform: translateX(10px);
}

/* 滚动条样式 */
.main-content::-webkit-scrollbar {
  width: 8px;
}

.main-content::-webkit-scrollbar-track {
  background: var(--color-background);
}

.main-content::-webkit-scrollbar-thumb {
  background: var(--color-border);
  border-radius: 4px;
}

.main-content::-webkit-scrollbar-thumb:hover {
  background: var(--color-text-secondary);
}

/* 响应式调整 */
@media (max-width: 768px) {
  .app-header {
    padding: 20px 16px;
  }
  
  .app-title {
    font-size: 24px;
  }
  
  .app-icon {
    font-size: 32px;
  }
  
  .main-content {
    padding: 30px 16px;
  }
  
  .features-grid {
    gap: 24px;
  }
  
  .group-header {
    padding: 16px 20px;
  }
  
  .group-title {
    font-size: 18px;
  }
  
  .cards-container {
    padding: 16px;
    gap: 16px;
  }
  
  .feature-card {
    width: 100%;
    max-width: 320px;
    padding: 32px 24px;
  }
  
  .card-icon {
    font-size: 56px;
  }
  
  .card-title {
    font-size: 20px;
  }
}

@media (max-width: 480px) {
  .app-title {
    font-size: 20px;
  }
  
  .app-icon {
    font-size: 28px;
  }
  
  .feature-card {
    padding: 24px 20px;
  }
  
  .card-icon {
    font-size: 48px;
  }
  
  .card-title {
    font-size: 18px;
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
  overflow: auto;
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