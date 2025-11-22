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
  background: linear-gradient(180deg, #fafbfc 0%, #ffffff 100%);
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
  display: flex;
  flex-direction: column;
  overflow: visible;
}

/* 头部样式 - 高端简洁版 */
.app-header {
  display: flex;
  justify-content: center;
  align-items: center;
  padding: 32px 20px 28px;
  background: transparent;
  border-bottom: none;
  position: relative;
}

.app-header::after {
  content: '';
  position: absolute;
  bottom: 0;
  left: 50%;
  transform: translateX(-50%);
  width: 60px;
  height: 2px;
  background: linear-gradient(90deg, transparent, var(--color-primary), transparent);
  opacity: 0.3;
}

.logo-container {
  display: flex;
  align-items: center;
  gap: 12px;
}

.app-icon {
  font-size: 32px;
  transition: transform 0.4s cubic-bezier(0.34, 1.56, 0.64, 1);
  filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.05));
}

.app-icon:hover {
  transform: scale(1.08) rotate(5deg);
}

.app-title {
  font-size: 26px;
  font-weight: 600;
  color: #1a1a1a;
  margin: 0;
  letter-spacing: -0.3px;
  background: linear-gradient(135deg, #1a1a1a 0%, #4a4a4a 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
}

/* 主内容区域 */
.main-content {
  flex: 1;
  display: flex;
  padding: 48px 24px;
  overflow-y: auto;
  max-width: 1400px;
  margin: 0 auto;
  width: 100%;
}

/* 功能卡片网格 */
.features-grid {
  width: 100%;
  display: flex;
  flex-direction: column;
  gap: 32px;
}

.features-section {
  width: 100%;
  background: #ffffff;
  border-radius: 16px;
  border: 1px solid rgba(0, 0, 0, 0.06);
  overflow: hidden;
  transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
}

.features-section:hover {
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
  border-color: rgba(0, 0, 0, 0.08);
  transform: translateY(-2px);
}

/* 分组标题样式 */
.group-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 20px 28px;
  background: transparent;
  border-bottom: 1px solid rgba(0, 0, 0, 0.04);
  cursor: pointer;
  transition: all 0.3s ease;
}

.group-header:hover {
  background: rgba(0, 0, 0, 0.01);
}

.group-title {
  font-size: 18px;
  font-weight: 600;
  color: #1a1a1a;
  margin: 0;
  letter-spacing: -0.2px;
}

.group-toggle {
  font-size: 20px;
  font-weight: 300;
  color: #666;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  width: 28px;
  height: 28px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 6px;
  background: rgba(0, 0, 0, 0.02);
}

.group-toggle:hover {
  background: rgba(0, 0, 0, 0.05);
  color: var(--color-primary);
}

.group-toggle.expanded {
  transform: rotate(0deg);
}

/* 功能卡片容器 */
.cards-container {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
  gap: 20px;
  padding: 28px;
  animation: slideDown 0.4s cubic-bezier(0.4, 0, 0.2, 1);
}

/* 展开动画 */
@keyframes slideDown {
  from {
    opacity: 0;
    transform: translateY(-10px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

/* 功能卡片样式 */
.feature-card {
  position: relative;
  padding: 32px 24px;
  background: #ffffff;
  border: 1px solid rgba(0, 0, 0, 0.06);
  border-radius: 12px;
  cursor: pointer;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
  display: flex;
  flex-direction: column;
  align-items: center;
  text-align: center;
  overflow: hidden;
}

.feature-card::before {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: linear-gradient(135deg, rgba(20, 184, 166, 0.02) 0%, rgba(20, 184, 166, 0) 100%);
  opacity: 0;
  transition: opacity 0.3s ease;
}

.feature-card:hover {
  transform: translateY(-4px);
  box-shadow: 0 12px 32px rgba(0, 0, 0, 0.12);
  border-color: rgba(20, 184, 166, 0.2);
}

.feature-card:hover::before {
  opacity: 1;
}

.card-icon {
  font-size: 48px;
  margin-bottom: 20px;
  transition: transform 0.4s cubic-bezier(0.34, 1.56, 0.64, 1);
  filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.08));
}

.feature-card:hover .card-icon {
  transform: scale(1.15) translateY(-4px);
}

.card-title {
  font-size: 18px;
  font-weight: 600;
  color: #1a1a1a;
  margin: 0 0 10px;
  transition: color 0.3s ease;
  letter-spacing: -0.2px;
}

.feature-card:hover .card-title {
  color: var(--color-primary);
}

.card-description {
  font-size: 13px;
  color: #666;
  line-height: 1.6;
  margin: 0;
  font-weight: 400;
}

.card-arrow {
  position: absolute;
  bottom: 20px;
  right: 20px;
  font-size: 16px;
  color: var(--color-primary);
  opacity: 0;
  transform: translate(8px, 8px);
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  width: 32px;
  height: 32px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 8px;
  background: rgba(20, 184, 166, 0.08);
}

.feature-card:hover .card-arrow {
  opacity: 1;
  transform: translate(0, 0);
}

/* 未实现功能的置灰样式 */
.feature-card-disabled {
  opacity: 0.5;
  cursor: not-allowed;
  filter: grayscale(100%);
  background: #f8f9fa;
}

.feature-card-disabled:hover {
  transform: none;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
  border-color: rgba(0, 0, 0, 0.06);
}

.feature-card-disabled:hover::before {
  opacity: 0;
}

.feature-card-disabled:hover .card-icon {
  transform: none;
}

.feature-card-disabled:hover .card-title {
  color: #1a1a1a;
}

.feature-card-disabled:hover .card-arrow {
  opacity: 0;
  transform: translate(8px, 8px);
}

/* 滚动条样式 */
.main-content::-webkit-scrollbar {
  width: 6px;
}

.main-content::-webkit-scrollbar-track {
  background: transparent;
}

.main-content::-webkit-scrollbar-thumb {
  background: rgba(0, 0, 0, 0.2);
  border-radius: 3px;
}

.main-content::-webkit-scrollbar-thumb:hover {
  background: rgba(0, 0, 0, 0.3);
}

/* 响应式调整 */
@media (max-width: 768px) {
  .app-header {
    padding: 24px 16px 20px;
  }
  
  .app-title {
    font-size: 22px;
  }
  
  .app-icon {
    font-size: 28px;
  }
  
  .main-content {
    padding: 32px 16px;
  }
  
  .features-grid {
    gap: 24px;
  }
  
  .group-header {
    padding: 16px 20px;
  }
  
  .group-title {
    font-size: 16px;
  }
  
  .cards-container {
    grid-template-columns: 1fr;
    padding: 20px;
    gap: 16px;
  }
  
  .feature-card {
    padding: 28px 20px;
  }
  
  .card-icon {
    font-size: 44px;
  }
  
  .card-title {
    font-size: 17px;
  }
}

@media (max-width: 480px) {
  .app-title {
    font-size: 20px;
  }
  
  .app-icon {
    font-size: 24px;
  }
  
  .feature-card {
    padding: 24px 18px;
  }
  
  .card-icon {
    font-size: 40px;
  }
  
  .card-title {
    font-size: 16px;
  }
  
  .card-description {
    font-size: 12px;
  }
}
</style>

<style>
/* 全局颜色变量定义 */
:root {
  /* 主色调 - 现代青色系列 */
  --color-primary: #14b8a6;
  --color-primary-dark: #0d9488;
  --color-primary-light: #e6f7f4;
  --color-primary-lighter: #f0fdfa;
  
  /* 中性色调 - 高端简洁 */
  --color-background: #ffffff;
  --color-border: rgba(0, 0, 0, 0.08);
  --color-border-light: rgba(0, 0, 0, 0.04);
  --color-text-primary: #1a1a1a;
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
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
}

#app {
  height: 100vh;
}

/* 移除默认按钮样式 */
button {
  background: none;
  border: none;
  font-family: inherit;
  cursor: pointer;
}
</style>