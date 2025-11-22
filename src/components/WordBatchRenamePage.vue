<template>
  <div class="word-batch-rename-container">
    <!-- 头部导航 -->
    <div class="header">
      <button class="back-button" @click="handleBack">
        ← 返回
      </button>
      <h1 class="page-title">Word 批量命名</h1>
    </div>

    <div class="content-wrapper">
      <!-- 操作流程步骤指示器 -->
      <div class="steps-indicator">
        <div v-for="(step, index) in steps" :key="index" class="step-item" :class="{
          'active': currentStep === index,
          'completed': currentStep > index
        }">
          <div class="step-number">{{ index + 1 }}</div>
          <div class="step-text">{{ step }}</div>
        </div>
      </div>

      <!-- 主内容区域 -->
      <div class="main-content">
        <!-- 步骤1：选择Excel文件 -->
        <div v-if="currentStep === 0" class="step-content">
          <div class="step-header">
            <h2>步骤 1：选择Excel文件</h2>
            <p class="step-description">请选择包含命名列表的Excel文件</p>
          </div>

          <div class="file-selection-area">
            <input type="file" ref="excelFileInput" @change="handleExcelFileChange" accept=".xlsx,.xls"
              style="display: none" />

            <div class="file-drop-area" @click="triggerExcelFileInput" @dragover.prevent @dragenter.prevent
              @drop="handleFileDrop">
              <div v-if="!excelFileName" class="drop-placeholder">
                <span class="drop-icon">📁</span>
                <p>点击或拖拽Excel文件到此处</p>
                <small>支持 .xlsx 和 .xls 格式</small>
              </div>

              <div v-else class="file-selected">
                <span class="file-icon">📊</span>
                <div class="file-info">
                  <p class="file-name">{{ excelFileName }}</p>
                  <button class="change-file-btn" @click.stop="triggerExcelFileInput">
                    更换文件
                  </button>
                </div>
              </div>
            </div>

            <!-- 工作表选择 -->
            <div v-if="worksheets.length > 0" class="worksheet-selection">
              <label for="worksheet-select">选择工作表：</label>
              <select id="worksheet-select" v-model="selectedWorksheet" @change="onWorksheetChange"
                class="select-input">
                <option v-for="sheet in worksheets" :key="sheet" :value="sheet">
                  {{ sheet }}
                </option>
              </select>
            </div>

            <!-- 命名列选择 -->
            <div v-if="columns.length > 0" class="column-selection">
              <label for="name-column-select">命名列：</label>
              <select id="name-column-select" v-model="renameColumn" class="select-input">
                <option v-for="col in columns" :key="col" :value="col">
                  {{ col }}
                </option>
              </select>
            </div>

            <!-- 命名列表预览 -->
            <div v-if="renameList.length > 0" class="name-list-preview">
              <h3>命名列表预览 ({{ renameList.length }} 条)</h3>
              <div class="preview-container">
                <div v-for="(name, index) in renameList.slice(0, 10)" :key="index" class="preview-item">
                  <span class="preview-index">{{ index + 1 }}.</span>
                  <span class="preview-name">{{ name }}</span>
                </div>
                <div v-if="renameList.length > 10" class="more-items">
                  还有 {{ renameList.length - 10 }} 条...
                </div>
              </div>
            </div>
          </div>

          <div class="button-group">
            <button class="primary-button" @click="nextStep" :disabled="renameList.length === 0">
              下一步
            </button>
          </div>
        </div>

        <!-- 步骤2：选择目标文件夹 -->
        <div v-else-if="currentStep === 1" class="step-content">
          <div class="step-header">
            <h2>步骤 2：选择目标文件夹</h2>
            <p class="step-description">请选择包含需要重命名的Word文件的文件夹</p>
          </div>

          <div class="folder-selection-area">
            <button class="folder-select-btn" @click="selectFolder" :disabled="isScanning">
              {{ isScanning ? '扫描中...' : '选择文件夹' }}
            </button>

            <div v-if="folderPath" class="folder-info">
              <span class="folder-icon">📂</span>
              <p class="folder-path">{{ folderPath }}</p>
            </div>

            <!-- Word文件列表 -->
            <div v-if="wordFiles.length > 0" class="word-files-list">
              <h3>Word文件列表 ({{ wordFiles.length }} 个)</h3>
              <div class="files-container">
                <div v-for="(file, index) in wordFiles.slice(0, 10)" :key="index" class="file-item">
                  <span class="file-icon">📄</span>
                  <span class="file-name">{{ file.name }}</span>
                </div>
                <div v-if="wordFiles.length > 10" class="more-items">
                  还有 {{ wordFiles.length - 10 }} 个文件...
                </div>
              </div>

              <!-- 文件数量与命名数量对比 -->
              <div class="comparison-info">
                <p v-if="wordFiles.length !== renameList.length" class="warning-message">
                  ⚠️ 警告：Word文件数量({{ wordFiles.length }})与命名数量({{ renameList.length }})不匹配！
                </p>
                <p v-else class="success-message">
                  ✓ 文件数量与命名数量匹配
                </p>
              </div>
            </div>
          </div>

          <div class="button-group">
            <button class="secondary-button" @click="prevStep">
              上一步
            </button>
            <button class="primary-button" @click="nextStep" :disabled="wordFiles.length === 0">
              下一步
            </button>
          </div>
        </div>

        <!-- 步骤3：设置命名规则 -->
        <div v-else-if="currentStep === 2" class="step-content">
          <div class="step-header">
            <h2>步骤 3：设置命名规则</h2>
            <p class="step-description">配置文件命名的格式和规则</p>
          </div>

          <div class="naming-settings">
            <!-- 前缀设置 -->
            <div class="setting-item">
              <label for="prefix-input">文件前缀：</label>
              <input id="prefix-input" type="text" v-model="filePrefix" placeholder="请输入前缀（可选）" class="text-input" />
            </div>

            <!-- 序号格式设置 -->
            <div class="setting-item">
              <label for="sequence-length">序号位数：</label>
              <input id="sequence-length" type="number" v-model.number="sequenceLength" min="1" max="10"
                class="number-input" />
            </div>

            <!-- 命名预览 -->
            <div class="preview-section">
              <h3>命名预览</h3>
              <div class="preview-container">
                <div v-for="(name, index) in renameList.slice(0, 5)" :key="index" class="preview-item">
                  <span class="preview-name">{{ generateFileNamePreview(name, index) }}</span>
                </div>
              </div>
            </div>
          </div>

          <div class="button-group">
            <button class="secondary-button" @click="prevStep">
              上一步
            </button>
            <button class="primary-button" @click="nextStep">
              下一步
            </button>
          </div>
        </div>

        <!-- 步骤4：执行重命名 -->
        <div v-else-if="currentStep === 3" class="step-content">
          <div class="step-header">
            <h2>步骤 4：执行重命名</h2>
            <p class="step-description">确认重命名设置并开始处理</p>
          </div>

          <div class="confirmation-section">
            <div class="summary-card">
              <h3>重命名概要</h3>
              <table class="summary-table">
                <tbody>
                  <tr>
                    <td>命名列表来源：</td>
                    <td>{{ excelFileName }}</td>
                  </tr>
                  <tr>
                    <td>目标文件夹：</td>
                    <td>{{ folderPath }}</td>
                  </tr>
                  <tr>
                    <td>文件总数：</td>
                    <td>{{ wordFiles.length }} 个</td>
                  </tr>
                  <tr>
                    <td>命名格式：</td>
                    <td>{{ filePrefix }}[命名]-{{ formatSequence(1, sequenceLength) }}</td>
                  </tr>
                </tbody>
              </table>
            </div>

            <div class="warnings-section" v-if="hasWarnings">
              <h3>⚠️ 注意事项</h3>
              <ul class="warning-list">
                <li v-if="wordFiles.length !== renameList.length">
                  Word文件数量与命名数量不匹配，将按照文件顺序重命名现有文件
                </li>
                <li v-if="filePrefix.trim() === ''">
                  未设置文件前缀，将直接使用命名+序号格式
                </li>
              </ul>
            </div>

            <!-- 重命名进度 -->
            <div v-if="isRenaming" class="progress-section">
              <div class="progress-bar">
                <div class="progress-fill" :style="{ width: progressPercentage + '%' }"></div>
              </div>
              <p class="progress-text">
                {{ renameProgress.current }} / {{ renameProgress.total }} 个文件
              </p>
            </div>

            <!-- 重命名结果 -->
            <div v-if="renameResult" class="result-section">
              <div :class="['result-status', renameResult.success ? 'success' : 'error']">
                <h3>{{ renameResult.success ? '✅ 重命名成功' : '❌ 重命名失败' }}</h3>
                <p>{{ renameResult.message }}</p>
              </div>

              <!-- 目标文件夹提示 -->
              <div v-if="renameResult.targetFolder" class="target-folder-info">
                <h4>📁 目标文件夹</h4>
                <p class="folder-info">
                  所有文件已复制到：<code>{{ renameResult.targetFolder.split(/[/\\]/).pop() }}</code>
                </p>
                <p class="folder-path">完整路径：{{ renameResult.targetFolder }}</p>
              </div>

              <!-- 重复值提示 -->
              <div v-if="renameResult.duplicates && renameResult.duplicates.count > 0" class="duplicate-warning">
                  <h4>⚠️ 重复值提示</h4>
                  <div class="duplicate-info">
                    检测到重复情况：
                    <ul style="margin: 8px 0; padding-left: 20px;">
                      <li v-if="renameResult.duplicates.names && renameResult.duplicates.names.length > 0">
                        Excel列表中的重复值：<strong>{{ renameResult.duplicates.names.join('、') }}</strong>
                      </li>
                      <li>Word文件夹中找到多个匹配文件的情况</li>
                    </ul>
                  </div>
                  
                  <!-- 详细的重复信息列表 -->
                  <div v-if="renameResult.duplicates.details && renameResult.duplicates.details.length > 0" class="duplicate-details">
                    <h5 style="margin-top: 15px; margin-bottom: 10px;">重复详情：</h5>
                    <div v-for="(detail, index) in renameResult.duplicates.details" :key="index" class="duplicate-item">
                      <div class="duplicate-value">
                        <strong>重复值：</strong>{{ detail.value }}
                      </div>
                      <div class="duplicate-reason">
                        <strong>重复原因：</strong>{{ detail.reason }}
                      </div>
                      <div v-if="detail.files && detail.files.length > 0" class="duplicate-files">
                        <strong>相关文件：</strong>
                        <ul style="margin: 5px 0; padding-left: 20px;">
                          <li v-for="(file, fileIndex) in detail.files" :key="fileIndex">{{ file }}</li>
                        </ul>
                      </div>
                    </div>
                  </div>
                  
                  <p class="duplicate-info">
                    已将 <strong>{{ renameResult.duplicates.count }}</strong> 个重复文件复制到 <code>{{ renameResult.targetFolder ? renameResult.targetFolder.split(/[/\\]/).pop() + '/00重名' : '00重名' }}</code> 文件夹（保持原文件名）
                  </p>
                <p class="duplicate-path">完整路径：{{ renameResult.duplicates.folderPath }}</p>
              </div>

              <div v-if="renameResult.logs && renameResult.logs.length > 0" class="logs-section">
                <h4>操作日志</h4>
                <div class="logs-container">
                  <div v-for="(log, index) in renameResult.logs.slice(0, 10)" :key="index" class="log-item"
              :class="log.success ? 'success-log' : 'error-log'">
              <span v-html="highlightKeywords(log.message)"></span>
            </div>
                  <div v-if="renameResult.logs.length > 10" class="more-logs">
                    <button class="more-logs-button" @click="showFullLogs">查看全部 {{ renameResult.logs.length }} 条日志 →</button>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div class="button-group">
            <button v-if="!isRenaming && !renameResult" class="secondary-button" @click="prevStep">
              返回修改
            </button>

            <button v-if="!isRenaming && !renameResult" class="primary-button" @click="startRename">
              开始重命名
            </button>

            <button v-if="renameResult" class="primary-button" @click="resetAndRestart">
              重新开始
            </button>
          </div>
        </div>
      </div>
    </div>
  </div>
  
  <!-- 操作日志弹出层 -->
  <div v-if="showLogsModal" class="modal-overlay" @click="closeLogsModal">
    <div class="modal-content" @click.stop>
      <div class="modal-header">
        <h3>完整操作日志</h3>
        <button class="modal-close-button" @click="closeLogsModal">×</button>
      </div>
      <div class="modal-body">
        <div class="full-logs-container">
          <div v-for="(log, index) in renameResult?.logs || []" :key="index" class="log-item" 
               :class="log.success ? 'success-log' : 'error-log'">
            <span v-html="highlightKeywords(log.message)"></span>
          </div>
        </div>
      </div>
      <div class="modal-footer">
        <button class="secondary-button" @click="closeLogsModal">关闭</button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch } from 'vue';
import * as XLSX from 'xlsx';
// 移除直接的Tauri API导入，改为通过Tauri的invoke调用后端

// 定义事件
const emit = defineEmits<{
  back: []
}>();

// 处理返回操作
function handleBack() {
  emit('back');
}

// 步骤定义
const steps = ['选择Excel文件', '选择目标文件夹', '设置命名规则', '执行重命名'];
const currentStep = ref(0);

// Excel相关状态
const excelFileInput = ref<HTMLInputElement | null>(null);
const excelFile = ref<File | null>(null);
const excelFileName = ref('');
const worksheets = ref<string[]>([]);
const selectedWorksheet = ref('');
const columns = ref<string[]>([]);
const renameColumn = ref('');
const renameList = ref<string[]>([]);

// 文件夹和文件相关状态
const folderPath = ref('');
const wordFiles = ref<{ name: string; path: string }[]>([]);
const isScanning = ref(false);

// 命名规则相关状态
const filePrefix = ref('');
const sequenceLength = ref(3);

// 重命名相关状态
const isRenaming = ref(false);
const renameProgress = ref({ current: 0, total: 0 });
const progressPercentage = computed(() => {
  if (renameProgress.value.total === 0) return 0;
  return (renameProgress.value.current / renameProgress.value.total) * 100;
});
// 重命名结果
const renameResult = ref<{
  success: boolean;
  message: string;
  logs?: Array<{ message: string; success: boolean }>;
  duplicates?: {
    names: string[];
    count: number;
    folderPath: string;
    details?: DuplicateInfo[];
  };
  targetFolder?: string;
} | null>(null);

// 操作日志弹出层控制
const showLogsModal = ref(false);

// 显示完整操作日志
// 显示完整日志
function showFullLogs() {
  showLogsModal.value = true;
}

// 关闭操作日志弹出层
function closeLogsModal() {
  showLogsModal.value = false;
}

// 高亮显示日志中的关键字（Excel关键字和Word匹配关键字）
function highlightKeywords(message: string): string {
  // 根据日志消息格式提取并高亮关键字
  // 处理格式："成功：已复制 "文件名" → "新文件名" 到文件夹"
  // 或 "成功：已复制 "文件名" 到文件夹（保持原文件名）"
  
  // 匹配格式中的文件名和新文件名
  const highlightedMessage = message
    // 高亮原始文件名（用"包围的第一个文件名）
    .replace(/"([^"]+)"(\s+→\s+)?/g, (_match, fileName, arrow) => {
      if (arrow) {
        // 第一个文件名（Word匹配的文件名）
        return `<span class="highlight-keyword">"${fileName}"</span> → `;
      } else {
        // 如果是保持原文件名的情况，高亮文件名
        return `<span class="highlight-keyword">"${fileName}"</span>`;
      }
    })
    // 高亮新文件名（用"包围的第二个文件名）
    .replace(/→\s+"([^"]+)"/g, (_match, newFileName) => {
      return `→ <span class="highlight-keyword">"${newFileName}"</span>`;
    });
  
  return highlightedMessage;
}

// 计算属性：是否有警告
const hasWarnings = computed(() => {
  return wordFiles.value.length !== renameList.value.length || filePrefix.value.trim() === '';
});

// 触发Excel文件选择
function triggerExcelFileInput() {
  excelFileInput.value?.click();
}

// 处理Excel文件选择
function handleExcelFileChange(event: Event) {
  const target = event.target as HTMLInputElement;
  if (target.files && target.files.length > 0) {
    const file = target.files[0];
    processExcelFile(file);
  }
}

// 处理文件拖放
function handleFileDrop(event: DragEvent) {
  event.preventDefault();
  if (event.dataTransfer?.files && event.dataTransfer.files.length > 0) {
    const file = event.dataTransfer.files[0];
    if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
      processExcelFile(file);
    } else {
      alert('请选择有效的Excel文件(.xlsx或.xls)');
    }
  }
}

// 处理Excel文件
function processExcelFile(file: File) {
  excelFile.value = file;
  excelFileName.value = file.name;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });

      // 获取所有工作表名称
      worksheets.value = workbook.SheetNames;
      if (worksheets.value.length > 0) {
        selectedWorksheet.value = worksheets.value[0];
        onWorksheetChange();
      }
    } catch (error) {
      console.error('解析Excel文件失败:', error);
      alert('解析Excel文件失败，请检查文件格式');
    }
  };
  reader.readAsArrayBuffer(file);
}

// 工作表变更处理
function onWorksheetChange() {
  if (!excelFile.value || !selectedWorksheet.value) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[selectedWorksheet.value];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      if (jsonData.length > 0) {
        // 获取第一行作为列名
        // 确保jsonData[0]存在且是数组
        if (jsonData[0] && Array.isArray(jsonData[0])) {
          columns.value = jsonData[0].map((col, index) => {
            // 安全地处理可能的null/undefined值
            return col != null ? String(col) : `列${index + 1}`;
          });
        } else {
          columns.value = [];
        }

        // 默认选择第一列
        if (columns.value && columns.value.length > 0) {
          renameColumn.value = columns.value[0];
          extractNames(jsonData as any[][]);
        }
      }
    } catch (error) {
      console.error('读取工作表数据失败:', error);
    }
  };
  reader.readAsArrayBuffer(excelFile.value);
}

// 监听命名列变更，提取命名
watch(renameColumn, () => {
  if (!excelFile.value || !selectedWorksheet.value || !renameColumn.value) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[selectedWorksheet.value];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      extractNames(jsonData as any[][]);
    } catch (error) {
      console.error('提取命名列表失败:', error);
    }
  };
  reader.readAsArrayBuffer(excelFile.value);
});

// 从Excel数据中提取命名
function extractNames(jsonData: any[][]) {
  // 安全检查
  if (!jsonData || !Array.isArray(jsonData) || jsonData.length <= 1) {
    renameList.value = [];
    return;
  }

  // 找到命名列的索引
  const headerRow = jsonData[0];
  if (!headerRow || !Array.isArray(headerRow)) {
    renameList.value = [];
    return;
  }

  // 安全地查找命名列索引
  const renameColumnIndex = headerRow.findIndex(col => {
    return col != null && String(col) === renameColumn.value;
  });

  if (renameColumnIndex === -1) {
    renameList.value = [];
    return;
  }

  // 提取命名列数据（跳过表头）
  renameList.value = jsonData.slice(1)
    .map(row => {
      // 确保row存在且是数组，然后安全地获取指定列的值
      if (!row || !Array.isArray(row)) return '';
      const value = row[renameColumnIndex];
      return value != null ? String(value).trim() : '';
    })
    .filter(name => name && name.length > 0);
}

// 选择文件夹并扫描Word文件
async function selectFolder() {
  try {
    isScanning.value = true;

    // 调用后端API来选择文件夹和扫描文件
    const { invoke } = await import('@tauri-apps/api/core');
    const result = await invoke<{
      folder_path?: string;
      word_files?: Array<{ name: string; path: string }>;
      error?: string;
    }>('select_and_scan_word_folder');

    if (result && typeof result === 'object') {
      const { folder_path, word_files, error } = result;

      if (error) {
        console.error('操作失败:', error);
        alert(`操作失败: ${error}`);
        wordFiles.value = [];
        folderPath.value = '';
      } else if (folder_path && Array.isArray(word_files)) {
        folderPath.value = folder_path;
        wordFiles.value = word_files.sort((a, b) => a.name.localeCompare(b.name));
      }
    } else {
      console.log('操作取消或失败');
    }
  } catch (error) {
    console.error('调用后端API失败:', error);
    alert('调用后端API失败，请重试');
    wordFiles.value = [];
  } finally {
    isScanning.value = false;
  }
}

// 生成文件名预览（带扩展名参数）
function generateFileNamePreview(name: string, index: number, originalFileName?: string): string {
  const sequence = formatSequence(index + 1, sequenceLength.value);
  // 如果提供了原文件名，提取其扩展名；否则默认使用 .docx
  let extension = '.docx';
  if (originalFileName) {
    const lastDotIndex = originalFileName.lastIndexOf('.');
    if (lastDotIndex > 0) {
      extension = originalFileName.substring(lastDotIndex);
    }
  }
  return `${filePrefix.value}${name}-${sequence}${extension}`;
}

// 格式化序号
function formatSequence(num: number, length: number): string {
  return num.toString().padStart(length, '0');
}

// 下一步
function nextStep() {
  if (currentStep.value < steps.length - 1) {
    currentStep.value++;
  }
}

// 上一步
function prevStep() {
  if (currentStep.value > 0) {
    currentStep.value--;
  }
}


// 定义重复文件信息接口
export interface DuplicateInfo {
  value: string;            // 重复的值
  indices: number[];        // 重复的索引列表
  reason: string;           // 重复理由
  files?: string[];         // 对应的文件名列表（可选）
}

// 检测重复值
function findDuplicates(list: string[]): DuplicateInfo[] {
  const duplicates: DuplicateInfo[] = [];
  const nameMap = new Map<string, number[]>();

  // 记录每个值出现的索引
  list.forEach((name, index) => {
    if (!nameMap.has(name)) {
      nameMap.set(name, []);
    }
    nameMap.get(name)!.push(index);
  });

  // 找出出现多次的值并添加重复理由
  nameMap.forEach((indices, name) => {
    if (indices.length > 1) {
      duplicates.push({
        value: name,
        indices: indices,
        reason: `Excel列表中重复出现（索引：${indices.join(', ')}）`,
        files: [] // 初始化空文件列表
      });
    }
  });

  return duplicates;
}

// 格式化日期时间为 YYYYMMDD-HHmmss
function formatDateTime(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${year}${month}${day}-${hours}${minutes}${seconds}`;
}

// 开始重命名
async function startRename() {
  isRenaming.value = true;
  renameProgress.value = { current: 0, total: renameList.value.length };
  renameResult.value = null;

  const logs: Array<{ message: string; success: boolean }> = [];
  let successCount = 0;
  let errorCount = 0;
  let skippedCount = 0;
  let duplicateCount = 0;

  // 记录已重命名的文件，避免重复处理
  const renamedFiles = new Set<string>();

  // 创建带索引的列表，并按长度排序（长的优先），保持原始索引用于后续处理
  const indexedNames = renameList.value.map((name, index) => ({ name, index }));
  indexedNames.sort((a, b) => {
    // 先按长度降序（长的在前）
    if (b.name.length !== a.name.length) {
      return b.name.length - a.name.length;
    }
    // 长度相同则按原始索引排序
    return a.index - b.index;
  });

  // 检测重复值（基于原始列表）
  const duplicateInfos = findDuplicates(renameList.value);
  const duplicateNames = duplicateInfos.map(info => info.value);
  const duplicateIndices = new Set<number>();
  
  // 收集所有重复值的索引（除了第一个）
  duplicateInfos.forEach((info) => {
    // 保留第一个，将其他索引标记为重复
    for (let i = 1; i < info.indices.length; i++) {
      duplicateIndices.add(info.indices[i]);
    }
  });
  
  // 用于存储Word文件匹配重复的信息
  const fileDuplicateInfos: DuplicateInfo[] = [];
  // 用于跟踪哪些文件已经被记录为重复

  // 创建重命名目标文件夹
  let targetFolderPath = '';
  if (folderPath.value) {
    const dateTime = formatDateTime(new Date());
    const folderName = `重命名-${dateTime}`;
    const pathParts = folderPath.value.split(/[/\\]/);
    pathParts.push(folderName);
    targetFolderPath = pathParts.join('/');
  }

  // 创建00重名文件夹 - 现在放在时间戳文件夹内
  let duplicateFolderPath = '';
  if (folderPath.value && targetFolderPath) {
    // 将00重名文件夹创建在时间戳文件夹内
    const pathParts = targetFolderPath.split(/[/\\]/);
    pathParts.push('00重名');
    duplicateFolderPath = pathParts.join('/');
  }

  try {
    // 导入Tauri API
    const { invoke } = await import('@tauri-apps/api/core');

    // 按照排序后的列表顺序进行处理（长的名字优先）
    for (let idx = 0; idx < indexedNames.length; idx++) {
      const { name, index: originalIndex } = indexedNames[idx];
      const isDuplicateInList = duplicateIndices.has(originalIndex);

      // 更新进度（基于原始索引）
      renameProgress.value.current = idx + 1;

      // 在Word文件列表中查找所有包含该命名的文件（可能多个）
      // 优先匹配更精确的（更长的名字）
      const matchedFiles = wordFiles.value.filter(file => {
        // 跳过已经处理过的文件
        if (renamedFiles.has(file.path)) {
          return false;
        }
        // 检查文件名是否包含该命名
        const fileNameNormalized = file.name.replace(/\s/g, '').toLowerCase();
        const nameNormalized = name.replace(/\s/g, '').toLowerCase();
        
        // 如果文件名包含该命名
        if (fileNameNormalized.includes(nameNormalized)) {
          // 检查是否有更长的名字应该匹配这个文件
          // 遍历已处理的名字（在排序列表中，前面的都是更长的）
          for (let prevIdx = 0; prevIdx < idx; prevIdx++) {
            const prevName = indexedNames[prevIdx].name;
            const prevNameNormalized = prevName.replace(/\s/g, '').toLowerCase();
            // 如果文件名包含更长的名字，且更长的名字也包含当前名字，则当前名字不应该匹配
            if (fileNameNormalized.includes(prevNameNormalized) && 
                prevNameNormalized.includes(nameNormalized) && 
                prevNameNormalized.length > nameNormalized.length) {
              // 更长的名字应该优先匹配，当前短名字不匹配
              return false;
            }
          }
          return true;
        }
        return false;
      });

      // 情况1：Excel列表中的值重复
      // 情况2：在Word文件夹中找到多个匹配文件
      const isDuplicate = isDuplicateInList || matchedFiles.length > 1;
      
      // 如果在Word文件夹中找到多个匹配文件且尚未记录过
      if (matchedFiles.length > 1) {
        // 创建新的文件重复信息
        const newFileDuplicateInfo: DuplicateInfo = {
          value: name,
          indices: [originalIndex],
          reason: `在Word文件夹中找到多个匹配文件`,
          files: matchedFiles.map(file => file.name)
        };
        fileDuplicateInfos.push(newFileDuplicateInfo);
      }

      if (matchedFiles.length > 0) {
        // 处理所有匹配的文件
        for (let j = 0; j < matchedFiles.length; j++) {
          const matchedFile = matchedFiles[j];
          // 为每个文件确定重复理由
          let duplicateReason = '';
          if (isDuplicateInList) {
            duplicateReason = `Excel列表中重复（索引：${duplicateInfos.find(info => info.value === name)?.indices.join(', ')}）`;
          } else if (matchedFiles.length > 1) {
            duplicateReason = `Word文件夹中找到${matchedFiles.length}个匹配文件`;
          }
          
          try {
            let newFileName = '';
            let targetFolder = '';

            if (isDuplicate) {
              // 两种情况都放到00重名文件夹，保持原文件名
              targetFolder = duplicateFolderPath;
              newFileName = matchedFile.name; // 保持原文件名
              duplicateCount++;
            } else {
              // 正常重命名，使用原逻辑生成文件名，放到目标文件夹
              targetFolder = targetFolderPath;
              newFileName = generateFileNamePreview(name, originalIndex, matchedFile.name);
            }

            // 复制文件到目标文件夹
            await invoke<string>('copy_file_to_folder', {
              filePath: matchedFile.path,
              targetFolder: targetFolder,
              newName: newFileName
            });

            renamedFiles.add(matchedFile.path);
            successCount++;
            
            const targetFolderName = targetFolderPath.split(/[/\\]/).pop();
            const folderName = isDuplicate ? `${targetFolderName}/00重名` : targetFolderName;
            if (isDuplicate) {
              logs.push({
                message: `成功：已复制 "${matchedFile.name}" 到${folderName}文件夹（保持原文件名）${duplicateReason ? ` - ${duplicateReason}` : ''}`,
                success: true
              });
            } else {
              logs.push({
                message: `成功：已复制 "${matchedFile.name}" → "${newFileName}" 到${folderName}文件夹`,
                success: true
              });
            }
          } catch (error: any) {
            // 复制失败
            errorCount++;
            const errorMsg = error?.toString() || '未知错误';
            logs.push({
              message: `失败：无法复制 "${matchedFile.name}" - ${errorMsg}`,
              success: false
            });
          }
        }
      } else {
        // 未找到匹配的文件，跳过
        skippedCount++;
        logs.push({
          message: `跳过：未找到包含 "${name}" 的文件`,
          success: true
        });
      }
    }

    // 完成所有重命名操作
    isRenaming.value = false;
    let message = '';

    // 构建消息，包含重复值提示
    const targetFolderName = targetFolderPath.split(/[/\\]/).pop() || '目标文件夹';
    // 重复文件夹现在在时间戳文件夹内创建，不再需要单独声明变量
    
    if (duplicateCount > 0) {
      const duplicateMsg = `检测到重复情况（Excel列表重复或Word文件夹中多个匹配），已复制${duplicateCount}个重复文件到"${targetFolderName}/00重名"文件夹（保持原文件名）。`;
      if (errorCount === 0 && skippedCount === 0) {
        message = `所有文件处理完成（成功：${successCount}个，重复：${duplicateCount}个）。正常文件已复制到"${targetFolderName}"文件夹，重复文件已复制到"${targetFolderName}/00重名"文件夹。${duplicateMsg}`;
      } else if (errorCount === 0) {
        message = `处理完成（成功：${successCount}个，跳过：${skippedCount}个，重复：${duplicateCount}个）。正常文件已复制到"${targetFolderName}"文件夹，重复文件已复制到"${targetFolderName}/00重名"文件夹。${duplicateMsg}`;
      } else {
        message = `部分文件处理失败（成功：${successCount}个，失败：${errorCount}个，跳过：${skippedCount}个，重复：${duplicateCount}个）。正常文件已复制到"${targetFolderName}"文件夹，重复文件已复制到"${targetFolderName}/00重名"文件夹。${duplicateMsg}`;
      }
    } else {
      if (errorCount === 0 && skippedCount === 0) {
        message = `所有文件处理成功（${successCount}个）。所有文件已复制到"${targetFolderName}"文件夹。`;
      } else if (errorCount === 0) {
        message = `处理完成（成功：${successCount}个，跳过：${skippedCount}个）。所有文件已复制到"${targetFolderName}"文件夹。`;
      } else {
        message = `部分文件处理失败（成功：${successCount}个，失败：${errorCount}个，跳过：${skippedCount}个）。所有文件已复制到"${targetFolderName}"文件夹。`;
      }
    }
    
    // 合并Excel重复信息和文件匹配重复信息
    const allDuplicateInfos = [...duplicateInfos, ...fileDuplicateInfos];
    
    renameResult.value = {
      success: errorCount === 0,
      message,
      logs,
      duplicates: duplicateCount > 0 ? {
        names: duplicateNames,
        count: duplicateCount,
        folderPath: duplicateFolderPath,
        details: allDuplicateInfos // 添加详细的重复信息
      } : undefined,
      targetFolder: targetFolderPath
    };
  } catch (error) {
    console.error('重命名过程出错:', error);
    isRenaming.value = false;
    renameResult.value = {
      success: false,
      message: `重命名过程出错: ${error}`,
      logs
    };
  }
}

// 重置并重新开始
function resetAndRestart() {
  currentStep.value = 0;
  excelFile.value = null;
  excelFileName.value = '';
  worksheets.value = [];
  selectedWorksheet.value = '';
  columns.value = [];
  renameColumn.value = '';
  renameList.value = [];
  folderPath.value = '';
  wordFiles.value = [];
  filePrefix.value = '';
  sequenceLength.value = 3;
  renameResult.value = null;

  // 清空文件输入
  if (excelFileInput.value) {
    excelFileInput.value.value = '';
  }
}
</script>

<style scoped>
.word-batch-rename-container {
  height: 100vh;
  display: flex;
  flex-direction: column;
  background-color: #f5f7fa;
}

/* 头部样式 */
.header {
  display: flex;
  align-items: center;
  padding: 20px;
  background-color: #ffffff;
  border-bottom: 1px solid #e4e7ed;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
}

.back-button {
  background: none;
  border: none;
  font-size: 16px;
  color: #606266;
  cursor: pointer;
  padding: 8px 12px;
  border-radius: 4px;
  transition: all 0.3s ease;
}

.back-button:hover {
  background-color: #f0f2f5;
  color: #409eff;
}

.page-title {
  margin: 0;
  margin-left: 20px;
  font-size: 24px;
  font-weight: 600;
  color: #303133;
}

/* 内容区域样式 */
.content-wrapper {
  flex: 1;
  padding: 30px;
  overflow-y: auto;
}

/* 步骤指示器 */
.steps-indicator {
  display: flex;
  justify-content: space-between;
  margin-bottom: 40px;
  max-width: 800px;
  margin-left: auto;
  margin-right: auto;
}

.step-item {
  display: flex;
  flex-direction: column;
  align-items: center;
  flex: 1;
  position: relative;
}

.step-item:not(:last-child)::after {
  content: '';
  position: absolute;
  top: 15px;
  right: -50%;
  width: 100%;
  height: 2px;
  background-color: #e4e7ed;
  z-index: 1;
}

.step-item.completed:not(:last-child)::after {
  background-color: #409eff;
}

.step-number {
  width: 32px;
  height: 32px;
  border-radius: 50%;
  background-color: #e4e7ed;
  color: #909399;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: 600;
  z-index: 2;
  margin-bottom: 8px;
}

.step-item.active .step-number {
  background-color: #409eff;
  color: #ffffff;
}

.step-item.completed .step-number {
  background-color: #67c23a;
  color: #ffffff;
}

.step-text {
  font-size: 14px;
  color: #606266;
}

.step-item.active .step-text {
  color: #409eff;
  font-weight: 500;
}

.step-item.completed .step-text {
  color: #67c23a;
}

/* 主内容区域 */
.main-content {
  max-width: 800px;
  margin: 0 auto;
  background-color: #ffffff;
  border-radius: 8px;
  box-shadow: 0 2px 12px 0 rgba(0, 0, 0, 0.1);
  padding: 30px;
}

/* 步骤内容 */
.step-content {
  display: flex;
  flex-direction: column;
  gap: 24px;
}

.step-header {
  text-align: center;
  margin-bottom: 16px;
}

.step-header h2 {
  margin: 0 0 8px 0;
  font-size: 20px;
  font-weight: 600;
  color: #303133;
}

.step-description {
  margin: 0;
  color: #909399;
  font-size: 14px;
}

/* 文件选择区域 */
.file-selection-area,
.folder-selection-area {
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.file-drop-area {
  border: 2px dashed #dcdfe6;
  border-radius: 8px;
  padding: 40px;
  text-align: center;
  cursor: pointer;
  transition: all 0.3s ease;
  background-color: #fafafa;
}

.file-drop-area:hover {
  border-color: #409eff;
  background-color: #ecf5ff;
}

.drop-placeholder .drop-icon {
  font-size: 48px;
  margin-bottom: 16px;
  display: block;
}

.drop-placeholder p {
  margin: 0 0 8px 0;
  font-size: 16px;
  color: #606266;
}

.drop-placeholder small {
  color: #909399;
  font-size: 14px;
}

.file-selected {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 16px;
}

.file-selected .file-icon {
  font-size: 32px;
}

.file-info {
  text-align: left;
}

.file-name {
  margin: 0 0 8px 0;
  font-size: 16px;
  color: #303133;
  max-width: 400px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.change-file-btn {
  background-color: #409eff;
  color: white;
  border: none;
  padding: 6px 12px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
  transition: background-color 0.3s ease;
}

.change-file-btn:hover {
  background-color: #66b1ff;
}

/* 工作表和列选择 */
.worksheet-selection,
.column-selection {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 16px;
  background-color: #f8f9fa;
  border-radius: 6px;
}

.worksheet-selection label,
.column-selection label {
  font-size: 14px;
  color: #606266;
  font-weight: 500;
  min-width: 80px;
}

.select-input {
  flex: 1;
  padding: 8px 12px;
  border: 1px solid #dcdfe6;
  border-radius: 4px;
  font-size: 14px;
  transition: border-color 0.3s ease;
}

.select-input:focus {
  outline: none;
  border-color: #409eff;
}

/* 预览列表 */
.name-list-preview,
.word-files-list,
.preview-section {
  padding: 16px;
  background-color: #f8f9fa;
  border-radius: 6px;
}

.name-list-preview h3,
.word-files-list h3,
.preview-section h3 {
  margin: 0 0 16px 0;
  font-size: 16px;
  font-weight: 600;
  color: #303133;
}

.preview-container,
.files-container {
  max-height: 200px;
  overflow-y: auto;
  background-color: #ffffff;
  border: 1px solid #e4e7ed;
  border-radius: 4px;
  padding: 8px;
}

.preview-item,
.file-item {
  padding: 8px 12px;
  display: flex;
  align-items: center;
  gap: 12px;
  border-bottom: 1px solid #f0f0f0;
}

.preview-item:last-child,
.file-item:last-child {
  border-bottom: none;
}

.preview-index {
  width: 30px;
  color: #909399;
  font-size: 14px;
}

.preview-name,
.file-name {
  flex: 1;
  font-size: 14px;
  color: #303133;
}

.more-items {
  text-align: center;
  padding: 8px;
  color: #909399;
  font-size: 14px;
  font-style: italic;
}

.more-logs {
  margin-top: 10px;
  text-align: center;
}

.more-logs-button {
  background-color: var(--color-primary-light);
  border: 1px solid var(--color-primary);
  border-radius: 6px;
  padding: 8px 16px;
  color: var(--color-primary-dark);
  font-size: 14px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.3s ease;
}

.more-logs-button:hover {
  background-color: var(--color-primary);
  color: white;
  transform: translateY(-1px);
}

/* 弹出层样式 */
.modal-overlay {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background-color: rgba(0, 0, 0, 0.5);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 1000;
  padding: 20px;
}

.modal-content {
  background-color: white;
  border-radius: 12px;
  width: 100%;
  max-width: 800px;
  max-height: 80vh;
  display: flex;
  flex-direction: column;
  box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
}

.modal-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 20px 24px;
  border-bottom: 1px solid #e5e7eb;
}

.modal-header h3 {
  margin: 0;
  font-size: 18px;
  font-weight: 600;
  color: #1a1a1a;
}

.modal-close-button {
  background: none;
  border: none;
  font-size: 24px;
  color: #6b7280;
  cursor: pointer;
  padding: 0;
  width: 32px;
  height: 32px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 4px;
  transition: all 0.2s ease;
}

.modal-close-button:hover {
  background-color: #f3f4f6;
  color: #374151;
}

.modal-body {
  flex: 1;
  padding: 0;
  overflow: hidden;
}

.full-logs-container {
  max-height: 60vh;
  overflow-y: auto;
  padding: 0 24px;
}

.modal-footer {
    padding: 16px 24px;
    border-top: 1px solid #e5e7eb;
    display: flex;
    justify-content: flex-end;
  }

  /* 关键字高亮样式 */
  .highlight-keyword {
    background-color: #dbeafe; /* 浅蓝色背景 */
    color: #1e40af; /* 深蓝色文字 */
    font-weight: bold;
    padding: 2px 4px;
    border-radius: 3px;
    border: 1px solid #bfdbfe;
  }

/* 文件夹选择按钮 */
.folder-select-btn {
  background-color: #409eff;
  color: white;
  border: none;
  padding: 12px 24px;
  border-radius: 6px;
  font-size: 16px;
  font-weight: 500;
  cursor: pointer;
  transition: background-color 0.3s ease;
  align-self: center;
}

.folder-select-btn:hover:not(:disabled) {
  background-color: #66b1ff;
}

.folder-select-btn:disabled {
  background-color: #c0c4cc;
  cursor: not-allowed;
}

.folder-info {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 16px;
  background-color: #f8f9fa;
  border-radius: 6px;
}

.folder-info .folder-icon {
  font-size: 24px;
}

.folder-path {
  margin: 0;
  font-size: 14px;
  color: #606266;
  word-break: break-all;
}

/* 比较信息 */
.comparison-info {
  padding: 16px;
  border-radius: 6px;
  border-left: 4px solid #e6a23c;
  background-color: #fdf6ec;
}

.warning-message {
  margin: 0;
  color: #e6a23c;
  font-size: 14px;
  font-weight: 500;
}

.success-message {
  margin: 0;
  color: #67c23a;
  font-size: 14px;
  font-weight: 500;
}

/* 命名设置 */
.naming-settings {
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.setting-item {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.setting-item label {
  font-size: 14px;
  color: #606266;
  font-weight: 500;
}

.text-input,
.number-input {
  padding: 8px 12px;
  border: 1px solid #dcdfe6;
  border-radius: 4px;
  font-size: 14px;
  transition: border-color 0.3s ease;
}

.text-input:focus,
.number-input:focus {
  outline: none;
  border-color: #409eff;
}

.number-input {
  width: 100px;
}

/* 确认区域 */
.confirmation-section {
  display: flex;
  flex-direction: column;
  gap: 24px;
}

.summary-card {
  padding: 20px;
  background-color: #f8f9fa;
  border-radius: 6px;
}

.summary-card h3 {
  margin: 0 0 16px 0;
  font-size: 16px;
  font-weight: 600;
  color: #303133;
}

.summary-table {
  width: 100%;
  border-collapse: collapse;
}

.summary-table td {
  padding: 8px 0;
  border-bottom: 1px solid #e4e7ed;
  font-size: 14px;
}

.summary-table td:first-child {
  color: #606266;
  font-weight: 500;
  width: 30%;
}

.summary-table td:last-child {
  color: #303133;
  word-break: break-all;
}

/* 警告区域 */
.warnings-section {
  padding: 16px;
  background-color: #fdf6ec;
  border: 1px solid #fde2e2;
  border-radius: 6px;
}

.warnings-section h3 {
  margin: 0 0 12px 0;
  font-size: 16px;
  font-weight: 600;
  color: #e6a23c;
}

.warning-list {
  margin: 0;
  padding-left: 20px;
  color: #e6a23c;
  font-size: 14px;
}

.warning-list li {
  margin-bottom: 8px;
}

/* 进度区域 */
.progress-section {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.progress-bar {
  height: 8px;
  background-color: #e4e7ed;
  border-radius: 4px;
  overflow: hidden;
}

.progress-fill {
  height: 100%;
  background-color: #409eff;
  transition: width 0.3s ease;
}

.progress-text {
  margin: 0;
  text-align: center;
  font-size: 14px;
  color: #606266;
}

/* 结果区域 */
.result-section {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.result-status {
  padding: 20px;
  border-radius: 6px;
  text-align: center;
}

.result-status.success {
  background-color: #f0f9ff;
  border: 1px solid #b3d8ff;
}

.result-status.error {
  background-color: #fef0f0;
  border: 1px solid #fde2e2;
}

.result-status h3 {
  margin: 0 0 8px 0;
  font-size: 18px;
  font-weight: 600;
}

.result-status.success h3 {
  color: #67c23a;
}

.result-status.error h3 {
  color: #f56c6c;
}

.result-status p {
  margin: 0;
  font-size: 14px;
  color: #606266;
}

/* 日志区域 */
.logs-section h4 {
  margin: 0 0 12px 0;
  font-size: 14px;
  font-weight: 600;
  color: #303133;
}

.logs-container {
  max-height: 200px;
  overflow-y: auto;
  background-color: #f8f9fa;
  border: 1px solid #e4e7ed;
  border-radius: 4px;
  padding: 8px;
}

.log-item {
  padding: 6px 12px;
  font-size: 13px;
  border-bottom: 1px solid #f0f0f0;
}

.log-item:last-child {
  border-bottom: none;
}

.success-log {
  color: #67c23a;
}

.error-log {
  color: #f56c6c;
}

/* 重复警告区域 */
.duplicate-warning {
  background-color: #fff9e6;
  border: 1px solid #ffd93d;
  border-radius: 6px;
  padding: 15px;
  margin-bottom: 20px;
}

.duplicate-warning h4 {
  margin-top: 0;
  color: #d97706;
}

.duplicate-info {
  margin-bottom: 10px;
}

.duplicate-path {
  color: #6b7280;
  font-size: 14px;
  margin-top: 10px;
}

.duplicate-details {
  background-color: #ffffff;
  border: 1px solid #e5e7eb;
  border-radius: 4px;
  padding: 10px;
  margin: 10px 0;
}

.duplicate-item {
  border-bottom: 1px solid #f3f4f6;
  padding: 8px 0;
}

.duplicate-item:last-child {
  border-bottom: none;
}

.duplicate-value, .duplicate-reason {
  margin-bottom: 5px;
  line-height: 1.5;
}

.duplicate-files {
  margin-top: 5px;
}

.duplicate-files li {
  color: #6b7280;
  font-size: 14px;
}

/* 按钮组 */
.button-group {
  display: flex;
  justify-content: center;
  gap: 16px;
  margin-top: 24px;
}

.primary-button,
.secondary-button {
  padding: 10px 24px;
  border-radius: 4px;
  font-size: 16px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.3s ease;
  border: none;
}

.primary-button {
  background-color: #409eff;
  color: white;
}

.primary-button:hover:not(:disabled) {
  background-color: #66b1ff;
}

.primary-button:disabled {
  background-color: #c0c4cc;
  cursor: not-allowed;
}

.secondary-button {
  background-color: #ffffff;
  color: #606266;
  border: 1px solid #dcdfe6;
}

.secondary-button:hover {
  color: #409eff;
  border-color: #c6e2ff;
  background-color: #ecf5ff;
}

/* 响应式设计 */
@media (max-width: 768px) {
  .content-wrapper {
    padding: 16px;
  }

  .main-content {
    padding: 20px;
  }

  .steps-indicator {
    flex-direction: column;
    gap: 16px;
  }

  .step-item:not(:last-child)::after {
    display: none;
  }

  .step-item {
    flex-direction: row;
    justify-content: flex-start;
    gap: 12px;
  }

  .file-drop-area {
    padding: 24px;
  }

  .button-group {
    flex-direction: column;
  }

  .summary-table td {
    display: block;
    width: 100%;
  }

  .summary-table td:first-child {
    font-weight: 600;
    margin-bottom: 4px;
  }
}
</style>