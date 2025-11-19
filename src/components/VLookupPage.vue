<script setup lang="ts">
import { ref, computed } from 'vue';
import * as XLSX from 'xlsx';
import './VLookupPage.css';

const emit = defineEmits<{
  back: []
}>();

// 文件和工作表选择状态
const targetFile = ref<string>('');
const targetFileObj = ref<File | null>(null); // 存储文件对象引用
const targetSheet = ref<string>('');
const sourceFile = ref<string>('');
const sourceFileObj = ref<File | null>(null); // 存储文件对象引用
const sourceSheet = ref<string>(''); // 源文件工作表选择
const lookupValue = ref<string>('');
const generatedFormula = ref<string>('');

// 公式单元格地址和匹配关键字
const targetCellAddress = ref<string>('B3'); // 要填充公式的单元格地址，默认值为B3
const dataToConfigure = ref<string>('姓名'); // 要匹配的关键字，默认值为姓名

// 扩展状态变量 - 用于显示更多信息

const dataSourceRange = ref<string>(''); // 数据表数据查找范围
const matchLookupValue = ref<string>(''); // MATCH函数要查找的值
const matchLookupRange = ref<string>(''); // MATCH函数的数据范围

// 标头匹配功能相关状态
const targetHeaderAddress = ref<string>('');
const targetHeaderCellValue = ref<string>(''); // 目标文件中的标头单元格值
const sourceHeaderAddress = ref<string>('');
const isSearchingHeaders = ref<boolean>(false);
const headerSearchError = ref<string>('');
const isHeaderSearchComplete = ref<boolean>(false);

// 姓名表头单元格信息
const sourceNameHeaderCellAddress = ref<string>(''); // 源文件中的姓名表头单元格地址（要处理表的姓名表头单元格）
const sourceNameHeaderCellValue = ref<string>(''); // 源文件中的姓名表头单元格值
const targetNameHeaderCellAddress = ref<string>(''); // 目标文件中的姓名表头单元格地址（操作表的姓名表头单元格）
const targetNameHeaderCellValue = ref<string>(''); // 目标文件中的姓名表头单元格值

// 工作表数据
const sheets = ref<string[]>([]);
const sourceSheets = ref<string[]>([]);

// 加载状态
const isLoadingTargetSheets = ref<boolean>(false);
const isLoadingSourceSheets = ref<boolean>(false);

// 处理返回按钮点击
function handleBack() {
  emit('back');
}

// 通用文件选择处理函数
function handleFileSelect(event: Event, type: 'target' | 'source') {
  const input = event.target as HTMLInputElement;
  if (input.files && input.files.length > 0) {
    // 即使选择相同文件也要更新状态
    const file = input.files[0];
    const isTarget = type === 'target';
    
    // 更新对应类型的文件状态
    if (isTarget) {
      targetFile.value = file.name;
      targetFileObj.value = file; // 保存文件对象引用
      targetSheet.value = '';
    } else {
      sourceFile.value = file.name;
      sourceFileObj.value = file; // 保存文件对象引用
      sourceSheet.value = '';
    }
    
    // 重置公式生成结果
    generatedFormula.value = '';
    
    // 设置加载状态
    const isLoadingRef = isTarget ? isLoadingTargetSheets : isLoadingSourceSheets;
    isLoadingRef.value = true;
    
    // 读取Excel文件获取实际工作表名称
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        // 获取所有工作表名称并更新到对应引用
        if (isTarget) {
          sheets.value = workbook.SheetNames;
        } else {
          sourceSheets.value = workbook.SheetNames;
        }
      } catch (error) {
        console.error('读取Excel文件失败:', error);
        // 如果读取失败，使用默认工作表
        if (isTarget) {
          sheets.value = ['Sheet1', 'Sheet2', 'Sheet3'];
        } else {
          sourceSheets.value = ['Sheet1', 'Sheet2', 'Sheet3'];
        }
      } finally {
        // 无论成功失败，都结束加载状态
        isLoadingRef.value = false;
      }
    };
    
    reader.onerror = () => {
      console.error('文件读取错误');
      // 如果读取失败，使用默认工作表
      if (isTarget) {
        sheets.value = ['Sheet1', 'Sheet2', 'Sheet3'];
      } else {
        sourceSheets.value = ['Sheet1', 'Sheet2', 'Sheet3'];
      }
      isLoadingRef.value = false;
    };
    
    reader.readAsArrayBuffer(file);
    
    // 重置输入值，允许重复选择同一个文件
    setTimeout(() => {
      input.value = '';
    }, 0);
  }
}

// 通用重置文件函数
function resetFile(type: 'target' | 'source') {
  const isTarget = type === 'target';
  
  if (isTarget) {
    targetFile.value = '';
    targetFileObj.value = null; // 清除文件对象引用
    sheets.value = [];
    targetSheet.value = '';
    isLoadingTargetSheets.value = false;
  } else {
    sourceFile.value = '';
    sourceFileObj.value = null; // 清除文件对象引用
    sourceSheets.value = [];
    sourceSheet.value = '';
    isLoadingSourceSheets.value = false;
  }
  
  // 重置公式生成结果
  generatedFormula.value = '';
  
  // 重置文件输入元素
  const nthOfType = isTarget ? 1 : 2;
  const fileInput = document.querySelector(`input[type="file"]:nth-of-type(${nthOfType})`) as HTMLInputElement;
  if (fileInput) {
    fileInput.value = '';
  }
}

// 验证是否可以生成公式 - 检查四个核心依赖项
const canGenerateFormula = computed(() => {
  // 检查四个必要的依赖参数：
  // 1. 要查找的姓名 (lookupValue)
  // 2. 数据表数据查找范围 (dataSourceRange)
  // 3. MATCH函数查找值 (matchLookupValue)
  // 4. MATCH函数数据范围 (matchLookupRange)
  console.log('targetHeaderAddress.value:', targetHeaderAddress.value);
  console.log('dataSourceRange.value:', dataSourceRange.value);
  console.log('matchLookupValue.value:', matchLookupValue.value);
  console.log('matchLookupRange.value:', matchLookupRange.value);

  const b = targetHeaderAddress.value !== '' && dataSourceRange.value !== '' && matchLookupValue.value !== '' && matchLookupRange.value !== '';
  console.log('b:', b);
  return (b);
});

// 生成VLOOKUP公式
function generateFormula() {
  if (!canGenerateFormula.value) {
    return;
  }
  
const addColumnAbsoluteReference = (address: string): string => {
  return address.replace(/([A-Z]+)(\d+)/, '$$$1$2'); // A3 -> $A3
};

// 辅助函数：为单元格地址添加行绝对引用（数字加$）
const addRowAbsoluteReference = (address: string): string => {
  return address.replace(/([A-Z]+)(\d+)/, '$1$$$2'); // A3 -> A$3
};

const addCellFullAbsoluteReference = (address: string): string => {
  return address.replace(/([A-Z]+)(\d+)/, '$$$1$$$2');
};

// 辅助函数：为范围添加绝对引用（列和行都加$）
const addFullAbsoluteReference = (range: string): string => {
  if (range.includes(':')) {
    return range.replace(/([A-Z]+)(\d+):([A-Z]+)(\d+)/, '$$$1$$$2:$$$3$$$4'); // A1:B2 -> $A$1:$B$2
  } else {
    return addCellFullAbsoluteReference(range);
  }
};
  
  // 处理各种地址的绝对引用
  const lookupValueWithDollar = addColumnAbsoluteReference(targetHeaderAddress.value);
  const matchLookupWithDollar = addRowAbsoluteReference(matchLookupValue.value);
  const dataSourceWithDollar = addFullAbsoluteReference(dataSourceRange.value);
  const matchRangeWithDollar = addFullAbsoluteReference(matchLookupRange.value);
  
  // 检查数据源和目标Excel是否是同一个文件
  const isSameFile = sourceFile.value === targetFile.value;
  
  // 根据是否是同一个文件构建不同的引用格式
  const tableArrayRef = isSameFile 
    ? `'${sourceSheet.value}'!${dataSourceWithDollar}` 
    : `'[${sourceFile.value}]${sourceSheet.value}'!${dataSourceWithDollar}`;
  
  const matchRangeRef = isSameFile 
    ? `'${sourceSheet.value}'!${matchRangeWithDollar}`
    : `'[${sourceFile.value}]${sourceSheet.value}'!${matchRangeWithDollar}`;
  
  // 使用MATCH函数动态确定列索引构建VLOOKUP公式
  const vlookupFormula = `VLOOKUP(${lookupValueWithDollar}, ${tableArrayRef}, MATCH(${matchLookupWithDollar}, ${matchRangeRef}, 0), FALSE)`;
  
  // 使用IF函数包装VLOOKUP公式，处理空值情况
  // =IF(LEN(VLOOKUP(...))=0,"",VLOOKUP(...))
  generatedFormula.value = `=IF(LEN(${vlookupFormula})=0,"",${vlookupFormula})`;
}

// 显示复制成功通知
function showCopyNotification(isButtonNotification: boolean = false) {
  if (isButtonNotification) {
    // 更新按钮文本
    const copyBtn = document.querySelector('.copy-button') as HTMLButtonElement;
    if (copyBtn) {
      const originalText = copyBtn.textContent;
      copyBtn.textContent = '已复制!';
      setTimeout(() => {
        copyBtn.textContent = originalText;
      }, 2000);
    }
  } else {
    // 创建浮动通知
    const notification = document.createElement('div');
    notification.textContent = '已复制!';
    notification.style.position = 'fixed';
    notification.style.bottom = '20px';
    notification.style.right = '20px';
    notification.style.padding = '8px 16px';
    notification.style.backgroundColor = 'var(--color-primary)';
    notification.style.color = 'white';
    notification.style.borderRadius = '4px';
    notification.style.zIndex = '1000';
    document.body.appendChild(notification);
    
    setTimeout(() => {
      document.body.removeChild(notification);
    }, 2000);
  }
}

// 通用复制到剪贴板函数
function copyToClipboard(text: string, isButtonNotification: boolean = false) {
  if (text) {
    navigator.clipboard.writeText(text)
      .then(() => {
        showCopyNotification(isButtonNotification);
      })
      .catch(err => {
        console.error('复制失败:', err);
      });
  }
}

// 复制公式到剪贴板
function copyFormula() {
  copyToClipboard(generatedFormula.value, true);
}

// 复制标头地址到剪贴板
function copyHeaderAddress(address: string) {
  if (address && address !== '未找到') {
    copyToClipboard(address, false);
  }
}

// 检查标头并获取单元格地址
async function checkHeaders() {
  // 验证输入
  if (!dataToConfigure.value.trim()) {
    headerSearchError.value = '请输入要匹配的关键字';
    return;
  }
  
  if (!targetFile.value || !targetSheet.value || !sourceFile.value || !sourceSheet.value) {
    headerSearchError.value = '请先选择目标文件和源文件及其工作表';
    return;
  }
  
  // 移除对targetFileObj和sourceFileObj的验证，确保即使文件对象引用丢失，只要文件路径存在就继续执行
  // 这样可以避免点击检查后要求重新选择文件的问题

  // 重置状态
  isSearchingHeaders.value = true;
  headerSearchError.value = '';
  targetHeaderAddress.value = '';
  targetHeaderCellValue.value = '';
  sourceHeaderAddress.value = '';
  sourceNameHeaderCellAddress.value = '';
  sourceNameHeaderCellValue.value = '';
  targetNameHeaderCellAddress.value = '';
  targetNameHeaderCellValue.value = '';

  dataSourceRange.value = '';
  matchLookupValue.value = '';
  matchLookupRange.value = '';
  isHeaderSearchComplete.value = false;
  
  try {
    // 使用保存的文件对象引用
    // 并行读取两个文件，使用dataToConfigure作为查找关键字
    const [targetResult, sourceResult] = await Promise.all([
      findHeaderInFile(targetFileObj.value, targetSheet.value, dataToConfigure.value.trim(), true),
      findHeaderInFile(sourceFileObj.value, sourceSheet.value, dataToConfigure.value.trim(), false)
    ]);
    
    // 保存目标文件的单元格值和地址
    targetHeaderCellValue.value = targetResult.cellValue || '';
    targetHeaderAddress.value = targetResult.address || '未找到';
    sourceHeaderAddress.value = sourceResult.address || '未找到';
    
    // 保存姓名表头单元格信息
    sourceNameHeaderCellAddress.value = sourceResult.nameHeaderCellAddress || '';
    sourceNameHeaderCellValue.value = sourceResult.nameHeaderCellValue || '';
    targetNameHeaderCellAddress.value = targetResult.nameHeaderCellAddress || '';
    targetNameHeaderCellValue.value = targetResult.nameHeaderCellValue || '';
    

    
    // 设置数据源范围（数据表姓名表头单元格到最后一个单元格）
    if (sourceResult.nameHeaderCellAddress) {
      const range = XLSX.utils.decode_range(sourceResult.range);
      // 数据表姓名表头单元格到最后一个单元格的范围
      dataSourceRange.value = `${sourceResult.nameHeaderCellAddress}:${XLSX.utils.encode_cell({r: range.e.r, c: range.e.c})}`;
    } else {
      // 如果没有找到姓名表头单元格，回退到使用源文件的范围
      dataSourceRange.value = sourceResult.range;
    }
    
    // 设置MATCH函数相关信息
    // MATCH(查找值, 查找范围, 匹配类型)
    // 查找值: 操作表的姓名表头单元格的行 + 用户输入的列
    // 查找范围: 数据源表中的姓名表头单元格到所在行的最后一列
    if (dataToConfigure.value.trim() && targetNameHeaderCellAddress.value && sourceNameHeaderCellAddress.value) {
      // MATCH函数第一个参数：操作表的姓名表头单元格的行和用户输入的列
      // 解析操作表的姓名表头单元格地址，获取行信息
      const nameCell = XLSX.utils.decode_cell(targetNameHeaderCellAddress.value);
      
      // 从用户输入的目标单元格地址中提取列信息
      const targetCell = XLSX.utils.decode_cell(targetCellAddress.value);
      
      // 构建新的单元格地址：使用操作表姓名表头单元格的行 + 用户指定的公式单元格列
      const newCellAddress = XLSX.utils.encode_cell({r: nameCell.r, c: targetCell.c});
      matchLookupValue.value = newCellAddress;
      
      // 设置lookupValue（要查找的姓名）为目标单元格地址
      lookupValue.value = targetCellAddress.value;
      
      // MATCH函数第二个参数：数据源表中的姓名表头单元格, 直到该行的最后一列
      // 注意：这里使用sourceResult的范围来确定工作表的最大列数
      const range = XLSX.utils.decode_range(sourceResult.range);
      // 解析数据源表的姓名表头单元格地址，获取行信息
      const sourceNameCell = XLSX.utils.decode_cell(sourceNameHeaderCellAddress.value);
      // 构建从数据源表姓名表头单元格到该行最后一列的范围
      matchLookupRange.value = `${sourceNameHeaderCellAddress.value}:${XLSX.utils.encode_cell({r: sourceNameCell.r, c: range.e.c})}`;
    } else if (targetHeaderAddress.value && targetHeaderAddress.value !== '未找到') {
      // 无论是否找到姓名表头单元格，始终使用目标单元格地址作为lookupValue
      lookupValue.value = targetCellAddress.value;
    }
    
    isHeaderSearchComplete.value = true;
  } catch (error) {
    headerSearchError.value = error instanceof Error ? error.message : '检查标头时发生错误';
  } finally {
    isSearchingHeaders.value = false;
  }
}



// 在单个文件中查找标头和姓名相关信息
interface HeaderSearchResult {
  address: string;
  range: string;
  nameColumn?: string; // 添加姓名列信息
  cellValue?: string; // 添加单元格值
  nameHeaderCellAddress?: string; // 姓名表头单元格地址
  nameHeaderCellValue?: string; // 姓名表头单元格值
}

function findHeaderInFile(file: File | null, sheetName: string, header: string, isTarget: boolean = false): Promise<HeaderSearchResult> {
  return new Promise((resolve, reject) => {
    // 添加文件对象空值检查
    if (!file) {
      // 如果文件对象丢失，返回一个默认结果而不是直接失败
      // 这样可以避免用户选择的文件和sheet状态被重置
      reject(new Error('文件读取时出错，可能是文件对象引用丢失，但您的选择不会被重置，请重试'));
      return;
    }
    
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[sheetName];
        
        if (!worksheet) {
          throw new Error(`工作表 ${sheetName} 不存在`);
        }

        // 获取工作表范围
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
        const rangeStr = worksheet['!ref'] || 'A1:A1';

        console.log('range:', range);
        
        let foundAddress = '';
        const nameAddresses: string[] = [];
        let nameHeaderCellAddress = '';
        let nameHeaderCellValue = '';
        
        // 按行和列遍历所有单元格
        for (let row = range.s.r; row <= range.e.r; row++) {
          for (let col = range.s.c; col <= range.e.c; col++) {
            // 转换行列索引为单元格地址
            const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
            // 获取单元格值
            const cell = worksheet[cellAddress];
            
            // 获取单元格值并进行安全的字符串转换
            const cellValue = cell && cell.v ? String(cell.v).trim() : '';
            
            // 删除单元格值和目标标头中的所有空格，以便进行准确比较
            const cellValueNoSpaces = cellValue.replace(/\s/g, '');
            const headerNoSpaces = header.replace(/\s/g, '');
            
            // 使用用户输入的关键字查找匹配
            if (cellValueNoSpaces.includes(headerNoSpaces)) {
              // 保存所有匹配的单元格地址
              nameAddresses.push(cellAddress);
              
              // 如果是表头行（通常是第一行或前几行），设置为姓名表头单元格
              if (row <= range.s.r + 5 && !nameHeaderCellAddress) {
                nameHeaderCellAddress = cellAddress;
                nameHeaderCellValue = cellValue;
              }
              
              // 设置找到的地址（第一个匹配）
              if (!foundAddress) {
                foundAddress = cellAddress;
                console.log(`找到匹配: ${cellAddress}`);
              }
            }
          }
        }
        
        // 提取第一个匹配字段的列（如果有）
        let nameColumn = '';
        let cellValue = '';
        if (nameHeaderCellAddress) {
          // 从姓名表头单元格提取列
          const nameMatch = nameHeaderCellAddress.match(/[A-Z]+/);
          if (nameMatch) {
            nameColumn = nameMatch[0];
            
            // 对于目标文件，如果有targetCellAddress并且找到姓名表头字段，使用targetCellAddress的行和姓名表头字段的列
            if (isTarget && targetCellAddress.value) {
              // 提取targetCellAddress的行号
              const rowMatch = targetCellAddress.value.match(/\d+/);
              if (rowMatch) {
                const finalCellAddress = `${nameColumn}${rowMatch[0]}`;
                // 获取该单元格的值
                const cell = worksheet[finalCellAddress];
                if (cell && cell.v) {
                  cellValue = String(cell.v).trim();
                }
              }
            }
          }
        } else if (nameAddresses.length > 0) {
          // 回退方案：使用第一个匹配的地址
          const nameMatch = nameAddresses[0].match(/[A-Z]+/);
          if (nameMatch) {
            nameColumn = nameMatch[0];
          }
        }
        
        // 对于目标文件，如果有targetCellAddress并且找到姓名表头字段，使用targetCellAddress的行和姓名表头字段的列
        let finalAddress = foundAddress;
        if (isTarget && targetCellAddress.value && nameColumn) {
          // 提取targetCellAddress的行号
          const rowMatch = targetCellAddress.value.match(/\d+/);
          if (rowMatch) {
            finalAddress = `${nameColumn}${rowMatch[0]}`;
          }
        }
        
        // 返回结果对象
        resolve({ 
          address: finalAddress, 
          range: rangeStr,
          nameColumn: nameColumn,
          cellValue: cellValue,
          nameHeaderCellAddress: nameHeaderCellAddress,
          nameHeaderCellValue: nameHeaderCellValue
        });
      } catch (error) {
        reject(error);
      }
    };
    
    reader.onerror = () => {
      reject(new Error('文件读取错误'));
    };
    
    reader.readAsArrayBuffer(file);
  });
}
</script>

<template>
  <div class="vlookup-container">
    <!-- 返回按钮 -->
    <div class="header">
      <button class="back-button" @click="handleBack">
        ← 返回
      </button>
      <h2 class="page-title">VLOOKUP 助手</h2>
    </div>
    
    <!-- VLOOKUP工具内容 -->
    <div class="vlookup-content">
      <!-- 文件选择区域 -->
      <div class="file-selection-section">
        <div class="file-selection-wrapper flex-layout">
          <div class="file-select-group target-excel">
            <div class="form-group">
              <label>目标Excel文件（生成公式的文件）</label>
              <div class="file-input-wrapper">
                <input 
                  type="file" 
                  accept=".xlsx,.xls" 
                  @change="(e) => handleFileSelect(e, 'target')"
                  class="file-input"
                />
                <label class="custom-file-upload">
                  <span>上传目标Excel文件</span>
                  <small>支持.xlsx和.xls格式</small>
                </label>
              </div>
              <div v-if="targetFile" class="file-info">
                <span class="file-name">{{ targetFile }}</span>
                <button class="reset-button" @click="() => resetFile('target')" title="重置文件选择">
                  ×
                </button>
              </div>
            </div>
            
            <div class="form-group" v-if="sheets.length > 0 || isLoadingTargetSheets">
              <label>选择工作表</label>
              <select v-model="targetSheet" class="sheet-select" :disabled="isLoadingTargetSheets">
                <option value="">请选择工作表</option>
                <option v-for="sheet in sheets" :key="sheet" :value="sheet">{{ sheet }}</option>
              </select>
              <span v-if="isLoadingTargetSheets" class="loading-text">正在加载工作表...</span>
            </div>
            
            <!-- 公式单元格地址和匹配关键字输入 -->
            <div class="form-group">
              <label>填充公式的单元格地址</label>
              <input 
                type="text" 
                v-model="targetCellAddress" 
                placeholder="例如: A3"
                class="param-input"
              />
            </div>
            
            <div class="form-group">
              <label>匹配的关键字</label>
              <input 
                type="text" 
                v-model="dataToConfigure" 
                placeholder="例如: 姓名"
                class="param-input"
              />
            </div>
          </div>
          
          <div class="file-select-group source-excel">
            <div class="form-group">
              <label>数据源Excel文件（查找数据的文件）</label>
              <div class="file-input-wrapper">
                <input 
                  type="file" 
                  accept=".xlsx,.xls" 
                  @change="(e) => handleFileSelect(e, 'source')"
                  class="file-input"
                />
                <label class="custom-file-upload">
                  <span>上传数据源Excel文件</span>
                  <small>支持.xlsx和.xls格式</small>
                </label>
              </div>
              <div v-if="sourceFile" class="file-info">
                <span class="file-name">{{ sourceFile }}</span>
                <button class="reset-button" @click="() => resetFile('source')" title="重置文件选择">
                  ×
                </button>
              </div>
            </div>
            
            <div class="form-group" v-if="sourceSheets.length > 0 || isLoadingSourceSheets">
              <label>选择工作表</label>
              <select v-model="sourceSheet" class="sheet-select" :disabled="isLoadingSourceSheets">
                <option value="">请选择工作表</option>
                <option v-for="sheet in sourceSheets" :key="sheet" :value="sheet">{{ sheet }}</option>
              </select>
              <span v-if="isLoadingSourceSheets" class="loading-text">正在加载工作表...</span>
            </div>
              <!-- 按钮区域 - 移至数据源Excel文件右下角 -->
              <div class="bottom-right-buttons">
                <button 
                  class="generate-button check-button" 
                  @click="checkHeaders"
                  :disabled="isSearchingHeaders || !targetFile || !targetSheet || !sourceFile || !sourceSheet || !dataToConfigure.trim()"
                >
                  {{ isSearchingHeaders ? '检查中...' : '检查' }}
                </button>
                <button 
                  class="generate-button formula-button" 
                  @click="generateFormula"
                  :disabled="!canGenerateFormula"
                >
                  生成公式
                </button>
              </div>
            </div>
          </div>
        </div>
      
      <!-- 标头匹配和地址查找区域 -->
      <div class="header-matching-section">
        <h4>检查结果</h4>
        
        <!-- 错误提示 -->
        <div class="error-message" v-if="headerSearchError">
          {{ headerSearchError }}
        </div>
        
        <!-- 检查结果显示 -->
        <div class="header-results" v-if="isHeaderSearchComplete">
          <h5>检查结果</h5>
          
          <!-- 标头地址结果 -->
          <div class="header-result-item">
            <span class="result-label">要查找的姓名: </span>
            <span class="result-value">{{ targetHeaderCellValue && targetHeaderAddress !== '未找到' ? `${targetHeaderCellValue} (${targetHeaderAddress})` : '未找到' }}</span>
            <button 
              class="copy-small-button" 
              @click="copyHeaderAddress(targetHeaderCellValue)"
              :disabled="!targetHeaderCellValue || targetHeaderAddress === '未找到'"
              title="复制姓名"
            >
              {{ targetHeaderCellValue && targetHeaderAddress !== '未找到' ? '复制' : '' }}
            </button>
          </div>
          <div class="header-result-item">
            <span class="result-label">操作表的姓名表头单元格：</span>
            <span class="result-value">{{ targetNameHeaderCellValue && targetNameHeaderCellAddress ? `${targetNameHeaderCellValue} (${targetNameHeaderCellAddress})` : '未找到' }}</span>
            <button 
              class="copy-small-button" 
              @click="copyHeaderAddress(targetNameHeaderCellAddress)"
              :disabled="!targetNameHeaderCellAddress"
              title="复制地址"
            >
              {{ targetNameHeaderCellAddress ? '复制' : '' }}
            </button>
          </div>
          <div class="header-result-item">
            <span class="result-label">数据表姓名表头单元格：</span>
            <span class="result-value">{{ sourceNameHeaderCellValue && sourceNameHeaderCellAddress ? `${sourceNameHeaderCellValue} (${sourceNameHeaderCellAddress})` : '未找到' }}</span>
            <button 
              class="copy-small-button" 
              @click="copyHeaderAddress(sourceNameHeaderCellAddress)"
              :disabled="!sourceNameHeaderCellAddress"
              title="复制地址"
            >
              {{ sourceNameHeaderCellAddress ? '复制' : '' }}
            </button>
          </div>
          <!-- 数据源范围 -->
          <div class="header-result-item" v-if="dataSourceRange">
            <span class="result-label">数据表数据查找范围：</span>
            <span class="result-value">{{ dataSourceRange }}</span>
            <button 
              class="copy-small-button" 
              @click="copyHeaderAddress(dataSourceRange)"
              title="复制范围"
            >
              复制
            </button>
          </div>
          
          <!-- MATCH函数相关信息 -->
          <div class="header-result-item" v-if="matchLookupValue">
            <span class="result-label">MATCH函数查找值：</span>
            <span class="result-value">{{ matchLookupValue }}</span>
            <button 
              class="copy-small-button" 
              @click="copyHeaderAddress(matchLookupValue)"
              :disabled="!matchLookupValue"
              title="复制查找值"
            >
              {{ matchLookupValue ? '复制' : '' }}
            </button>
          </div>
          <div class="header-result-item" v-if="matchLookupRange">
            <span class="result-label">MATCH函数数据范围：</span>
            <span class="result-value">{{ matchLookupRange }}</span>
            <button 
              class="copy-small-button" 
              @click="copyHeaderAddress(matchLookupRange)"
              title="复制范围"
            >
              复制
            </button>
          </div>
        </div>
      </div>
      

      
      <!-- 生成的公式显示区域 -->
      <div class="formula-result" v-if="generatedFormula">
        <h4>生成的公式</h4>
        <div class="formula-box">
          <code>{{ generatedFormula }}</code>
          <button class="copy-button" @click="copyFormula">复制</button>
        </div>
      </div>
    </div>
  </div>
</template>