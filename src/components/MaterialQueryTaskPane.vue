<template>
  <div class="global">
    <h2 class="title">物料查询</h2>
    <hr class="divider" />
    
    <!-- 物料名称输入框 -->
    <div class="input-section">
      <div class="input-group">
        <label class="input-label">物料名称:</label>
        <input 
          type="text" 
          class="input-field" 
          v-model="materialName" 
          placeholder="请输入物料名称（支持模糊）" 
          @input="handleInput"
        />
      </div>
    </div>
    
    <!-- 自动查询提示 -->
    <div class="auto-query-hint">
      <small>提示：输入内容变化时会自动查询</small>
    </div>
    
    <!-- 结果显示区域 -->
    <div class="result-section">
      <div v-if="materialData.length > 0" class="result-container">
        <div class="result-header">查询结果 ({{ materialData.length }})</div>
        <div class="result-table-container">
          <table class="result-table">
                  <thead>
                    <tr>
                      <th>序号</th>
                      <th>物料编码</th>
                      <th>物料名称</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr v-for="(item, index) in materialData" :key="item.code" class="result-row">
                      <td>{{ index + 1 }}</td>
                      <td>{{ item.code }}</td>
                      <td v-html="highlightKeyword(item.name)"></td>
                    </tr>
                  </tbody>
                </table>
        </div>
      </div>
      
      <div v-else-if="queryError" class="error-message">
        {{ queryError }}
      </div>
      
      <div v-else-if="hasSearched" class="info-message">
        未找到符合条件的物料
      </div>

      
      <div v-else class="hint-message">
        请输入物料名称进行查询
      </div>
    </div>

  </div>
</template>

<script>
import { ref, computed } from 'vue'
import materialQuery from './js/materialQuery.js'

export default {
  name: 'MaterialQueryTaskPane',
  setup() {
    const materialName = ref('')
    const materialData = ref([])
    const queryError = ref('')
    const isLoading = ref(false)
    const hasSearched = ref(false)
    let debounceTimer = null

    // 处理输入事件，添加防抖并自动查询
    const handleInput = () => {
      clearTimeout(debounceTimer)
      debounceTimer = setTimeout(() => {
        // 自动执行查询
        queryMaterial()
      }, 500)
    }

    const queryMaterial = async () => {
      // 检查查询条件
      if (!materialName.value) {
        queryError.value = '请输入物料名称'
        materialData.value = []
        return
      }

      try {
        isLoading.value = true
        hasSearched.value = true
        queryError.value = ''
        
        console.log('开始查询物料:', { 
          name: materialName.value
        })
        
        // 调用物料查询函数
        const result = await materialQuery.queryMaterial({
          name: materialName.value
        })
        
        if (result && result.length > 0) {
          materialData.value = result
          queryError.value = ''
        } else {
          queryError.value = ''
          materialData.value = []
        }
      } catch (error) {
        console.error('查询物料失败:', error)
        queryError.value = `查询失败: ${error.message}`
        materialData.value = []
      } finally {
        isLoading.value = false
      }
    }

    // 高亮显示匹配的关键词
    const highlightKeyword = (text) => {
      if (!materialName.value) return text
      
      const keyword = materialName.value.toLowerCase().trim()
      if (!keyword) return text
      
      // 1. 将关键词分解为不同单元：连续数字、连续字母、单个汉字
      const keywordUnits = []
      let currentUnit = ''
      let currentType = ''
      
      for (let i = 0; i < keyword.length; i++) {
        const char = keyword[i]
        let type = ''
        
        // 确定当前字符的类型
        if (/\d/.test(char)) {
          type = 'number'
        } else if (/[a-zA-Z]/.test(char)) {
          type = 'letter'
        } else if (/[\u4e00-\u9fa5]/.test(char)) {
          type = 'chinese'
        } else {
          type = 'other'
        }
        
        // 如果当前字符类型与上一个不同，或者是中文/其他字符，保存当前单元
        if (type === 'chinese' || type === 'other') {
          if (currentUnit) {
            keywordUnits.push({ type: currentType, value: currentUnit })
            currentUnit = ''
            currentType = ''
          }
          if (type === 'chinese') {
            keywordUnits.push({ type: 'chinese', value: char })
          }
        } else {
          if (currentType === '' || currentType === type) {
            currentUnit += char
            currentType = type
          } else {
            keywordUnits.push({ type: currentType, value: currentUnit })
            currentUnit = char
            currentType = type
          }
        }
      }
      
      // 保存最后一个单元
      if (currentUnit) {
        keywordUnits.push({ type: currentType, value: currentUnit })
      }
      
      // 2. 创建正则表达式，匹配所有单元，不区分大小写
      let result = text
      keywordUnits.forEach(unit => {
        if (unit.value) {
          const regex = new RegExp(unit.value, 'gi')
          result = result.replace(regex, match => 
            `<span class="highlight">${match}</span>`
          )
        }
      })
      
      return result
    }

    return {
      materialName,
      materialData,
      queryError,
      isLoading,
      hasSearched,
      handleInput,
      queryMaterial,
      highlightKeyword
    }
  }
}
</script>

<style scoped>
.global {
  font-size: 14px;
  padding: 15px;
  min-height: 400px;
  background-color: #ffffff;
}

/* 标题样式 */
.title {
  color: #333333;
  margin-bottom: 10px;
  text-align: center;
  font-size: 16px;
  font-weight: bold;
}

/* 分隔线 */
.divider {
  margin: 10px 0;
  border: none;
  border-top: 1px solid #e0e0e0;
}

/* 输入区域 */
.input-section {
  margin-bottom: 12px;
  padding: 8px;
  background-color: #f9f9f9;
  border-radius: 4px;
  border: 1px solid #e0e0e0;
}

.input-group {
  display: flex;
  align-items: center;
  width: 100%;
}

.input-label {
  margin-right: 10px;
  display: inline-block;
  width: 80px;
  font-weight: 500;
  color: #333333;
  font-size: 13px;
}

.input-field {
  flex: 1;
  padding: 6px 10px;
  border: 1px solid #d0d0d0;
  border-radius: 4px;
  font-size: 13px;
  transition: border-color 0.2s ease;
}

.input-field:focus {
  outline: none;
  border-color: #1890ff;
  box-shadow: 0 0 0 2px rgba(24, 144, 255, 0.1);
}

/* 自动查询提示 */
.auto-query-hint {
  margin: 5px 0 15px 0;
  color: #666666;
  font-size: 11px;
  font-style: italic;
}

/* 结果区域 */
.result-section {
  margin: 15px -15px -15px -15px;
  padding: 15px;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  background-color: #f9f9f9;
}

.result-container {
  width: 100%;
}

.result-header {
  font-weight: 500;
  color: #333333;
  margin-bottom: 10px;
  font-size: 13px;
}

.result-table-container {
  margin: 8px -15px -15px -15px;
  padding: 0 15px;
  overflow-x: auto;
  max-height: 400px;
  overflow-y: auto;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  background-color: #ffffff;
}

.result-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
}

.result-table th, .result-table td {
  border: 1px solid #e0e0e0;
  padding: 6px 8px;
  text-align: left;
}

.result-table th {
  background-color: #f5f5f5;
  font-weight: 500;
  color: #333333;
  position: sticky;
  top: 0;
  z-index: 1;
}

.result-table th:first-child {
  width: 40px;
  text-align: center;
}

.result-table th:nth-child(2) {
  width: 120px;
}

.result-table th:nth-child(3) {
  flex: 1;
}

.result-table td:first-child {
  text-align: center;
  font-weight: 500;
}

.result-row:hover {
  background-color: #f0f8ff;
}

/* 高亮样式 */
:deep(.highlight) {
  background-color: #fff3cd;
  color: #856404;
  padding: 0 2px;
  border-radius: 2px;
  font-weight: 500;
}

/* 消息样式 */
.error-message {
  color: #ff4d4f;
  padding: 10px;
  background-color: #f8d7da;
  border: 1px solid #f5c6cb;
  border-radius: 4px;
  font-size: 13px;
}

.info-message {
  color: #17a2b8;
  padding: 10px;
  background-color: #d1ecf1;
  border: 1px solid #bee5eb;
  border-radius: 4px;
  font-size: 13px;
}

.hint-message {
  color: #6c757d;
  padding: 10px;
  background-color: #f8f9fa;
  border: 1px solid #e9ecef;
  border-radius: 4px;
  font-size: 13px;
  text-align: center;
}


</style>