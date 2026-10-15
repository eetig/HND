import masterMaterialsData from './materials.json'

// 物料数据管理系统
// 使用内存缓存 + IndexedDB备份的方案

// 物料数据（直接嵌入，避免JSON导入问题）
const materialsData = masterMaterialsData;

function normalizeSearchText(value) {
  return String(value ?? '')
    .toLowerCase()
    .replace(/[\s_\-()/、，,.:：]+/g, '');
}

class MaterialManager {
  constructor() {
    this.materials = [];
    this.isLoaded = false;
    this.ready = this.init();
  }

  // 初始化
  async init() {
    try {
      // 1. 加载内存缓存
      await this.loadMaterials();
      
      console.log('物料管理系统初始化完成');
    } catch (error) {
      console.error('初始化物料管理系统失败:', error);
    }
  }

  // 加载物料数据
  async loadMaterials() {
    try {
      // 直接使用嵌入的物料数据
      this.materials = materialsData;
      this.isLoaded = true;
      console.log('物料数据加载完成，共', this.materials.length, '条记录');
    } catch (error) {
      console.error('加载物料数据失败:', error);
      // 使用备用数据
      this.materials = [];
    }
  }
  // 模糊查询物料
  fuzzySearch(keyword) {
    if (!keyword || !this.isLoaded) {
      return [];
    }

    const normalizedKeyword = normalizeSearchText(keyword);
    if (!normalizedKeyword) {
      return [];
    }
    
    // 内存缓存查询
    return this.materials.filter(material => {
      const searchableText = normalizeSearchText(
        [material.name, material.spec].filter(Boolean).join(' ')
      );
      const normalizedCode = normalizeSearchText(material.code);

      if (/^[a-z]+\d+[a-z\d]*$/i.test(normalizedKeyword)) {
        return searchableText.includes(normalizedKeyword);
      }

      if (/^\d+$/.test(normalizedKeyword)) {
        return normalizedCode.includes(normalizedKeyword) ||
          searchableText.includes(normalizedKeyword);
      }

      if (searchableText.includes(normalizedKeyword)) {
        return true;
      }
      
      // 严格检查：物料名称必须包含所有输入的关键词部分
      // 对于"200塑桶"，必须包含"200"、"塑"、"桶"全部三个部分
      let containsAllKeywords = true;
      
      // 1. 检查数字部分（作为整体）
      const numberParts = normalizedKeyword.match(/\d+/g) || [];
      for (const numberPart of numberParts) {
        if (!searchableText.includes(numberPart)) {
          containsAllKeywords = false;
          break;
        }
      }
      
      // 2. 检查汉字部分（逐个）
      const chineseParts = normalizedKeyword.match(/[\u4e00-\u9fa5]/g) || [];
      for (const chinesePart of chineseParts) {
        if (!searchableText.includes(chinesePart)) {
          containsAllKeywords = false;
          break;
        }
      }
      
      // 3. 检查其他部分（如字母）
      const otherParts = normalizedKeyword
        .replace(/\d+/g, '')
        .replace(/[\u4e00-\u9fa5]/g, '')
        .trim();
      if (otherParts) {
        if (!searchableText.includes(otherParts)) {
          containsAllKeywords = false;
        }
      }
      
      return containsAllKeywords;
    });
  }

  // 根据编码查询物料
  getByCode(code) {
    if (!code) {
      return null;
    }
    const normalizedCode = String(code).trim();
    return this.materials.find(material => material.code === normalizedCode);
  }

  // 根据名称查询物料
  getByName(name) {
    if (!name) {
      return [];
    }
    const lowerName = name.toLowerCase();
    return this.materials.filter(material => 
      material.name.toLowerCase().includes(lowerName)
    );
  }

  // 获取所有物料
  getAllMaterials() {
    return this.materials;
  }

  // 添加物料
  addMaterial(material) {
    this.materials.push(material);
  }

  // 批量添加物料
  addMaterials(materials) {
    this.materials = [...this.materials, ...materials];
      
  }
}

// 创建单例实例
const materialManager = new MaterialManager();

// 导出查询函数
async function queryMaterial(params) {
  await materialManager.ready;

  let result = [];

  if (params.code) {
    // 按编码查询
    const material = materialManager.getByCode(params.code);
    if (material) {
      result = [toQueryResult(material)];
    }
  } else if (params.name) {
    // 按名称模糊查询
    const materials = materialManager.fuzzySearch(params.name);
    result = materials.map(toQueryResult);
  }

  // 按物料编码排序
  result.sort((a, b) => {
    // 提取编码中的数字部分进行排序
    const getCodeNumber = (code) => {
      // 提取编码中的所有数字
      const numbers = code.match(/\d+/g);
      if (numbers) {
        return parseInt(numbers.join(''));
      }
      return 0;
    };
    
    return getCodeNumber(a.code) - getCodeNumber(b.code);
  });

  // 模拟异步查询
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve(result);
    }, 100);
  });
}

function toQueryResult(material) {
  return {
    code: material.code,
    name: material.name,
    spec: material.spec || "",
    unit: material.unit || "",
    materialGroup: material.materialGroup || "",
    productGroup: material.productGroup || "",
    sourceSheet: material.sourceSheet || "",
    sourceRow: material.sourceRow || 0
  };
}

// 隐藏TaskPane
function hideTaskPane() {
  try {
    // 尝试隐藏物料查询TaskPane
    if (window.Application && window.Application.PluginStorage) {
      let tsId = window.Application.PluginStorage.getItem("material_query_taskpane_id");
      if (tsId && window.Application.GetTaskPane) {
        let tskpane = window.Application.GetTaskPane(tsId);
        if (tskpane) {
          tskpane.Visible = false;
        }
      }
    }
  } catch (error) {
    console.error("隐藏TaskPane失败:", error);
  }
}

export default {
  queryMaterial,
  hideTaskPane,
  materialManager
};