// 物料数据管理系统
// 使用内存缓存 + IndexedDB备份的方案

// 物料数据（直接嵌入，避免JSON导入问题）
const materialsData = [
  {"code": "114001895", "name": "HND-TMOS(H)"},
  {"code": "114001896", "name": "HND-V171"},
  {"code": "114001897", "name": "HND-V150"},
  {"code": "114001898", "name": "HND-TMOS(L)"},
  {"code": "114001899", "name": "HND-D171"},
  {"code": "114001900", "name": "HND-D150"},
  {"code": "114001901", "name": "HND-D2150"},
  {"code": "114001902", "name": "HND-N113"},
  {"code": "114001903", "name": "电石渣"},
  {"code": "114001904", "name": "硅渣"},
  {"code": "114001905", "name": "盐酸"},
  {"code": "113000129", "name": "三甲氧基氢硅烷"},
  {"code": "113000130", "name": "乙炔"},
  {"code": "115055979", "name": "HND-V171_200kg_塑料桶_蓝色_华耐德"},
  {"code": "115055980", "name": "HND-V150_200kg_塑料桶_蓝色_华耐德"},
  {"code": "115055981", "name": "HND-TMOS(H)_200kg_塑料桶_蓝色_华耐德"},
  {"code": "115055982", "name": "HND-D171_200kg_塑料桶_蓝色_华耐德"},
  {"code": "115055983", "name": "HND-D150_200KG_塑料桶_蓝色_华耐德"},
  {"code": "115055984", "name": "HND-D2150_200KG_塑料桶_蓝色_华耐德"},
  {"code": "115055985", "name": "HND-N113_200kg_塑料桶_蓝色_华耐德"},
  {"code": "115055989", "name": "HND-V171_200kg_内衬PVF铁桶_蓝色_华耐德"},
  {"code": "115055990", "name": "HND-V150_200kg_内衬PVF铁桶_蓝色_华耐德"},
  {"code": "115055991", "name": "HND-TMOS(H)_200kg_内衬PVF铁桶_蓝色_华耐德"},
  {"code": "115055992", "name": "HND-D171_200kg_内衬PVF铁桶_蓝色_华耐德"},
  {"code": "115055993", "name": "HND-D150_200KG_内衬PVF铁桶_蓝色_华耐德"},
  {"code": "115055994", "name": "HND-D2150_200KG_内衬PVF铁桶_蓝色_华耐德"},
  {"code": "115055995", "name": "HND-N113_200kg_内衬PVF铁桶_蓝色_华耐德"},
  {"code": "115056001", "name": "HND-TMOS(L)_1000kg_吨桶PE_白色_华耐德"},
  {"code": "115056003", "name": "HND-D150_1000kg_吨桶PE_白色_华耐德"},
  {"code": "115056004", "name": "HND-D2150_1000kg_吨桶PE_白色_华耐德"},
  {"code": "115056008", "name": "盐酸_1000kg_吨桶PE_白色_华耐德"},
  {"code": "115056009", "name": "盐酸_槽车"},
  {"code": "111001785", "name": "硅粉"},
  {"code": "111001786", "name": "电石"},
  {"code": "111001787", "name": "三氯氢硅"},
  {"code": "111001788", "name": "甲醇"},
  {"code": "111001789", "name": "次氯酸钠"},
  {"code": "111001790", "name": "液碱"},
  {"code": "111001791", "name": "无水氯化钙"},
  {"code": "111001792", "name": "氯铂酸"},
  {"code": "111001793", "name": "氯苯"},
  {"code": "112004289", "name": "三甲氧基氢硅烷辅料包"},
  {"code": "112004290", "name": "HND-TMOS(H)_辅料包"},
  {"code": "112004291", "name": "HND-V171_辅料包"},
  {"code": "112004292", "name": "HND-V150_辅料包"},
  {"code": "111001808", "name": "1.2-双三甲氧基硅基乙烷（溶剂）"},
  {"code": "111001831", "name": "四甲基四乙烯基环四硅氧烷"},
  {"code": "111001832", "name": "咔唑"},
  {"code": "111001833", "name": "150抗氧化剂"},
  {"code": "111001834", "name": "吩噻嗪"},
  {"code": "111001835", "name": "乙烯基双封头"},
  {"code": "111001836", "name": "碳酸氢钠"},
  {"code": "111001838", "name": "甲醇钠"},
  {"code": "111001839", "name": "硫酸"},
  {"code": "111001840", "name": "双氧水"},
  {"code": "111001841", "name": "聚合氯化铝"},
  {"code": "111001842", "name": "聚丙烯酰胺"},
  {"code": "111001843", "name": "硫酸亚铁"},
  {"code": "220000440", "name": "5kg 塑料桶 白色_华耐德"},
  {"code": "220000441", "name": "25kg 塑料桶 蓝色_华耐德"},
  {"code": "220000442", "name": "25kg 内衬铁桶 蓝色_华耐德"},
  {"code": "220000443", "name": "200kg 塑料桶 蓝色_华耐德"},
  {"code": "220000444", "name": "200kg 内衬PVF铁桶 蓝色_华耐德"},
  {"code": "220000445", "name": "吨桶PE 白色_华耐德"},
  {"code": "220000463", "name": "25kg_塑料外贸桶_蓝色_华耐德"},
  {"code": "220000464", "name": "20kg_塑料内桶_白色_华耐德"},
  {"code": "220000465", "name": "190kg_铁桶_蓝色_华耐德"},
  {"code": "220000466", "name": "950kg_IBC桶_白色_华耐德"},
  {"code": "220000467", "name": "5kg_内涂桶_蓝色_华耐德"},
  {"code": "220000468", "name": "28kg_塑料桶_蓝色_华耐德"},
  {"code": "220000486", "name": "190kg_塑料桶_蓝色_华耐德"},
  {"code": "115056200", "name": "HND-TMOS(H)_5kg_塑料桶_白色_华耐德"},
  {"code": "115056201", "name": "HND-TMOS(H)_25kg_塑料桶_蓝色_华耐德"},
  {"code": "115056202", "name": "HND-TMOS(H)_25kg_内衬铁桶_蓝色_华耐德"},
  {"code": "115056203", "name": "HND-TMOS(H)_1000kg_吨桶PE_白色_华耐德"},
  {"code": "115056204", "name": "HND-V171_5kg_塑料桶_白色_华耐德"},
  {"code": "115056205", "name": "HND-V171_25kg_塑料桶_蓝色_华耐德"},
  {"code": "115056206", "name": "HND-V171_25kg_内衬铁桶_蓝色_华耐德"},
  {"code": "115056207", "name": "HND-V171_1000kg_吨桶PE_白色_华耐德"},
  {"code": "115056208", "name": "HND-V150_5kg_塑料桶_白色_华耐德"},
  {"code": "115056209", "name": "HND-V150_25kg_塑料桶_蓝色_华耐德"},
  {"code": "115056210", "name": "HND-V150_25kg_内衬铁桶_蓝色_华耐德"},
  {"code": "115056211", "name": "HND-V150_1000kg_吨桶PE_白色_华耐德"},
  {"code": "115056212", "name": "HND-D171_5kg_塑料桶_白色_华耐德"},
  {"code": "115056213", "name": "HND-D171_25kg_塑料桶_蓝色_华耐德"},
  {"code": "115056214", "name": "HND-D171_25kg_内衬铁桶_蓝色_华耐德"},
  {"code": "115056215", "name": "HND-D171_1000kg_吨桶PE_白色_华耐德"},
  {"code": "115056216", "name": "HND-N113_5kg_塑料桶_白色_华耐德"},
  {"code": "115056217", "name": "HND-N113_25kg_塑料桶_蓝色_华耐德"},
  {"code": "115056218", "name": "HND-N113_25kg_内衬铁桶_蓝色_华耐德"},
  {"code": "115056219", "name": "HND-N113_1000kg_吨桶PE_白色_华耐德"},
  {"code": "111001856", "name": "二甲基二苯基醚"},
  {"code": "114001928", "name": "四氯化硅"},
  {"code": "112004336", "name": "HND-TMOS(L)_辅料包"},
  {"code": "112004337", "name": "HND-N113_辅料包"},
  {"code": "114001938", "name": "HND-171合成粗品"},
  {"code": "114001940", "name": "HND-150合成粗品"},
  {"code": "115056298", "name": "HND-TMOS(L)_5kg_塑料桶_白色_华耐德"},
  {"code": "115056299", "name": "HND-TMOS(L)_25kg_塑料桶_蓝色_华耐德"},
  {"code": "115056300", "name": "HND-TMOS(L)_25kg_内衬铁桶_蓝色_华耐德"},
  {"code": "115056301", "name": "HND-TMOS(L)_200kg_塑料桶_蓝色_华耐德"},
  {"code": "115056302", "name": "HND-TMOS(L)_200kg_内衬PVF铁桶_蓝色_华耐德"},
  {"code": "114001944", "name": "HND-TMOS粗品"},
  {"code": "220000469", "name": "1000ml氟化塑料瓶"},
  {"code": "220000470", "name": "500ml氟化塑料瓶"},
  {"code": "220000471", "name": "250ml氟化塑料瓶"},
  {"code": "220000472", "name": "100ml氟化塑料瓶"},
  {"code": "115056406", "name": "HND-V171_20kg_塑料内桶_白色_华耐德"},
  {"code": "115056407", "name": "HND-V171_25kg_塑料桶_蓝色_华耐德(外贸)"},
  {"code": "115056408", "name": "HND-V171_190kg_铁桶_蓝色_华耐德"},
  {"code": "115056409", "name": "HND-V171_950kg_IBC桶_白色_华耐德"},
  {"code": "115056410", "name": "HND-TMOS(H)_5kg_内涂桶_蓝色_华耐德"},
  {"code": "115056411", "name": "HND-TMOS(L)_5kg_内涂桶_蓝色_华耐德"},
  {"code": "115056412", "name": "HND-N113_28kg_塑料桶_蓝色_华耐德"},
  {"code": "115056413", "name": "HND-N113_190kg_塑料桶_蓝色_华耐德"},
  {"code": "115056414", "name": "HND-N113_190kg_铁桶_蓝色_华耐德"},
  {"code": "115056415", "name": "HND-N113_900kg_IBC桶_白色_华耐德"},
  {"code": "111001411", "name": "正丁醇"},
  {"code": "111000363", "name": "异丙醇"}
];

class MaterialManager {
  constructor() {
    this.materials = [];
    this.isLoaded = false;
    this.db = null;
    this.init();
  }

  // 初始化
  async init() {
    try {
      // 1. 加载内存缓存
      await this.loadMaterials();
      
      // 2. 初始化IndexedDB
      await this.initIndexedDB();
      
      // 3. 同步数据到IndexedDB
      await this.syncToIndexedDB();
      
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

  // 初始化IndexedDB
  initIndexedDB() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open('MaterialDB', 1);
      
      request.onupgradeneeded = (event) => {
        const db = event.target.result;
        if (!db.objectStoreNames.contains('materials')) {
          const store = db.createObjectStore('materials', { keyPath: 'id', autoIncrement: true });
          store.createIndex('name', 'name', { unique: false });
          store.createIndex('code', 'code', { unique: false });
        }
      };
      
      request.onsuccess = (event) => {
        this.db = event.target.result;
        resolve();
      };
      
      request.onerror = (event) => {
        reject(event.target.error);
      };
    });
  }

  // 同步数据到IndexedDB
  syncToIndexedDB() {
    return new Promise((resolve, reject) => {
      if (!this.db || !this.materials.length) {
        resolve();
        return;
      }

      const transaction = this.db.transaction(['materials'], 'readwrite');
      const store = transaction.objectStore('materials');
      
      // 清空现有数据
      store.clear();
      
      // 插入新数据
      this.materials.forEach((material, index) => {
        const request = store.add({
          id: index + 1,
          code: material.code,
          name: material.name
        });
      });
      
      transaction.oncomplete = () => {
        console.log('物料数据同步到IndexedDB完成');
        resolve();
      };
      
      transaction.onerror = (event) => {
        reject(event.target.error);
      };
    });
  }

  // 模糊查询物料
  fuzzySearch(keyword) {
    if (!keyword || !this.isLoaded) {
      return [];
    }

    const lowerKeyword = keyword.toLowerCase().trim();
    
    // 内存缓存查询
    return this.materials.filter(material => {
      const materialName = material.name.toLowerCase();
      
      // 严格检查：物料名称必须包含所有输入的关键词部分
      // 对于"200塑桶"，必须包含"200"、"塑"、"桶"全部三个部分
      let containsAllKeywords = true;
      
      // 1. 检查数字部分（作为整体）
      const numberParts = lowerKeyword.match(/\d+/g) || [];
      for (const numberPart of numberParts) {
        if (!materialName.includes(numberPart)) {
          containsAllKeywords = false;
          break;
        }
      }
      
      // 2. 检查汉字部分（逐个）
      const chineseParts = lowerKeyword.match(/[\u4e00-\u9fa5]/g) || [];
      for (const chinesePart of chineseParts) {
        if (!materialName.includes(chinesePart)) {
          containsAllKeywords = false;
          break;
        }
      }
      
      // 3. 检查其他部分（如字母）
      const otherParts = lowerKeyword.replace(/\d+/g, '').replace(/[\u4e00-\u9fa5]/g, '').trim();
      if (otherParts) {
        if (!materialName.includes(otherParts)) {
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
    return this.materials.find(material => material.code === code);
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
    this.syncToIndexedDB();
  }

  // 批量添加物料
  addMaterials(materials) {
    this.materials = [...this.materials, ...materials];
    this.syncToIndexedDB();
  }
}

// 创建单例实例
const materialManager = new MaterialManager();

// 导出查询函数
async function queryMaterial(params) {
  // 确保数据已加载
  if (!materialManager.isLoaded) {
    await materialManager.init();
  }

  let result = [];

  if (params.code) {
    // 按编码查询
    const material = materialManager.getByCode(params.code);
    if (material) {
      result = [{
        code: material.code,
        name: material.name,
        spec: "",
        unit: "个",
        stock: 1000
      }];
    }
  } else if (params.name) {
    // 按名称模糊查询
    const materials = materialManager.fuzzySearch(params.name);
    result = materials.map(material => ({
      code: material.code,
      name: material.name,
      spec: "",
      unit: "个",
      stock: 1000
    }));
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