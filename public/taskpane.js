// 物料模糊查询核心逻辑

// 初始化函数
function init() {
    console.log("初始化物料模糊查询功能");
    
    // 绑定搜索按钮点击事件
    document.getElementById("searchBtn").onclick = searchMaterials;
    
    // 可选：实时搜索
    document.getElementById("keyword").oninput = debounce(searchMaterials, 300);
    
    console.log("物料模糊查询功能初始化完成");
}

// 防抖函数
function debounce(func, wait) {
    let timeout;
    return function() {
        const context = this;
        const args = arguments;
        clearTimeout(timeout);
        timeout = setTimeout(() => {
            func.apply(context, args);
        }, wait);
    };
}

// 搜索物料
async function searchMaterials() {
    const keyword = document.getElementById("keyword").value.trim().toLowerCase();
    if (!keyword) {
        document.getElementById("resultlist").innerHTML = "<div>请输入物料名称关键词</div>";
        return;
    }
    
    try {
        console.log("开始搜索物料:", keyword);
        
        // 读取表格数据
        let data = [];
        
        // 尝试使用WPS/Office JS API读取表格数据
        if (window.wps && window.wps.EtApplication) {
            console.log("使用WPS JS API读取表格数据");
            data = await readDataWithWPSAPI();
        } else if (window.Application) {
            console.log("使用Application对象读取表格数据");
            data = await readDataWithApplication();
        } else {
            console.log("无法访问WPS应用对象，使用模拟数据");
            data = getMockData();
        }
        
        // 模糊匹配
        const matches = fuzzyMatch(data, keyword);
        
        // 渲染结果
        renderResults(matches);
        
        console.log("物料搜索完成，找到", matches.length, "个匹配项");
        
    } catch (error) {
        console.error("搜索物料失败:", error);
        console.error("错误堆栈:", error.stack);
        document.getElementById("resultlist").innerHTML = `<div style="color: red;">搜索失败: ${error.message}</div>`;
    }
}

// 使用WPS JS API读取表格数据
async function readDataWithWPSAPI() {
    return new Promise((resolve, reject) => {
        try {
            const app = window.wps.EtApplication();
            const workbook = app.ActiveWorkbook;
            const worksheet = workbook.ActiveSheet;
            
            // 获取使用范围
            const usedRange = worksheet.UsedRange;
            const rowCount = usedRange.Rows.Count;
            const columnCount = usedRange.Columns.Count;
            
            console.log(`表格数据范围: ${rowCount}行, ${columnCount}列`);
            
            // 读取数据
            const data = [];
            for (let i = 1; i <= rowCount; i++) {
                const row = [];
                for (let j = 1; j <= columnCount; j++) {
                    const cell = usedRange.Cells.Item(i, j);
                    row.push(cell.Value2 || "");
                }
                data.push(row);
            }
            
            resolve(data);
            
        } catch (error) {
            reject(error);
        }
    });
}

// 使用Application对象读取表格数据
async function readDataWithApplication() {
    return new Promise((resolve, reject) => {
        try {
            const app = window.Application;
            const workbook = app.ActiveWorkbook;
            const worksheet = workbook.ActiveSheet;
            
            // 获取使用范围
            const usedRange = worksheet.UsedRange;
            const rowCount = usedRange.Rows.Count;
            const columnCount = usedRange.Columns.Count;
            
            console.log(`表格数据范围: ${rowCount}行, ${columnCount}列`);
            
            // 读取数据
            const data = [];
            for (let i = 1; i <= rowCount; i++) {
                const row = [];
                for (let j = 1; j <= columnCount; j++) {
                    const cell = usedRange.Cells(i, j);
                    row.push(cell.Value2 || "");
                }
                data.push(row);
            }
            
            resolve(data);
            
        } catch (error) {
            reject(error);
        }
    });
}

// 获取模拟数据
function getMockData() {
    // 模拟物料数据
    return [
        ["物料编码", "物料名称", "规格型号", "单位", "库存数量"],
        ["115055982", "HND-D171(1.2-双三甲氧基硅基乙烷)", "200/塑桶", "个", 1000],
        ["115056207", "HND-V171(乙烯基三甲氧基硅烷)", "950/吨桶", "个", 800],
        ["115055989", "HND-V171(乙烯基三甲氧基硅烷)", "190/铁桶", "个", 1200],
        ["115058270", "HND-V171(乙烯基三甲氧基硅烷)", "28/塑桶", "个", 500],
        ["115056204", "HND-V171(乙烯基三甲氧基硅烷)", "5/塑桶", "个", 1500],
        ["115056207", "HND-V171(乙烯基三甲氧基硅烷)", "吨桶", "个", 1000],
        ["115056205", "HND-V171(乙烯基三甲氧基硅烷)", "25/塑桶", "个", 800],
        ["115055980", "HND-V150(乙烯基三氯硅烷)", "220/塑桶", "个", 1200],
        ["115057979", "HND-V150(乙烯基三氯硅烷)", "槽罐", "个", 500],
        ["115056001", "HND-TMOS(L)(正硅酸甲酯)", "吨桶", "个", 1500],
        ["115056301", "HND-TMOS(L)(正硅酸甲酯)", "190/塑桶", "个", 1000],
        ["115056302", "HND-TMOS(L)(正硅酸甲酯)", "200/铁桶", "个", 800],
        ["115056301", "HND-TMOS(L)(正硅酸甲酯)", "200/塑桶", "个", 1200],
        ["114001903", "电石渣", "散装", "个", 500],
        ["115055993", "HND-D150(1.2-双三氯硅基乙烷)", "250/铁桶", "个", 1500],
        ["115055983", "HND-D150(1.2-双三氯硅基乙烷)", "200L塑桶", "个", 1000],
        ["115055993", "HND-D150(1.2-双三氯硅基乙烷)", "200L铁桶", "个", 800]
    ];
}

// 模糊匹配算法
function fuzzyMatch(data, keyword) {
    console.log("开始模糊匹配，数据行数:", data.length);
    
    let matches = [];
    
    // 跳过表头，从第二行开始匹配
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const code = row[0] || "";
        const name = (row[1] || "").toLowerCase();
        
        // 简单包含匹配
        if (name.includes(keyword)) {
            matches.push({
                code: code.toString(),
                name: row[1],
                spec: row[2] || "",
                unit: row[3] || "",
                stock: row[4] || 0
            });
        }
    }
    
    console.log("模糊匹配完成，找到", matches.length, "个匹配项");
    return matches;
}

// 渲染结果
function renderResults(matches) {
    const resultList = document.getElementById("resultlist");
    
    if (matches.length === 0) {
        resultList.innerHTML = "<div>未找到匹配的物料</div>";
        return;
    }
    
    let html = "";
    for (const match of matches) {
        html += `
            <div class="result-item" onclick="selectMaterial('${match.code}', '${match.name}', '${match.spec}', '${match.unit}', ${match.stock})">
                <span class="result-code">${match.code}</span>
                <span class="result-name">${match.name}</span>
                <br>
                <small>规格: ${match.spec}, 单位: ${match.unit}, 库存: ${match.stock}</small>
            </div>
        `;
    }
    
    resultList.innerHTML = html;
}

// 选择物料（可选：点击结果 → 写入单元格）
function selectMaterial(code, name, spec, unit, stock) {
    console.log("选择物料:", code, name);
    
    // 尝试将物料信息写入当前选中的单元格
    try {
        if (window.wps && window.wps.EtApplication) {
            const app = window.wps.EtApplication();
            const range = app.ActiveCell;
            range.Value2 = name;
            range.Offset(0, 1).Value2 = code;
            range.Offset(0, 2).Value2 = spec;
            range.Offset(0, 3).Value2 = unit;
            range.Offset(0, 4).Value2 = stock;
            alert("物料信息已写入单元格");
        } else if (window.Application) {
            const app = window.Application;
            const range = app.ActiveCell;
            range.Value2 = name;
            range.Offset(0, 1).Value2 = code;
            range.Offset(0, 2).Value2 = spec;
            range.Offset(0, 3).Value2 = unit;
            range.Offset(0, 4).Value2 = stock;
            alert("物料信息已写入单元格");
        } else {
            alert(`选择物料: ${name} (${code})`);
        }
    } catch (error) {
        console.error("写入单元格失败:", error);
        alert(`选择物料: ${name} (${code})`);
    }
}

// 使用wps.ready或Office.onReady进行初始化
if (window.wps && window.wps.ready) {
    window.wps.ready(init);
} else if (window.Office && window.Office.onReady) {
    window.Office.onReady(init);
} else {
    // 浏览器环境下直接初始化
    window.onload = init;
}