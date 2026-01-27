import Util from './js/util.js'
import SystemDemo from './js/systemdemo.js'

// 生成动态时钟唯一标识的hash函数
function generateClockHash(workbookFullName, sheetName, address) {
    const data = `${workbookFullName}:${sheetName}:${address}`;
    // 使用简单高效的hash算法
    let hash = 0;
    for (let i = 0; i < data.length; i++) {
        const char = data.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash; // 转换为32位整数
    }
    return Math.abs(hash).toString(16);
}

// 文件操作函数
function saveDynamicTimeConfig(configs) {
    try {
        // 去重处理，确保每个单元格只保存一次配置
        const uniqueConfigs = [];
        const seen = new Set();
        
        for (const config of configs) {
            const hash = generateClockHash(config.workbookFullName, config.sheetName, config.address);
            if (!seen.has(hash)) {
                seen.add(hash);
                uniqueConfigs.push({
                    ...config,
                    hash: hash // 存储hash值，方便后续查找
                });
            }
        }
        
        console.log('去重后的配置:', uniqueConfigs);
        console.log('保存动态时间配置:', uniqueConfigs);
        
        // 使用localStorage存储配置
        localStorage.setItem('hnd_dynamic_time_configs', JSON.stringify(uniqueConfigs));
        console.log('配置保存成功');
    } catch (e) {
        console.error('保存配置失败:', e);
    }
}

function loadDynamicTimeConfig() {
    try {
        const configs = localStorage.getItem('hnd_dynamic_time_configs');
        const parsedConfigs = configs ? JSON.parse(configs) : [];
        console.log('加载动态时间配置:', parsedConfigs);
        return parsedConfigs;
    } catch (e) {
        console.error('加载配置失败:', e);
        return [];
    }
}

//这个函数在整个HND插件中是第一个执行的
function OnAddinLoad(ribbonUI){
    console.log('OnAddinLoad函数开始执行');
    if (typeof (window.Application.ribbonUI) != "object"){
        window.Application.ribbonUI = ribbonUI
    }
    
    if (typeof (window.Application.Enum) != "object") { // 如果没有内置枚举值
        window.Application.Enum = Util.WPS_Enum
    }

    //这几个导出函数是给外部业务系统调用的
    window.openOfficeFileFromSystemDemo = SystemDemo.openOfficeFileFromSystemDemo
    window.InvokeFromSystemDemo = SystemDemo.InvokeFromSystemDemo

    // 初始化动态时钟数组
    window.dynamicTimeClocks = [];
    console.log('初始化动态时钟数组');

    // 加载并恢复之前的动态时钟配置
    function restoreDynamicTimeClocks() {
        try {
            const savedConfigs = loadDynamicTimeConfig();
            if (savedConfigs.length > 0) {
                // 获取当前工作簿
                const currentWorkbook = window.Application.ActiveWorkbook;
                if (currentWorkbook) {
                    const workbookName = currentWorkbook.Name;
                    const currentWorkbookFullName = currentWorkbook.FullName;
                    
                    // 检查是否有匹配当前工作簿的配置
                    for (const config of savedConfigs) {
                        if (config.workbookFullName === currentWorkbookFullName) {
                            try {
                                // 检查工作表是否存在
                                const sheet = currentWorkbook.Sheets.Item(config.sheetName);
                                if (sheet) {
                                    // 创建新的动态时钟配置
                                    const clockConfig = {
                                        address: config.address,
                                        sheetName: config.sheetName,
                                        workbook: currentWorkbook,
                                        intervalId: null
                                    };
                                    
                                    // 为当前单元格创建定时器
                                    clockConfig.intervalId = setInterval(function() {
                                        // 获取当前时间字符串
                                        const timeString = getNowTimeString();
                                        
                                        // 更新单元格
                                        try {
                                            // 检查Workbook对象是否有效
                                            if (clockConfig.workbook) {
                                                try {
                                                    // 尝试访问Workbook的属性，检查是否有效
                                                    const workbookName = clockConfig.workbook.Name;
                                                    
                                                    // Workbook有效，执行更新
                                                    const sheet = clockConfig.workbook.Sheets.Item(clockConfig.sheetName);
                                                    const range = sheet.Range(clockConfig.address);
                                                    // 尝试使用Value2设置值
                                                    range.Value2 = timeString;
                                                    
                                                    // 设置单元格格式
                                                    range.NumberFormatLocal = "yyyy/m/d hh:mm:ss";
                                                } catch (workbookError) {
                                                    // Workbook对象失效，清除定时器
                                                    clearInterval(clockConfig.intervalId);
                                                    // 从数组中删除
                                                    if (window.dynamicTimeClocks) {
                                                        const index = window.dynamicTimeClocks.indexOf(clockConfig);
                                                        if (index > -1) {
                                                            window.dynamicTimeClocks.splice(index, 1);
                                                        }
                                                    }
                                                }
                                            }
                                        } catch (e) {
                                            // 静默处理错误
                                        }
                                    }, 1000); // 每秒执行一次
                                    
                                    // 将新的动态时钟添加到数组中
                                    window.dynamicTimeClocks.push(clockConfig);
                                    console.log('恢复动态时钟:', {
                                        workbook: config.workbookName,
                                        workbookFullPath: config.workbookFullName,
                                        sheet: config.sheetName,
                                        cell: config.address
                                    });
                                }
                            } catch (e) {
                                // 静默处理错误
                            }
                        }
                    }
                }
            }
        } catch (e) {
            // 静默处理错误
        }
    }

    // 调用恢复函数
    restoreDynamicTimeClocks();

    return true
}

function OnAction(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnFillImage":
            {
                // 填入图片功能：打开文件管理器，支持多选图片，并从选中单元格开始依次嵌入到所在列
                try {
                    console.log("开始执行填入图片功能");
                    
                    // 检查WPS应用环境
                    if (!window.Application) {
                        throw new Error("无法访问WPS应用对象");
                    }
                    
                    // 检查是否是Excel文档
                    if (!window.Application.ActiveWorkbook) {
                        throw new Error("此功能仅支持Excel文档");
                    }
                    
                    // 检查是否有选中的单元格
                    let selectedCell;
                    try {
                        selectedCell = window.Application.ActiveCell;
                        if (!selectedCell) {
                            throw new Error("没有选中的单元格");
                        }
                        console.log("选中的单元格:", selectedCell.Address);
                    } catch (e) {
                        alert("请先选中一个单元格，然后再执行此操作");
                        return true;
                    }
                    
                    // 1. 打开文件选择对话框，支持多选图片
                    let selectedFiles = [];
                    
                    // 使用WPS API的FileDialog方法打开文件选择对话框
                    if (window.Application.FileDialog) {
                        console.log("尝试使用Application.FileDialog方法打开文件选择对话框");
                        
                        const fileDialog = window.Application.FileDialog(3); // 3代表msoFileDialogFilePicker
                        
                        // 设置文件对话框属性
                        fileDialog.AllowMultiSelect = true; // 允许多选
                        fileDialog.Title = "选择图片文件"; // 对话框标题
                        fileDialog.Filters.Clear(); // 清除默认过滤器
                        // 添加图片文件过滤器
                        fileDialog.Filters.Add("图片文件", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.svg");
                        fileDialog.Filters.Add("所有文件", "*.*");
                        
                        // 显示文件对话框
                        const userSelected = fileDialog.Show();
                        if (userSelected) {
                            // 获取选中的文件路径
                            for (let i = 1; i <= fileDialog.SelectedItems.Count; i++) {
                                selectedFiles.push(fileDialog.SelectedItems.Item(i));
                            }
                            console.log("选中的图片文件:", selectedFiles);
                        } else {
                            console.log("用户取消了文件选择");
                            return true;
                        }
                    }                    
                    // 2. 将选中的图片从选中单元格开始依次嵌入到所在列
                    insertImagesToExcel(selectedFiles);
                    
                    console.log("填入图片功能执行完成");
                    
                } catch (error) {
                    console.error("填入图片失败:", error);
                    console.error("错误堆栈:", error.stack);
                    alert(`填入图片失败:\n详细错误: ${error.message}`);
                }
            }
            break
        case "btnAddComment":
            {
                // 添加批注功能：打开文件管理器，支持单选图片
                try {
                    console.log("开始执行添加批注功能");
                    
                    // 检查WPS应用环境
                    if (!window.Application) {
                        throw new Error("无法访问WPS应用对象");
                    }
                    
                    // 检查是否是Excel文档
                    if (!window.Application.ActiveWorkbook) {
                        throw new Error("此功能仅支持Excel文档");
                    }
                    
                    // 检查是否有选中的单元格
                    let selectedCell;
                    try {
                        selectedCell = window.Application.ActiveCell;
                        if (!selectedCell) {
                            throw new Error("没有选中的单元格");
                        }
                        console.log("选中的单元格:", selectedCell.Address);
                    } catch (e) {
                        alert("请先选中一个单元格，然后再执行此操作");
                        return true;
                    }
                    
                    // 打开文件选择对话框，允许多选图片
                    let selectedImages = [];
                    
                    // 使用WPS API的FileDialog方法打开文件选择对话框
                    if (window.Application.FileDialog) {
                        console.log("尝试使用Application.FileDialog方法打开文件选择对话框");
                        
                        const fileDialog = window.Application.FileDialog(3); // 3代表msoFileDialogFilePicker
                        
                        // 设置文件对话框属性
                        fileDialog.AllowMultiSelect = true; // 允许多选
                        fileDialog.Title = "选择批注图片"; // 对话框标题
                        fileDialog.Filters.Clear(); // 清除默认过滤器
                        // 添加图片文件过滤器
                        fileDialog.Filters.Add("图片文件", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.svg");
                        fileDialog.Filters.Add("所有文件", "*.*");
                        
                        // 显示文件对话框
                        const userSelected = fileDialog.Show();
                        if (userSelected) {
                            // 获取选中的图片路径
                            selectedImages = [];
                            for (let i = 1; i <= fileDialog.SelectedItems.Count; i++) {
                                selectedImages.push(fileDialog.SelectedItems.Item(i));
                            }
                            console.log("选中的批注图片:", selectedImages);
                            
                            // 检查是否有选中的图片
                            if (selectedImages.length === 0) {
                                alert("没有选中任何图片");
                                return true;
                            }
                            
                            // 将选中的图片以选中单元格为起点依次向下方单元格添加批注
                            try {
                                // 获取当前选中单元格
                                const activeSheet = window.Application.ActiveSheet;
                                const startCell = window.Application.ActiveCell;
                                
                                // 获取起始行列
                                const startRow = startCell.Row;
                                const startColumn = startCell.Column;
                                
                                // 遍历所有选中的图片，依次向下单元格添加批注
                                for (let i = 0; i < selectedImages.length; i++) {
                                    // 计算当前单元格的行列
                                    const currentRow = startRow + i;
                                    const currentColumn = startColumn;
                                    
                                    // 获取当前单元格
                                    const currentCell = activeSheet.Cells.Item(currentRow, currentColumn);
                                    if (!currentCell) {
                                        throw new Error(`无法获取单元格: 行${currentRow}，列${currentColumn}`);
                                    }
                                    
                                    // 1. 先删除单元格现有的批注（如果有）
                                    if (currentCell.Comment) {
                                        currentCell.Comment.Delete();
                                    }
                                    
                                    // 2. 为单元格添加新批注
                                    currentCell.AddComment("");
                                    
                                    // 3. 获取批注对象
                                    const comment = currentCell.Comment;
                                    if (!comment) {
                                        throw new Error(`无法为单元格添加批注: 行${currentRow}，列${currentColumn}`);
                                    }
                                    
                                    // 4. 在批注中插入图片
                                    // 获取批注的Shape对象
                                    const commentShape = comment.Shape;
                                    if (!commentShape) {
                                        throw new Error(`无法获取批注Shape对象: 行${currentRow}，列${currentColumn}`);
                                    }
                                    
                                    // 调整批注大小
                                    commentShape.Width = 680; // 设置批注宽度
                                    commentShape.Height = 420; // 设置批注高度
                                    
                                    // 5. 在批注中添加图片
                                    const selectedImage = selectedImages[i];

                                    // 7.使用AddPicture方法在批注中插入图片
                                    const shapes = commentShape.Shapes;
                                    if (shapes && typeof shapes.AddPicture === 'function') {
                                        // 在批注中插入图片
                                        const picture = shapes.AddPicture(
                                            selectedImage, // 图片路径
                                            false, // 不链接到文件
                                            true, // 保存到文档
                                            0, // 图片左上角X坐标
                                            0, // 图片左上角Y坐标
                                            commentShape.Width, // 图片宽度，与批注宽度相同
                                            commentShape.Height // 图片高度，与批注高度相同
                                        );
                                        
                                        console.log(`图片已添加到批注中: 行${currentRow}，列${currentColumn}`);
                                    } else {
                                        // 备选方案：使用Fill.UserPicture方法填充批注背景
                                        if (commentShape.Fill && typeof commentShape.Fill.UserPicture === 'function') {
                                            commentShape.Fill.UserPicture(selectedImage);
                                            console.log(`图片已作为背景添加到批注中: 行${currentRow}，列${currentColumn}`);
                                        } else {
                                            throw new Error(`当前WPS版本不支持在批注中插入图片: 行${currentRow}，列${currentColumn}`);
                                        }
                                    }
                                }
                                
                                alert(`成功为 ${selectedImages.length} 个单元格添加了批注图片！`);
                                
                            } catch (commentError) {
                                console.error("添加图片到批注失败:", commentError);
                                alert(`添加图片到批注失败:\n详细错误: ${commentError.message}`);
                            }
                        } else {
                            console.log("用户取消了文件选择");
                            return true;
                        }
                    } else {
                        throw new Error("当前WPS版本不支持FileDialog方法");
                    }
                    
                    console.log("添加批注功能执行完成");
                    
                } catch (error) {
                    console.error("添加批注失败:", error);
                    console.error("错误堆栈:", error.stack);
                    alert(`添加批注失败:\n详细错误: ${error.message}`);
                }
            }
                break
            case "btnMaterialQuery":
                {
                    // 物料查询功能：切换物料查询TaskPane的显示/隐藏
                    try {
                        console.log("开始执行物料查询功能");
                        
                        // 检查WPS应用环境
                        if (!window.Application) {
                            throw new Error("无法访问WPS应用对象");
                        }
                        
                        // 检查是否已存在任务窗格
                        let taskPaneId = window.Application.PluginStorage.getItem("material_query_taskpane_id");
                        let tskpane = null;
                        
                        if (taskPaneId) {
                            // 尝试获取已存在的任务窗格
                            try {
                                tskpane = window.Application.GetTaskPane(taskPaneId);
                            } catch (e) {
                                console.log("无法获取已存在的任务窗格，将创建新的任务窗格");
                                tskpane = null;
                            }
                        }
                        
                        if (tskpane && tskpane.Visible) {
                            // 如果任务窗格存在且可见，则隐藏它
                            tskpane.Visible = false;
                            console.log("隐藏任务窗格，ID:", taskPaneId);
                        } else {
                            // 如果任务窗格不存在或不可见，则创建或显示它
                            if (tskpane) {
                                // 任务窗格存在但不可见，直接显示
                                tskpane.Visible = true;
                                console.log("显示已有任务窗格，ID:", taskPaneId);
                            } else {
                                // 任务窗格不存在，创建新的
                                let baseUrl = window.location.href;
                                let taskPaneUrl = baseUrl.replace(/index\.html.*$/, 'index.html#/materialquery');
                                tskpane = window.Application.CreateTaskPane(taskPaneUrl)
                                let id = tskpane.ID
                                window.Application.PluginStorage.setItem("material_query_taskpane_id", id)
                                tskpane.Visible = true;
                                
                                console.log("创建新TaskPane，ID:", id, "URL:", taskPaneUrl);
                            }
                        }
                        
                    } catch (error) {
                        console.error("物料查询功能失败:", error);
                        console.error("错误堆栈:", error.stack);
                        alert(`物料查询功能失败:\n详细错误: ${error.message}`);
                    }
                }
                break

            case "btnDynamicTime":
                {
                    // 动态时间功能：在选中单元格生成每秒跳动的动态时间，再次点击则删除
                    try {
                        console.log('点击动态时间按钮');
                        // 检查WPS应用环境
                        if (!window.Application) {
                            throw new Error("无法访问WPS应用对象");
                        }
                        
                        // 检查是否是Excel文档
                        if (!window.Application.ActiveWorkbook) {
                            throw new Error("此功能仅支持Excel文档");
                        }
                        
                        // 检查是否有选中的单元格
                        let selectedCell;
                        try {
                            selectedCell = window.Application.ActiveCell;
                            if (!selectedCell) {
                                throw new Error("没有选中的单元格");
                            }
                            console.log('选中的单元格:', selectedCell);
                        } catch (e) {
                            console.log('没有选中单元格，返回');
                            return true;
                        }
                        
                        // 保存选中单元格的信息
                        const row = selectedCell.Row;
                        const col = selectedCell.Column;
                        const colLetter = String.fromCharCode(64 + col);
                        const cellAddress = `${colLetter}${row}`;
                        const sheetName = window.Application.ActiveSheet.Name;
                        const workbook = window.Application.ActiveWorkbook;
                        const workbookName = workbook.Name;
                        const workbookFullName = workbook.FullName;
                        
                        console.log('单元格信息:', {
                            workbook: workbookName,
                            workbookFullPath: workbookFullName,
                            sheet: sheetName,
                            cell: cellAddress
                        });
                        
                        // 初始化动态时钟数组（如果不存在）
                        if (!window.dynamicTimeClocks) {
                            window.dynamicTimeClocks = [];
                            console.log('初始化动态时钟数组');
                        }
                        
                        // 检查是否已经存在相同的动态时钟
                        let foundClockIndex = -1;
                        const targetHash = generateClockHash(workbookFullName, sheetName, cellAddress);
                        console.log('目标时钟hash:', targetHash);
                        for (let i = 0; i < window.dynamicTimeClocks.length; i++) {
                            const existingClock = window.dynamicTimeClocks[i];
                            const existingHash = generateClockHash(
                                existingClock.workbook.FullName,
                                existingClock.sheetName,
                                existingClock.address
                            );
                            console.log('检查现有时钟:', {
                                address: existingClock.address,
                                sheetName: existingClock.sheetName,
                                workbookName: existingClock.workbook.Name,
                                workbookFullPath: existingClock.workbook.FullName,
                                hash: existingHash
                            });
                            if (existingHash === targetHash) {
                                foundClockIndex = i;
                                console.log('找到匹配的动态时钟，准备删除');
                                break;
                            }
                        }
                        
                        if (foundClockIndex > -1) {
                            // 找到匹配的动态时钟，执行删除操作
                            const clockToDelete = window.dynamicTimeClocks[foundClockIndex];
                            
                            // 清除定时器
                            if (clockToDelete.intervalId) {
                                clearInterval(clockToDelete.intervalId);
                                console.log('定时器已清除');
                            }
                            
                            // 从数组中删除
                            window.dynamicTimeClocks.splice(foundClockIndex, 1);
                            console.log('从数组中删除成功');
                            
                            // 清空单元格内容
                            try {
                                const sheet = window.Application.ActiveSheet;
                                const range = sheet.Range(cellAddress);
                                range.Value2 = '';
                                console.log('单元格内容已清空');
                            } catch (e) {
                                console.error('清空单元格失败:', e);
                            }
                            
                            // 更新保存的配置
                            try {
                                // 先加载所有已有的配置
                                const allConfigs = loadDynamicTimeConfig();
                                // 准备当前工作簿的配置
                                const currentConfigs = [];
                                for (const clock of window.dynamicTimeClocks) {
                                    currentConfigs.push({
                                        workbookName: clock.workbook.Name,
                                        workbookFullName: clock.workbook.FullName,
                                        sheetName: clock.sheetName,
                                        address: clock.address
                                    });
                                }
                                // 合并配置：保留其他工作簿的配置，只更新当前工作簿的配置
                                const mergedConfigs = allConfigs.filter(config => 
                                    config.workbookFullName !== currentWorkbook.FullName
                                ).concat(currentConfigs);
                                // 保存到localStorage
                                saveDynamicTimeConfig(mergedConfigs);
                            } catch (e) {
                                console.error('保存配置失败:', e);
                            }
                            
                            console.log('动态时钟删除成功');
                        } else {
                            // 没有找到匹配的动态时钟，执行添加操作
                            console.log('没有找到匹配的动态时钟，准备添加');
                            
                            // 创建新的动态时钟配置
                            const clockConfig = {
                                address: cellAddress,
                                sheetName: sheetName,
                                workbook: workbook,
                                intervalId: null
                            };
                            
                            // 为当前单元格创建定时器
                            clockConfig.intervalId = setInterval(function() {
                                // 获取当前时间字符串
                                const timeString = getNowTimeString();
                                
                                // 更新单元格
                                try {
                                    // 检查Workbook对象是否有效
                                    if (clockConfig.workbook) {
                                        try {
                                            // 尝试访问Workbook的属性，检查是否有效
                                            const workbookName = clockConfig.workbook.Name;
                                            
                                            // Workbook有效，执行更新
                                            const sheet = clockConfig.workbook.Sheets.Item(clockConfig.sheetName);
                                            const range = sheet.Range(clockConfig.address);
                                            // 尝试使用Value2设置值
                                            range.Value2 = timeString;
                                            
                                            // 设置单元格格式
                                            range.NumberFormatLocal = "yyyy/m/d hh:mm:ss";
                                        } catch (workbookError) {
                                            // Workbook对象失效，清除定时器
                                            clearInterval(clockConfig.intervalId);
                                            // 从数组中删除
                                            if (window.dynamicTimeClocks) {
                                                const index = window.dynamicTimeClocks.indexOf(clockConfig);
                                                if (index > -1) {
                                                    window.dynamicTimeClocks.splice(index, 1);
                                                }
                                            }
                                        }
                                    }
                                } catch (e) {
                                    // 静默处理错误
                                }
                            }, 1000); // 每秒执行一次
                            
                            // 将新的动态时钟添加到数组中
                            window.dynamicTimeClocks.push(clockConfig);
                            console.log('添加新的动态时钟:', {
                                workbook: workbookName,
                                sheet: sheetName,
                                cell: cellAddress
                            });
                            
                            // 保存配置到localStorage
                            try {
                                // 先加载所有已有的配置
                                const allConfigs = loadDynamicTimeConfig();
                                // 准备当前工作簿的配置
                                const currentConfigs = [];
                                for (const clock of window.dynamicTimeClocks) {
                                    currentConfigs.push({
                                        workbookName: clock.workbook.Name,
                                        workbookFullName: clock.workbook.FullName,
                                        sheetName: clock.sheetName,
                                        address: clock.address
                                    });
                                }
                                // 合并配置：保留其他工作簿的配置，只更新当前工作簿的配置
                                const mergedConfigs = allConfigs.filter(config => 
                                    config.workbookFullName !== workbookFullName
                                ).concat(currentConfigs);
                                // 保存到localStorage
                                saveDynamicTimeConfig(mergedConfigs);
                            } catch (e) {
                                console.error('保存配置失败:', e);
                            }
                            
                            console.log('动态时钟添加成功');
                        }
                        
                    } catch (error) {
                        console.error('动态时间功能失败:', error);
                    }
                }
                break

            default:
                break
    }
    return true
}



// 将图片嵌入到Excel文档，使用WPS官方支持的方法
function insertImagesToExcel(imageFiles) {
    try {
        // 正确获取WPS应用对象和选中区域
        const app = window.Application;
        if (!app) {
            throw new Error("无法获取WPS应用对象");
        }
        
        // 获取当前选中区域
        let selection = app.Selection;
        if (!selection) {
            throw new Error("请先选中一个单元格");
        }        
        // 获取选中区域的起始行列
        const startRow = selection.Row;
        const startColumn = selection.Column;
        
        console.log(`开始从单元格 ${startRow} 行， ${startColumn} 列，嵌入图片，共 ${imageFiles.length} 个图片`);
        
        // 遍历所有选中的图片文件
        for (let i = 0; i < imageFiles.length; i++) {
            const imagePath = imageFiles[i];
            
            // 计算当前图片要插入的单元格位置
            const currentRow = startRow + i;
            const currentColumn = startColumn;
            
            // 获取当前单元格的地址
            const cellAddress = `${String.fromCharCode(64 + currentColumn)}${currentRow}`;
            console.log(`正在嵌入图片: ${imagePath} 到Excel单元格 ${cellAddress}`);
            
            // 验证图片路径
            if (!imagePath || typeof imagePath !== 'string') {
                console.error("图片路径无效:", imagePath);
                throw new Error(`图片路径无效: ${imagePath}`);
            }         
            
            // 根据用户提供的方法：Application.Selection.RangeEx.InsertCellPicture(imgPath)
            try {
                // 获取当前要插入图片的单元格
                const currentCell = app.ActiveSheet.Cells.Item(currentRow, currentColumn);
                if (!currentCell) {
                    throw new Error(`无法获取单元格: ${cellAddress}`);
                }
                
                // 选中当前单元格
                currentCell.Select();
                // 更新selection对象
                selection = app.Selection;
                
                // 使用用户提供的方法插入图片
                if (selection.RangeEx && typeof selection.RangeEx.InsertCellPicture === 'function') {
                    selection.RangeEx.InsertCellPicture(imagePath);
                    console.log(`图片嵌入成功: ${imagePath} 到Excel单元格 ${cellAddress}（使用Application.Selection.RangeEx.InsertCellPicture方法）`);
                } else {
                    throw new Error(`当前WPS版本不支持Selection.RangeEx.InsertCellPicture方法`);
                }
            } catch (e) {
                // 如果用户提供的方法失败，尝试其他备选方案
                console.log(`使用Application.Selection.RangeEx.InsertCellPicture方法失败: ${e.message}`);
                
                // 获取当前单元格范围
                const cellRange = app.ActiveSheet.Range(cellAddress);
                if (!cellRange) {
                    throw new Error(`无法获取单元格范围: ${cellAddress}`);
                }
                
                // 备选方案：使用getRangeEx方法
                if (cellRange.getRangeEx && typeof cellRange.getRangeEx === 'function') {
                    cellRange.getRangeEx().InsertCellPicture(imagePath);
                    console.log(`图片嵌入成功: ${imagePath} 到Excel单元格 ${cellAddress}（使用Range.getRangeEx().InsertCellPicture方法）`);
                } else {
                    throw new Error(`当前WPS版本不支持嵌入图片功能`);
                }
            }
        }
        
        // 选中起始单元格，方便用户查看结果
        app.ActiveSheet.Cells.Item(startRow, startColumn).Select();
        
        alert(`成功嵌入 ${imageFiles.length} 个图片到Excel表格中`);
    } catch (error) {
        console.error("嵌入图片到Excel失败:", error);
        console.error("错误堆栈:", error.stack);
        alert(`嵌入图片失败:\n详细错误: ${error.message}`);
        throw error;
    }
}

function GetImage(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnFillImage":
            // 使用images目录下的1.svg作为"填入图片"按钮的图标
            return "images/1.svg"
        case "btnAddComment":
            // 使用images目录下的2.svg作为"添加批注"按钮的图标
            return "images/2.svg"
        case "btnMaterialQuery":
            // 使用images目录下的search.svg作为"物料查询"按钮的图标
            return "images/search.svg"
        case "btnFuzzySearch":
            // 使用images目录下的search.svg作为"物料模糊查询"按钮的图标
            return "images/search.svg"
        case "btnDynamicTime":
            // 使用数字时钟样式的图标
            return "images/digital-clock.svg"

        default:
            return "images/newFromTemp.svg"
    }
}

function OnGetEnabled(control) {
    return true
}

function OnGetVisible(control){
    const eleId = control.Id
    // 显示"填入图片"、"添加批注"、"物料查询"和"动态时间"按钮
    return eleId === "btnFillImage" || eleId === "btnAddComment" || eleId === "btnMaterialQuery" || eleId === "btnDynamicTime"
}

function OnGetLabel(control){
    const eleId = control.Id
    switch (eleId) {
        case "btnFillImage":
            return "填入图片"
        case "btnAddComment":
            return "添加批注"
        case "btnMaterialQuery":
            return "物料查询"
        case "btnDynamicTime":
            return "动态时间"

        default:
            return ""
    }
}

// 控制标签的可见性
function OnGetTabVisible(control) {
    const tabId = control.Id
    // 只显示我们的HND插件标签
    return tabId === "wpsAddinTab"
}



// 获取当前时间字符串的函数
function getNowTimeString() {
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;
    const day = now.getDate();
    const hours = now.getHours().toString().padStart(2, '0');
    const minutes = now.getMinutes().toString().padStart(2, '0');
    const seconds = now.getSeconds().toString().padStart(2, '0');
    return `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
}

// 确保函数在全局作用域中可用
window.getNowTimeString = getNowTimeString;

//这些函数是给wps客户端调用的
export default {
    OnAddinLoad,
    OnAction,
    GetImage,
    OnGetEnabled,
    OnGetVisible,
    OnGetLabel,
    OnGetTabVisible
};