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
        console.log('保存当前时间配置:', uniqueConfigs);
        
        // 使用localStorage存储配置
        localStorage.setItem('hnd_date_time_configs', JSON.stringify(uniqueConfigs));
        console.log('配置保存成功');
    } catch (e) {
        console.error('保存配置失败:', e);
    }
}

function loadCurrentTimeConfig() {
    try {
        const configs = localStorage.getItem('hnd_date_time_configs');
        const parsedConfigs = configs ? JSON.parse(configs) : [];
        console.log('加载当前时间配置:', parsedConfigs);
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
    function updateClockCell(clockConfig) {
        try {
            // 检查Workbook对象是否有效
            if (!clockConfig.workbook) return;
            
            // 尝试访问Workbook的属性，检查是否有效
            clockConfig.workbook.Name;
            
            // Workbook有效，执行更新
            const sheet = clockConfig.workbook.Sheets.Item(clockConfig.sheetName);
            const range = sheet.Range(clockConfig.address);
            // 尝试使用Value2设置值
            const timeString = getNowTimeString();
            
            // 只在值发生变化时更新，减少对撤回栈的影响
            if (range.Value2 !== timeString) {
                range.Value2 = timeString;
            }
            
            // 只在第一次设置格式，避免重复操作影响撤回栈
            if (!clockConfig.formatSet) {
                range.NumberFormatLocal = "yyyy/mm/dd hh:mm:ss";
                clockConfig.formatSet = true;
            }
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

    function createClockConfig(config, currentWorkbook) {
        try {
            // 检查工作表是否存在
            const sheet = currentWorkbook.Sheets.Item(config.sheetName);
            if (!sheet) return null;
            
            // 创建新的动态时钟配置
            const clockConfig = {
                address: config.address,
                sheetName: config.sheetName,
                workbook: currentWorkbook,
                intervalId: null,
                formatSet: false
            };
            
            // 为当前单元格创建定时器
            clockConfig.intervalId = setInterval(function() {
                updateClockCell(clockConfig);
            }, 1000); // 每秒执行一次
            
            return clockConfig;
        } catch (e) {
            // 静默处理错误
            return null;
        }
    }

    function restoreCurrentTimeClocks() {
        try {
            const savedConfigs = loadCurrentTimeConfig();
            if (savedConfigs.length === 0) {
                console.log('没有保存的当前时间配置');
                return;
            }
            
            // 获取当前工作簿
            const currentWorkbook = window.Application.ActiveWorkbook;
            if (!currentWorkbook) {
                console.log('没有活动工作簿');
                return;
            }
            
            const currentWorkbookFullName = currentWorkbook.FullName;
            let restoredCount = 0;
            
            // 检查是否有匹配当前工作簿的配置
            for (const config of savedConfigs) {
                if (config.workbookFullName !== currentWorkbookFullName) continue;
                
                const clockConfig = createClockConfig(config, currentWorkbook);
                if (!clockConfig) continue;
                
                // 将新的动态时钟添加到数组中
                window.dynamicTimeClocks.push(clockConfig);
                console.log('恢复动态时钟:', {
                    workbook: config.workbookName,
                    workbookFullPath: config.workbookFullName,
                    sheet: config.sheetName,
                    cell: config.address
                });
                restoredCount++;
            }
            
            console.log(`共恢复了 ${restoredCount} 个动态时钟`);
        } catch (e) {
            console.error('恢复动态时钟失败:', e);
        }
    }



    // 调用恢复函数
    restoreCurrentTimeClocks();

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
                    
                    // 检查FileDialog支持
                    if (!window.Application.FileDialog) {
                        throw new Error("当前WPS版本不支持FileDialog方法");
                    }
                    
                    // 打开文件选择对话框
                    console.log("尝试使用Application.FileDialog方法打开文件选择对话框");
                    const fileDialog = window.Application.FileDialog(3); // 3代表msoFileDialogFilePicker
                    
                    // 设置文件对话框属性
                    fileDialog.AllowMultiSelect = true; // 允许多选
                    fileDialog.Title = "选择批注图片"; // 对话框标题
                    fileDialog.Filters.Clear(); // 清除默认过滤器
                    fileDialog.Filters.Add("图片文件", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.svg");
                    fileDialog.Filters.Add("所有文件", "*.*");
                    
                    // 显示文件对话框
                    const userSelected = fileDialog.Show();
                    if (!userSelected) {
                        console.log("用户取消了文件选择");
                        return true;
                    }
                    
                    // 获取选中的图片路径
                    const selectedImages = [];
                    for (let i = 1; i <= fileDialog.SelectedItems.Count; i++) {
                        selectedImages.push(fileDialog.SelectedItems.Item(i));
                    }
                    console.log("选中的批注图片:", selectedImages);
                    
                    // 检查是否有选中的图片
                    if (selectedImages.length === 0) {
                        alert("没有选中任何图片");
                        return true;
                    }
                    
                    // 获取当前选中单元格和起始位置
                    const activeSheet = window.Application.ActiveSheet;
                    const startCell = window.Application.ActiveCell;
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
                        const commentShape = comment.Shape;
                        if (!commentShape) {
                            throw new Error(`无法获取批注Shape对象: 行${currentRow}，列${currentColumn}`);
                        }
                        
                        // 调整批注大小
                        commentShape.Width = 680;
                        commentShape.Height = 420;
                        
                        // 5. 在批注中添加图片
                        const selectedImage = selectedImages[i];

                        // 使用AddPicture方法在批注中插入图片
                        const shapes = commentShape.Shapes;
                        if (shapes && typeof shapes.AddPicture === 'function') {
                            shapes.AddPicture(
                                selectedImage,
                                false,
                                true,
                                0,
                                0,
                                commentShape.Width,
                                commentShape.Height
                            );
                            console.log(`图片已添加到批注中: 行${currentRow}，列${currentColumn}`);
                        } else if (commentShape.Fill && typeof commentShape.Fill.UserPicture === 'function') {
                            // 备选方案：使用Fill.UserPicture方法填充批注背景
                            commentShape.Fill.UserPicture(selectedImage);
                            console.log(`图片已作为背景添加到批注中: 行${currentRow}，列${currentColumn}`);
                        } else {
                            throw new Error(`当前WPS版本不支持在批注中插入图片: 行${currentRow}，列${currentColumn}`);
                        }
                    }
                    
                    console.log(`成功为 ${selectedImages.length} 个单元格添加了批注图片！`);
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
                    // 物料查询功能：使用统一管理函数
                    manageTaskPane("btnMaterialQuery");
                }
                break

            case "btnDynamicTime":
                {
                    // 当前时间功能：在选中单元格生成每秒跳动的当前时间，再次点击则删除
                    try {
                        console.log('点击当前时间按钮');
                        // 内部辅助函数
                        const updateClockCell = function(clockConfig) {
                            try {
                                // 检查Workbook对象是否有效
                                if (!clockConfig.workbook) return;
                                
                                // 尝试访问Workbook的属性，检查是否有效
                                clockConfig.workbook.Name;
                                
                                // Workbook有效，执行更新
                                const sheet = clockConfig.workbook.Sheets.Item(clockConfig.sheetName);
                                const range = sheet.Range(clockConfig.address);
                                // 尝试使用Value2设置值
                                const timeString = getNowTimeString();
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
                        };
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
                        const workbookFullPath = workbook.FullName;
                        
                        console.log('单元格信息:', {
                            workbook: workbookName,
                            workbookFullPath: workbookFullPath,
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
                        const targetHash = generateClockHash(workbookFullPath, sheetName, cellAddress);
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
                                const allConfigs = loadCurrentTimeConfig();
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
                                    config.workbookFullName !== workbook.FullName
                                ).concat(currentConfigs);
                                // 保存到localStorage
                                saveDynamicTimeConfig(mergedConfigs);
                                console.log('配置已更新:', mergedConfigs);
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
                                intervalId: null,
                                formatSet: false
                            };
                            
                            // 为当前单元格创建定时器
                            clockConfig.intervalId = setInterval(function() {
                                try {
                                    // 检查Workbook对象是否有效
                                    if (!clockConfig.workbook) return;
                                    
                                    // 尝试访问Workbook的属性，检查是否有效
                                    clockConfig.workbook.Name;
                                    
                                    // Workbook有效，执行更新
                                    const sheet = clockConfig.workbook.Sheets.Item(clockConfig.sheetName);
                                    const range = sheet.Range(clockConfig.address);
                                    const timeString = getNowTimeString();
                                    
                                    // 只在值发生变化时更新，减少对撤回栈的影响
                                    if (range.Value2 !== timeString) {
                                        range.Value2 = timeString;
                                    }
                                    
                                    // 只在第一次设置格式，避免重复操作影响撤回栈
                                    if (!clockConfig.formatSet) {
                                        range.NumberFormatLocal = "yyyy/mm/dd hh:mm:ss";
                                        clockConfig.formatSet = true;
                                    }
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
                                const allConfigs = loadCurrentTimeConfig();
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
                                    config.workbookFullName !== workbookFullPath
                                ).concat(currentConfigs);
                                // 保存到localStorage
                                saveDynamicTimeConfig(mergedConfigs);
                            } catch (e) {
                                console.error('保存配置失败:', e);
                            }
                            
                            console.log('动态时钟添加成功');
                        }
                        
                    } catch (error) {
                        console.error('当前时间功能失败:', error);
                    }
                }
                break

            case "btnScientificCalculator":
                {
                    // 科学计算器功能：使用统一管理函数
                    manageTaskPane("btnScientificCalculator");
                }
                break

            case "btnCurrencyConverter":
                {
                    // 汇率换算功能：使用统一管理函数
                    manageTaskPane("btnCurrencyConverter");
                }
                break
                
            case "btnTankVolume":
                {
                    // 储罐容积计算功能：使用统一管理函数
                    manageTaskPane("btnTankVolume");
                }
                break
            case "btnQRCode":
                {
                    // 二维码功能：使用统一管理函数
                    manageTaskPane("btnQRCode");
                }
                break
            case "btnMoyu":
                {
                    // 摸鱼功能：使用统一管理函数
                    manageTaskPane("btnMoyu");
                }
                break
            case "btnTetris":
                {
                    // 俄罗斯方块功能：使用统一管理函数
                    manageTaskPane("btnTetris");
                }
                break

            default:
                break
    }
    return true
}

// 完全清除所有当前时间配置的函数
function clearAllCurrentTimeClocks() {
    try {
        // 清除所有定时器
        if (window.dynamicTimeClocks) {
            for (const clock of window.dynamicTimeClocks) {
                if (clock.intervalId) {
                    clearInterval(clock.intervalId);
                    console.log('定时器已清除:', clock.address);
                }
            }
            // 清空动态时钟数组
            window.dynamicTimeClocks = [];
            console.log('当前时间数组已清空');
        }
        
        // 清空localStorage中的配置
        localStorage.removeItem('hnd_date_time_configs');
        console.log('localStorage中的当前时间配置已清空');
        
        console.log('所有当前时间配置已完全清除');
    } catch (e) {
        console.error('清除当前时间配置失败:', e);
    }
}

// 确保函数在全局作用域中可用
window.clearAllCurrentTimeClocks = clearAllCurrentTimeClocks;



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
        
        console.log(`成功嵌入 ${imageFiles.length} 个图片到Excel表格中`);
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
            // 使用images目录下的add-image.svg作为"填入图片"按钮的图标
            return "images/add-image.svg"
        case "btnAddComment":
            // 使用images目录下的add-comment.svg作为"添加批注"按钮的图标
            return "images/add-comment.svg"
        case "btnMaterialQuery":
            // 使用images目录下的query.svg作为"物料查询"按钮的图标
            return "images/query.svg"
        case "btnFuzzySearch":
            // 使用images目录下的query.svg作为"物料模糊查询"按钮的图标
            return "images/query.svg"
        case "btnDynamicTime":
            // 使用日期时间样式的图标
            return "images/rqsj.svg"
        case "btnScientificCalculator":
            // 使用计算器样式的图标
            return "images/calculator.svg"
        case "btnCurrencyConverter":
            // 使用汇率换算样式的图标
            return "images/exchange-rate.svg"
        case "btnTankVolume":
            // 使用液位体积计算样式的图标
            return "images/horizontal-tank.svg"
        case "btnQRCode":
            // 使用二维码样式的图标
            return "images/ewm.svg"
        case "btnMoyu":
            // 使用数独样式的图标
            return "images/shudu.svg"
        case "btnTetris":
            // 使用俄罗斯方块样式的图标
            return "images/tile.svg"

        default:
            return "images/newFromTemp.svg"
    }
}

function OnGetEnabled(control) {
    return true
}

function OnGetVisible(control){
    const eleId = control.Id
    // 显示"填入图片"、"添加批注"、"物料查询"、"当前时间"、"科学计算器"、"汇率换算"、"储罐容积计算"、"二维码"、"数独"和"俄罗斯方块"按钮
    return eleId === "btnFillImage" || eleId === "btnAddComment" || eleId === "btnMaterialQuery" || eleId === "btnDynamicTime" || eleId === "btnScientificCalculator" || eleId === "btnCurrencyConverter" || eleId === "btnTankVolume" || eleId === "btnQRCode" || eleId === "btnMoyu" || eleId === "btnTetris"
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
            return "当前时间"
        case "btnScientificCalculator":
            return "科学计算器"
        case "btnCurrencyConverter":
            return "汇率换算"
        case "btnTankVolume":
            return "液位体积计算"
        case "btnQRCode":
            return "二维码"
        case "btnMoyu":
            return "数独"
        case "btnTetris":
            return "俄罗斯方块"

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
    const month = (now.getMonth() + 1).toString().padStart(2, '0');
    const day = now.getDate().toString().padStart(2, '0');
    const hours = now.getHours().toString().padStart(2, '0');
    const minutes = now.getMinutes().toString().padStart(2, '0');
    const seconds = now.getSeconds().toString().padStart(2, '0');
    return `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
}

// 确保函数在全局作用域中可用
window.getNowTimeString = getNowTimeString;

// 任务窗格配置信息 - 分为两组
const taskPaneGroups = {
    right: [
        {
            id: "fill_image_taskpane_id",
            key: "btnFillImage",
            url: "index.html#/fill-image",
            name: "填入图片"
        },
        {
            id: "add_comment_taskpane_id",
            key: "btnAddComment",
            url: "index.html#/add-comment",
            name: "添加批注"
        },
        {
            id: "material_query_taskpane_id",
            key: "btnMaterialQuery",
            url: "index.html#/materialquery",
            name: "物料查询"
        },
        {
            id: "dynamic_time_taskpane_id",
            key: "btnDynamicTime",
            url: "index.html#/dynamic-time",
            name: "动态时间"
        },
        {
            id: "scientific_calculator_taskpane_id",
            key: "btnScientificCalculator",
            url: "index.html#/calculator",
            name: "科学计算器"
        },
        {
            id: "currency_converter_taskpane_id",
            key: "btnCurrencyConverter",
            url: "index.html#/currency",
            name: "汇率换算"
        },
        {
            id: "tank_volume_taskpane_id",
            key: "btnTankVolume",
            url: "index.html#/tank-volume",
            name: "液位体积计算"
        },
        {
            id: "qrcode_taskpane_id",
            key: "btnQRCode",
            url: "index.html#/qrcode",
            name: "二维码"
        }
    ],
    left: [
        {
            id: "tetris_taskpane_id",
            key: "btnTetris",
            url: "index.html#/tetris",
            name: "俄罗斯方块"
        },
        {
            id: "moyu_taskpane_id",
            key: "btnMoyu",
            url: "index.html#/moyu",
            name: "数独"
        }
    ]
};

// 统一管理任务窗格的函数
function manageTaskPane(buttonKey) {
  try {
    if (!window.Application) {
      throw new Error("无法访问WPS应用对象");
    }
    
    console.log(`开始管理任务窗格: ${buttonKey}`);
    
    // 找到当前按钮对应的任务窗格配置和组
    let currentConfig = null;
    let currentGroup = null;
    
    // 查找当前按钮属于哪个组
    for (const [groupName, groupConfigs] of Object.entries(taskPaneGroups)) {
        const config = groupConfigs.find(c => c.key === buttonKey);
        if (config) {
            currentConfig = config;
            currentGroup = groupName;
            break;
        }
    }
    
    if (!currentConfig || !currentGroup) {
      console.error(`未找到任务窗格配置: ${buttonKey}`);
      return;
    }
    
    // 获取当前组的所有配置
    const groupConfigs = taskPaneGroups[currentGroup];
    
    // 1. 检查当前窗格是否开启，如果开启，则关闭
    console.log(`检查当前任务窗格状态: ${currentConfig.name}`);
    const currentTaskPaneId = window.Application.PluginStorage.getItem(currentConfig.id);
    let tskpane = null;
    let isVisible = false;
    
    if (currentTaskPaneId) {
      try {
        tskpane = window.Application.GetTaskPane(currentTaskPaneId);
        isVisible = tskpane && tskpane.Visible;
        console.log(`当前任务窗格状态: ${isVisible ? "开启" : "关闭"}`);
      } catch (e) {
        console.log(`无法获取已存在的任务窗格: ${currentConfig.name}`);
        tskpane = null;
        isVisible = false;
      }
    }
    
    if (isVisible) {
      // 当前窗格已开启，关闭它
      tskpane.Visible = false;
      console.log(`关闭任务窗格: ${currentConfig.name}`);
    } else {
      // 2. 如果当前窗格未开启，则先关闭组内的其他任务窗格，再打开当前窗格
      console.log(`关闭组内其他任务窗格`);
      for (const config of groupConfigs) {
        if (config.key === buttonKey) continue; // 跳过当前任务窗格
        
        const taskPaneId = window.Application.PluginStorage.getItem(config.id);
        if (taskPaneId) {
          try {
            const groupTskpane = window.Application.GetTaskPane(taskPaneId);
            if (groupTskpane && groupTskpane.Visible) {
              groupTskpane.Visible = false;
              console.log(`隐藏任务窗格: ${config.name}`);
            }
          } catch (e) {
            console.log(`无法获取任务窗格: ${config.name}`);
          }
        }
      }
      
      // 3. 打开当前任务窗格
      console.log(`打开任务窗格: ${currentConfig.name}`);
      if (tskpane) {
        // 任务窗格存在但不可见，直接显示
        tskpane.Visible = true;
        // 根据组设置不同的停靠位置
        if (window.Application.Enum) {
          const dockPosition = currentGroup === 'left' ? 
            window.Application.Enum.msoCTPDockPositionLeft : 
            window.Application.Enum.msoCTPDockPositionRight;
          tskpane.DockPosition = dockPosition;
          console.log(`设置任务窗格${currentGroup === 'left' ? '左侧' : '右侧'}停靠: ${currentConfig.name}`);
        }
        console.log(`显示任务窗格: ${currentConfig.name}`);
      } else {
        // 任务窗格不存在，创建新的
        let baseUrl = window.location.href;
        let taskPaneUrl = baseUrl.replace(/index\.html.*$/, currentConfig.url);
        tskpane = window.Application.CreateTaskPane(taskPaneUrl);
        let id = tskpane.ID;
        window.Application.PluginStorage.setItem(currentConfig.id, id);
        // 根据组设置不同的停靠位置
        if (window.Application.Enum) {
          const dockPosition = currentGroup === 'left' ? 
            window.Application.Enum.msoCTPDockPositionLeft : 
            window.Application.Enum.msoCTPDockPositionRight;
          tskpane.DockPosition = dockPosition;
          console.log(`设置任务窗格${currentGroup === 'left' ? '左侧' : '右侧'}停靠: ${currentConfig.name}`);
        }
        tskpane.Visible = true;
        console.log(`创建并显示任务窗格: ${currentConfig.name}, ID: ${id}`);
      }
    }
    
  } catch (error) {
    console.error("管理任务窗格失败:", error);
    console.error("错误堆栈:", error.stack);
    alert(`管理任务窗格失败:\n详细错误: ${error.message}`);
  }
}

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