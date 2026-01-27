<template>
  <div class="hello">
    <div class="global">
      <div class="divItem">
        这是一个网页，按<span style="font-weight: bolder">"F12"</span>可以打开调试器。
      </div>
      <div class="divItem">
        这个示例展示了wps加载项的相关基础能力，与B/S业务系统的交互，请用浏览器打开：
        <span style="font-weight: bolder; color: slateblue; cursor: pointer" @click="onOpenWeb()">{{
          DemoSpan
        }}</span>
      </div>
      <div class="divItem">
        开发文档:
        <span style="font-weight: bolder; color: slateblue">https://open.wps.cn/docs/office</span>
      </div>
      <hr />
      <div class="divItem">
        <button style="margin: 3px" @click="onDocNameClick()">取文件名</button>
        <button style="margin: 3px" @click="onbuttonclick('createTaskPane')">创建任务窗格</button>
        <button style="margin: 3px" @click="onbuttonclick('newDoc')">新建文件</button>
        <button style="margin: 3px" @click="onbuttonclick('addString')">文档开头添加字符串</button>
        <button style="margin: 3px" @click="onbuttonclick('closeDoc')">关闭文件</button>
        <button style="margin: 3px" @click="insertPictures">批量嵌入图片</button>
      </div>
      <hr />
      <div class="divItem">
        文档文件名为：<span>{{ docName }}</span>
      </div>
    </div>
  </div>
</template>

<script>
import { onMounted } from 'vue'
import dlgFunc from './js/dialog.js'
import axios from 'axios'

export default {
  name: 'Dialog',
  data() {
    return {
      DemoSpan: '',
      docName: ''
    }
  },
  methods: {
    onbuttonclick(id) {
      return dlgFunc.onbuttonclick(id)
    },
    onDocNameClick() {
      this.docName = dlgFunc.onbuttonclick('getDocName')
    },
    onOpenWeb() {
      dlgFunc.onbuttonclick('openWeb', this.DemoSpan)
    },
    insertPictures() {
      try {
        // 正确获取WPS应用对象和选中区域
        const app = window.Application;
        let selection = app.Selection;
        
        // 获取当前工作表
        const activeSheet = app.ActiveSheet;
        if (!activeSheet) {
          throw new Error("无法获取当前工作表");
        }
        
        // 获取选中区域的起始行列
        const startRow = selection.Row;
        const startColumn = selection.Column;
        
        console.log(`开始从Excel单元格 ${selection.Address} 嵌入图片`);
        
        // 打开文件选择对话框，支持多选图片
        let selectedFiles = [];
        
        // 使用WPS API的FileDialog方法打开文件选择对话框
        if (app.FileDialog) {
          console.log("尝试使用Application.FileDialog方法打开文件选择对话框");
          
          const fileDialog = app.FileDialog(3); // 3代表msoFileDialogFilePicker
          
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
            return;
          }
        } else {
          alert("当前WPS版本不支持FileDialog方法，无法选择图片文件");
          return;
        }
        
        // 遍历所有选中的图片文件
        for (let i = 0; i < selectedFiles.length; i++) {
          const imagePath = selectedFiles[i];
          
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
            const currentCell = activeSheet.Cells.Item(currentRow, currentColumn);
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
            const cellRange = activeSheet.Range(cellAddress);
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
        activeSheet.Cells.Item(startRow, startColumn).Select();
        
        alert(`成功嵌入 ${selectedFiles.length} 个图片到Excel表格中`);
      } catch (error) {
        console.error("嵌入图片到Excel失败:", error);
        console.error("错误堆栈:", error.stack);
        alert(`嵌入图片失败:\n详细错误: ${error.message}`);
      }
    }
  }
}
onMounted(() => {
  axios.get('/.debugTemp/NotifyDemoUrl').then((res) => {
    this.DemoSpan = res.data
  })
})</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
.global {
  font-size: 15px;
  min-height: 95%;
}
.divItem {
  margin-left: 5px;
  margin-bottom: 18px;
  font-size: 15px;
  word-wrap: break-word;
}
</style>
