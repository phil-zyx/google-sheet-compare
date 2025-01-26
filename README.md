# Google Sheets Helper - google 表格工具

## 介绍
该仓库实现了一些基于 Google Sheets 表格 Apps Script 的工具类功能，主要专注于解决将 Goole Sheets 表格作为配置表需求的业务处理方案。

## 功能
- [x] Base64 编解码函数
- [ ] 支持两个表格页签的数据对比
  - [x] 提供页签对比界面
  - [ ] 对比完成像 `git` 一样标识出差异项
- [ ] 页签合并功能
  - [ ] 将对比差异修改后，提供将差异合并的功能，有点类似 git merge 
- [ ] ID 冲突检查
  - [ ] 新增 `id` 后全页签检查，标记出冲突的ID，从基础上移除 ID 冲突

## 安装及使用方法

### 方案1
从商店直接搜索 google sheets help 使用。

### 方案2（可自行修改代码）
1. 打开 Google Sheets
2. 点击 "Extensions" > "Apps Script"
3. 复制以下文件到对应位置：
   - Code.gs
   - Compare.gs
   - CompareDialog.html
4. 保存后刷新界面