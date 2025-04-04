# 聯盟管理系統 (Alliance Management System)

這是一個用於Google Sheets的聯盟管理系統，可以幫助您追蹤成員參與度、管理活動、計算分數並提供等級晉升建議。

## 系統初始化流程

要開始使用聯盟管理系統，請按照以下步驟初始化：

1. **下載專案檔案**
   - 下載本專案的所有檔案到您的電腦
   - 解壓縮下載的檔案

2. **創建新的Google Sheets檔案**
   - 前往[Google Sheets](https://sheets.google.com)
   - 創建一個新的空白試算表

3. **開啟Apps Script編輯器**
   - 在Google Sheets中，點選「擴充功能」→「Apps Script」
   - 這將開啟Apps Script編輯器

4. **創建專案檔案**
   - 在Apps Script編輯器中，您會看到一個預設的空白檔案
   - 按一下檔案旁邊的「+」按鈕創建新檔案
   - 為每個新檔案輸入與專案中相同的檔案名稱，確保使用正確的副檔名（.gs或.html）
   - 需要創建的檔案包括：
     - `initialize.gs`
     - `MenuManagement.gs`
     - `alliance manager.gs`
     - `ActivityManager.html`
     - `ActivityRecording.html`
     - `ActivitySidebar.html`
     - `ActivityViewer.html`
     - `MembersDialog.html`
     - `RankedMembers.html`

5. **複製程式碼**
   - 打開您下載的每個檔案
   - 將各檔案的內容複製到Apps Script編輯器中相應的檔案中
   - 按Ctrl+S（或Cmd+S）儲存每個檔案

6. **初始化系統**
   - 在Apps Script編輯器中，選擇`initialize.gs`檔案
   - 從上方的函數下拉選單中選擇`runSetup`
   - 點擊「執行」按鈕（▶️）
   - 系統將會建立所有必要的工作表和設定

7. **授權與啟用**
   - 第一次執行時，系統會要求您授權Apps Script存取您的Google Sheets
   - 按照螢幕上的指示完成授權流程
   - 完成後，回到Google Sheets並重新整理頁面
   - 您現在應該可以看到頂部選單中出現「聯盟系統」選項

## 系統功能概述

聯盟管理系統提供以下主要功能：

- **成員管理**：追蹤成員資訊、等級和戰力
- **活動管理**：設定多層級活動結構和權重
- **活動記錄**：記錄成員參與活動情況和表現
- **分數計算**：根據活動參與和表現計算成員總分
- **等級建議**：根據設定的標準提供成員晉升或降級建議
- **數據視覺化**：顯示成員排名和活動參與統計資料

## 基本使用指南

初始化系統後，您可以通過頂部的「聯盟系統」選單存取各種功能：

1. **記錄活動**：記錄成員參與特定活動的情況
2. **管理成員**：添加、編輯或停用成員
3. **管理活動**：設定活動結構和評分標準
4. **成員排名**：查看基於總分的成員排名
5. **查看活動記錄**：檢視特定日期或活動的參與記錄
6. **更新儀表板**：更新儀表板上的概覽資訊
7. **計算分數**：為所有成員計算最新總分
8. **生成等級建議**：根據設定的標準生成晉升/降級建議

系統會自動設定每日和每週的自動化任務，以保持分數和建議的更新。

## 系統結構

系統包含以下工作表：

- **Dashboard**：提供系統概覽和重要指標
- **Members**：存儲所有成員資訊
- **Activities**：定義活動層級結構和權重
- **Participation**：記錄所有活動參與數據
- **WeightConfig**：配置各種計分權重
- **EvaluationCriteria**：定義晉升和降級的標準

這個系統設計為靈活可配置，可以適應不同類型的組織或聯盟需求。
