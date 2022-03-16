# PerformanceData
效能數據視覺化



### 介面示意圖：

<img src=https://user-images.githubusercontent.com/76928680/158659231-d8b0ed94-55ec-488c-bb86-483075de0df0.png width=50% />



## [使用說明]
目錄：

```
    ├── main.py
    ├── PerformanceTestingReport.py
    ├── Report/
        └── report.html
```

- `PerformanceTestingReport.py`：進行資料處理、視覺化

- `main.py`：GUI主介面邏輯
 > 步驟
 > 1. 將原始Excel文件 (xlsx格式) 拖動到區域內，顯示該文件檔名
 > 2. 點選 *Start* 自動執行`PerformanceTestingReport.py`，且子資料夾Report生成.html檔案
 
 



## [Change Log]
### (2021)
針對小工具做了對應的修改及優化：
- 修改執行檔：解壓後將相關模組暫存到本機減少載入時間
- 修改場景標籤(Bar chart)、圖表的軸距計算判斷
- 新增Android機型參考數據Stutter、InterFrame
- 新增html檔案檢查、更新命名規則
- 新增按鈕觸發事件：可一鍵開啟html及測試報告文件對應路徑
