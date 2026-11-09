# Com Port Scanner

列出 Windows 上所有序列埠（COM Port），檢查是否被其他程式佔用，並可移除已拔除但還佔著 COM 號碼的殘留裝置。

## 檔案

| 檔案 | 說明 |
|---|---|
| `ScanSerialPorts.ps1` | 主要腳本 |
| `RunSerialPorts.bat` | 雙擊執行用，會自動以系統管理員身分開啟 PowerShell 並執行腳本 |

## 使用方式

### 方法一：雙擊執行（建議）

1. 把兩個檔案放在同一個資料夾
2. 雙擊 `RunSerialPorts.bat`
3. 出現 UAC 視窗時按「是」，允許以系統管理員身分執行

### 方法二：手動執行

以**系統管理員身分**開啟 PowerShell，切換到腳本所在資料夾後執行：

```powershell
cd "C:\Path\To\Script"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\ScanSerialPorts.ps1
```

`Set-ExecutionPolicy -Scope Process` 只對這次開啟的 PowerShell 視窗有效，關閉後就會恢復原本的設定。

### 只列出清單

```powershell
.\ScanSerialPorts.ps1 -ListOnly
```

只顯示清單，不會詢問任何問題，也不會開啟或移除任何埠，不需要系統管理員權限。

## 執行流程

1. **列出所有 COM 埠**，包含：
   - 裝置管理員「連接埠 (COM 和 LPT)」裡的所有裝置，包括已拔除的裝置（USB 轉序列線、內建序列埠等）
   - 系統有登記、但不在上述類別裡的 COM 埠（通常是虛擬序列埠軟體建立的），狀態會顯示 `Virtual/Other`

2. **詢問是否檢查佔用狀態**（輸入 `y` 才會檢查）：會嘗試短暫開啟每個連接中的埠
   - `Free`：可以使用
   - `In use`：正在被其他程式使用
   - `Unknown`：無法判斷

3. **顯示結果表格**

   | 欄位 | 說明 |
   |---|---|
   | Port | COM 號碼 |
   | Present | 裝置是否還接著 |
   | Status | `OK`、`Error`、`Unplugged`（已拔除）、`Virtual/Other` |
   | InUse | 第 2 步的檢查結果，沒有檢查時顯示 `-` |
   | Name | 裝置名稱 |

4. **移除殘留裝置**：列出已拔除（`Present = False`）的裝置，輸入 `Yes` 後用 Windows 內建的 `pnputil /remove-device` 移除，釋放它們佔用的 COM 號碼

5. 按 Enter 結束

## 注意事項

- **只會移除已拔除的裝置**，目前接著的裝置不會被移除
- 移除裝置需要系統管理員權限，以及 Windows 10 2004（組建 19041）以上版本。版本較舊時，腳本會提示改用：裝置管理員 → 檢視 → 顯示隱藏的裝置 → 連接埠 (COM 和 LPT)
- 檢查佔用狀態時會短暫開啟序列埠，**部分開發板（例如 Arduino）可能會因此重新啟動**，所以預設不檢查
- 移除後，下次插上同一個裝置時 Windows 會重新安裝驅動，並分配新的 COM 號碼
- 支援 Windows PowerShell 5.1 和 PowerShell 7

## 更新紀錄

### 2026-11-09
- 修正移除功能：原本用錯裝置 ID，而且呼叫了不存在的方法，所以從來沒有成功移除過。改用 `pnputil /remove-device`
- 改成只移除已拔除的殘留裝置，不會再對所有埠嘗試移除
- 改為從「連接埠 (COM 和 LPT)」類別找裝置，USB 轉序列線（CH340、FTDI、PL2303 等）不會再被誤判成虛擬序列埠
- 新增佔用狀態檢查和結果表格
- 新增 `-ListOnly` 參數
- 改用 `Get-PnpDevice`，PowerShell 7 也可以執行
