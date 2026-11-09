# Com Port Scanner

列出 Windows 上所有序列埠（COM Port），並可選擇移除殘留或被佔用的裝置（例如沒有正確移除的虛擬 COM Port）的 PowerShell 腳本。

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

1. 以**系統管理員身分**開啟 PowerShell
2. 切換到腳本所在資料夾，然後執行：

   ```powershell
   cd "C:\Path\To\Script"
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\ScanSerialPorts.ps1
   ```

   `Set-ExecutionPolicy -Scope Process` 只對這次開啟的 PowerShell 視窗有效，關閉後就會恢復原本的設定。

## 執行流程

1. **列出序列埠**：用 WMI（`Win32_SerialPort`）查詢，顯示每個埠的 `DeviceID` 和說明
2. **詢問是否移除裝置**：輸入 `Yes` 才會移除，輸入其他任何內容都會略過
   ```
   Do you want to remove occupied devices? (Type 'Yes' to confirm, any other key to skip)
   ```
3. **列出可用的 COM Port 名稱**：用 `[System.IO.Ports.SerialPort]::GetPortNames()` 取得
4. **判斷虛擬序列埠**：如果第 1 步顯示 `No available serial ports.`，但第 3 步仍列出 COM Port，代表這些是虛擬序列埠
5. 按 Enter 結束

## 注意事項

- 移除裝置需要系統管理員權限
- 輸入 `Yes` 會嘗試移除**第 1 步列出的所有序列埠**，不是只移除其中一個。正在使用的實體序列埠也會被移除，可能造成裝置暫時無法使用。移除後可以到「裝置管理員」→「掃描硬體變更」，或重新插拔裝置來恢復
- 只支援 Windows（使用 WMI 和 .NET）
