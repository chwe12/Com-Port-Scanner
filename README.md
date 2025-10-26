# Serial Port 管理腳本（PowerShell）

## 簡介
這是一個用於 **列出並管理 Windows 系統中所有序列埠（COM Port）** 的 PowerShell 腳本
除了能顯示所有可用的序列埠外，可選擇是否移除被佔用或殘留的裝置（例如未正確移除的虛擬 COM Port）

---

## 功能說明
- 查詢所有序列埠（使用 `Win32_SerialPort` WMI）
- 顯示每個埠的名稱與描述（DeviceID、Caption）
- 可選擇是否移除已佔用或殘留的裝置
- 顯示目前可用的 COM Port 名稱（使用 `.NET SerialPort`）
- 自動判斷是否存在虛擬序列埠

---

## 腳本流程

1. **查詢序列埠**
   - 使用 `Get-WmiObject -Query "SELECT DeviceID, Caption FROM Win32_SerialPort"`  
     取得所有偵測到的序列埠資訊

2. **顯示序列埠清單**
   - 逐一列出所有埠的名稱與描述
   - 若未找到任何埠，顯示提示訊息

3. **選擇是否移除佔用裝置**
   - 提示使用者輸入：
     ```
     Do you want to remove occupied devices? (Type 'Yes' to confirm, any other key to skip)
     ```
   - 若輸入 `Yes`，則嘗試使用 WMI (`Win32_PnPEntity`) 來移除該裝置

4. **列出可用 COM Port**
   - 透過 `[System.IO.Ports.SerialPort]::GetPortNames()`  
     顯示系統中可用的 COM Port 名稱

5. **辨識虛擬 COM Port**
   - 若 WMI 查無裝置，但仍有可用的 COM Port，則表示這些為虛擬埠

---

## 執行方式

1. 將腳本儲存為 `ManageSerialPorts.ps1`
2. 以 **系統管理員身分** 開啟 PowerShell
3. 切換至腳本所在資料夾：
   ```powershell
   cd "C:\Path\To\Script"
4. 執行腳本
   ```
   .\ManageSerialPorts.ps1
   ```
---

## 注意事項

1. 請使用 系統管理員權限 執行腳本，否則無法移除裝置
2. 不建議移除正在使用中的實體埠，以免造成驅動程式或裝置異常
3. 若出現執行權限錯誤，可暫時允許 PowerShell 執行腳本
   ```
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```

