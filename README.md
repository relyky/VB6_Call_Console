# Write to the Console Using VB6
* 兩種應用
 * 應用程式 - TRACE MODE
 * 丟訊息出來
* 此範例原碼改良自：Write to the Console Using VB6 http://www.freevbcode.com/ShowCode.asp?ID=4618
* 強化項目
 * 可正確顯示中文，原版顯示中文字串會截斷不完整。
 * 強化一些API，使更易用。
 * 調校程式碼，使更符合物件導向精神。
* 開發環境：
 * IDE: VB6
 * OS: Windows XP - in VMware9
* 其中核心程式碼有二個
 * Console.bas --- 導入Console 相關WIN32 SDK
 * console.cls --- 物件化包裝成 clsConsole.cls 類別以操作Console。
 * Console.frm --- 應用demo，非核心 
 
