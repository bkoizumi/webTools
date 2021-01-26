Attribute VB_Name = "LinkChecker"

'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function xxxxxxxxxx()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
'  On Error GoTo catchError

  targetURL = "https://www.yahoo.co.jp/"
  Call init.setting
  Call Chrome起動
  
  driver.Get targetURL
  driver.n


  Exit Function
'エラー発生時====================================
catchError:
  
  
  
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'*******************************************************************************************************
' * Chrome起動
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function Chrome起動()
  
  Call init.setting
  
  'chrome 拡張機能読み込み
  Call ProgressBar.showCount("WebCapture", 0, 10, "Chrome起動")
  With driver
    .AddArgument ("--lang=ja")
'    .AddArgument ("--user-data-dir=" & BrowserProfiles("default"))
    .AddArgument ("--window-size=1200,600")
    .AddArgument ("--hide-scrollbars")
    .AddArgument ("--disable-gpu")
    
    
    .start "chrome"
    .Wait 1000
  End With
  
 
End Function




