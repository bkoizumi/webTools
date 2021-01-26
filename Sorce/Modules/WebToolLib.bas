Attribute VB_Name = "WebToolLib"

'*******************************************************************************************************
' * optionForm表示
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function showOptionForm()
  
  'ダイアログへ表示
  With optionForm
    .StartUpPosition = 0
    .Top = Application.Top + 120
    .Left = Application.Left + 200
  End With

  optionForm.Show vbModeless

End Function


'*******************************************************************************************************
' * Chrome起動
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function Chrome起動()
  
  Call init.setting

  Call Library.chkShellEnd(Shell(setVal("appInstDir") & "\bin\SeleniumBasic\updateChromeDriver.bat"))
  
'  With openingDriver
'    .AddArgument ("--lang=ja")
'    .AddArgument ("--window-size=1200,600")
'    .AddArgument ("--hide-scrollbars")
'    .AddArgument ("--disable-gpu")
'    .AddArgument ("--app=file:///" & setVal("appInstDir") & "/var/WebCapture/opening/index.html")
''    .AddArgument ("--user-data-dir=" & BrowserProfiles("default"))
''    .AddArgument ("--disable-javascript")
''    .AddArgument ("--headless")
''    .AddArgument ("--kiosk")
'    .start "chrome"
'  End With
'
  With driver
    .AddArgument ("--lang=ja")
    .AddArgument ("--window-size=1200,600")
    .AddArgument ("--hide-scrollbars")
    .AddArgument ("--disable-gpu")
'    .AddArgument ("--headless")
    .AddArgument ("--enable-logging")
    .AddArgument ("--log-level=0")
    .AddArgument ("--dump-histograms-on-exit ")
    
    If setVal("InstNetwork") = "TCI" Then
      .AddArgument ("--proxy-server=" & setVal("ProxyURL") & ":" & setVal("ProxyPort"))
    End If
    
    .start "chrome"
    .Timeouts.PageLoad = 60000
  End With

End Function


'*******************************************************************************************************
' * Chrome終了
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function Chrome終了()
  
'  openingDriver.Close
'  openingDriver.Quit
'  Set openingDriver = Nothing
  
  driver.Close
  driver.Quit
  Set driver = Nothing
    
    
  ThisWorkbook.Save

End Function


'*******************************************************************************************************
' * Proxy設定取得
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function getProxy()



End Function




'*******************************************************************************************************
' * Test
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function Test()
  
  Call init.setting
  
  With driver
    .AddArgument ("--lang=ja")
    .AddArgument ("--window-size=1200,600")
    .AddArgument ("--hide-scrollbars")
    .AddArgument ("--disable-gpu")
    
    .start "chrome"
    .Timeouts.PageLoad = 60000
  End With

  driver.Get "https://news.yahoo.co.jp/"


Call Chrome終了
  

End Function
