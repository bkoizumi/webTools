Attribute VB_Name = "WebToolLib"

'*******************************************************************************************************
' * Chrome�N��
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function Chrome�N��()
  
  Call init.setting
  
  Call ProgressBar.showCount("WebCapture", 0, 10, "Chrome�N��")
  With driver
    .AddArgument ("--lang=ja")
'    .AddArgument ("--user-data-dir=" & BrowserProfiles("default"))
    .AddArgument ("--window-size=1200,600")
    .AddArgument ("--hide-scrollbars")
    .AddArgument ("--disable-gpu")
    
    If setVal("debugMode") = "develop" Then
'      .AddArgument ("--incognito")
    Else
      .AddArgument ("--headless")
    End If
    
    If setVal("InstNetwork") = "TCI" Then
      .AddArgument ("--proxy-server=" & setVal("ProxyURL") & ":" & setVal("ProxyPort"))
    End If
    
    .start "chrome"
    .Wait 1000
    .Timeouts.PageLoad = 60000
  End With
  
  Call ProgressBar.showCount("WebCapture", 10, 10, "Chrome�N��")
  Call Library.showDebugForm("WebCapture", "Chrome�N������")
  
End Function


'*******************************************************************************************************
' * Proxy�ݒ�擾
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function getProxy()



End Function
