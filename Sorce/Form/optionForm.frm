VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} optionForm 
   Caption         =   "オプション"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11310
   OleObjectBlob   =   "optionForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "optionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'初期設定------------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()

  Call init.setting
  '初期表示ページ
  'マルチページの表示
  Me.MultiPage1.Value = 0
  Me.MultiPage2.Value = 0
  Me.MultiPage3.Value = 0


  'WebCapture--------------------------------------------------------------------------------------
  Me.UserName.Value = sheetSetting.Range("UserName")
  Me.JobName.Value = sheetSetting.Range("JobName")
  Me.sheetName.Value = sheetSetting.Range("sheetName")
  
  'ログイン情報1-----------------------------------------------------------------------------------
  Me.Login1IDVal.Value = sheetSetting.Range("Login1IDVal")
  Me.Login1IDTagName.Value = sheetSetting.Range("Login1IDTagName")
  Me.Login1IDTagID.Value = sheetSetting.Range("Login1IDTagID")
  Me.Login1IDTagClass.Value = sheetSetting.Range("Login1IDTagClass")
  Me.Login1PWVal.Value = sheetSetting.Range("Login1PWVal")
  Me.Login1PWTagName.Value = sheetSetting.Range("Login1PWTagName")
  Me.Login1PWTagID.Value = sheetSetting.Range("Login1PWTagID")
  Me.Login1PWTagClass.Value = sheetSetting.Range("Login1PWTagClass")
  Me.Login1Btn1Val.Value = sheetSetting.Range("Login1Btn1Val")
  Me.Login1Btn1TagName.Value = sheetSetting.Range("Login1Btn1TagName")
  Me.Login1Btn1TagID.Value = sheetSetting.Range("Login1Btn1TagID")
  Me.Login1Btn1TagClass.Value = sheetSetting.Range("Login1Btn1TagClass")
  Me.Login1Btn2Val.Value = sheetSetting.Range("Login1Btn2Val")
  Me.Login1Btn2TagName.Value = sheetSetting.Range("Login1Btn2TagName")
  Me.Login1Btn2TagID.Value = sheetSetting.Range("Login1Btn2TagID")
  Me.Login1Btn2TagClass.Value = sheetSetting.Range("Login1Btn2TagClass")
  
  'ログイン情報2-----------------------------------------------------------------------------------
  Me.Login2IDVal.Value = sheetSetting.Range("Login2IDVal")
  Me.Login2IDTagName.Value = sheetSetting.Range("Login2IDTagName")
  Me.Login2IDTagID.Value = sheetSetting.Range("Login2IDTagID")
  Me.Login2IDTagClass.Value = sheetSetting.Range("Login2IDTagClass")
  Me.Login2PWVal.Value = sheetSetting.Range("Login2PWVal")
  Me.Login2PWTagName.Value = sheetSetting.Range("Login2PWTagName")
  Me.Login2PWTagID.Value = sheetSetting.Range("Login2PWTagID")
  Me.Login2PWTagClass.Value = sheetSetting.Range("Login2PWTagClass")
  Me.Login2Btn1Val.Value = sheetSetting.Range("Login2Btn1Val")
  Me.Login2Btn1TagName.Value = sheetSetting.Range("Login2Btn1TagName")
  Me.Login2Btn1TagID.Value = sheetSetting.Range("Login2Btn1TagID")
  Me.Login2Btn1TagClass.Value = sheetSetting.Range("Login2Btn1TagClass")
  Me.Login2Btn2Val.Value = sheetSetting.Range("Login2Btn2Val")
  Me.Login2Btn2TagName.Value = sheetSetting.Range("Login2Btn2TagName")
  Me.Login2Btn2TagID.Value = sheetSetting.Range("Login2Btn2TagID")
  Me.Login2Btn2TagClass.Value = sheetSetting.Range("Login2Btn2TagClass")
  
  'ログイン情報3-----------------------------------------------------------------------------------
  Me.Login3IDVal.Value = sheetSetting.Range("Login3IDVal")
  Me.Login3IDTagName.Value = sheetSetting.Range("Login3IDTagName")
  Me.Login3IDTagID.Value = sheetSetting.Range("Login3IDTagID")
  Me.Login3IDTagClass.Value = sheetSetting.Range("Login3IDTagClass")
  Me.Login3PWVal.Value = sheetSetting.Range("Login3PWVal")
  Me.Login3PWTagName.Value = sheetSetting.Range("Login3PWTagName")
  Me.Login3PWTagID.Value = sheetSetting.Range("Login3PWTagID")
  Me.Login3PWTagClass.Value = sheetSetting.Range("Login3PWTagClass")
  Me.Login3Btn1Val.Value = sheetSetting.Range("Login3Btn1Val")
  Me.Login3Btn1TagName.Value = sheetSetting.Range("Login3Btn1TagName")
  Me.Login3Btn1TagID.Value = sheetSetting.Range("Login3Btn1TagID")
  Me.Login3Btn1TagClass.Value = sheetSetting.Range("Login3Btn1TagClass")
  Me.Login3Btn2Val.Value = sheetSetting.Range("Login3Btn2Val")
  Me.Login3Btn2TagName.Value = sheetSetting.Range("Login3Btn2TagName")
  Me.Login3Btn2TagID.Value = sheetSetting.Range("Login3Btn2TagID")
  Me.Login3Btn2TagClass.Value = sheetSetting.Range("Login3Btn2TagClass")
  
  '検索情報1-----------------------------------------------------------------------------------
  Me.search1Val.Value = sheetSetting.Range("search1Val")
  Me.search1TagName.Value = sheetSetting.Range("search1TagName")
  Me.search1TagID.Value = sheetSetting.Range("search1TagID")
  Me.search1TagClass.Value = sheetSetting.Range("search1TagClass")
  Me.search1BtnVal.Value = sheetSetting.Range("search1BtnVal")
  Me.search1BtnTagName.Value = sheetSetting.Range("search1BtnTagName")
  Me.search1BtnTagID.Value = sheetSetting.Range("search1BtnTagID")
  Me.search1BtnTagClass.Value = sheetSetting.Range("search1BtnTagClass")
  
  '検索情報2-----------------------------------------------------------------------------------
  Me.search2Val.Value = sheetSetting.Range("search2Val")
  Me.search2TagName.Value = sheetSetting.Range("search2TagName")
  Me.search2TagID.Value = sheetSetting.Range("search2TagID")
  Me.search2TagClass.Value = sheetSetting.Range("search2TagClass")
  Me.search2BtnVal.Value = sheetSetting.Range("search2BtnVal")
  Me.search2BtnTagName.Value = sheetSetting.Range("search2BtnTagName")
  Me.search2BtnTagID.Value = sheetSetting.Range("search2BtnTagID")
  Me.search2BtnTagClass.Value = sheetSetting.Range("search2BtnTagClass")
  
  '検索情報3-----------------------------------------------------------------------------------
  Me.search3Val.Value = sheetSetting.Range("search3Val")
  Me.search3TagName.Value = sheetSetting.Range("search3TagName")
  Me.search3TagID.Value = sheetSetting.Range("search3TagID")
  Me.search3TagClass.Value = sheetSetting.Range("search3TagClass")
  Me.search3BtnVal.Value = sheetSetting.Range("search3BtnVal")
  Me.search3BtnTagName.Value = sheetSetting.Range("search3BtnTagName")
  Me.search3BtnTagID.Value = sheetSetting.Range("search3BtnTagID")
  Me.search3BtnTagClass.Value = sheetSetting.Range("search3BtnTagClass")


  'Sitemap
  Me.siteMapURL.Value = sheetSetting.Range("siteMapURL")
  
End Sub


'キャンセルボタン押下------------------------------------------------------------------------------
Private Sub BtnCancel_Click()
  Unload Me
End Sub

'OKボタン押下--------------------------------------------------------------------------------------
Private Sub BtnSubmit_Click()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

  'WebCapture--------------------------------------------------------------------------------------
  sheetSetting.Range("UserName") = Me.UserName.Value
  sheetSetting.Range("JobName") = Me.JobName.Value
  sheetSetting.Range("sheetName") = Me.sheetName.Value
  
  sheetSetting.Range("Login1IDVal") = Me.Login1IDVal.Value
  sheetSetting.Range("Login1IDVal") = Me.Login1IDVal.Value
  sheetSetting.Range("Login1IDTagName") = Me.Login1IDTagName.Value
  sheetSetting.Range("Login1IDTagID") = Me.Login1IDTagID.Value
  sheetSetting.Range("Login1IDTagClass") = Me.Login1IDTagClass.Value
  sheetSetting.Range("Login1PWVal") = Me.Login1PWVal.Value
  sheetSetting.Range("Login1PWTagName") = Me.Login1PWTagName.Value
  sheetSetting.Range("Login1PWTagID") = Me.Login1PWTagID.Value
  sheetSetting.Range("Login1PWTagClass") = Me.Login1PWTagClass.Value
  sheetSetting.Range("Login1Btn1Val") = Me.Login1Btn1Val.Value
  sheetSetting.Range("Login1Btn1TagName") = Me.Login1Btn1TagName.Value
  sheetSetting.Range("Login1Btn1TagID") = Me.Login1Btn1TagID.Value
  sheetSetting.Range("Login1Btn1TagClass") = Me.Login1Btn1TagClass.Value
  sheetSetting.Range("Login1Btn2Val") = Me.Login1Btn2Val.Value
  sheetSetting.Range("Login1Btn2TagName") = Me.Login1Btn2TagName.Value
  sheetSetting.Range("Login1Btn2TagID") = Me.Login1Btn2TagID.Value
  sheetSetting.Range("Login1Btn2TagClass") = Me.Login1Btn2TagClass.Value
  sheetSetting.Range("Login2IDVal") = Me.Login2IDVal.Value
  sheetSetting.Range("Login2IDTagName") = Me.Login2IDTagName.Value
  sheetSetting.Range("Login2IDTagID") = Me.Login2IDTagID.Value
  sheetSetting.Range("Login2IDTagClass") = Me.Login2IDTagClass.Value
  sheetSetting.Range("Login2PWVal") = Me.Login2PWVal.Value
  sheetSetting.Range("Login2PWTagName") = Me.Login2PWTagName.Value
  sheetSetting.Range("Login2PWTagID") = Me.Login2PWTagID.Value
  sheetSetting.Range("Login2PWTagClass") = Me.Login2PWTagClass.Value
  sheetSetting.Range("Login2Btn1Val") = Me.Login2Btn1Val.Value
  sheetSetting.Range("Login2Btn1TagName") = Me.Login2Btn1TagName.Value
  sheetSetting.Range("Login2Btn1TagID") = Me.Login2Btn1TagID.Value
  sheetSetting.Range("Login2Btn1TagClass") = Me.Login2Btn1TagClass.Value
  sheetSetting.Range("Login2Btn2Val") = Me.Login2Btn2Val.Value
  sheetSetting.Range("Login2Btn2TagName") = Me.Login2Btn2TagName.Value
  sheetSetting.Range("Login2Btn2TagID") = Me.Login2Btn2TagID.Value
  sheetSetting.Range("Login2Btn2TagClass") = Me.Login2Btn2TagClass.Value
  sheetSetting.Range("Login3IDVal") = Me.Login3IDVal.Value
  sheetSetting.Range("Login3IDTagName") = Me.Login3IDTagName.Value
  sheetSetting.Range("Login3IDTagID") = Me.Login3IDTagID.Value
  sheetSetting.Range("Login3IDTagClass") = Me.Login3IDTagClass.Value
  sheetSetting.Range("Login3PWVal") = Me.Login3PWVal.Value
  sheetSetting.Range("Login3PWTagName") = Me.Login3PWTagName.Value
  sheetSetting.Range("Login3PWTagID") = Me.Login3PWTagID.Value
  sheetSetting.Range("Login3PWTagClass") = Me.Login3PWTagClass.Value
  sheetSetting.Range("Login3Btn1Val") = Me.Login3Btn1Val.Value
  sheetSetting.Range("Login3Btn1TagName") = Me.Login3Btn1TagName.Value
  sheetSetting.Range("Login3Btn1TagID") = Me.Login3Btn1TagID.Value
  sheetSetting.Range("Login3Btn1TagClass") = Me.Login3Btn1TagClass.Value
  sheetSetting.Range("Login3Btn2Val") = Me.Login3Btn2Val.Value
  sheetSetting.Range("Login3Btn2TagName") = Me.Login3Btn2TagName.Value
  sheetSetting.Range("Login3Btn2TagID") = Me.Login3Btn2TagID.Value
  sheetSetting.Range("Login3Btn2TagClass") = Me.Login3Btn2TagClass.Value
  sheetSetting.Range("search1Val") = Me.search1Val.Value
  sheetSetting.Range("search1TagName") = Me.search1TagName.Value
  sheetSetting.Range("search1TagID") = Me.search1TagID.Value
  sheetSetting.Range("search1TagClass") = Me.search1TagClass.Value
  sheetSetting.Range("search1BtnVal") = Me.search1BtnVal.Value
  sheetSetting.Range("search1BtnTagName") = Me.search1BtnTagName.Value
  sheetSetting.Range("search1BtnTagID") = Me.search1BtnTagID.Value
  sheetSetting.Range("search1BtnTagClass") = Me.search1BtnTagClass.Value
  sheetSetting.Range("search2Val") = Me.search2Val.Value
  sheetSetting.Range("search2TagName") = Me.search2TagName.Value
  sheetSetting.Range("search2TagID") = Me.search2TagID.Value
  sheetSetting.Range("search2TagClass") = Me.search2TagClass.Value
  sheetSetting.Range("search2BtnVal") = Me.search2BtnVal.Value
  sheetSetting.Range("search2BtnTagName") = Me.search2BtnTagName.Value
  sheetSetting.Range("search2BtnTagID") = Me.search2BtnTagID.Value
  sheetSetting.Range("search2BtnTagClass") = Me.search2BtnTagClass.Value
  sheetSetting.Range("search3Val") = Me.search3Val.Value
  sheetSetting.Range("search3TagName") = Me.search3TagName.Value
  sheetSetting.Range("search3TagID") = Me.search3TagID.Value
  sheetSetting.Range("search3TagClass") = Me.search3TagClass.Value
  sheetSetting.Range("search3BtnVal") = Me.search3BtnVal.Value
  sheetSetting.Range("search3BtnTagName") = Me.search3BtnTagName.Value
  sheetSetting.Range("search3BtnTagID") = Me.search3BtnTagID.Value
  sheetSetting.Range("search3BtnTagClass") = Me.search3BtnTagClass.Value



  'Sitemap
  If Me.siteMapURL.Value Like "*/" Then
    sheetSetting.Range("siteMapURL") = Library.cutRight(Me.siteMapURL.Value, 1)
  Else
    sheetSetting.Range("siteMapURL") = Me.siteMapURL.Value
  End If
  Unload Me
End Sub

