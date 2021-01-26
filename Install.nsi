; ユーザー変数宣言
Var setNetwork

; インストーラーの識別子
!define PRODUCT_NAME "WebTools"
; インストーラーのバージョン。
!define PRODUCT_VERSION "0.0.4.1"
!define APPDIR "ExcelMethod"

; 多言語で使用する場合はここをUnicodeにすることを推奨
Unicode true

; インストーラーのアイコン
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\orange-install.ico"

; アンインストーラーのアイコン
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\orange-uninstall.ico"

; インストーラの見た目
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_RIGHT
!define MUI_HEADERIMAGE_BITMAP          "${NSISDIR}\Contrib\Graphics\Header\orange-r.bmp"
!define MUI_HEADERIMAGE_UNBITMAP        "${NSISDIR}\Contrib\Graphics\Header\orange-uninstall-r.bmp"

!define MUI_WELCOMEFINISHPAGE_BITMAP    "${NSISDIR}\Contrib\Graphics\Wizard\orange.bmp"
!define MUI_UNWELCOMEFINISHPAGE_BITMAP  "${NSISDIR}\Contrib\Graphics\Wizard\orange-uninstall.bmp"


; 使用する外部ライブラリ
!include Sections.nsh
!include MUI2.nsh
!include LogicLib.nsh
; !include nsProcess.nsh


; 圧縮設定。通常は/solid lzmaが最も圧縮率が高い
SetCompressor /solid lzma

; インストーラー名
Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
; 出力されるファイル名
OutFile "${PRODUCT_NAME}-${PRODUCT_VERSION}-setup.exe"

; インストール/アンインストール時の進捗画面
ShowInstDetails   show
ShowUnInstDetails show


; インストーラーフィアルのバージョン情報記述
VIProductVersion ${PRODUCT_VERSION}
VIAddVersionKey ProductName     "${PRODUCT_NAME}"
VIAddVersionKey ProductVersion  "${PRODUCT_VERSION}"
VIAddVersionKey Comments        "WebTools for Excel"
VIAddVersionKey LegalTrademarks ""
VIAddVersionKey LegalCopyright  "Copyright 2020 Bumpei.Koizumi"
VIAddVersionKey FileDescription "${PRODUCT_NAME}"
VIAddVersionKey FileVersion     "${PRODUCT_VERSION}"

; デフォルトのファイルのインストール先
InstallDir "C:\ExcelMethod"
; InstallDir "$appData\ExcelMethod"

;実行権限 [user/admin]
RequestExecutionLevel admin

;インストール画面構成
; !define MUI_LICENSEPAGE_RADIOBUTTONS      ; 「ライセンスに同意する」をラジオボタンにする
; !define MUI_FINISHPAGE_NOAUTOCLOSE        ; インストール完了後自動的に完了画面に遷移しないようにする

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "C:\WorkSpace\VBA\webTools\LICENSE.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
# アンインストール画面構成
UninstPage uninstConfirm
UninstPage instfiles

!insertmacro MUI_LANGUAGE "Japanese"

; インストール処理---------------------------------------------------------------------------------------
Section "WebTools本体" sec_Main
  SetOutPath $INSTDIR

  ; ディレクトリ/ファイルをコピー
  File    "C:\WorkSpace\VBA\webTools\WebTools.xlsm"
  File /r "${APPDIR}\Downloads"
  File /r "${APPDIR}\work"
  File /r "${APPDIR}\logs"
  File /r "${APPDIR}\bin"
  File /r "${APPDIR}\var"

  SetShellVarContext current

  ; アクセス権の設定
  AccessControl::GrantOnFile "$INSTDIR\bin\SeleniumBasic" "(S-1-1-0)" "FullAccess"
  AccessControl::GrantOnFile "$INSTDIR\logs" "(S-1-5-32-545)" "FullAccess"
  AccessControl::GrantOnFile "$INSTDIR\var" "(S-1-5-32-545)" "FullAccess"

  ; SeleniumBasicをExcel参照設定に追加
  ExecWait 'C:\Windows\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe /tlb $INSTDIR\bin\SeleniumBasic\Selenium.dll /codebase'
  Pop $0
  DetailPrint "SeleniumBasicをExcel参照設定に追加: $0"

  ; レジストリキーの設定
  WriteRegStr HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}\Main" "InstDir"     $INSTDIR
  WriteRegStr HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}\Main" "InstVersion" ${PRODUCT_VERSION}

  ; アンインストーラを出力
  WriteUninstaller "$INSTDIR\Uninstall.exe"


  ;スタートメニューの作成
  CreateDirectory "$SMPROGRAMS\ExcelMethod"
  CreateShortCut  "$SMPROGRAMS\ExcelMethod\WebTools.lnk"           "$INSTDIR\WebTools.xlsm"
  CreateShortCut  "$SMPROGRAMS\ExcelMethod\アンインストール.lnk"   "$INSTDIR\Uninstall.exe"
SectionEnd

SectionGroup /e "ネットワーク" Network
    Section /o "Proxyあり" ProxyOn
      WriteRegStr HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}\Main" "InstNetwork" "ProxyOn"
    SectionEnd

    Section  "Proxyなし" ProxyOff
      WriteRegStr HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}\Main" "InstNetwork" "ProxyOff"
    SectionEnd
SectionGroupEnd

Section "Uninstall"
  SetShellVarContext all

  ; Seleniumの登録解除
  ExecWait 'C:\Windows\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe $INSTDIR\bin\SeleniumBasic\Selenium.dll /u'
  Pop $0
  DetailPrint "SeleniumBasicをExcel参照設定を削除: $0"

  ;スタートメニューから削除
  Delete "$SMPROGRAMS\ExcelMethod\WebTools.lnk"
  Delete "$SMPROGRAMS\ExcelMethod\アンインストール.lnk"

  RMDir /r  "$SMPROGRAMS\ExcelMethod"

  ; ディレクトリ削除
  RMDir /r "$INSTDIR"

  ; レジストリキー削除
  DeleteRegKey HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}"
SectionEnd

; セクションの説明文を入力
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
    !insertmacro MUI_DESCRIPTION_TEXT ${sec_Main}       "Webサイトの画面キャプチャ取得"
    !insertmacro MUI_DESCRIPTION_TEXT ${Network}        "ネットワークを選択してください"
    !insertmacro MUI_DESCRIPTION_TEXT ${ProxyOn}        "Proxyが必要なネットワークに接続"
    !insertmacro MUI_DESCRIPTION_TEXT ${ProxyOff}       "Proxyが不要なネットワークに接続"
!insertmacro MUI_FUNCTION_DESCRIPTION_END



Function .onInit
  call  BootingCheck
  call  isInstalled
  call  IsDotNetFramework

  StrCpy $setNetwork ${ProxyOn}

FunctionEnd

; セクションの選択が変わったときの処理
Function .onSelChange
    !insertmacro StartRadioButtons $setNetwork
        !insertmacro RadioButton ${ProxyOn}
        !insertmacro RadioButton ${ProxyOff}
    !insertmacro EndRadioButtons
FunctionEnd


; Excelの起動確認---------------------------------------------------------------------------------------
Function BootingCheck

; reCheck:
  ; ${nsProcess::FindProcess} "EXCEL.EXE" $R0
  ; MessageBox MB_OK "[$R0]"
  ; ${If} $R0 = 0
  ; ${Else}
  ;   MessageBox MB_OK "Excel が起動しています"
  ;   ; nsProcess::_KillProcess "$1"
  ;   Pop $R0
  ;   Sleep 500
  ;   Goto reCheck
  ; ${EndIf}
FunctionEnd


; インストール済みかどうか------------------------------------------------------------------------------
Function isInstalled
  ReadRegStr $0 HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}\Main" "InstVersion"
  ReadRegStr $1 HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}\Main" "InstDir"

  ${If} $0 == ${PRODUCT_VERSION}
    MessageBox MB_OK "同一バージョンがインストールされています"
    Abort

  ; ${Else}
  ;   SetOutPath $1
  ;   File "${APPDIR}\WebTools.xlsm"
  ;   File "${APPDIR}\var\WebCapture\新規Book.xlsm"
  ;   WriteRegStr HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}\Main" "Version" ${PRODUCT_VERSION}
  ;   MessageBox MB_OK "既にバージョン $0 がインストールされているため、Excelのみ更新しました"
  ;   Abort
  ${EndIf}

FunctionEnd

; .NET Frameworkバージョンチェック------------------------------------------------------------------------------
Function IsDotNetFramework
  ReadRegStr $1 HKLM "SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727" "Install"
  StrCmp $1 "" noDotNet yesDotNet1

  yesDotNet1:
    ReadRegStr $1 HKLM "SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727" "Version"
    StrCpy $2 $1 3
    StrCmp $2 "2.0" yesDotNet noDotNet
  noDotNet:
    MessageBox MB_OK ".NET Framework 2.0 がインストールされていません。"
    Exec  '"fondue.exe" /enable-feature:netfx3'
    Abort

  yesDotNet:
    ; OK の場合
    ; MessageBox MB_OK ".NET Framework 2.0 がインストールされています。"
    ; Exec  '"fondue.exe" /enable-feature:netfx3'
    ; Abort

FunctionEnd
