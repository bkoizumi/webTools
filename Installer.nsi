﻿; ユーザー変数宣言
Var tciNetwork

; インストーラーの識別子
!define PRODUCT_NAME "WebTools"
; インストーラーのバージョン。
!define PRODUCT_VERSION "3.1.0.0"
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
!define MUI_HEADERIMAGE_BITMAP "${NSISDIR}\Contrib\Graphics\Header\orange-r.bmp"
!define MUI_HEADERIMAGE_UNBITMAP "${NSISDIR}\Contrib\Graphics\Header\orange-uninstall-r.bmp"

!define MUI_WELCOMEFINISHPAGE_BITMAP "${NSISDIR}\Contrib\Graphics\Wizard\orange.bmp"
!define MUI_UNWELCOMEFINISHPAGE_BITMAP "${NSISDIR}\Contrib\Graphics\Wizard\orange-uninstall.bmp"


; 使用する外部ライブラリ
!include Sections.nsh
!include MUI2.nsh
!include LogicLib.nsh
!include nsProcess.nsh


; 圧縮設定。通常は/solid lzmaが最も圧縮率が高い
SetCompressor /solid lzma

; インストーラー名
Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
; 出力されるファイル名
OutFile "${PRODUCT_NAME}_${PRODUCT_VERSION}.exe"

; インストール/アンインストール時の進捗画面
ShowInstDetails show
ShowUnInstDetails show


; インストーラーフィアルのバージョン情報記述
VIProductVersion ${PRODUCT_VERSION}
VIAddVersionKey ProductName "${PRODUCT_NAME}"
VIAddVersionKey ProductVersion "${PRODUCT_VERSION}"
VIAddVersionKey Comments "WebTool for Excel"
VIAddVersionKey LegalTrademarks ""
VIAddVersionKey LegalCopyright "Copyright 2020 Bumpei.Koizumi"
VIAddVersionKey FileDescription ""
VIAddVersionKey FileVersion "${PRODUCT_VERSION}"

; デフォルトのファイルのインストール先
InstallDir "C:\ExcelMethod"
; InstallDir "$appData\ExcelMethod"

;実行権限（user/admin）
RequestExecutionLevel admin

;インストール画面構成
!define MUI_LICENSEPAGE_RADIOBUTTONS      ; 「ライセンスに同意する」をラジオボタンにする
!define MUI_FINISHPAGE_NOAUTOCLOSE        ; インストール完了後自動的に完了画面に遷移しないようにする

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "C:\WorkSpace\VBA\webTools\LICENSE.txt"
; !insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

# アンインストーラ ページ
UninstPage uninstConfirm
UninstPage instfiles

!insertmacro MUI_LANGUAGE "Japanese"

; インストール処理---------------------------------------------------------------------------------------
Section "WebTools本体" sec_Main

  SetOutPath $INSTDIR
  ; ディレクトリ/ファイルをコピー
  File    "${APPDIR}\WebTools.xlsm"
  File /r "${APPDIR}\Downloads"
  File /r "${APPDIR}\var"
  File /r "${APPDIR}\logs"

  SetOutPath $INSTDIR\bin
  File /r "${APPDIR}\bin\7zip"
  File /r "${APPDIR}\bin\SeleniumBasic"
  File    "${APPDIR}\bin\aria2c.exe"
  File    "${APPDIR}\bin\nkf.exe"


  AccessControl::GrantOnFile "$INSTDIR\bin\SeleniumBasic" "(S-1-1-0)" "FullAccess"
  AccessControl::GrantOnFile "$INSTDIR\logs" "(S-1-5-32-545)" "FullAccess"

  SetShellVarContext all
  CreateShortCut "$DESKTOP\WebTools.lnk" "$INSTDIR\WebTools.xlsm"
  WriteRegStr HKCU "Software\VB and VBA Program Settings\B.Koizumi\${PRODUCT_NAME}" "InstDir" $INSTDIR
  WriteRegStr HKCU "Software\VB and VBA Program Settings\B.Koizumi\${PRODUCT_NAME}" "InstVersion" ${PRODUCT_VERSION}

  ; SeleniumBasicをExcel参照設定に追加
  ExecWait  '"%SystemRoot%\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe"  /tlb /codebase  "$INSTDIR\bin\SeleniumBasic\Selenium.dll"'
  Pop $0
  DetailPrint "SeleniumBasicをExcel参照設定に追加: $0"

  # アンインストーラを出力
  WriteUninstaller "$INSTDIR\Uninstall.exe"


  ;スタートメニューの作成
  CreateDirectory "$SMPROGRAMS\${PRODUCT_NAME}"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\WebTools.lnk"           "$INSTDIR\WebTools.xlsm"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\Chrome起動.lnk"         "$INSTDIR\var\BrowserProfile\Chrome起動_default.bat"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\アンインストール.lnk"   "$INSTDIR\Uninstall.exe"
SectionEnd

SectionGroup /e "ネットワーク" Network
    Section "TCIネットワーク" TCINW
      WriteRegStr HKCU "Software\VB and VBA Program Settings\B.Koizumi\${PRODUCT_NAME}" "InstNetwork" "tci"
    SectionEnd

    Section /o "その他ネットワーク" otherNW
      WriteRegStr HKCU "Software\VB and VBA Program Settings\B.Koizumi\${PRODUCT_NAME}" "InstNetwork" "other"
    SectionEnd
SectionGroupEnd

Section "Uninstall"
  ; Seleniumの登録解除
  ExecWait '"%SystemRoot%\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe $INSTDIR\bin\SeleniumBasic\Selenium.dll /u"' $0

  ; ディレクトリ削除
  RMDir /r "$INSTDIR"
  RMDir /r "$SMPROGRAMS\${PRODUCT_NAME}"

  ; レジストリキー削除
  DeleteRegKey HKCU "Software\VB and VBA Program Settings\B.Koizumi\${PRODUCT_NAME}"
SectionEnd

; セクションの説明文を入力
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
    !insertmacro MUI_DESCRIPTION_TEXT ${sec_Main}       "Web用ツール"
    !insertmacro MUI_DESCRIPTION_TEXT ${Network}        "ネットワークを選択してください"
    !insertmacro MUI_DESCRIPTION_TEXT ${TCINW}          "TCIネットワークに接続"
    !insertmacro MUI_DESCRIPTION_TEXT ${otherNW}        "TCIネットワーク以外に接続"
!insertmacro MUI_FUNCTION_DESCRIPTION_END



Function .onInit
  call  BootingCheck
  call  isInstalled
  call  IsDotNetFramework


  StrCpy $tciNetwork ${TCINW}

FunctionEnd

; セクションの選択が変わったときの処理
Function .onSelChange
    !insertmacro StartRadioButtons $tciNetwork
        !insertmacro RadioButton ${TCINW}
        !insertmacro RadioButton ${otherNW}
    !insertmacro EndRadioButtons

FunctionEnd


; Excelの起動確認---------------------------------------------------------------------------------------
Function BootingCheck

; reCheck:
;   StrCpy $1 "EXCEL.EXE"
;   nsProcess::_FindProcess "$1"
;   Pop $R0
;   ${If} $R0 = 0
;     MessageBox MB_OK "Excel :[$R0]"
;     ; nsProcess::_KillProcess "$1"
;     Pop $R0
;     Sleep 500
;     Goto reCheck
;   ${EndIf}
FunctionEnd


; インストール済みかどうか------------------------------------------------------------------------------
Function isInstalled
  ; Var /GLOBAL majorVer
  ; Var /GLOBAL minorVer
  ; Var /GLOBAL revisionVer
  ; Var /GLOBAL buildVer
  ; Var /GLOBAL instDir

  ReadRegStr $0 HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}" "InstVersion"
  ReadRegStr $1 HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}" "InstDir"

  ${If} $0 == ${PRODUCT_VERSION}
    MessageBox MB_OK "同一バージョンがインストールされています"
    Abort

  ; ${Else}
  ;   SetOutPath $1
  ;   File "${APPDIR}\Koetol.xlsm"
  ;   WriteRegStr HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}" "Version" ${PRODUCT_VERSION}
  ;   MessageBox MB_OK "既にバージョン $0 がインストールされているため、Koetol本体のみ更新しました"
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
    ExecWait  '"fondue.exe" /enable-feature:netfx3'
    Pop $0
    MessageBox MB_OK ".NET Framework インストール => $0"
    Abort

  yesDotNet:
    ; OK の場合

FunctionEnd
