@ECHO OFF

REM ********************************************************************************
REM * @Author      : B.Koizumi
REM * @Create Date : 2019/01/21
REM * @Description : updateChromeDriver
REM ********************************************************************************

REM 設定ファイルの読み込み----------------------------------------------------------
SET ThisBatFileDir=%~dp0
title update Chrome Driver
REM --------------------------------------------------------------------------------

cd %ThisBatFileDir%
powershell -ExecutionPolicy RemoteSigned -File %ThisBatFileDir%\updateChromeDriver.ps1
