@ECHO OFF

REM ********************************************************************************
REM * @Author      : B.Koizumi
REM * @Create Date : 2019/01/21
REM * @Description : updateChromeDriver
REM ********************************************************************************

REM �ݒ�t�@�C���̓ǂݍ���----------------------------------------------------------
SET ThisBatFileDir=%~dp0
REM --------------------------------------------------------------------------------

cd %ThisBatFileDir%
powershell -ExecutionPolicy RemoteSigned -File %ThisBatFileDir%\updateChromeDriver.ps1
