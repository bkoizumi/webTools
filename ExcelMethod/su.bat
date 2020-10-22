@echo off

powershell start-process cmd -ArgumentList '/k ""cd /d %CD%""' -verb runas
