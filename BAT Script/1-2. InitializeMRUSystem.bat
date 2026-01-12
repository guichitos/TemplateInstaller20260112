@echo off

set "MRUTools=%BaseDirectoryPath%1-2. MRU-PathResolver.bat"
echo Initializing MRU System...

rem ============================================================
rem ===             InitMruSystem (archivo externo)           ===
rem ============================================================
:InitMruSystem
echo [DEBUG - InitMruSystem] Inicializando sistema MRU...

rem ------------------------------------------------------------
rem WORD
rem ------------------------------------------------------------
call "%MRUTools%" :DetectMRUPath WORD ADAL
call "%MRUTools%" :DetectMRUPath WORD LIVEID

rem ------------------------------------------------------------
rem POWERPOINT
rem ------------------------------------------------------------
call "%MRUTools%" :DetectMRUPath POWERPOINT ADAL
call "%MRUTools%" :DetectMRUPath POWERPOINT LIVEID

rem ------------------------------------------------------------
rem EXCEL
rem ------------------------------------------------------------
call "%MRUTools%" :DetectMRUPath EXCEL ADAL
call "%MRUTools%" :DetectMRUPath EXCEL LIVEID

echo [DEBUG - InitMruSystem] Sistema MRU inicializado correctamente.
exit /b 0
