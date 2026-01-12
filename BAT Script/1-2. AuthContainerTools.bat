@echo off
goto :EOF

:CleanPath
echo [DEBUG - CleanPath] CleanPath called with args: %*
setlocal EnableDelayedExpansion

set "VAR=%~2"
set "VALUE=!%VAR%!"

echo [DEBUG] Entering CleanPath with values VAR=[%VAR%], VALUE=[%VALUE%]
rem Salida si está realmente vacío
if "!VALUE!"=="" (
    echo [ERROR] El path a limpiar está vacío. Se recibieron los paramatros: VAR=[%VAR%], VALUE=[%VALUE%]
    echo .
    endlocal & exit /b 0
)

rem Quitar comillas exteriores

if "!VALUE:~0,1!"=="\"" if "!VALUE:~-1!"=="\"" (
    echo quitando comillas
    set "VALUE=!VALUE:~1,-1!"
)

rem Quitar espacios inicio
for /f "tokens=* delims= " %%A in ("!VALUE!") do set "VALUE=%%A"

rem Quitar espacios al final
:trim_end
if "!VALUE:~-1!"==" " (
    set "VALUE=!VALUE:~0,-1!"
    goto trim_end
)

rem Si empieza con "\" agregar C:
if "!VALUE:~0,1!"=="\" set "VALUE=C:!VALUE!"

rem Si termina con "\" quitarlo
if "!VALUE:~-1!"=="\" set "VALUE=!VALUE:~0,-1!"

endlocal & set "%VAR%=%VALUE%"
exit /b 0

:HandleRecentTemplateSubkey
rem Args: APP_NAME SUBKEY_NAME SUBKEY_PATH
set "__DAC_APP=%~1"
set "__DAC_ID=%~2"
set "__DAC_PATH=%~3"

if /I "%IsDesignModeEnabled%"=="true" (
    call :Log "%LogFilePath%" "[TRACE] Checking subkey leaf: %__DAC_ID% (%__DAC_PATH%)"
)

if not defined __DAC_ID exit /b 0


set "__DAC_PREFIX5=!__DAC_ID:~0,5!"
set "__DAC_PREFIX7=!__DAC_ID:~0,7!"
set "__DAC_IS_TARGET="
if /I "!__DAC_PREFIX5!"=="ADAL_" set "__DAC_IS_TARGET=1"
if not defined __DAC_IS_TARGET if /I "!__DAC_PREFIX7!"=="LIVEID_" set "__DAC_IS_TARGET=1"
if not defined __DAC_IS_TARGET exit /b 0

call :RegisterAuthContainer "!__DAC_APP!" "!__DAC_ID!" "!__DAC_PATH!"
exit /b 0

:RegisterAuthContainer
rem Args: APP_NAME CONTAINER_ID CONTAINER_PATH
set "__DAC_APP=%~1"
set "__DAC_ID=%~2"
set "__DAC_PATH=%~3"
if not defined __DAC_APP exit /b 0
if not defined __DAC_ID exit /b 0
if not defined __DAC_PATH exit /b 0

set "__DAC_MATCH_FOUND="
if defined __DAC_COUNT if !__DAC_COUNT! GTR 0 (
    set /a __DAC_LAST=!__DAC_COUNT!-1
    for /L %%I in (0,1,!__DAC_LAST!) do (
        if /I "!__DAC_ID[%%I]!"=="%__DAC_ID%" if /I "!__DAC_PATH[%%I]!"=="%__DAC_PATH%" set "__DAC_MATCH_FOUND=1"
    )
)

if defined __DAC_MATCH_FOUND exit /b 0

set "__DAC_APP[!__DAC_COUNT!]=%__DAC_APP%"
set "__DAC_ID[!__DAC_COUNT!]=%__DAC_ID%"
set "__DAC_PATH[!__DAC_COUNT!]=%__DAC_PATH%"

if not defined __DAC_PRIMARY_PATH (
    set "__DAC_PRIMARY_APP=%__DAC_APP%"
    set "__DAC_PRIMARY_ID=%__DAC_ID%"
    set "__DAC_PRIMARY_PATH=%__DAC_PATH%"
)

set /a __DAC_COUNT+=1
exit /b 0

:CollectAuthContainerPaths
rem Args: OUT_VAR APP_NAME
set "__CAP_TARGET_VAR=%~1"
set "__CAP_APP_FILTER=%~2"
if "%__CAP_TARGET_VAR%"=="" exit /b 0

setlocal EnableDelayedExpansion

set "__CAP_RESULT="

call :BuildAuthContainerCache "!__CAP_APP_FILTER!"

if defined __DAC_COUNT if !__DAC_COUNT! GTR 0 (
    set /a __DAC_LAST=!__DAC_COUNT!-1
    for /L %%I in (0,1,!__DAC_LAST!) do (
        set "__CAP_ENTRY_APP=!__DAC_APP[%%I]!"
        set "__CAP_ENTRY_PATH=!__DAC_PATH[%%I]!"
        if defined __CAP_ENTRY_PATH (
            if not defined __CAP_APP_FILTER (
                call :AppendUniquePath __CAP_RESULT "!__CAP_ENTRY_PATH!"
            ) else if /I "!__CAP_ENTRY_APP!"=="!__CAP_APP_FILTER!" (
                call :AppendUniquePath __CAP_RESULT "!__CAP_ENTRY_PATH!"
            )
        )
    )
)

set "__CAP_OUTPUT=!__CAP_RESULT!"

for %%# in (1) do (
    endlocal
    if not "%__CAP_TARGET_VAR%"=="" set "%__CAP_TARGET_VAR%=%__CAP_OUTPUT%"
)

exit /b 0

:AppendUniquePath
rem Args: VAR_NAME NEW_PATH
set "VAR_NAME=%~1"
set "NEW_PATH=%~2"
if "%VAR_NAME%"=="" exit /b 0
if "%NEW_PATH%"=="" exit /b 0
setlocal EnableDelayedExpansion
set "CURRENT=!%VAR_NAME%!"
set "NEED=1"
if defined CURRENT (
    for %%P in (!CURRENT!) do (
        if /I "%%~P"=="%~2" set "NEED=0"
    )
)
if "!NEED!"=="1" (
    if defined CURRENT (
        set "CURRENT=!CURRENT! ""%~2"""
    ) else (
        set "CURRENT=""%~2"""
    )
)
set "UPDATED=!CURRENT!"
for %%# in (1) do (
    endlocal
    set "%VAR_NAME%=%UPDATED%"
)
exit /b 0



