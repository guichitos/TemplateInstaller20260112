@echo off
setlocal EnableDelayedExpansion
rem =======================================================
rem == FLAG PARA MODO DISEÑO (VER SALIDA DE CONSOLA) ======
rem =======================================================
rem   true  = imprime todo
rem   false = consola completamente silenciosa
set "IsDesignModeEnabled=false"


:: --------------------------------------------------------
:: ENCABEZADO
:: --------------------------------------------------------
if /I "%IsDesignModeEnabled%"=="true" (
    echo ==========================================================
    echo   RESET OFFICE TEMPLATE MRU LISTS
    echo ==========================================================
    echo.
    echo Targeting template MRU keys for Word, PowerPoint, and Excel...
    echo.
)


set "OFFICE_APPS=Word PowerPoint Excel"

for %%A in (%OFFICE_APPS%) do (

    if /I "%IsDesignModeEnabled%"=="true" (
        echo ----------------------------------------------------------
        echo Borrando llaves de %%A...
    )
    set "MRU_TARGETS="
    call :GetAppMruTargets MRU_TARGETS "%%A"
    if not defined MRU_TARGETS (

        if /I "%IsDesignModeEnabled%"=="true" (
            echo     No se encontraron listas MRU para %%A.
        )
    ) else (
        for %%T in (!MRU_TARGETS!) do (
            set "CURRENT_TARGET=%%~T"
            if not "!CURRENT_TARGET!"=="" (
                call :ClearMruKey "!CURRENT_TARGET!" "%%A MRU list"
            )
        )

    )
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo ----------------------------------------------------------
    echo Finalizado.
    echo ----------------------------------------------------------
)
exit /b 0


:: ==========================================================
:: SUBRUTINAS
:: ==========================================================

:GetAppMruTargets
rem Args: OUT_VAR APP_NAME
set "GAM_OUT=%~1"
set "GAM_APP=%~2"
if "%GAM_OUT%"=="" exit /b 0

setlocal EnableDelayedExpansion
set "APP_NAME=%GAM_APP%"
set "MRU_TARGET_PATHS="

call :ResolveAppProperties "!APP_NAME!"
if defined PROP_REG_NAME (
    set "MRU_VAR=!PROP_MRU_VAR!"
    set "MRU_PATH="
    set "MRU_CONTAINER_PATH="
    set "MRU_CONTAINER_ID="

    call :DetectAdalContainer MRU_CONTAINER_ID MRU_CONTAINER_PATH "!PROP_REG_NAME!"
    if not errorlevel 1 if defined MRU_CONTAINER_PATH set "MRU_PATH=!MRU_CONTAINER_PATH!\File MRU"

    if not defined MRU_PATH (
        call :DetectMRUPath "!APP_NAME!"
        for /f "tokens=2 delims==" %%V in ('set !MRU_VAR! 2^>nul') do set "MRU_PATH=%%V"
    )

    if not defined MRU_PATH (
        set "MRU_PATH=HKCU\Software\Microsoft\Office\16.0\!PROP_REG_NAME!\Recent Templates\File MRU"
    )

    set "!MRU_VAR!=!MRU_PATH!"
    if defined MRU_PATH set "MRU_TARGET_PATHS=""!MRU_PATH!"""
    if defined MRU_TARGET_PATHS set "MRU_TARGET_PATHS=!MRU_TARGET_PATHS:""="!"

    set "AUTH_CONTAINER_TARGETS="
    call :CollectAuthContainerPaths AUTH_CONTAINER_TARGETS "!PROP_REG_NAME!"
    if defined AUTH_CONTAINER_TARGETS (
        set "AUTH_CONTAINER_TARGETS=!AUTH_CONTAINER_TARGETS:""="!"
        for %%C in (!AUTH_CONTAINER_TARGETS!) do (
            set "CURRENT_CONTAINER=%%~C"
            if not "!CURRENT_CONTAINER!"=="" call :AppendUniquePath MRU_TARGET_PATHS "!CURRENT_CONTAINER!\File MRU"
        )
    )

    if not defined MRU_TARGET_PATHS set "MRU_TARGET_PATHS=""!MRU_PATH!"""
    if defined MRU_TARGET_PATHS set "MRU_TARGET_PATHS=!MRU_TARGET_PATHS:""="!"
)

set "RESULT=%MRU_TARGET_PATHS%"

for %%# in (1) do (
    endlocal
    if not "%GAM_OUT%"=="" set "%GAM_OUT%=%RESULT%"
)

exit /b 0



:DetectAdalContainer
rem Args: OUT_ID_VAR OUT_PATH_VAR [APP_REG_NAME]
set "TARGET_ID=%~1"
set "TARGET_PATH=%~2"
set "TARGET_APP=%~3"
setlocal EnableDelayedExpansion

call :BuildAuthContainerCache "!TARGET_APP!"

if not defined __DAC_PRIMARY_PATH goto :dac_not_found

set "FOUND_ID=!__DAC_PRIMARY_ID!"
set "FOUND_PATH=!__DAC_PRIMARY_PATH!"

:dac_found
for %%# in (1) do (
    endlocal
    if not "%TARGET_ID%"=="" set "%TARGET_ID%=%FOUND_ID%"
    if not "%TARGET_PATH%"=="" set "%TARGET_PATH%=%FOUND_PATH%"
    exit /b 0
)

:dac_not_found
for %%# in (1) do (
    endlocal
    if not "%TARGET_ID%"=="" set "%TARGET_ID%="
    if not "%TARGET_PATH%"=="" set "%TARGET_PATH%="
    exit /b 1
)



:BuildAuthContainerCache
rem Args: APP_FILTER
set "DAC_REQUESTED_APP=%~1"
call :ResetAuthContainerCache

if defined DAC_REQUESTED_APP (
    set "__DAC_APP_LIST=%DAC_REQUESTED_APP%"
) else (
    set "__DAC_APP_LIST=Word PowerPoint Excel"
)

for %%V in (16.0 15.0 14.0 12.0) do (
    for %%A in (!__DAC_APP_LIST!) do (
        call :ScanRecentTemplateKey "%%~A" "%%~V"
    )
)

set "__DAC_APP_LIST="
set "DAC_REQUESTED_APP="

exit /b 0



:ResetAuthContainerCache
for /f "tokens=1 delims==" %%R in ('set __DAC_ 2^>nul') do set "%%R="
set "__DAC_COUNT=0"
set "__DAC_PRIMARY_ID="
set "__DAC_PRIMARY_PATH="
set "__DAC_PRIMARY_APP="
exit /b 0



:ScanRecentTemplateKey
rem Args: APP_NAME APP_VERSION
set "__DAC_CURRENT_APP=%~1"
set "__DAC_CURRENT_VER=%~2"
if not defined __DAC_CURRENT_APP exit /b 0
if not defined __DAC_CURRENT_VER exit /b 0

set "__DAC_CURRENT_KEY=HKCU\Software\Microsoft\Office\%__DAC_CURRENT_VER%\%__DAC_CURRENT_APP%\Recent Templates"

for /f "skip=2 tokens=*" %%S in ('2^>nul reg query "%__DAC_CURRENT_KEY%"') do (
    set "__DAC_SUBKEY=%%~S"
    if defined __DAC_SUBKEY (
        for %%T in ("!__DAC_SUBKEY!") do set "__DAC_LEAF=%%~nxT"
        call :HandleRecentTemplateSubkey "!__DAC_CURRENT_APP!" "!__DAC_LEAF!" "!__DAC_SUBKEY!"
    )
)

exit /b 0



:HandleRecentTemplateSubkey
rem Args: APP_NAME SUBKEY_NAME SUBKEY_PATH
set "__DAC_APP=%~1"
set "__DAC_ID=%~2"
set "__DAC_PATH=%~3"
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



:DetectMRUPath
rem Args: APP_NAME
setlocal enabledelayedexpansion
set "APP_NAME=%~1"
call :ResolveAppProperties "!APP_NAME!"
if not defined PROP_REG_NAME (
    endlocal
    exit /b 1
)
set "MRU_VAR=!PROP_MRU_VAR!"
set "MRU_PATH="
set "MRU_CONTAINER_PATH="

call :DetectAdalContainer MRU_CONTAINER_ID MRU_CONTAINER_PATH "!PROP_REG_NAME!"
if not errorlevel 1 if defined MRU_CONTAINER_PATH (
    set "MRU_PATH=!MRU_CONTAINER_PATH!\File MRU"
)

for %%V in (16.0 15.0 14.0 12.0) do (
    if not defined MRU_PATH (
        set "BASE=HKCU\Software\Microsoft\Office\%%V\!PROP_REG_NAME!\Recent Templates"
        for /f "delims=" %%K in ('reg query "!BASE!" /s /v "File MRU" 2^>nul ^| findstr /I "HKEY_CURRENT_USER"') do (
            set "MRU_PATH=%%K\File MRU"
            goto :found
        )
    )
)
:found
if not defined MRU_PATH (
    set "MRU_PATH=HKCU\Software\Microsoft\Office\16.0\!PROP_REG_NAME!\Recent Templates\File MRU"
)
endlocal & set "%MRU_VAR%=%MRU_PATH%"
exit /b



:ResolveAppProperties
rem Internal helper. Args: APP_NAME
set "APP_UP=%~1"
if /I "%APP_UP%"=="WORD" (
    set "PROP_REG_NAME=Word"
    set "PROP_MRU_VAR=WORD_MRU_PATH"
    set "PROP_COUNTER_VAR=GLOBAL_ITEM_COUNT_WORD"
) else if /I "%APP_UP%"=="POWERPOINT" (
    set "PROP_REG_NAME=PowerPoint"
    set "PROP_MRU_VAR=PPT_MRU_PATH"
    set "PROP_COUNTER_VAR=GLOBAL_ITEM_COUNT"
) else if /I "%APP_UP%"=="EXCEL" (
    set "PROP_REG_NAME=Excel"
    set "PROP_MRU_VAR=EXCEL_MRU_PATH"
    set "PROP_COUNTER_VAR=GLOBAL_ITEM_COUNT_EXCEL"
) else (
    set "PROP_REG_NAME="
    set "PROP_MRU_VAR="
    set "PROP_COUNTER_VAR="
)
exit /b



:ClearMruKey
rem Args: KEY_PATH, LABEL
setlocal DisableDelayedExpansion
set "CMK_KEY=%~1"
set "CMK_LABEL=%~2"
:: Limpiar comillas y saltos de línea
set "CMK_KEY=%CMK_KEY:"=%"
for /f "tokens=* delims=" %%A in ('echo %CMK_KEY%') do set "CMK_KEY=%%A"

if "%CMK_KEY%"=="" (
    endlocal & exit /b 0
)
powershell -NoLogo -NoProfile -Command ^
    "if (Test-Path 'Registry::%CMK_KEY%') { exit 0 } else { exit 1 }"

if errorlevel 1 (
    if /I "%IsDesignModeEnabled%"=="true" echo     %CMK_LABEL% no presente.
    endlocal & exit /b 0
)
powershell -NoLogo -NoProfile -Command ^
    "$path='Registry::%CMK_KEY%';" ^
    "if (Test-Path -LiteralPath $path) {" ^
    "    $item = Get-Item -LiteralPath $path;" ^
    "    foreach ($name in $item.GetValueNames()) {" ^
    "        if ([string]::IsNullOrEmpty($name)) {" ^
    "            Set-ItemProperty -LiteralPath $path -Name '(default)' -Value $null -ErrorAction SilentlyContinue;" ^
    "        } else {" ^
    "            Remove-ItemProperty -LiteralPath $path -Name $name -ErrorAction SilentlyContinue;" ^
    "        }" ^
    "    }" ^
    "}"
set "CMK_REMOVED=0"
for /f "usebackq tokens=*" %%R in (`powershell -NoLogo -NoProfile -Command ^
    "$path='Registry::%CMK_KEY%';" ^
    "if (-not (Test-Path -LiteralPath $path)) { Write-Output -1; exit }" ^
    "$item = Get-Item -LiteralPath $path;" ^
    "$removed = 0;" ^
    "foreach ($name in $item.GetValueNames()) {" ^
    "    if ([string]::IsNullOrEmpty($name)) {" ^
    "        Set-ItemProperty -LiteralPath $path -Name '(default)' -Value $null -ErrorAction SilentlyContinue;" ^
    "        continue" ^
    "    }" ^
    "    $value = $item.GetValue($name);" ^
    "    $filePath = ($value -split '\\*')[-1];" ^
    "    if ([string]::IsNullOrWhiteSpace($filePath) -or -not (Test-Path -LiteralPath $filePath)) {" ^
    "        Remove-ItemProperty -LiteralPath $path -Name $name -ErrorAction SilentlyContinue;" ^
    "        $removed++" ^
    "    }" ^
    "}" ^
    "Write-Output $removed;"`) do set "CMK_REMOVED=%%R"
if "%CMK_REMOVED%"=="-1" (
    if /I "%IsDesignModeEnabled%"=="true" echo     %CMK_LABEL% no presente.
    endlocal & exit /b 0
)
if /I "%IsDesignModeEnabled%"=="true" (
    if "%CMK_REMOVED%"=="0" (
        echo     %CMK_LABEL% sin cambios - todas las rutas siguen existiendo.
    ) else (
        echo     %CMK_LABEL% limpiada - %CMK_REMOVED% entradas eliminadas.
    )
)
endlocal & exit /b 0
