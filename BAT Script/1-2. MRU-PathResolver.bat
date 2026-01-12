rem 1-2.MRU
@echo off

:DetectMRUPath

setlocal enabledelayedexpansion

rem ------------------------------------------------------------
rem PARAMETROS
rem   %~1 = APP_NAME   (WORD | POWERPOINT | EXCEL)
rem   %~2 = AUTH_MODE  (ADAL | LIVEID)
rem ------------------------------------------------------------

set "APP_NAME=%~2"
set "AUTH_MODE=%~3"

rem Normalizar AUTH_MODE
if /I "%AUTH_MODE%"=="ADAL"   set "AUTH_MODE=ADAL"
if /I "%AUTH_MODE%"=="LiveID" set "AUTH_MODE=LIVEID"

rem ------------------------------------------------------------
rem Resolver propiedades base
rem ------------------------------------------------------------
call "%ResolveAppProps%" "%APP_NAME%"
if not defined PROP_REG_NAME (
    echo [ERROR] Aplicación desconocida: %APP_NAME%
    endlocal & exit /b 1
)

rem Nombre corto para componer variable final
if /I "%APP_NAME%"=="WORD"       set "PROP_SHORT_VAR=WORD"
if /I "%APP_NAME%"=="POWERPOINT" set "PROP_SHORT_VAR=PPT"
if /I "%APP_NAME%"=="EXCEL"      set "PROP_SHORT_VAR=EXCEL"

set "OUT_VAR=%PROP_SHORT_VAR%_MRU_%AUTH_MODE%"

set "MRU_PATH="

rem ============================================================
rem 1. AUTH_MODE = ADAL → buscar contenedor ADAL
rem ============================================================
if /I "%AUTH_MODE%"=="ADAL" (
    call :DetectAdalContainer DAC_ID DAC_PATH "!PROP_REG_NAME!"
    if defined DAC_PATH (
        set "MRU_PATH=!DAC_PATH!\File MRU"
    )
)

rem ============================================================
rem 2. AUTH_MODE = LIVEID → buscar contenedor LiveID
rem ============================================================
if /I "%AUTH_MODE%"=="LIVEID" (
    call :DetectLiveIdContainer LID_ID LID_PATH "!PROP_REG_NAME!"
    if defined LID_PATH (
        set "MRU_PATH=!LID_PATH!\File MRU"
    )
)

rem ============================================================
rem 3. FALLBACK: si aún no se definió MRU_PATH, buscar File MRU general
rem ============================================================
if not defined MRU_PATH (
    for %%V in (16.0 15.0 14.0 12.0) do (
        if not defined MRU_PATH (
            set "BASE=HKCU\Software\Microsoft\Office\%%V\!PROP_REG_NAME!\Recent Templates"
            for /f "delims=" %%K in ('reg query "!BASE!" /s /v "File MRU" 2^>nul ^| findstr /I "HKEY_CURRENT_USER"') do (
                set "MRU_PATH=%%K\File MRU"
                goto :found_roml
            )
        )
    )
)

:found_roml

rem ============================================================
rem 4. Si sigue vacío, fallback final
rem ============================================================
if not defined MRU_PATH (
    set "MRU_PATH=HKCU\Software\Microsoft\Office\16.0\!PROP_REG_NAME!\Recent Templates\File MRU"
)

endlocal & set "%OUT_VAR%=%MRU_PATH%"
exit /b 0



rem ============================================================
rem ===           Función: DetectLiveIdContainer              ===
rem ============================================================
:DetectLiveIdContainer
rem Args: OUT_ID_VAR OUT_PATH_VAR APP_REG_NAME

set "OUT_ID=%~1"
set "OUT_PATH=%~2"
set "APP=%~3"

setlocal EnableDelayedExpansion

set "__LIVE_ID="
set "__LIVE_PATH="

for %%V in (16.0 15.0 14.0 12.0) do (
    set "__BASE=HKCU\Software\Microsoft\Office\%%V\!APP!\Recent Templates"
    for /f "skip=2 tokens=*" %%S in ('reg query "!__BASE!" 2^>nul') do (
        set "__SUB_KEY=%%~S"
        for %%X in ("!__SUB_KEY!") do set "__LEAF=%%~nxX"
        if /I "!__LEAF:~0,7!"=="LIVEID_" (
            set "__LIVE_ID=!__LEAF!"
            set "__LIVE_PATH=!__SUB_KEY!"
            goto :live_found
        )
    )
)

:live_found
endlocal & (
    if not "%OUT_ID%"=="" set "%OUT_ID%=%__LIVE_ID%"
    if not "%OUT_PATH%"=="" set "%OUT_PATH%=%__LIVE_PATH%"
)
exit /b 0



rem ============================================================
rem ===          ResolveAppProperties (sin cambios)           ===
rem ============================================================
:ResolveAppProperties
echo [DEBUG - en MRU - ResolveAppProperties] Entered ResolveAppProperties with args: %*
set "APP_UP=%~1"

if /I "%APP_UP%"=="WORD" (
    set "PROP_REG_NAME=Word"
) else if /I "%APP_UP%"=="POWERPOINT" (
    set "PROP_REG_NAME=PowerPoint"
) else if /I "%APP_UP%"=="EXCEL" (
    set "PROP_REG_NAME=Excel"
) else (
    set "PROP_REG_NAME="
)

exit /b 0


rem ============================================================
rem ===        DetectAdalContainer (tu versión original)      ===
rem ===        (no la cambio para no romper nada)             ===
rem ============================================================

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
reg query "%__DAC_CURRENT_KEY%" >nul 2>&1
if errorlevel 1 (
    exit /b 0
)

for /f "skip=2 tokens=*" %%S in ('reg query ^"%__DAC_CURRENT_KEY%^"') do (

    set "__DAC_SUBKEY=%%~S"

    rem Extraer correctamente el nombre final de la clave de registro
    for %%A in ("!__DAC_SUBKEY!") do set "__DAC_LEAF=%%~nxA"
    if /I "%IsDesignModeEnabled%"=="true" (
        set "__PREFIX5=!__DAC_LEAF:~0,5!"
        set "__PREFIX7=!__DAC_LEAF:~0,7!"

        rem Mostrar solo ADAL_ o LIVEID_
        if /I "!__PREFIX5!"=="ADAL_"  echo [DEBUG - ScanRecentTemplateKey] Encontrado contenedor ADAL: !__DAC_LEAF!
        if /I "!__PREFIX7!"=="LiveId_" echo [DEBUG - ScanRecentTemplateKey] Encontrado contenedor LiveID: !__DAC_LEAF!
    )

    if defined __DAC_LEAF (
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