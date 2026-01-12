@echo off
setlocal enabledelayedexpansion

:: Params:
:: 1     - Design mode flag (true/false) controls debug output
:: 2-6   - Booleans to open each folder (theme, custom, roaming, Excel startup, alt custom)
:: 7     - Global boolean to attempt file selection instead of only opening the folder
:: 8-12  - Filenames to optionally select per folder in the same order
set "TFO_DESIGN_MODE=%~1"
if not defined TFO_DESIGN_MODE set "TFO_DESIGN_MODE=false"
rem set "TFO_DESIGN_MODE=true" descomentar para debug local
set "OPEN_THEME=%~2"
set "OPEN_CUSTOM=%~3"
set "OPEN_ROAMING=%~4"
set "OPEN_EXCEL=%~5"
set "OPEN_CUSTOM_ALT=%~6"
set "SELECT_FILES=%~7"
set "THEME_FILE=%~8"
set "CUSTOM_FILE=%~9"
shift
set "ROAMING_FILE=%~9"
shift
set "EXCEL_FILE=%~9"
shift
set "CUSTOM_ALT_FILE=%~9"

if /I "%TFO_DESIGN_MODE%"=="true" (
    echo ================= WORKER RAW PARAMETERS =================
    echo Arg  Indesign mode = [%TFO_DESIGN_MODE%]
    echo Arg  Open theme = [%OPEN_THEME%]
    echo Arg  Open custom = [%OPEN_CUSTOM%]
    echo Arg  Open roaming = [%OPEN_ROAMING%]
    echo Arg  Open Excel startup = [%OPEN_EXCEL%]
    echo Arg  Open custom alt = [%OPEN_CUSTOM_ALT%]
    echo Arg  Select files = [%SELECT_FILES%]
    echo Arg  THEME_FILE = [%THEME_FILE%]
    echo Arg  CUSTOM_FILE = [%CUSTOM_FILE%]
    echo Arg  ROAMING_FILE = [%ROAMING_FILE%]
    echo Arg  EXCEL_FILE = [%EXCEL_FILE%]
    echo Arg  CUSTOM_ALT_FILE = [%CUSTOM_ALT_FILE%]
    echo =========================================================
    echo.

    pause
)

if not defined OPEN_THEME set "OPEN_THEME=1"
if not defined OPEN_CUSTOM set "OPEN_CUSTOM=1"
if not defined OPEN_ROAMING set "OPEN_ROAMING=1"
if not defined OPEN_EXCEL set "OPEN_EXCEL=1"
if not defined OPEN_CUSTOM_ALT set "OPEN_CUSTOM_ALT=1"
if not defined SELECT_FILES set "SELECT_FILES=0"

set "ScriptDirectory=%~dp0"
set "OfficeTemplateLib=%ScriptDirectory%1-2. AuthContainerTools.bat"

set "APPDATA_EXPANDED="
for /f "delims=" %%T in ('powershell -NoLogo -Command "$app=(Get-ItemProperty -Path \"HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\" -Name AppData -ErrorAction SilentlyContinue).AppData; if ($app) {[Environment]::ExpandEnvironmentVariables($app)}"') do set "APPDATA_EXPANDED=%%T"
if not defined APPDATA_EXPANDED set "APPDATA_EXPANDED=%APPDATA%"

set "THEME_PATH=%APPDATA_EXPANDED%\Microsoft\Templates\Document Themes"
set "ROAMING_TEMPLATE_PATH=%APPDATA_EXPANDED%\Microsoft\Templates"
set "EXCEL_STARTUP_PATH=%APPDATA_EXPANDED%\Microsoft\Excel\XLSTART"

set "DOCUMENTS_PATH="
for /f "delims=" %%D in ('powershell -NoLogo -Command "$path=(Get-ItemProperty -Path \"HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\" -Name Personal -ErrorAction SilentlyContinue).Personal; if ($path) {[Environment]::ExpandEnvironmentVariables($path)}"') do set "DOCUMENTS_PATH=%%D"
if defined DOCUMENTS_PATH (
    if "!DOCUMENTS_PATH:~-1!"=="\" set "DOCUMENTS_PATH=!DOCUMENTS_PATH:~0,-1!"
    set "DEFAULT_CUSTOM_DIR=!DOCUMENTS_PATH!\Custom Templates"
) else (
    set "DEFAULT_CUSTOM_DIR=%USERPROFILE%\Documents\Custom Templates"
)
if not defined DEFAULT_CUSTOM_DIR set "DEFAULT_CUSTOM_DIR=%USERPROFILE%\Documents\Custom Templates"

set "CUSTOM_OFFICE_TEMPLATE_PATH="
for %%V in (16.0 15.0 14.0 12.0) do (
    if not defined CUSTOM_OFFICE_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Word\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "CUSTOM_OFFICE_TEMPLATE_PATH=%%C"
    )
)
for %%V in (16.0 15.0 14.0 12.0) do (
    if not defined CUSTOM_OFFICE_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "CUSTOM_OFFICE_TEMPLATE_PATH=%%C"
    )
)
if not defined CUSTOM_OFFICE_TEMPLATE_PATH if defined DEFAULT_CUSTOM_DIR set "CUSTOM_OFFICE_TEMPLATE_PATH=%DEFAULT_CUSTOM_DIR%"
if not defined CUSTOM_OFFICE_TEMPLATE_PATH set "CUSTOM_OFFICE_TEMPLATE_PATH=%USERPROFILE%\Documents\Custom Templates"

set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH="
for %%V in (16.0 15.0 14.0 12.0) do (
    if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\PowerPoint\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=%%C"
    )
)

for %%V in (16.0 15.0 14.0 12.0) do (
    if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=%%C"
    )
)
if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH if defined CUSTOM_OFFICE_TEMPLATE_PATH set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=%CUSTOM_OFFICE_TEMPLATE_PATH%"
if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH if defined DEFAULT_CUSTOM_DIR set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=%DEFAULT_CUSTOM_DIR%"
if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=%USERPROFILE%\Documents\Custom Templates"

if exist "%OfficeTemplateLib%" (
    call "%OfficeTemplateLib%" :CleanPath APPDATA_EXPANDED
    call "%OfficeTemplateLib%" :CleanPath THEME_PATH
    call "%OfficeTemplateLib%" :CleanPath ROAMING_TEMPLATE_PATH
    call "%OfficeTemplateLib%" :CleanPath EXCEL_STARTUP_PATH
    call "%OfficeTemplateLib%" :CleanPath DOCUMENTS_PATH
    call "%OfficeTemplateLib%" :CleanPath DEFAULT_CUSTOM_DIR
    call "%OfficeTemplateLib%" :CleanPath CUSTOM_OFFICE_TEMPLATE_PATH
    call "%OfficeTemplateLib%" :CleanPath CUSTOM_OFFICE_TEMPLATE_ALT_PATH
) else (
    if /I "%TFO_DESIGN_MODE%"=="true" echo [DEBUG] There is no OfficeTemplateLib
    if "!APPDATA_EXPANDED:~-1!"=="\" set "APPDATA_EXPANDED=!APPDATA_EXPANDED:~0,-1!"
    if "!THEME_PATH:~-1!"=="\" set "THEME_PATH=!THEME_PATH:~0,-1!"
    if "!ROAMING_TEMPLATE_PATH:~-1!"=="\" set "ROAMING_TEMPLATE_PATH=!ROAMING_TEMPLATE_PATH:~0,-1!"
    if "!EXCEL_STARTUP_PATH:~-1!"=="\" set "EXCEL_STARTUP_PATH=!EXCEL_STARTUP_PATH:~0,-1!"
    if "!DOCUMENTS_PATH:~-1!"=="\" set "DOCUMENTS_PATH=!DOCUMENTS_PATH:~0,-1!"
    if "!DEFAULT_CUSTOM_DIR:~-1!"=="\" set "DEFAULT_CUSTOM_DIR=!DEFAULT_CUSTOM_DIR:~0,-1!"
    if "!CUSTOM_OFFICE_TEMPLATE_PATH:~-1!"=="\" set "CUSTOM_OFFICE_TEMPLATE_PATH=!CUSTOM_OFFICE_TEMPLATE_PATH:~0,-1!"
    if "!CUSTOM_OFFICE_TEMPLATE_ALT_PATH:~-1!"=="\" set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=!CUSTOM_OFFICE_TEMPLATE_ALT_PATH:~0,-1!"
)

if /I "%TFO_DESIGN_MODE%"=="true" (
    echo.
    echo ================= CALCULATED FOLDER PATHS =================
    echo THEME_PATH                     = [!THEME_PATH!]
    echo CUSTOM_OFFICE_TEMPLATE_PATH    = [!CUSTOM_OFFICE_TEMPLATE_PATH!]
    echo ROAMING_TEMPLATE_PATH          = [!ROAMING_TEMPLATE_PATH!]
    echo EXCEL_STARTUP_PATH             = [!EXCEL_STARTUP_PATH!]
    echo CUSTOM_OFFICE_TEMPLATE_ALT_PATH= [!CUSTOM_OFFICE_TEMPLATE_ALT_PATH!]
    echo ===========================================================
    echo.
)

call :OpenIfEnabled "!OPEN_THEME!" "%THEME_PATH%" "!SELECT_FILES!" "!THEME_FILE!"
call :OpenIfEnabled "!OPEN_CUSTOM!" "%CUSTOM_OFFICE_TEMPLATE_PATH%" "!SELECT_FILES!" "!CUSTOM_FILE!"
call :OpenIfEnabled "!OPEN_ROAMING!" "%ROAMING_TEMPLATE_PATH%" "!SELECT_FILES!" "!ROAMING_FILE!"
call :OpenIfEnabled "!OPEN_EXCEL!" "%EXCEL_STARTUP_PATH%" "!SELECT_FILES!" "!EXCEL_FILE!"
call :OpenIfEnabled "!OPEN_CUSTOM_ALT!" "%CUSTOM_OFFICE_TEMPLATE_ALT_PATH%" "!SELECT_FILES!" "!CUSTOM_ALT_FILE!"

goto :EOF

:OpenIfEnabled
if /I "%TFO_DESIGN_MODE%"=="true" (
    echo ---------------------------
    echo Abriendo con parametros:
    echo Flag: [%~1]
    echo Target: [%~2]
    echo SelectFlag: [%~3]
    echo Filename: [%~4]
    echo ---------------------------
)

set "FLAG=%~1"
set "TARGET_RAW=%~2"
set "SELECT_FLAG=%~3"
set "FILENAME=%~4"

set "SHOULD_OPEN=0"
for %%B in (1 true yes on) do if /I "!FLAG!"=="%%B" set "SHOULD_OPEN=1"

if "!SHOULD_OPEN!"=="1" (
    set "TARGET=!TARGET_RAW!"
    call :NormalizePath "!TARGET!" TARGET_COMPARE
    set "TARGET=!TARGET_COMPARE!"

    rem === detectar OneDrive ===
    set "IS_ONEDRIVE=0"
    if not "!TARGET:\OneDrive\=!"=="!TARGET!" set "IS_ONEDRIVE=1"

    set "SHOULD_SELECT=0"
    for %%B in (1 true yes on) do if /I "!SELECT_FLAG!"=="%%B" set "SHOULD_SELECT=1"
    if "!SHOULD_SELECT!"=="0" if defined FILENAME if not "!FILENAME!"=="" set "SHOULD_SELECT=1"

    if "!SHOULD_SELECT!"=="1" if "!IS_ONEDRIVE!"=="0" (
        if /I "%TFO_DESIGN_MODE%"=="true" echo Intentando seleccionar archivo: "!FILENAME!" en carpeta: "!TARGET!"
        if defined FILENAME (
            set "FILE_TARGET=!TARGET!\!FILENAME!"
            if exist "!FILE_TARGET!" (
                if /I "%TFO_DESIGN_MODE%"=="true" (
                    echo Ejecutando:
                    echo start "" explorer.exe /select,"!FILE_TARGET!"
                )
                start "" explorer.exe /select,"!FILE_TARGET!"
            ) else (
                if /I "%TFO_DESIGN_MODE%"=="true" (
                    echo Archivo no existe. Abriendo solo carpeta.
                    echo start "" "!TARGET!"
                )
                start "" "!TARGET!"
            )
            set "FILE_TARGET="
        ) else (
            if /I "%TFO_DESIGN_MODE%"=="true" (
                echo No se proporciono archivo. Abriendo solo carpeta.
                echo start "" "!TARGET!"
            )
            start "" "!TARGET!"
        )
    ) else (
        if /I "%TFO_DESIGN_MODE%"=="true" (
            if "!IS_ONEDRIVE!"=="1" (
                echo Ruta bajo OneDrive detectada. Omitiendo seleccion de archivo.
            ) else (
                echo Seleccion deshabilitada o sin archivo. Abriendo solo carpeta.
            )
        )
        if /I "%TFO_DESIGN_MODE%"=="true" echo start "" "!TARGET!"
        start "" "!TARGET!"
    )
)

set "FLAG="
set "TARGET_RAW="
set "SELECT_FLAG="
set "SHOULD_SELECT="
set "FILENAME="
set "SHOULD_OPEN="
set "IS_ONEDRIVE="

goto :EOF

:NormalizePath
set "NP_INPUT=%~1"
set "NP_OUTPUT_VAR=%~2"
if "%NP_OUTPUT_VAR%"=="" exit /b
setlocal enabledelayedexpansion
set "NP_WORK=!NP_INPUT!"
:_TrimLoop
if defined NP_WORK if "!NP_WORK:~-1!"==" " set "NP_WORK=!NP_WORK:~0,-1!" & goto _TrimLoop
if defined NP_WORK if "!NP_WORK:~-1!"=="\\" set "NP_WORK=!NP_WORK:~0,-1!" & goto _TrimLoop
endlocal & set "%NP_OUTPUT_VAR%=%NP_WORK%"
exit /b
