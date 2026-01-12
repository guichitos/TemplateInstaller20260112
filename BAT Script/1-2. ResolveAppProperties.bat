@echo off
rem ============================================================
rem ===            1-2. ResolveAppProperties.bat             ===
rem ===           Biblioteca de propiedades de la app        ===
rem ===  Uso: call "1-2. ResolveAppProperties.bat" APP1 APP2 ===
rem ===       Devuelve: PROP_REG_NAME                        ===
rem ============================================================

rem Convertir parámetros a mayúsculas
set "APP_UP=%~1"
set "APP_ALT=%~2"

if defined APP_UP  set "APP_UP=%APP_UP:"=%"
if defined APP_ALT set "APP_ALT=%APP_ALT:"=%"

for %%A in (APP_UP APP_ALT) do (
    for /f "delims=" %%B in ("!%%A!") do set "%%A=%%~B"
)

set "APP_UP=%APP_UP%"
set "APP_ALT=%APP_ALT%"

set "APP_UP=%APP_UP:~0,256%"
set "APP_ALT=%APP_ALT:~0,256%"

set "APP_UP=%APP_UP%"
set "APP_ALT=%APP_ALT%"

set "APP_UP=%APP_UP%"
set "APP_ALT=%APP_ALT%"

set "AU=%APP_UP%"
set "AL=%APP_ALT%"

rem Normalizar a mayúsculas
set "AU=%AU:"=%"
set "AL=%AL:"=%"

set "AU=%AU%"
set "AL=%AL%"

for %%C in (AU AL) do (
    for /f "delims=" %%D in ("!%%C!") do set "%%C=%%~D"
)

set "AU=%AU%"
set "AL=%AL%"

rem Forzar uppercase
for %%X in (AU AL) do (
    for /f "delims=" %%Y in ("!%%X!") do set "%%X=%%Y"
)
set "AU=%AU%"
set "AL=%AL%"

rem Convertir a uppercase real
for %%X in (AU AL) do (
    for /f "tokens=*" %%Z in ('echo !%%X!^| powershell -noprofile -command "$input.toupper()"') do set "%%X=%%Z"
)

rem ============================================================
rem === LÓGICA OR (cualquiera de los dos parámetros sirve)  ===
rem ============================================================

set "PROP_REG_NAME="

if /I "%AU%"=="WORD"       set "PROP_REG_NAME=Word"
if /I "%AL%"=="WORD"       set "PROP_REG_NAME=Word"

if /I "%AU%"=="POWERPOINT" set "PROP_REG_NAME=PowerPoint"
if /I "%AL%"=="POWERPOINT" set "PROP_REG_NAME=PowerPoint"

if /I "%AU%"=="EXCEL"      set "PROP_REG_NAME=Excel"
if /I "%AL%"=="EXCEL"      set "PROP_REG_NAME=Excel"

rem Si ninguno coincidió, PROP_REG_NAME queda vacío
exit /b 0
