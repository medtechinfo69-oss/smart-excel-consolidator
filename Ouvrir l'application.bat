@echo off
rem ============================================================
rem  Smart Excel - ouverture locale
rem  Demarre le petit serveur local (cle Gemini lue dans .env,
rem  ou saisie dans l'application) puis ouvre le navigateur.
rem ============================================================
setlocal
set "APP_DIR=%~dp0"
set "PORT=3000"
set "NODE_EXE=node.exe"

where node >nul 2>nul
if errorlevel 1 (
  for %%P in ("%ProgramFiles%\nodejs\node.exe" "%ProgramFiles(x86)%\nodejs\node.exe" "%LocalAppData%\Programs\nodejs\node.exe") do (
    if exist %%P set "NODE_EXE=%%~P"
  )
)
where node >nul 2>nul
if errorlevel 1 if "%NODE_EXE%"=="node.exe" (
  echo Node.js est introuvable sur cet ordinateur.
  echo Installez la version LTS depuis https://nodejs.org puis relancez ce fichier.
  echo L'application reste utilisable en double-cliquant index.html, sans assistant Gemini.
  pause
  exit /b 1
)

echo Demarrage de Smart Excel sur http://127.0.0.1:%PORT%
start "" "http://127.0.0.1:%PORT%/"
"%NODE_EXE%" "%APP_DIR%server\server.js"
pause
