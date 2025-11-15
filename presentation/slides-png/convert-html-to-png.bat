@echo off
REM Convert HTML slides to PNG using Chrome headless mode (Windows)
REM Usage: convert-html-to-png.bat

echo Converting HTML slides to PNG...
echo.

REM Create output directory
if not exist png-output mkdir png-output

REM Find Chrome installation
set CHROME=""
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    set CHROME="%ProgramFiles%\Google\Chrome\Application\chrome.exe"
) else if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
    set CHROME="%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
) else if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" (
    set CHROME="%LocalAppData%\Google\Chrome\Application\chrome.exe"
) else (
    echo Error: Google Chrome not found. Please install Chrome or update the path.
    pause
    exit /b 1
)

echo Using Chrome: %CHROME%
echo.

REM Convert each slide
for /l %%i in (1,1,17) do (
    set "num=0%%i"
    set "num=!num:~-2!"
    if exist "slide%%i.html" (
        echo Converting slide%%i.html to slide-!num!.png
        %CHROME% --headless --screenshot="png-output\slide-!num!.png" --window-size=1920,1080 --default-background-color=0 "slide%%i.html" 2>nul
    ) else (
        echo Warning: slide%%i.html not found
    )
)

echo.
echo Conversion complete!
echo PNG files saved to: png-output\
dir /b png-output\*.png
pause
