@echo off
rem Сборка ConverterReport.exe (один файл, без установки .NET) в папку publish
chcp 65001 >nul
dotnet test || exit /b 1
dotnet publish src\ConverterReport.Gui -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publish || exit /b 1
echo.
echo Готово: publish\ConverterReport.exe
