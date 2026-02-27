@echo off
dotnet publish NetworkDiagramBuilder.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -p:DebugType=embedded -o ./publish
if %ERRORLEVEL% neq 0 (
  echo BUILD FAILED
  pause
  exit /b 1
)
echo.
echo BUILD SUCCESS
echo Output: publish\NetworkDiagramBuilder.exe
echo.
pause
