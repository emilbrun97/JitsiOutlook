@echo off

rem Set the path to msbuild.exe
set MSBUILD_PATH="C:\Windows\Microsoft.NET\Framework\v4.0.30319\msbuild.exe"

rem Set the path to the solution file
set SOLUTION_PATH="C:\MyProjects\OutlookAddIn1\OutlookAddIn1.sln"

rem Build the solution in release mode
"%MSBUILD_PATH%" "%SOLUTION_PATH%" /p:Configuration=Release /verbosity:minimal

echo Build complete
