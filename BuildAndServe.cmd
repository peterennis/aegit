@ECHO OFF
set root_folder=c:\ae\aegit\
set project_folder=aefx

:: Change to the root folder
cd %root_folder%

:: Print the current directory
cd
pause
.\docfx\docfx.exe init --quiet --output %project_folder%

echo Press any key to exit . . .
pause > nul

