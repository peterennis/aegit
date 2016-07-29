@ECHO OFF
set root_folder=c:\ae\aegit\
set project_folder=aefx

:: Change to the root folder
cd %root_folder%

:: Print the current directory
cd
pause
.\docfx\docfx.exe init --quiet --output %project_folder%

.\docfx\docfx.exe serve %root_folder%\_site



echo Press any key to exit . . .
pause > nul

