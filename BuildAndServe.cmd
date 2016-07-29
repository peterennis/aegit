@ECHO OFF
:: Ref https://dotnet.github.io/docfx/tutorial/walkthrough/walkthrough_create_a_docfx_project.html

set root_folder=c:\ae\aegit\
set project_folder=aefx

:: Change to the root folder
cd %root_folder%

:: Print the current directory
cd
pause
.\docfx\docfx.exe init --quiet --output %project_folder%

cd %project_folder%
cd

:: Build the site
..\docfx\docfx.exe docfx.json

:: Serve the site
..\docfx\docfx.exe serve .\_site



echo Press any key to exit . . .
pause > nul

