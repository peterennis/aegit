@ECHO OFF
:: Ref https://dotnet.github.io/docfx/tutorial/walkthrough/walkthrough_create_a_docfx_project.html

set root_folder=c:\ae\aegit\
set project_folder=aefx
set theme=C:\ae\aegit\themes\ae\

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

:: Copy the assets
copy %theme%images\favicon.ico .\_site
copy %theme%images\logo.svg .\_site

:: Serve the site
..\docfx\docfx.exe serve .\_site



echo Press any key to exit . . .
pause > nul

