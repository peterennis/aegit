; The name of the installer
Name aegit

; The file to write
OutFile "aegit.exe"

; The default installation directory
InstallDir $DESKTOP\aegit

; Request application privileges for Windows Vista
RequestExecutionLevel user

;--------------------------------

; Pages

Page directory
Page instfiles

;--------------------------------

; The stuff to install
Section "" ;No components page, name is not important

  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ; Put files there
  File "adaept revision control.accdb"
  File "adaept revision control.bmp"
  File "adaept64.ico"
  File "lgpl-3.0.txt"

  # create a shortcut named "adaept revision control" in the desktop
  # point the new shortcut at the app
  CreateShortCut "$DESKTOP\adaept revision control.lnk" "$INSTDIR\adaept revision control.accdb"

  
SectionEnd ; end the section
