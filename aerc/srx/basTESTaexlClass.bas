Attribute VB_Name = "basTESTaexlClass"
Option Explicit

' Default Usage:
' The following folders are used if no custom configuration is provided:
' aegitType.SourceFolder = "C:\ae\aegit\aerc\srx\"
' aegitType.ImportFolder = "C:\ae\aegit\aerc\imx\"
' Run in immediate window:                  aexlClassTest
' Show debug output in immediate window:    aexlClassTest("debug")
'
' Custom Usage:
' Public Const THE_SOURCE_FOLDER = "Z:\The\Source\Folder\srx.MYPROJECT\"
' For custom configuration of the output source folder in aexlClassTest use:
' oDbObjects.SourceFolder = THE_SOURCE_FOLDER
' Run in immediate window: MYPROJECT_TEST
'

Public Function MYPROJECT_TEST()
    aexlClassTest
    'aexlClassTest ("debug")
End Function

Private Function aexlClassTest(Optional Debugit As Variant) As Boolean
    
    Debug.Print "Function aexlClassTest"

End Function
