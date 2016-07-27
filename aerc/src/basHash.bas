Option Compare Database
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub HashAllModules()
' mwolfe02, Ref: https://github.com/rubberduck-vba/Rubberduck/issues/1966

    On Error GoTo 0

    Dim s As Long
    Dim LineCount As Long

    s = GetTickCount

    Dim enc As Object
    Set enc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")

    Dim obj As AccessObject
    For Each obj In CurrentProject.AllModules
        Dim objLoaded As Boolean
        objLoaded = obj.IsLoaded

        If Not objLoaded Then DoCmd.OpenModule obj.Name
        Dim mdl As Module
        Set mdl = Application.Modules(obj.Name)
        LineCount = LineCount + mdl.CountOfLines

        'Convert the module contents to a byte array...
        Dim Content() As Byte
        Content = mdl.Lines(1, mdl.CountOfLines)

        '...and hash it
        Dim HashedBytes As Variant
        HashedBytes = enc.ComputeHash_2((Content))

        'Convert the hashed byte array to a hex string
        Dim Pos As Integer
        Dim HexedHash As String
        HexedHash = vbNullString
        For Pos = 1 To LenB((HashedBytes))
            HexedHash = HexedHash & LCase$(Right$("0" & Hex$(AscB(MidB$(HashedBytes, Pos, 1))), 2))
        Next
        Debug.Print HexedHash, mdl.Name    'Returns a 40 byte/character hex string

        Set mdl = Nothing
        If Not objLoaded Then DoCmd.Close acModule, obj.Name

    Next obj
    Set enc = Nothing
    Debug.Print "Done...processed "; CurrentProject.AllModules.Count; "modules with"; _
                LineCount; "lines of code in"; GetTickCount - s; "ms"

End Sub