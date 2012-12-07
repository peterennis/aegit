Option Compare Database
Option Explicit

' Problems:
' ReadDocDatabase debug output when custom test folder given
' Exists function debug output
' Test for expected references when class first created
' Import of class source code into a new database creates a module
' http://www.trigeminal.com/usenet/usenet026.asp - Fix DISAMBIGUATION?
' http://access.mvps.org/access/modules/mdl0022.htm - test the References Wizard?
' Pass Fail of the tests should be associated to True False of the function, any error should return False
' Fix error when tst folder not set, it is intended for import testing to recreate a database
'


'
'20121205 - v022 - Print linked table path in output, GetLinkedTableCurrentPath
    ' Table name header output fix for length and linked table path
'20121205 - v021 - Centralize code comments in basChangeLog
    ' output for append on tables setup, kill files before export
    ' Move historical comments in the code into the log here:
                '20121204 v020  intFailCount for TableInfo, output sql text for queries
                '               output table setup
                '20121203 v019  LongestFieldPropsName()
                '20121201 v018  Fix err=0 and error=0
                '               Add SizeString from Chip Pearson for help formatting TableInfo from Allen Browne
                '               Include LGPL license
                '               Ref: http://www.gnu.org/licenses/gpl-howto.html
                '               Ref: http://blogs.sourceallies.com/2011/07/creating-an-open-source-project/
                '20121129 v017  Output error messages to the immediate window when debug is turned on
                '               Pass Fail test results and debug output cleanup
                '20121128 v016  Use strSourceLocation to allow custom path and test for error,
                '               Cleanup debug messages code
                '               Include GetReferences from aeladdin (tm) and fix it
                '20121127 v015  Update version, export using OASIS and commit to github
                '               Reverse order of version comments so newest is at the top
                '               Skip ~TMP* names for scripts (macros)
                '20110303 v014  Make class PublicNotCreatable, project name aegitClassProvider
                '               http://support.microsoft.com/kb/555159
                '20110303 v013  Initialize class using Private Type
                '20110303 v012  Fix bug in skip export of all zzz objects, must use doc.Name
                '20110303 v011  Skip export of all zzz objects, create module basTESTaegitClass
                '20110303 v010  Add Option blnDebug to ReadDocDatabase property
                '20110302 v010  Delete basRevisionControl
                '20110302 v009  Skip export of ~TMP queries, debug message output singular and plural
                '20110302 v008  Move other finctions from basRevisionControl to asgitClass
                '20110302 v007  Add private function aeDocumentTheDatabase from DocumentTheDatabase
                '               Test with updated aegitClassTest
                '20110226 v006  TEST_FOLDER=>THE_FOLDER, TEST_DRIVE=>THE_DRIVE, BuildTestDirectory=>BuildTheDirectory
                '               Objects have obj prefix, use For Each qdf, output "Macros EXPORTED" (not Scripts)
                '20110222 v004  Create aegitClass shell and basTestRevisionControl
                '               Use ?aegitClassTest of basTestRevisionControl in the immediate window to check basic operation
                '
                '====================================================================
                'Private Function aeDocumentTheDatabase(Optional varDebug As Variant) As Boolean
                ' 20121128: Use strSourceLocation to allow custom path and test for error,
                '           Cleanup debug messages code
                ' 20121127: Reverse comment order, newest at top
                '           Skip export of ~TMP macros
                ' 20110303: Add Optional blnDebug parameter
                '           Skip export of all zzz objects (using doc.Name)
                ' 20110302: Skip export of ~TMP queries
                '           debug message output singular and plural
                ' 20110302: Change to aeDocumentTheDatabase for use in aegitClass
                ' 20110226: Skip export of MSys (hiddem system queries) and
                '           ~sq_ (hidden ODBC queries) objects
                '           Add count of objects in debug output
                ' 20110224: Make this a function. Add optional debug flag
                ' 20110218: Forms->frm, Reports->rpt, Scripts->mac
                '           Modules->bas, Queries->qry
                '           Error handler
                '====================================================================
                '
                '====================================================================
                'Private Function BuildTheDirectory(FSO As Scripting.FileSystemObject, _
                '                                        Optional varDebug As Variant) As Boolean
                ' 20110302: Add error handler and include in aegitClass
                '====================================================================
                '
                '====================================================================
                'Private Function aeReadDocDatabase(Optional varDebug As Variant) As Boolean
                ' 20121128: Fix debugging output
                ' 20110303: Add Debug.Print output for Skipping: message
                '           Output VERSION and VERSION_DATE for debug
                ' 20110302: Change to aeReadDocDatabase for use in aegitClass
                '           Add Skipping: to MsgBox for existing objects
                ' 20110224: Make this a function
                '====================================================================
                '
                '====================================================================
                'Private Function aeExists(strAccObjType As String, _
                '                        strAccObjName As String, Optional varDebug As Variant) As Boolean
                ' 20121128:   Fix debugging output
                ' 20110302:   Make aeExists private in aegitClass
                '====================================================================
                '
                '====================================================================
                'Private Function TableInfo(strTableName As String, Optional varDebug As Variant) As Boolean
                'Original Code Provided by Allen Browne. Last updated: April 2010.
                'TableInfo() function
                'This function displays in the Immediate Window (Ctrl+G) the structure of any table in the current database
                'For Access 2000 or 2002, make sure you have a DAO reference
                'The Description property does not exist for fields that have no description, so a separate function handles that error
                'Update:   Peter Ennis
                '20121201  SizeString(), LongestTableName()
                '====================================================================
                '
'20121203 - v020 - Output positioning of TableInfo, use debug flag
    ' Output query sql to a text file
    ' Output table configuration to a text file
'20121203 - v019 - LongestFieldPropsName()
'20121201 - v018 -
'Move old research comments from basTESTaegitClass
'' RESEARCH:
'' Ref: http://stackoverflow.com/questions/47400/best-way-to-test-a-ms-access-application#70572
'' Ref: http://sourceforge.net/projects/vb-lite-unit/
'' Using VB 2008 to access a Microsoft Access .accdb database
'' Ref: http://boards.straightdope.com/sdmb/showthread.php?t=514884
'
'Public Function New_aegitClass() As aegitClass
'' Ref: http://support.microsoft.com/kb/555159#top
''===========================================================================================
'' Author:   Peter F. Ennis
'' Date:     March 3, 2011
'' Comment:  Instantiation of PublicNotCreatable aegitClass
'' Updated:  November 27, 2012
''           Added project to github and fixed aegitClassTest configuration for the new setup
''===========================================================================================
'    Set New_aegitClass = New aegitClass
'End Function
'
'Public Sub aegitClass_EarlyBinding()
''    Dim my_aegitSetup As aegitClassProvider.aegitClass
''    Set my_aegitSetup = aegitClassProvider.
''    anEmployee.Name = "Tushar Mehta"
''    MsgBox anEmployee.Name
'End Sub
'
'Public Sub aegitClass_LateBinding()
''    Dim anEmployee As Object
''    Set anEmployee = Application.Run("'g:\temp\class provider.xls'!new_clsEmployee")
''    anEmployee.Name = "Tushar Mehta"
''    MsgBox anEmployee.Name
'End Sub
'
' 20121129 - v017 - Pass Fail test results and debug output cleanup
    ' Working on documenting the tables and relations
' 20121128 - v016 - SourceFolder property updated to allow passing the path into the class
    ' Cleanup debug messages code, include GetReferences from aeladdin (tm)
    ' Public Function aeReadDocDatabase - does it need a Get property call to make the function Private? - fixed, set to Private
' 20121127 - v015 - delete mac1 from accdb and manually delete the S_mac1.def file as the data export does not delete files
    ' version number continues from zip files stored in OLD folder
    ' basChangeLog added, export with OASIS and commit new changes to github