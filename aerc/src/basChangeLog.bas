Option Compare Database
Option Explicit

' Problems:
' Modify ObjectCounts to provide more details and export results for development tracking, charting
' Fix table output field description to max of each table
' Document custom tabs - adaept sample tab displayed, but no output indication
' ReadDocDatabase debug output when custom test folder given
' Import of class source code into a new database creates a module
' http://www.trigeminal.com/usenet/usenet026.asp - Fix DISAMBIGUATION?
' http://access.mvps.org/access/modules/mdl0022.htm - test the References Wizard?
' http://stackoverflow.com/questions/2019605/why-does-msysnavpanegroupcategories-show-up-in-a-net-oledbprovider-initiated
'   This has useful information about using tdf.attributes
'
' Guides:
' Office VBA Basic Debugging Techniques
' Ref: http://pubs.logicalexpressions.com/pub0009/LPMArticle.asp?ID=410
' *** Ref: http://www.vb123.com/toolshed/02_accvb/remotequeries.htm - Remote Queries In Microsoft Access
' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/f8a050b9-3e12-465e-9448-36be59827581/vba-code-redirect-results-from-immediate-window-to-an-access-table-or-csv-file?forum=accessdev


'20140207 - v070 - Allow array of tables for xml data export to be provided via let property
    ' Office 2010: File > Options > Client Settings > General > Show add-in user interface errors
    ' Ref: http://msdn.microsoft.com/en-us/library/office/dd548010(v=office.12).aspx
    ' Ribbon development: By default, you won't see any errors if there are problems with the XML that you've defined
    ' for a ribbon customization. To display errors during development, be sure to set the Show add-in user interface errors
    ' option in the Advanced page of the Access Options dialog box. Without this option set, ribbon customizations
    ' may not load and it may not necessarily be clear why.
    ' There is no VBA setting for Show add-in user interface errors
    ' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/c3c54c12-6cbd-404a-8709-dba485f82377/access-2010-set-option-show-addin-user-interface-errors-via-vba?forum=accessdev
    ' Need to develop a registry settings test/output routine
    ' Add code to test for reg key "ReportAddinCustomUIErrors"
'20140206 - v069 - Load tlkpStates for test of xml output, from Census 1999 data
'20140205 - v068 - Xml output of tables to xsd
    ' Delete src\xml\* files on export
    ' KillProperly TryAgain for Err=70 permission denied
'20140203 - v067 - Tables and queries are the same container
    ' Ref: http://msdn.microsoft.com/en-us/library/office/bb177484(v=office.12).aspx
    ' Ref: http://www.office-archive.com/16-ms-access/8a3929131ae55e2e.htm
    ' Merge code - output of container properties for tables and queries
    ' Set varDebug for limiting output
    ' Force retry on Err=2220
'20140130 - v066 - Fix debug in ListOfAccessApplicationOptions
'20140129 - v065 - Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=247561 => Error 5 and 9
'20140123 - v064 - Move licence text after option statements so that it appears in a printout
    ' Merge changes for fix table output field description to max of each table and test
'20140121 - v063 - Fixing LoadRibbons
    ' Add error handling to aeReadWriteStream
    ' Fix aeDocumentTablesXML
'20140119 - v062 - ListOrCloseAllOpenQueries added to aegitClass
'20140117 - v061 - Debug output for longest field name etc.
'20140115 - v060 - Test for fixing table output field description to max of each table
'20140114 - v059 - OutputFieldLookupControlTypeList added to aegitClass
    ' Remove ExecSql from zTesting
'20140113 - v058 - FieldLookupControlTypeList outputs table, field, control type and count for lookup tab
    ' Office 2007 AcControlType Enumeration
    ' http://msdn.microsoft.com/en-us/library/office/bb225848(v=office.12).aspx
    ' ExportAllModulesToFile
    ' Ref: http://www.cpearson.com/excel/Enums.aspx - removed, use AcControlType instead
    ' Prep FieldLookupControlTypeList for inclusion in aegitClass
'20140112 - v057 - adaept sample tab displayed, but no output indication
    ' Research for documenting field lookup value list
    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=160994
    ' Ref: http://www.devhut.net/tag/ms-access-vba-programming/
    ' GetVBEDeatils
    ' Ref: http://www.devhut.net/tag/ms-access-vba-programming/page/2/
    ' Ref: http://www.functionx.com/vbaccess/Lesson25.htm
    ' *** Ref: http://www.accessmvp.com/twickerath/articles/multiuser.htm ***
    ' Ref: http://www.codeproject.com/Articles/380870/Microsoft-Access-Application-Development-Guideline
    ' Ref: http://www.eraserve.com/tutorials/MS_ACCESS_VBA_Get_Indexed_Fields.asp
    ' !!! Ref: http://allenbrowne.com/ser-27.html !!!
    ' "There is no safe, reliable way for users to add items to the Value List in the form without messing up the integrity of the data."
    ' The  Evils of Lookup Fields in Tables - Ref: http://access.mvps.org/access/lookupfields.htm
    ' Ref: http://www.utteraccess.com/forum/Lookup-fields-table-pro-t269783.html
    ' Ref: http://improvingsoftware.com/2009/10/02/blog-response-lookup-fields-in-access-are-evil/
'20140110 - v056 - Use latebinding everywhere
    ' Ref: http://www.granite.ab.ca/access/latebinding.htm
    ' DONE: Test for expected references when class first created OR fix to use late binding
    ' INFO: Say you want to display a list of reports available in your database to a user in one of your forms.  Simply add a combo-box to your form, then set the Row Source property as follows:
    '   SELECT [Name] FROM [MSysObjects] WHERE [Type] = -32764 AND Left([Name],1) <> "~" ORDER BY [Name]
    '   Ref: http://www.opengatesw.net/ms-access-tutorials/Access-Articles/Microsoft-Access-System-Tables.htm
'20131219 - v055 - Output GUID for project references
'20131219 - v054 - Note about need for reference to Microsoft Scripting Runtime
    ' Explicity define Dim dbs As Database
    ' Erl(0) Error 2950 if the ouput location does not exist so test for it first in aeDocumentTheDatabase
'20131025 - v053 - Improve message for err 70 KillAllFiles
'20131022 - v052 - Create KillAllFiles outside of aeDocumentTheDatabase with varDebug pass through parameter
    ' err 2220 can't open the file... still occurring with forms
    ' Use strTheCurrentPathAndFile in DocumentTheContainer
    ' Include test for err 70 permission denied in KillAllFiles and STOP with critical message
    ' DONE:
    ' Need proper fix for internationalization SaveAsText UTF-16 for Access 2013 and proper set of github to deal with it
    '   Ref: http://www.git-scm.com/docs/gitattributes.html and look at:
    '   If you want to interoperate with a source code management system that enforces end-of-line normalization,
    '   or you simply want all text files in your repository to be normalized, you should instead set the text attribute
    '   to "auto" for all files.
'20131003 - v051 - Add ListOfApplicationProperties to aegitClass
    ' Prepend Output... to some filenames
'20131003 - v050 - Bump
'20131003 - v049 - Ref: http://msdn.microsoft.com/en-us/library/gg435977(v=office.14).aspx
    ' Add table macro for output test
    ' Discontinued features and modified functionality in Access 2013
    ' Ref: http://office.microsoft.com/en-us/access-help/discontinued-features-and-modified-functionality-in-access-2013-HA102749226.aspx
    ' Ref: http://stackoverflow.com/questions/9206153/how-to-export-access-2010-data-macros
    ' Setting VBA Module Options Properly
    ' Ref: http://www.fmsinc.com/free/newtips/vba/Option/index.html
'20131002 - v048 - Add ListAccessApplicationOptions to aegitClass for options txt output
'20130920 - v047 - Add aeReadWriteStream to aegitClass to fix SaveAsText from UTF-16 to txt
    ' Keeps compatibility with older versions of Access
    ' Will need proper fix for internationalization
    ' Access 2013 is saving text as UTF-16 (FFFE at the start of file)
    ' This is a test sample for fixing it
    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=241996
    ' Add frm_Dummy, rpt_Dummy
'20130917 - v046 - If it exists then kill export.ini in class terminate
    ' Add qpt_Dummy as test, UTF2TXT_TestFunction added to deal with Access 2013 saving text as UTF-16 and resulting github diff challenges
'20130916 - v045 - Include sub ListAllHiddenQueries and ExportTheTableData
    ' Use OutputTo command as export.ini created with the export method
    ' Revert to TransferText - no header or text formatting
'20130906 - v044 - Start dev for creating a file with a list of hidden queries
'20130820 - v043 - Hidden attributes e.g. queries, not exported - needs fix
    ' *** Ref: http://stackoverflow.com/questions/10882317/get-list-of-queries-in-project-ms-access
'20130816 - v042 - err 2220 again. Query with "<" in the name always causes the error.
    ' Test creating a file with ">" in the name gives the OS error message:
    ' A filename can't contain any of the following characters:     \/:*?"<>|
    ' Ref: http://support.microsoft.com/kb/177506
    ' WONTFIX - Solution is to change the query name in the db or on exporting the code.
    ' Let the user do this in the db and fix it if needed. Source code export to file will follow the OS naming conventions.
'20130715 - v042 - Change location of source folder debug output in aeDocumentTheDatabase so that Stop will show details before active code runs
    ' Get diff and add this to v043
'20130711 - v041 - err 2220 again. Happens after new file copy and rename then Security Warning! with "Enable Content" button appears in VBA when opened
    ' and run code export. Related to the initial protection status?
    ' The error occurred a lot. Deleted contents of src folder and export again with no error! Test this again if future iterations and if it fixes the
    ' issue then code related to file deletion could be the problem. Consider a test and warning if the folder has content after deletion?
    ' Also seems related to file rename/compact and repair.
    ' Best current process: Rename file, open, enable content, update log/version, compact and repair, close then open the new version and export.
'20130708 -v0402 - Decrease pause to 0.25 secs. err 2220 possibly related to bad parameter in OpenForm command.
'20130708 -v0401 - Add and use Pause function. err 2220 appeared again. Increase pause to 0.5 secs.
'20130702 - v040 - Remove Stop for err 2220. Does not work with geh. Add WaitSeconds procedure.
'20130702 - v039 - Add Stop for err 2220 in Function DocumentTheContainer. See code comment.
'20130613 - v038 - Allow True/False option for CompactAndRepair
'20130320 - v037 - Add CompactAndRepair (tested for Access 2010) to aegitClass using SendKeys
'20130315 - v036 - Add Linked=> indicator to linked tables description in output file TblSetupForTables.txt
    ' Do not define zzz tables for Schema.txt export
'20130226 - v035 - Change Test folder references and variables to Import folder
    ' wsh.CurrentDirectory = aestrImportLocation
    ' Fix places where erl missing
    ' Get values for aestrSourceLocation, aestrImportLocation in aeReadDocdatabase
'20130225 - v034 - Add UseTestFolder with associated Let property to explicitly turn on the routines for import testing
'20130222 - v033 - Use MYPROJECT_TEST to run zzzaegitClassTest with a value for THE_SOURCE_FOLDER
    ' Ref: http://stackoverflow.com/questions/7907255/hide-access-options
    ' Add fix for properties output, Ref: http://www.granite.ab.ca/access/settingstartupoptions.htm
'20130215 - v032 - Merge ContainerObjectX proc into ListOfContainers
'20130212 - v031 - Add geh to aexists, BuildTheDirectory, aeDocumentTheDatabase, OutputBuiltInPropertiesText, OutputQueriesSqlText,
    ' aeDocumentRelations, aeDocumentTables, TableInfo, LongestFieldPropsName, LongestTableName, aeGetReferences
    ' Output ListOfContainers.txt
'20130118 - v030 - Create the WriteErrorToFile procedure
    ' Add global error handler to aeReadDocDatabase and test
    ' Output to aegitErrorLog.txt in My Documents
'20130117 - v029 - Add global error handler sample
    ' Ref: http://msdn.microsoft.com/en-us/library/office/ee358847(v=office.12).aspx#odc_ac2007_ta_ErrorHandlingAndDebuggingTipsForAccessVBAndVBA_WritingCodeForDebugging
'20121226 - v028 - Use DocumentTheContainer for Forms, Reports, Scripts (Macros), Modules - removes duplicate code
    ' FIXED - aeDocumentTheDatabase breakout cnt container operations into a function
    ' Delete all TEMP queries,  Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=160994
    ' Enumeration for SaveAsText, http://bytes.com/topic/access/answers/190534-saveastext-syntax
'20121219 - v027 - List or close all open queries added to test module
'20121218 - v026 - Fix for test folder
    ' FIXED - Fix error when tst folder not set, it is intended for import testing to recreate a database
    ' FIXED - Pass Fail of the tests should be associated to True False of the function, any error should return False
'20121209 - v025 - Output built in properties to text file
    ' Fix error 3251 in OutputBuiltInPropertiesText
'20121207 - v024 - bump version, compact and repair, create zip, add tag
    ' KillProperly
'20121206 - v023 - Fix for aeintFDLen < Len("DESCRIPTION")
    ' FIXED - Exists function debug output
    ' Error trap for LongestFieldPropsName
    ' Return boolean for DocumentTheDatabase to fix Test 1 result
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