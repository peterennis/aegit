Option Compare Database
Option Explicit

' Tools:
' MZ-Tools 8.0 for VBA - Ref: http://www.mztools.com/index.aspx
' TM VBA-Inspector - Ref: http://www.team-moeller.de/en/?Add-Ins:TM_VBA-Inspector
' RibbonX Visual Designer 2010 - Ref: http://www.andypope.info/vba/ribboneditor_2010.htm
' IDBE RibbonCreator 2016 (Office 2016) - Ref: http://www.ribboncreator2016.de/en/?Download
' V-Tools - Ref: http://www.skrol29.com/us/vtools.php
' Bill Mosca - Ref: http://www.thatlldoit.com/Pages/utilsaddins.aspx
' Rubberduck - Ref: https://github.com/rubberduck-vba/Rubberduck
' DataNumen Access Repair - Ref: https://www.datanumen.com/access-repair/
'
'
' Research:
' Ref: http://www.msoutlook.info/question/482 - officeUI-files
' The Ribbon and QAT settings - C:\Users\%username%\AppData\Local\Microsoft\Office
' Ref: http://msdn.microsoft.com/en-us/library/ee704589(v=office.14).aspx
' *** Windows API help - replacing As Any declaration
' Ref: http://allapi.mentalis.org/vbtutor/api1.shtml
' Ref: http://programmersheaven.com/discussion/237489/passing-an-array-as-an-optional-parameter
' *** Example of SQL INSERT / UPDATE using ADODB.Command and ADODB.Parameters objects
' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=219149
' *** CreateObject("System.Collections.ArrayList")
' Ref: http://www.ozgrid.com/forum/showthread.php?t=167349
' Microsoft Access - Really useful queries - Ref: http://www.sqlquery.com/Microsoft_Access_useful_queries.html
' Ref: http://www.micronetservices.com/manage_remote_backend_access_database.htm
' Microsoft Access Tips and Tricks - Ref: http://www.datagnostics.com/tips.html
'
'
' Guides:
' Office VBA Basic Debugging Techniques
' Ref: http://pubs.logicalexpressions.com/pub0009/LPMArticle.asp?ID=410
' *** Ref: http://www.vb123.com/toolshed/02_accvb/remotequeries.htm - Remote Queries In Microsoft Access
' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/f8a050b9-3e12-465e-9448-36be59827581/vba-code-redirect-results-from-immediate-window-to-an-access-table-or-csv-file?forum=accessdev
' Access Articles- Ref: http://www.databasejournal.com/article.php/1464721
' Long Binary Data - Ref: http://www.ammara.com/support/technologies/long-binary-data.html
'
'
'=============================================================================================================================
' Tasks:
' %140 -
' %139 -
' %128 - Add snips to doc for Webspeak/Newspeak links as local reference
' %127 - exp tables should respect ODBC flag - NOTE: be careful of export sql table data via this method, only use for test data
' %126 - Add Export folder exp, integrate with aegit class
' %125 - Update basGDIPlus with latest code from aeGDIPlusDemo
' %124 - USysApplicationLog Table - Ref: https://www.amazon.com/Microsoft-Access-2010-Inside-Out/dp/0735626855#reader_0735626855
'           The USysApplicationLog table is used to record any data macro execution errors - Ref: http://www.accessjunkie.com/Pages/faq2010_32.aspx
' %123 - VBA How to Hide a table - http://hitechcoach.com/index.php/component/content/article/61-access-databases/tables/63-how-to-hide-a-table
' %122 - USysRegInfo is a table that is normally created in databases intended as add-ins, to be managed by the Access Add-In Manager
' %121 - Fix needed for export like tables_~TMPCLP132461.xsd
' %120 - Relates to %116, Create function to output list of MSys tables then work through steps to understand conditions when they are created
' %119 - Use HashAllModules for hash signed release of 2.0.0 on GitHhub
' %118 - Use HashAllModules for code output hash signed export files
' %116 - MSysIMEXColumns and MSysIMEXSpecs - These two tables contain information about any Import/Export Specifications you have created in Access.
'           The MSysIMEXSpecs contains general information about the specification, while MSysIMEXColumns includes the column mapping for each specification.
'           Ref: http://www.opengatesw.net/ms-access-tutorials/Access-Articles/Microsoft-Access-System-Tables.htm
'           Ref: http://stackoverflow.com/questions/143420/how-can-i-modify-a-saved-microsoft-access-2007-or-2010-import-specification
'           Ref: http://www.blueclaw-db.com/export-specifications.htm
'           DoCmd.RunSavedImportExport Method - Ref: https://msdn.microsoft.com/en-us/library/office/ff834375.aspx
'           DoCmd.TransferText Method - Ref: https://msdn.microsoft.com/en-us/library/office/ff835958.aspx
'           *** Research: VBA Manipulation Of Import/Export Specification - Ref: http://www.utteraccess.com/forum/VBA-Manipulation-Saved-E-t1990584.html ***
' %115 - MSysAccessXML ??? - Ref: http://www.databasejournal.com/features/msaccess/article.php/3528491/Use-System-Tables-to-Manage-Objects.htm
' %114 - How to clean up MSysObjects, Ref: http://www.office-forums.com/threads/how-to-clean-up-msysobjects.681442/ - MSysComplexColumns holds information and references about the "Attachments" type field, and Multi-Valued-Fields
'           Follow up with this technical discussion - Ref: https://social.msdn.microsoft.com/Forums/office/en-US/f3fc50e9-0c2d-45fc-b9fa-bc3b852cda6b/is-it-possible-to-get-at-the-table-underlying-an-attachment-field?forum=accessdev
' %112 - Test ExportXML method, allows developers to export XML data, schemas, and presentation information from ... or the Microsoft Access database engine, Ref: https://msdn.microsoft.com/en-us/library/office/ff193212.aspx
' %101 - Importing and Exporting XML Data Using Microsoft Access, Ref: https://technet.microsoft.com/en-us/library/ee692914.aspx
' %100 - Automation Error on CreateObject("System.Collections.ArrayList"), install .Net 3.5 SP1 (includes .NET 2) for W10, Ref: https://social.msdn.microsoft.com/Forums/sqlserver/en-US/9bfcd001-5168-4cff-b2ba-6b8e8d465138/excel-2010-vb-runtime-error-2146232576-80131700-automation-error-on?forum=exceldev
' %095 - Set minimum support to 2010 - Ref: http://stackoverflow.com/documentation/vba/3364/conditional-compilation/11558/using-declare-imports-that-work-on-all-versions-of-office#t=201607251720494878701, Rubberduck
' %093 - Create a results folder to hold test output and configure the test system to use it
' %089 - Add timing and test ouptut to a file; identify PC, OS, Office
' %087 - Allow setup options to be loaded from a file. Consider yaml, json, other
' %086 - Improve debug ouput, provide trace levels and logging to file. Ref: http://pixcels.nl/debug-and-trace-in-vba/
' %085 - Automation goal - Ref: https://github.com/blog/1271-how-we-ship-github-for-windows
' %084 - On Error Goto 0 - Ref: http://www.peterssoftware.com/t_fixerr.htm (k)
' %062 - Fix backend testing for existence of srcbe, srcbe/xml, srcbe/xmldata folders
' %057 - MS Access Control Property.Type not making sense, Ref: http://stackoverflow.com/questions/27682177/ms-access-control-property-type-not-making-sense
' %052 - Create property to define text encoding output
' %051 - UTF-16 to UTF-8, http://www.di-mgt.com.au/howto-convert-vba-unicode-to-utf8.html, David Ireland
'           Internationalization - Potential source issue #033
'           Ref: http://www.vb-helper.com/tip_internationalization.html
'           *** Ref: http://blog.nkadesign.com/2013/vba-unicode-strings-and-the-windows-api/
'           *** Ref: http://accessblog.net/2007/06/how-to-write-out-unicode-text-files-in.html
'           *** Ref: https://github.com/timabell/msaccess-vcs-integration/blob/master/MSAccess-VCS/VCS_IE_Functions.bas
' %041 - Relates to %039, %040, Create Set property so that mblnUTF16 is not Const and can be changed outside of the aegit class
' %007 - Make varDebug work as optional parameter to Let property
' Issues:
' #040 - Picture for command button is stored in MSysResources, include option to export the records of this table
'=============================================================================================================================
'
'
'20170727 - v200 -
    ' FIXED - %138 - Zip up v200, put in zips and commit to GitHub
    ' FIXED - %137 - Decompile the accdb to minimal size
    ' FIXED - %136 - Bump  to version 2.0.0 as it looks good!
'20170727 - v19916 -
    ' FIXED - %135 - Fix export tool for separate procedure queries for more detailed timing output and messages
    ' FIXED - %134 - Create Setup_The_Source_Location
'20170630 - v19915 -
    ' FIXED - %133 - aestrXMLDataLocationBe is not converted correctly from relative to absolute path (SVIPAZSQL011.PNG), TestForRelativePath (SVIPAZSQL012.PNG)
'20170616 - v19914 -
    ' FIXED - %132 - Add report txt export capability
'20170509 - v19913 -
    ' FIXED - %131 - Do not export unwanted temp tabes e.g. src\xml\tables_~TMPCLP117951.xsd
'20170411 - v19912 -
    ' FIXED - %130 - Add RenameSQLinkedTable
'20170324 - v19911 -
    ' FIXED - %129 - Mark development testing export types as EXPERIMENTAL
'20161107 - v19905 -
    ' FIXED - %117 - Make schema output optional and default is false
    ' FIXED - %113 - Lovefield export unknown type Hyperlink, Attachment, Currency - from Customers; Orders; Order Details - tables imported from Northwind 2007 for testing
    ' ONHOLD - %096 - Relates to %117, Bug - CreateDb schema is treating relationships as an index. Need different SQL for the relationships.
    ' ONHOLD - %078 - Relates to %117, Parse output error for table aeItems, it has no index or primary key and is missing the semicolon for LF creation
    ' ONHOLD - %071 - Relates to %117, Add varDebug for schema output debugging
    ' ONHOLD - %070 - Relates to %117, Create multi field index sample then fix access and lf schema output; tblDummy3 is test case, Relates to %080, %098
'20161007 - v19901 -
    ' FIXED - %110 - aeDescribeIndexField should also return the index name - this is useful also for ODBC linked SQL Server tables
'20161007 - v199 -
    ' FIXED - %111 - dbo_studentAttendances index studentId:date shows Iii instead of Ii
    ' FIXED - %109 - Schema export error for primary and index when field order changed in tables. Need to remember the field names and pass to create db schema. Access will sort alphabetically by index name.
    ' FIXED - %099 - Multi field index statement, Ref: https://msdn.microsoft.com/en-us/library/office/ff823109(v=office.15).aspx
    ' FIXED - %098 - tblDummy3 shows incorrect output for primary key; Use modified version of DescribeIndexField from Allan Browne; Relates to %099
    ' FIXED - %094 - Comment out debug print statements when varDebug is not used
    ' WONTFIX - %090 - Consider using build number
    ' FIXED - (Rubberduck) %083 - Access closing down and restarting, Ref: http://answers.microsoft.com/en-us/office/forum/office_2010-access/access-wont-shut-downkeeps-restarting/b8295bca-bfc8-4b59-8747-a609f3ba466b?auth=1
'20160925 - v199 -
    ' FIXED - %108 - Relates to %105, Create function to check for linked ODBC tables
    ' FIXED - %107 - Relates to %108, Export spinning after third KillAllFiles, investigate and resolve - related to ODBC linked tables
    ' FIXED - %105 - Set option to export all tables info or only Access tables, i.e. skip ODBC linked tables is possible
'20160907 - v196 -
    ' FIXED - %106 - Do not show ~TMP* or MSys* tables in OutputListOfTables
    ' FIXED - %104 - Parse field string fails when there are spaces in the name
    ' FIXED - %103 - Get DbIssueChecker.mdb from Allen Browne and use it to identify issues
    ' FIXED - %097 - Pass the actual value "varDebug" for debugging in all cases so potentially strange variant parameters will not be transferred
    ' NOTABUG - %081 - Cannot redim dimensioned array (Inefficient file reading, Ref: https://github.com/rubberduck-vba/Rubberduck/issues/2004)
    ' DUPLICATE %080 - Relates to %098, Wrong output for tblDummy3 in OutputListOfIndexes.txt and therefore OutputLovefieldSchema.txt is incorrect
    ' OBSOLETE - #042 - Global Error Handler Routines, Ref: http://msdn.microsoft.com/en-us/library/office/ee358847(v=office.12).aspx#odc_ac2007_ta_ErrorHandlingAndDebuggingTipsForAccessVBAndVBA_WritingCodeForDebugging
    ' FIXED - #039 - x64 support - https://github.com/peterennis/aegit/issues/3
'20160831 - v193 -
    ' FIXED - %102 - Add extra Lovefield types to match SQL Database tables from Azure
    ' FIXED - %061 - Fix OutputCatalogUserCreatedObjects so that it does not need to leave zzz query
'20160724 - v190 -
    ' FIXED - %092 - Add relationship example from MONDIAL (no data), Ref: http://www.dbis.informatik.uni-goettingen.de/Mondial/
    '           Ref: http://databases.about.com/od/sampleaccessdatabases/a/Microsoft-Access-Sample-Database-Countries-Cities-And-Provinces.htm
    ' FIXED - %091 - Update aegitClassTest to be one export, move tests to module TestsCodeForA~Z
'20160716 - v187 -
    ' FIXED - %088 - Use timing, Ref: https://bytes.com/topic/access/insights/618175-timegettime-vs-timer
    '           NOTE: Timer() rolls over every 24 hours. timeGetTime() keeps on ticking for up to 49 days before it resets the returned tick count to 0 => Use in Excel chart player?
    ' OBSOLETE - %082 - aegit drowns Rubberduck, Ref: https://github.com/rubberduck-vba/Rubberduck/issues/2018
'20160713 - v185 -
    ' WONTFIX - %044 - Generalizing Form Behavior e.g. global error handling ??? Ref: http://www.dymeng.com/techblog/generalizing-form-behavior-refined/
    '           Not needed for aegit but possibly useful for other forms development work
'20160707 - v180 -
    ' FIXED - %079 - Create tblDummy3 with multi field primary key and multi field index
    ' FIXED - %077 - Fix addPrimaryKey parsing for Lovefield schema
    ' FIXED - %076 - OutputContainerTablesProperties remove GUID number and replace with text "GUID"
    ' FIXED - %075 - Change mPRIMARYKEY to not include name of primary key in parsing
    ' FIXED - %074 - Add State Name index for aetlkpStates, change primary key name to ID for aetblThisTableHasSomeReallyLongNameButItCouldBeMuchLonger
'20160706 - v179 -
    ' FIXED - %073 - Unknown Access Field Type in GetLovefieldType, strAccessFieldType=Double => running on Access 2013
    ' FIXED - %072 - Unknown Access Field Type in GetLovefieldType, strAccessFieldType=Integer => running on Access 2013
'20160705 - v178 -
    ' FIXED - %068 - Creation of schema adds index to wrong field
'20160705 - v177 -
    ' FIXED - %069 - Fix name case of id for indexes, remove superfluous reference to index on id when it is already the PrimaryKey
'20160622 - v171 -
    ' FIXED - %067 - Wrong field picked up for primary key
    ' FIXED - %066 - Syntax error for OLE Object
    ' FIXED - %065 - CreateDb script (OutputTheSchemaFile) fail when SQL string too long
    ' FIXED - %064 - CreateDb script (OutputTheSchemaFile) missing Erl
'20160505 - v168 -
    ' FIXED - %063 - Trap Error 31532 in aeDocumentTablesXML when the linked Azure tables are not available
'20160126 - v166 -
    ' OBSOLETE - %028 - Relates to %020, Compare export time for standalone vs. flag set for split
    ' FIXED - #002 - v093 using export aegit 1.6.5 deletes files in root folder of C:\PETER\SVIP\SVIPDB
'20151231 - v164 -
    ' OBSOLETE - IDBE RibbonCreator 2013 (Office 2013) - Ref: http://www.ribboncreator2013.de/en/?Download
    ' OBSOLETE - MZ-Tools 3.0 for VBA - Ref: http://www.mztools.com/v3/download.aspx
    ' OBSOLETE? - %058 - Windows 10, GetEdition in aeVer reports Office Home and Student Guid and but Office 2010 Pro is installed
    ' WONTFIX - %043 - Enable communication between VBA and HTML5/JavaScript via the Access 2010+ native Web Browser control - Ref: http://www.dymeng.com/browseEmbed/
    '           Ref: http://www.dymeng.com/techblog/browseembed-html5javascript-for-your-access-projects/
    ' WONTFIX - %035 - Relates to #004, Integrate with baem - Ref: https://www.youtube.com/watch?v=960UNEiOdTo, research media players
    ' FIXED - #041 - Output list of visible/hidden macros
    ' FIXED - #026 - Output list of visible/hidden reports
    ' FIXED - #025 - Output list of visible/hidden modules
    ' FIXED - #024 - Output list of visible/hidden forms
    ' FIXED - #023 - Output list of visible/hidden tables
    ' OBSOLETE - #012 - Document custom tabs - adaept sample tab displayed, but no output indication => Not clear. Need more detail.
'20151222 - v161 -
    ' FIXED - %060 - Fix output list of forms so that it does not need to use a temp table
    ' FIXED - %059 - Database opened exclusive error when exporting hidden queries, Ref: http://stackoverflow.com/questions/18121099/trying-to-read-data-out-of-msysobjects-with-odbc-in-c-but-getting-no-permissio
'20151013 - v158 - Bump
'20150816 - v153 -
    ' FIXED - %056 - Control can't be edited; it's bound to a replication system column - OutputListOfAllHiddenQueries
    '           Ref: https://groups.google.com/forum/#!topic/microsoft.public.access.replication/-78b5ZRqPCQ
    '           Creating the target table first and then runing the query as an append query seems to have solved the problem
    ' OBSOLETE - #021 - Relates to %051, %052 - Caption ="Gr??e" - Language display problem on output - GDIPlusDemo | possibly locale/UTF-16 related - TBD
    ' OBSOLETE - #010 - Check if ViewAppProperties includes anything new - not sure what this is referring to anymore
'20150807 - v150 -
    ' FIXED - %055 - Take care of Err 53 File not found in KillProperly
    ' FIXED - %030 - Some Output* files need to send results to srcbe when back end is exported
'20150806 - v149 -
    ' FIXED - %054 - Ignore internal date changes for export
    ' FIXED - %053 - Files stored in srcbe are deleted, fix config so that KillAllFiles is intelligent about src vs. srcbe
'20150804 - v148 -
    ' FIXED - %050 - Error 58 (File already exists) USysRibbons.xml in OutputTheTableDataAsXML - xml folder contents has not been deleted
    ' FIXED - %015 - Set default table font Ref: http://superuser.com/questions/416860/how-can-i-change-the-default-datasheet-font-in-ms-access-2010
    ' OBSOLETE - Sendkeys module Ref: http://www.codeguru.com/vb/gen/vb_system/keyboard/article.php/c14629/SendKeys.htm#page-1 => VB6, needs too many changes
    ' OBSOLETE - SendInput Module Ref: http://vb.mvps.org/samples/SendInput/
    ' OBSOLETE - Relates to %043, *** Ref: http://blogs.office.com/2013/01/22/visualize-your-access-2013-web-app-data-in-excel/ - Visualize your Access 2013 web app data in Excel
    ' FIXED - Ref: http://blogs.office.com/2013/07/02/the-access-2013-runtime-now-available-for-download/
    '           Easy test for db app with Access runtime without having to install the runtime - start Access with the runtime switch or better yet,
    '           just rename your ACCDB file to an ACCDR. When you double click on the ACCDR file it will start the Access client in the runtime mode.
'20150804 - v147 -
    ' FIXED - %049 - NameMap for azure linked tables shows strange characters - treat like GUID
    ' FIXED - %048 - Test for existence of _frmPersist
    ' FIXED - %047 - XML table info not exported for aegit in v145 since relative path introduced - Reinstate Test 7 fixes it
    ' WONTFIX - %002 - Ref: http://access.mvps.org/access/modules/mdl0022.htm - test the References Wizard?
    ' WONTFIX - #035 - Configure git with diff for UTF16 files - GitHub diff now seems capable to deal with it
    '           e.g. http://blog.xk72.com/post/31456986659/diff-strings-files-in-git
    '           but... Ref: http://gigliwood.com/blog/to-hell-with-utf-16-strings.html
    '           and... Ref: https://coderwall.com/p/yka9da/better-diffs-with-sql-files
    '           Reopen if internationalization becomes an issue
    '           Consider conversion to UTF8, Ref: https://github.com/timabell/msaccess-vcs-integration/commit/82e56a4df23b74cc57b7c4fd353babadd96c0ed4
    ' WONTFIX - #017 - Relates to %035, %043, KPI chart test not working
    ' OBSOLETE - #011 - Modify ObjectCounts to provide more details and export results for development tracking, charting
'20150803 - v146 -
    ' FIXED - Error 76 for relative path on xml export
    ' FIXED - Error 2220 on xml export
    ' FIXED - %046 - Test NoBOM stream writing - NoBOM seems to be much faster
    ' FIXED - %045 - Erl=170, Err=76, Path not found, OutputListOfApplicationOptions
    ' FIXED - %042 - Relates to %046, NOTE: Use of mblnUTF16 really slows down export - investigate faster option than read/write stream
    ' FIXED - %039 - BOM, UTF-8, UTF-16, Access Export - Ref: http://axlr8r.blogspot.nl/2011/05/how-to-export-data-into-utf-8-without.html
    '           Ref: http://blog.nkadesign.com/2013/vba-unicode-strings-and-the-windows-api/
    '           Mojibake - Ref: https://en.wikipedia.org/wiki/Mojibake
    ' CLOSED - %023 - Access source control options Ref: http://stackoverflow.com/questions/187506/how-do-you-use-version-control-with-access-development
    ' CLOSED - %012 - https://support.office.com/en-za/article/Discontinued-features-and-modified-functionality-in-Access-2013-bc006fc3-5b48-499e-8c7d-9a2dfef68e2f
'20150730 - v143 -
    ' FIXED - %043 - Relates to %042, Keep connection(?) open with frmPersist method may be a solution
    '           No performance improvement for single dbs export
    '           NOTE: GitHub seems to intelligently run diff when UTF-16 files are committed and recognize ansi, so leaving mblnUTF16 as True could be a workaround, allowing international files - TBD
    ' FIXED - %040 - Reinstate removal of BOM (FF FE) from exported *.txt form files - consider Ref: http://www.experts-exchange.com/Programming/Languages/Visual_Basic/VB_Script/Q_25105941.html
    ' FIXED - %038 - Load to aegit_Template and test %037
    ' FIXED - %037 - Relates to %009, Allow relative path
    ' FIXED - %036 - Create help page with GitHub MarkDown - Ref: https://github.com/peterennis/aegit_Template
    ' FIXED - %009 - Move "Default Usage" and "Custom Usage" from test module - done in aegit_Template
    ' WONTFIX - #004 - Change the color of each interior (histogram) chart vba access - Ref: http://stackoverflow.com/questions/16819859/change-the-color-of-each-interior-histogram-chart-vba-access
    '           Work with bemb/baemb for JS charts integration
    ' WONTFIX - %003 - Ref: http://www.trigeminal.com/usenet/usenet026.asp - Fix DISAMBIGUATION? - Reopen if it ever becomes a problem
'20150717 - v141 -
    ' FIXED - %034 - Set value of backend to "NONE" to allow single user dbs other than aegit
    ' WONTFIX - Leave control with the developer - %029 - Split db export use causes exclusive lock, faster export but requires access restart
    '           Consider automated restart here - Ref: http://blog.nkadesign.com/2008/ms-access-restarting-the-database-programmatically/
    ' FIXED - %026 - Remove Last Updated ouput from table properties, it is just noise - OutputCatalogUserCreatedObjects.txt
    ' FIXED - %018 - With split db add export tool to back end, e.g. save to srcbe, test
    ' OBSOLETE - %017 - Linked tables still hang on output, use test and only export linked tables as tblName.Linked.txt
    '           Ref: http://p2p.wrox.com/access-vba/37117-finding-linked-tables.html
    ' FIXED - %016 - Create OutputCatalogUserCreatedObjects as a text file list with all objects
    ' WONTFIX - This project focus is access export for source control - %014 - Set default forms, report, database Ref: http://allenbrowne.com/ser-43.html
    ' WONTFIX - Relates to %035, #005 - How to Format Your Graphs Using Visual Basic for Microsoft Access - Ref: http://www.brighthub.com/computing/windows-platform/articles/116946.aspx#imgn_1
'20150717 - v140 -
    ' FIXED - %033 - Add Properties_ to the table names when exporting properties
'20150714 - v139 -
    ' FIXED - %032 - Do not output DateCreated or LastUpdated when exporting table properties
    ' FIXED - %031 - Output the tables properties
'20150508 - v136 -
    ' FIXED - %027 - Hidden queries output is listing all queries for split db. Possible error in IsQryHidden function
    ' FIXED - %025 - Relates to %024, Error 3011, Could not find object 'zzzTmpTblQueries' => it is in the front end and not the back end
    ' FIXED - %024 - Error 3167, Record is deleted. OutputListOfAllHiddenQueries
    ' FIXED - %022 - Set size of MRU list to 'Not Tracked' for export, it is noise
    ' WONTFIX - %021 - Create Let property to transfer password for code export when back end is encrypted
    '           Develop without password then add it at deployment when possible, reopen if special case needs it
    ' FIXED - %020 - Relates to %010, Add code for bypass of OpenAllDatabases routine when using standalone (not split) database
    ' RESOLVED - %013 - Ref: http://architects.dzone.com/articles/20-database-design-best
    ' FIXED - %010 - Relates to %020, Export locks up easily when linked tables are on a network drive - need to fail gracefully and/or give warning
'20150505 - v133 -
    ' FIXED - %019 - Fix GetLinkedTableCurrentPath to not include password in link result on export
'20150303 - v129
    ' FIXED - #038 - Office 2013 export takes 5 minutes for ITILRDA at OutputListOfContainers
    ' CLOSED - Ref: https://github.com/peterennis/aegit/issues/1 - Thanks Jason Zhu!
    ' Change Test2: to look for aegit_expClass so that it will pass (class was renamed and test not updated)
'20150204 - v127 - Fix error on export when USysRibbons does not exist
    ' FIXED - %011 - Add an exclusion list of objects other than zzz marker, useful so that aegit objects do not need renaming
'20150123 - v126 - CLOSED #036 - USysRibbons content export
    ' Expanded row height of USysRibbons shows the formatted data text contents
    ' FIXED - #037 - Adding USysRibbons form and adjusting xml causes the ribbon bar logo to disappear
    ' CLOSED - #029 - Error 3270 Property not found - in OutputListOfAccessApplicationOptions - occurs if "Break on All Errors" is set
'20150121 - v125 - USysRibbons.xml export
    ' OutputTableDataAsFormattedText "USysRibbons"
    ' Ref: http://www.dbforums.com/showthread.php?1043254-Is-there-a-way-to-export-a-access-table-to-a-text-file-with-VBA
    ' Add TransferText command to OutputTableDataAsFormattedText to bypass text cutoff problem
'20150113 - v125 - Testing with excluding rename of aegit files to zzz* not needed
'20140930 - v123
    ' WONTFIX - #033 - OutputListOfCommandBarIDs showing FS in Notepad++ followed by ???? in descriptions - Access 2010
    ' The code related to #033 is flakey so set option to turn off this output by default
'20140805 - v122
    ' Raymond reported lockup when backend is on network. Added task %010 - workaround is dev/link with local tables.
'20140709 - v122 - Bump
'20140709 - v121 - Remove test 8, not used
    ' Remove test 3, not used
    ' Move unused procedures from the export test module to zTesting2
    ' Update description for running the export
    ' Fix results for Not Used tests
'20140708 - v120 - Bump
'20140708 - v119 - Reorganize code for optional export settings
'20140708 - v118 - FIXED - %008 - Implement varDebug in export procedure
    ' Add test for ExportQAT. Implement Let property for ExportQAT
    ' FIXED - #036 - Error with Debug.Print UBound(gvarMyTablesForExportToXML) when debug is on. Variable not initialized.
    ' CLOSED - #022 - Ref: http://www.hanselman.com/blog/YoureJustAnotherCarriageReturnLineFeedInTheWall.aspx
    ' FIXED - #008 - Export QAT - Ref: http://winaero.com/blog/how-to-make-a-backup-of-your-quick-access-toolbar-settings-in-windows-8-1/
'20140707 - v117 - Bump
    ' Array not initialized error Ref: http://www.vbforums.com/showthread.php?654880-How-do-I-tell-if-an-array-is-quot-empty-quot&highlight=array+initialised
'20140701 - v116 - Code tidy
    ' Create aegitExport of type myExportType to configure optional outputs
'20140627 - v115
    ' CLOSED - %005 - Ref: http://stackoverflow.com/questions/3313561/what-are-the-limitations-of-git-on-windows
    ' Research for Let property with optional parameter - Ref: ' Ref: http://www.ozgrid.com/forum/showthread.php?t=82338
'20140624 - v114 - FIXED - #034 - QAT not exported for aegit
'20140623 - v113 - Clean up unused error handler variables
'20140620 - v112 - OutputTheQAT message for Err 3270 when no AppTitle is set
    ' Ref: http://bytes.com/topic/access/answers/205495-late-binding
    ' FIXED - #032 - Late Binding for OutputListOfCommandBarIDs
    ' FIXED - ONGOING - %001 - Fix hints from TM VBA-Inspector and track progress
    ' WONTFIX - Reopen if development on import continues - #014 - ReadDocDatabase debug output when custom test folder given - applies to aegit_impClass
    ' WONTFIX - Reopen if development on import continues - #013 - Import of class source code into a new database creates a module - applies to aegit_impClass
    ' OLD - Ref: http://www.codematic.net/excel-development/excel-xll/excel-xll.htm => Look at XLW Ref: http://xlw.sourceforge.net/
'20140618 - v110 - Make all comments of the form '<space> so that ' followed by no space is code
    ' that has been commented out.
    ' FIXED - #027 - Cannot run the macro or callback function OnActionButton
    ' Change MsgBox for avarTableNames to debug statement
    ' FIXED - #031 - Sort OutputListOfCommandBarIDs.txt
'20140612 - v109 - Remove old code related to Global Error Handler
    ' OutputListOfCommandIDs
    ' Access Basics by Crystal - Good intro for beginners
    ' Ref: http://allenbrowne.com/casu-22.html
    ' Ref: Conrad Systems Development - http://www.accessmvp.com/JConrad/index.html
'20140611 - v108 - Testing for #001
    ' FIXED - #001 - Property Let TablesExportToXML(ByVal varTablesArray As Variant)
    ' FIXED - #009 - Create Let property for setting aegitExportDataToXML
'20140604 - v105 - DONE - %006 - Add adaept ribbon
    ' Remove module aeribpng - basGDIPlus will be used
    ' Add web location to aegit_expClass - https://github.com/peterennis/aegit
    ' Old Research: Convert Access 2007 forms to 97 - Ref: http://www.esotechnica.co.uk/2011/02/convert-access-2007-forms-to-97/
    ' Old Research: Ref: http://stackoverflow.com/questions/2019605/why-does-msysnavpanegroupcategories-show-up-in-a-net-oledbprovider-initiated
    ' This has useful information about using tdf.attributes
    ' Remove "If aeDEBUG_PRINT Then " globally
    ' FIXED - #028 - Err 55 file already open in WriteStringToFile
    ' DONE - %004 - Learning git on the command line - Ref: http://cheat.errtheblog.com/s/git
    ' FIXED - #030 - Testing GDIPlusDemo2013 - Error -2147319779 (Multiple-step OLE DB operation generated errors, Check
    ' each OLE DB status value, if available. No work was done.) in procedure
    ' aeGetReferences of Class aegit_expClass
    ' => Add reference to MSCOMCTL.OCX, compile, remove reference - fixes ActiveX registration problem in GDIPlusDemo2013 forms.
'20140530 - v104 - Import basGDIPlus
    ' Add basaeRibbonCallbacks modules
    ' Load simple ADAEPT ribbon using GDIPlus and example from Avenius IDBE RibbonCreator 2013
    ' 32x32 PNG Logo stored in tblBinary
'20140529 - v103 - Add to issues list
    ' Ref: http://blog.vishalon.net/index.php/change-ms-access-application-title-and-icon-using-vba/
'20140525 - v102 - FIXED #020 - Run-time error 3011 at IsQryHidden when testing GDIPlus
'20140523 - v100 - FIXED #019 - Testing GDIPlus module showed need to separate dev of aegit_exp and aegit_imp classes
'20140523 - v099 - #021 International encoding - Ref: http://stackoverflow.com/questions/8038729/github-using-utf-8-encoding-for-files
    ' s/gcfHandleErrors/mblnHandleErrors/g - It is only used in the class and is not global
    ' s/gblnOutputPrinterInfo/mblnOutputPrinterInfo/g - It is only used in the class and is not global
    ' Set mblnUTF16=True to force UTF16 output for testing i18n
    ' Using TortoiseGit diff opens TortoiseGitMerge with the message "The text is identical, but the files do not match!
    ' The following differences were found: Encoding (ASCII, UTF-16LE BOM)"
    ' FIXED #006 - Rewrite UTF-16 files to standard text as optional
'20140523 - v098 - #020 showed output like strQueryName=~sq_ffrmImages_35BF4C8896444268BB942DACE45FF252
    ' MSComctlLib (Microsoft Windows Common Controls 6.0 (SP6))causes error in aeGetReferences
    ' Work Around - Remove the reference
'20140424 - v097 - Add paramater to PopCallStack to trace #018 - problem is in LongestTableName
    ' LongestTableName used in class initialize - give default value 11 instead
    ' LongestFieldPropsName raises Error 9
    ' FIXED - #018 - Error 9, subscript out of range in PopCallStack
    ' Delete some old code, add varDebug to PrettyXML
    ' Fix varDebug for OutputListOfContainers
'20140418 - v096 - Bump
    ' Use aegit_expVERSION in the class and not gstrVERSION
'20140410 - v095 - Big charts test, KPI chart test
    ' Chart big, HasModule set to no
'20140407 - v094 -
    ' Ref: http://www.vbmigration.com/whitepapers/apicalls.aspx
    ' Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destAddress As Long, ByVal destAddress As Long, ByVal numBytes As Long)
    ' The solution does not work  due to HOSTENT type, test with overloading is a possibility, too much work
    ' ? WONTFIX - #015 - Declare statement does not support parameters of type 'As Any' - Ref: http://msdn.microsoft.com/en-us/library/wccc9bx3(v=vs.71).aspx
    ' Table field types missing, Ref: http://allenbrowne.com/ser-49.html
    ' Update app title bar with version when export is run
    ' Fixed #016 - Unknown field type message 104 x1, 101 x3 for aetrak test in OutputTheSchemaFile
'20140404 - v093 - Use aegit_expClass, aegit_impClass, aeDEBUG_PRINT
    ' Reorganize changelog tasks, issues, research, use Right$
    ' Use chart enum
    ' Fixed #003 - Changing The Microsoft Access Graph Type - Ref: http://www.vb123.com.au/toolbox/99_graphs/msgraph1.htm
    ' Verified fixed in an earlier version #007 - Fix table output field description to max of each table
'20140403 - v092 -
    ' Research on #002 - Ref: http://answers.microsoft.com/en-us/office/forum/office_2010-access/graphs-crashing-microsoft-office-2010-component/9196aa27-ee0f-426d-bb7f-5c6e8858f6de
    ' Ref: http://answers.microsoft.com/en-us/office/forum/office_2013_release-access/access-crashes-when-editing-pie-chart/c3178d1f-91a8-4dfd-98b3-86c5465546ec
    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=245033
    ' Ref: http://www.accessforums.net/queries/perplexing-scatter-chart-x-axis-problem-18287.html
    ' NT Command Script and Documented Steps to Decompile / Compact / Compile an Access DB
    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=219948
    '   Suggested Decompile Steps from Michael D Lueck, mlueck@lueckdatasystems.com
    '   1) The database should be in the same directory on the C: drive as the decompile.cmd script attached. Update the script as necessary to make correct for your working environment.
    '   2) Run the decompile.cmd script which will start Access, the database, and Access will decompile it.
    '   3) Next close the database – not all of Access.
    '   4) Then reopen the database. (Remember the shift key if you have an autoexec macro!)
    '   5) Compact the database at this point (MS Office icon \ Manage \ Compact and Repair) (Remember the shift key if you have an autoexec macro!)
    '   6) Press Ctrl+G to open the VBA window
    '   7) Click the Debug menu \ Clear All Breakpoints
    '   8) Click the Debug menu \ Compile - ONLY do this step the FIRST time!
    '   9) Then Compact again as in Step 5 (Remember the shift key if you have an autoexec macro!)
    '   10) Completely exit Access (Remember the shift key if you have an autoexec macro!)
    '   When the before / after file size are finally the same, the decompile.cmd script will end.
    '   Note: I found that only doing step 8) the first time through results in a completely decompiled database.
    '   This is much smaller (in my case an FE DB compiled of 18MB reduces to 13MB completely decompiled) and also completely avoids the nasty "big bang upgrade"
    '   encountered when crossing the Windows 7 SP0 to Windows 7 SP 1 / Office 2010 SP0 to Office 2010 SP1 divide.
    ' Fixed #002 - v091 Crash on edit chart - run decompile and test
'20140402 - v091 - Run TM VB-Inspector => 1460 hints in 19 objects
    ' Verify then turn off "Forgotten command:'Stop'"
    ' Fix 'Select Case' without 'Case Else'
    ' Add aeDEBUG_PRINT boolean to explicitly verify output of Debug.Print statements
    ' Use Hex$ for string result, use vbNullString, use Chr$, use Left$
    ' Use VBA-Inspector:Ignore for zzOLD module
    ' Use UCase$, use Mid$, use Trim$, use On Error GoTo 0 for turning on default VBA error handling
    ' 1100 hints in 15 objects
    ' Use Format$, use Dir$
    ' Fix "Multiple commands divided by a colon"
    ' Fix "Data access object declared without library"
    ' Fix "Scope of a procedure is not explicitly declared"
    ' Start to use aeDEBUG_PRINT throughout
    ' Fix pass through of varDebug in aeDocumentTheDatabase to aeExists
    ' ByVal vs ByRef - Ref: http://msdn.microsoft.com/en-us/library/ddck1z30.aspx - VBA default is ByVal
    ' How to: Force an Argument to Be Passed by Value (Visual Basic) - Ref: http://msdn.microsoft.com/en-us/library/chy4288y.aspx
    ' add decompile shortcut, start to add ByVal, test #002, strXMLResDoc As Variant
'20140324 - v090 - Add KPI chart test
    ' Show KPI chart as bar or pie with command button click
    ' Ref: http://answers.microsoft.com/en-us/office/forum/office_2010-access/ms-access-2010-windows-7-part-of-chart-is-not/9f8359bf-acf8-49a6-8f76-fe90d332a653
    ' This is a size limit "feature" - set the Width property to 23.9cm or 9.4"
    ' DONE: Create an adaept Chart Object as ["_tblChart", "_tlkpChart", _qryChart, _chtChart]
    ' Tidy output for no debugging, use () explicitly for functions
    ' Remove all blnDebug references, replace with IsMissing and standardize the messages
    ' Add Tasks/Issues section to change log for simple tracking
'20140318 - v089 - Bump
    ' Use GetRowSource to reduce code in _chtChart
    ' Add debug statement to test ChartType, changing Requery to Refresh gives better chart redraw response
    ' Use constants to display for different ChartType values
    ' Show command button numbers only using SetButtonsCaption - Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=157256
    ' Test for control type acLabel on chart form
    ' Set label caption to chart title
    ' Fix ribbons, test HideExport ribbon, add Access 2010 Controls IDs to doc - Ref: http://www.microsoft.com/en-us/download/details.aspx?id=6627
    ' Add AJP Ribbon Edit 2010 xlam to doc - Ref: http://www.andypope.info/vba/ribboneditor_2010.htm
    ' imageMso values from here Ref: http://soltechs.net/CustomUI/imageMso01.asp
    ' imageMso size="large" - case sensitive, add chart images to sample tab
'20140317 - v088 - Add CreateFormReportTextFile and FoundkeyWordInLine to aegit_expClass
    ' Add varDebug option to CreateFormReportTextFile
    ' Use varDebug in aeDocumentTheDatabase
    ' PrettyXML for table data macros
    ' Expand chart to 36 items
'20140313 - v087 - Solving embedded binary code in forms and reports
    ' OLE data change when chart saved: I changed the design view "Row Source Type" to "List" and then set "Row Source" to "1".
    ' Then in code, for the "Form_Current" event, I set "Row Source" to my SQL string, then set "Row Source Type" to "Table/Query".
    ' Ref: http://www.tek-tips.com/viewthread.cfm?qid=1092848
    ' There seems no reasonable way to manager PrtDevMode and OLEData in forms and reports.
    ' Strategy is to rewrite, remove hex chunks, and save new files a FormName_frm.txt
    ' Start TestForCreateFormReportTextFile testing
    ' Output line numbers of markers for checksum and hex data
    ' CreateFormReportTextFile working and diff tested with WinMerge
    ' Make chart more generic
    ' Chart Title in Form_Load Event
'20140307 - v086 - Start initial import testing
    ' QAQC template structure, _tblQAQC, _qryQAQC, _chtQAQC
    ' _tlkpQAQC for _chtQAQC X axis descriptions
    ' Testing ReportUseDefaultPrinter for PrtDevMode settings in form and report output
    ' Ref: http://msdn.microsoft.com/en-us/library/office/ff845464(v=office.15).aspx
'20140305 - v085 - Trap Err 75 in KillProperly after rename of database
    ' Split class to aegit_expClass and aegit_impClass
    ' Testing updates to ExportToExcel by James Kauffman, using late binding
'20140304 - v084 - Bump
'20140303 - v083 - Use mblnOutputPrinterInfo to determine if printer output info will be exported
    ' Write pretty xml
'20140226 - v082 - OutputTableDataMacros included in aegitClass
    ' ExportTableDataAsFormattedText test
    ' OutputTableDataAsFormattedText added to aegitClass with aetlkpStates as hardcoded example - to be fixed
    ' Only output aetlkpStates if it exists
    ' Fix error message when aetlkpStates does not exist
'20140224 - v081 - Fix file list output
    ' Add USysRibbons table - Ref: http://office.microsoft.com/en-us/access-help/customize-the-ribbon-HA010211415.aspx
    ' Create HideRibbon XML in USysRibbons table and set as default Ribbon Name for test
'20140221 - v080 - Early vs. Late Binding Ref: http://support.microsoft.com/kb/245115
    ' Add path length to folder list output
    ' Module aelan started
    ' Ref: http://www.exceltrick.com/formulas_macros/filesystemobject-in-vba/ - example of file size, modified date, etc.
    ' Fix output for files list
'20140220 - v079 - Format folder level with leading zero
    ' Add fLevelArrow to show levels in output, wsh not used
'20140219 - v078 - Add alternative test to output under the temp folder
    ' Formatting, source code review
    ' Code reorg, aever - office edition version info, aefs - file system related
    ' Change name of ListFilesRecursively to ListFileSystemRecursively
    ' Set default to show folders only
'20140216 - v077 - OutputTableDataMacros test
'20140214 - v076 - Output PrinterInfo to file
    ' Integrate PrinterInfo in aegitClass, turn on global error handler
    ' Export OfficeUI files
    ' Ref: http://msdn.microsoft.com/en-us/library/ee704589(v=office.14).aspx
    ' Add some data handlers Ref: http://www.saplsmw.com
    ' Testing GetSQLServerData
'20140212 - v075 - PrinterInfo procedure
'20140211 - v074 - Fix some dim statements as DAO, varDebug for output of properties
'20140210 - v073 - Use aegitExportDataToXML as flag for xml data output
'20140210 - v072 - Office 2013 test for "Show add-in user interface errors" reg key setting
'20140208 - v071 - Office 2010: File > Options > Client Settings > General > Show add-in user interface errors
    ' Ref: http://msdn.microsoft.com/en-us/library/office/dd548010(v=office.12).aspx
    ' Ribbon development: By default, you won't see any errors if there are problems with the XML that you've defined
    ' for a ribbon customization. To display errors during development, be sure to set the Show add-in user interface errors
    ' option in the Advanced page of the Access Options dialog box. Without this option set, ribbon customizations
    ' may not load and it may not necessarily be clear why.
    ' There is no VBA setting for Show add-in user interface errors
    ' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/c3c54c12-6cbd-404a-8709-dba485f82377/access-2010-set-option-show-addin-user-interface-errors-via-vba?forum=accessdev
    ' Need to develop a registry settings test/output routine
    ' Add code to test for reg key "ReportAddinCustomUIErrors"
'20140207 - v070 - Allow array of tables for xml data export to be provided via let property
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
    '   SELECT [Name] FROM [MSysObjects] WHERE [Type] = -32764 AND Left$([Name],1) <> "~" ORDER BY [Name]
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