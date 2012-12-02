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
'

' 20121201 - v018 -
' 20121129 - v017 - Pass Fail test results and debug output cleanup
    ' Working on documenting the tables and relations
' 20121128 - v016 - SourceFolder property updated to allow passing the path into the class
    ' Cleanup debug messages code, include GetReferences from aeladdin (tm)
    ' Public Function aeReadDocDatabase - does it need a Get property call to make the function Private? - fixed, set to Private
' 20121127 - v015 - delete mac1 from accdb and manually delete the S_mac1.def file as the data export does not delete files
    ' version number continues from zip files stored in OLD folder
    ' basChangeLog added, export with OASIS and commit new changes to github