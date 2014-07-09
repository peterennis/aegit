Option Compare Database
Option Explicit

Private Ref As Reference

' 0 if Late Binding
' 1 if Reference to FSO set.
#Const FSORef = 0

Public Sub ExportIt(ByVal strTableName As String)
' Exports the table to the default folder
' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=65762

    On Error GoTo 0
    Application.ExportXML ObjectType:=acExportTable, DataSource:=strTableName, _
                    DataTarget:=strTableName & ".xml"

End Sub

Public Sub TestHideQueryDef()
' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/25d9dafd-b446-40ba-8dbd-a0efa983f2ff/how-to-programatically-hide-a-querydef

    On Error GoTo 0
    ' Query1 returns the list of all queries
    ' Ref: http://stackoverflow.com/questions/10882317/get-list-of-queries-in-project-ms-access
    Const strQueryName As String = "Query1"

    Dim fIsHidden As Boolean

    ' To determine if the query is hidden:
    fIsHidden = GetHiddenAttribute(acQuery, strQueryName)
    Debug.Print strQueryName, fIsHidden

    ' To show the query:
    SetHiddenAttribute acQuery, strQueryName, False

    ' To hide the query:
    SetHiddenAttribute acQuery, strQueryName, True

End Sub

Public Sub OutputListOfAllQueries()
' Ref: http://www.pcreview.co.uk/forums/runtime-error-7874-a-t2922352.html

    On Error GoTo 0
    Const strSQL As String = "SELECT m.Name " & vbCrLf & _
                                "FROM MSysObjects AS m " & vbCrLf & _
                                "WHERE (((m.Name) Not ALike ""~%"") AND ((m.Type)=5)) " & vbCrLf & _
                                "ORDER BY m.Name;"
    ' NOTE: Use zzz* for the query name so that it will be ignored by aegit code export
    Const strTempQuery As String = "zzz___MyTempQuery___"

    Debug.Print strSQL

    Dim qdfCurr As DAO.QueryDef

    ' Create the temp query that will have the query names
    On Error Resume Next
    Set qdfCurr = CurrentDb.QueryDefs(strTempQuery)
    If Err.Number = 3265 Then ' 3265 is "Item not found in this collection."
        Set qdfCurr = CurrentDb.CreateQueryDef(strTempQuery)
    End If
    qdfCurr.sql = strSQL
    'Debug.Print """" & strTempQuery & """"
    DoCmd.OpenQuery strTempQuery
    'DoCmd.Close acQuery, strTempQuery
    Debug.Print "The number of queries in the database is: " & DCount("Name", strTempQuery)

End Sub

Public Sub MakeTableWithListOfAllQueries()

    On Error GoTo 0
    Const strTempTable As String = "zzzTmpTblQueries"
    ' NOTE: Use zzz* for the table name so that it will be ignored by aegit code export
    Const strSQL As String = "SELECT m.Name, 0 AS Hidden INTO " & strTempTable & " " & vbCrLf & _
                                "FROM MSysObjects AS m " & vbCrLf & _
                                "WHERE (((m.Name) Not ALike ""~%"") AND ((m.Type)=5)) " & vbCrLf & _
                                "ORDER BY m.Name;"

    ' RunSQL works for Action queries
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
    Debug.Print "The number of queries in the database is: " & DCount("Name", strTempTable)

End Sub

Public Function ExportTheTableData(ByVal strTbl As String, ByVal strSpec As String, _
                    ByVal strPathFileName As String, ByVal blnHasHeaders As Boolean) As Boolean
' Ref: http://www.btabdevelopment.com/ts/2010ExpSpec

    On Error GoTo 0
    DoCmd.TransferText acExportDelim, strSpec, strTbl, strPathFileName, blnHasHeaders

End Function

';Function:
'; GetExtFileProperties()
';
';Author:
'; Allen Powell
';
';Version:
'; 1.2.2
';
';Version History
'; 1.2.2 2013/03/06 Updated to work with Windows 8, Added Table for Windows 8
'; 1.2.1 2009/06/11 Added Table for Windows 7, no changes in code
'; 1.2.0 2007/03/02 Updated to work with Vista, can now use the name or number of the Attribute/Property
'; 1.0.1 2006/11/18 Updated with a more efficient method of determining the same information
'; 1.0.0 2006/04/18 Original
';
';Action:
'; Get Extended File Properties / Attributes of Files
';
';Syntax:
'; GetExtFileProperties($FQFN, $attribute)
';
';Parameters:
'; $FQFN - (Required) Full Path Name and File name
'; $attribute - (Required) Number or Name from Attribute Table below.
';
';Returns:
'; String of Attribute(s) or nothing if there is an error.  Check @error for problems.
';
';Dependencies
';  Tested with Kixtart 4.53
';
';Examples:
';Using Windows XP and Attribute Number
';$MP3="C:\MP3s\Ac--Dc - Back In Black.mp3"
';$duration=GetExtFileProperties($mp3,21)
';? $duration
';
';Using Windows XP and Attribute Name
';$MP3="C:\MP3s\Ac--Dc - Back In Black.mp3"
';$duration=GetExtFileProperties($mp3,"Duration")
';? $duration
';
';Using Windows Vista and Attribute Name
';$MP3="C:\MP3s\Ac--Dc - Back In Black.mp3"
';$duration=GetExtFileProperties($mp3,"Length")
';? $duration
';
';Comments:
';  Please notice that attributes are specific to the OS.  This means that not only can the
';  Attribute number be different from one OS to the next, but the Attribute Name can be
';  different as well.
';

'Public Function GetExtFileProperties($FQFN, $attribute)
'  dim $objShell, $objFolder,$i,$found
'  if exist($FQFN)
'    $objShell=CreateObject("Shell.Application")
'    $objFolder=$objShell.Namespace(left($FQFN,instrrev($FQFN,"\")))
'    if $objFolder
'      if vartypename($attribute)="string"
'        While $i<298 and $found=0
'          if $attribute=$objFolder.GetDetailsOf($objFolder.Items, $i)
'            $attribute=$i
'            $found=1
'          End If
'          $i=$i+1
'        Loop
'      End If
'      if vartypename($attribute)="long" ; Number
'        $GetExtFileProperties=$objFolder.GetDetailsOf($objFolder.ParseName(Right$$FQFN,len($FQFN)-instrrev($FQFN,"\"))),$attribute)
'      Else
'        exit -1
'      End If
'    Else
'      exit @error
'    End If
'  Else
'    exit 2
'  End If
'endfunction

';
';Attribute tables
';
';    Windows 8/2012              Windows 7/2008 R2           Windows Vista/2008          Windows XP/2003     Windows 2000
';--------------------------------------------------------------------------------------------------------------------------
';  0 Name                        Name                        Name                        Name                Name
';  1 Size                        Size                        Size                        Size                Size
';  2 Item type                   Item type                   Type                        Type                Type
';  3 Date modified               Date modified               Date modified               Date Modified       Date Modified
';  4 Date created                Date created                Date created                Date Created        Attributes
';  5 Date accessed               Date accessed               Date accessed               Date Accessed       Comment
';  6 Attributes                  Attributes                  Attributes                  Attributes          Date Created
';  7 Offline status              Offline status              Offline status              Status              Date Accessed
';  8 Offline availability        Offline availability        Offline availability        Owner               Owner
';  9 Perceived type              Perceived type              Perceived type              Author              ???
'; 10 Owner                       Owner                       Owner                       Title               Author
'; 11 Kind                        Kind                        Kinds                       Subject             Title
'; 12 Date taken                  Date taken                  Date taken                  Category            Subject
'; 13 Contributing artists        Contributing artists        Artists                     Pages               Category
'; 14 Album                       Album                       Album                       Comments            Pages
'; 15 Year                        Year                        Year                        Copyright           Copyright
'; 16 Genre                       Genre                       Genre                       Artist              Company Name
'; 17 Conductors                  Conductors                  Conductors                  Album Title         Module Desription
'; 18 Tags                        Tags                        Tags                        Year                Module Version
'; 19 Rating                      Rating                      Rating                      Track Number        Product Name
'; 20 Authors                     Authors                     Authors                     Genre               Product Version
'; 21 Title                       Title                       Title                       Duration            Sender Name
'; 22 Subject                     Subject                     Subject                     Bit Rate            Recipient Name
'; 23 Categories                  Categories                  Categories                  Protected           Recipient Number
'; 24 Comments                    Comments                    Comments                    Camera Model        Csid
'; 25 Copyright                   Copyright                   Copyright                   Date Picture Taken  Tsid
'; 26 #                           #                           #                           Dimensions          Transmission Time
'; 27 Length                      Length                      Length                      ???                 Caller Id
'; 28 Bit rate                    Bit rate                    Bit rate                    ???                 Routing
'; 29 Protected                   Protected                   Protected                   Episode Name        Audio Format
'; 30 Camera model                Camera model                Camera model                Program Description Sample Rate
'; 31 Dimensions                  Dimensions                  Dimensions                  Description         Audio Sample Size
'; 32 Camera maker                Camera maker                Camera maker                Audio sample size   Channels
'; 33 Company                     Company                     Company                     Audio sample rate   Play Length
'; 34 File description            File description            File description            Channels            Frame Count
'; 35 Program name                Program name                Program name                Company             Frame Rate
'; 36 Duration                    Duration                    Duration                    Description         Video Sample Size
'; 37 Is online                   Is online                   Is online                   File Version        Video Compression
'; 38 Is recurring                Is recurring                Is recurring                Product Name        ???
'; 39 Location                    Location                    Location                    Product Version     ???
'; 40 Optional attendee addresses Optional attendee addresses Optional attendee addresses Keywords (XP only)  ???
'; 41 Optional attendees          Optional attendees          Optional attendees
'; 42 Organizer address           Organizer address           Organizer address
'; 43 Organizer name              Organizer name              Organizer name
'; 44 Reminder time               Reminder time               Reminder time
'; 45 Required attendee addresses Required attendee addresses Required attendee addresses
'; 46 Required attendees          Required attendees          Required attendees
'; 47 Resources                   Resources                   Resources
'; 48 Meeting status              Meeting status              Free/busy status
'; 49 Free/busy status            Free/busy status            Total size
'; 50 Total size                  Total size                  Account name
'; 51 Account name                Account name                Computer
'; 52 Task status                 Task status                 Anniversary
'; 53 Computer                    Computer                    Assistant's name
'; 54 Anniversary                 Anniversary                 Assistant's phone
'; 55 Assistant's name            Assistant's name            Birthday
'; 56 Assistant's phone           Assistant's phone           Business address
'; 57 Birthday                    Birthday                    Business city
'; 58 Business address            Business address            Business country/region
'; 59 Business city               Business city               Business P.O. box
'; 60 Business country/region     Business country/region     Business postal code
'; 61 Business P.O. box           Business P.O. box           Business state or province
'; 62 Business postal code        Business postal code        Business street
'; 63 Business state or province  Business state or province  Business fax
'; 64 Business street             Business street             Business home page
'; 65 Business fax                Business fax                Business phone
'; 66 Business home page          Business home page          Callback number
'; 67 Business phone              Business phone              Car phone
'; 68 Callback number             Callback number             Children
'; 69 Car phone                   Car phone                   Company main phone
'; 70 Children                    Children                    Department
'; 71 Company main phone          Company main phone          E-mail Address
'; 72 Department                  Department                  E-mail2
'; 73 E-mail address              E-mail address              E-mail3
'; 74 E-mail2                     E-mail2                     E-mail list
'; 75 E-mail3                     E-mail3                     E-mail display name
'; 76 E-mail list                 E-mail list                 File as
'; 77 E-mail display name         E-mail display name         First name
'; 78 File as                     File as                     Full name
'; 79 First name                  First name                  Gender
'; 80 Full name                   Full name                   Given name
'; 81 Gender                      Gender                      Hobbies
'; 82 Given name                  Given name                  Home address
'; 83 Hobbies                     Hobbies                     Home city
'; 84 Home address                Home address                Home country/region
'; 85 Home city                   Home city                   Home P.O. box
'; 86 Home country/region         Home country/region         Home postal code
'; 87 Home P.O. box               Home P.O. box               Home state or province
'; 88 Home postal code            Home postal code            Home street
'; 89 Home state or province      Home state or province      Home fax
'; 90 Home street                 Home street                 Home phone
'; 91 Home fax                    Home fax                    IM addresses
'; 92 Home phone                  Home phone                  Initials
'; 93 IM addresses                IM addresses                Job title
'; 94 Initials                    Initials                    Label
'; 95 Job title                   Job title                   Last name
'; 96 Label                       Label                       Mailing address
'; 97 Last name                   Last name                   Middle name
'; 98 Mailing address             Mailing address             Cell phone
'; 99 Middle name                 Middle name                 Nickname
';100 Cell phone                  Cell phone                  Office location
';101 Nickname                    Nickname                    Other address
';102 Office location             Office location             Other city
';103 Other address               Other address               Other country/region
';104 Other city                  Other city                  Other P.O. box
';105 Other country/region        Other country/region        Other postal code
';106 Other P.O. box              Other P.O. box              Other state or province
';107 Other postal code           Other postal code           Other street
';108 Other state or province     Other state or province     Pager
';109 Other street                Other street                Personal title
';110 Pager                       Pager                       City
';111 Personal title              Personal title              Country/region
';112 City                        City                        P.O. box
';113 Country/region              Country/region              Postal code
';114 P.O. box                    P.O. box                    State or province
';115 Postal code                 Postal code                 Street
';116 State or province           State or province           Primary e-mail
';117 Street                      Street                      Primary phone
';118 Primary e-mail              Primary e-mail              Profession
';119 Primary phone               Primary phone               Spouse
';120 Profession                  Profession                  Suffix
';121 Spouse/Partner              Spouse/Partner              TTY/TTD phone
';122 Suffix                      Suffix                      Telex
';123 TTY/TTD phone               TTY/TTD phone               Webpage
';124 Telex                       Telex                       Status
';125 Webpage                     Webpage                     Content type
';126 Content status              Content status              Date acquired
';127 Content type                Content type                Date archived
';128 Date acquired               Date acquired               Date completed
';129 Date archived               Date archived               Date imported
';130 Date completed              Date completed              Client ID
';131 Device category             Device category             Contributors
';132 Connected                   Connected                   Content created
';133 Discovery method            Discovery method            Last printed
';134 Friendly name               Friendly name               Date last saved
';135 Local computer              Local computer              Division
';136 Manufacturer                Manufacturer                Document ID
';137 Model                       Model                       Pages
';138 Paired                      Paired                      Slides
';139 Classification              Classification              Total editing time
';140 Status                      Status                      Word count
';141 Status                      Client ID                   Due date
';142 Client ID                   Contributors                End date
';143 Contributors                Content created             File count
';144 Content created             Last printed                Filename
';145 Last printed                Date last saved             File version
';146 Date last saved             Division                    Flag color
';147 Division                    Document ID                 Flag status
';148 Document ID                 Pages                       Space free
';149 Pages                       Slides                      Bit depth
';150 Slides                      Total editing time          Horizontal resolution
';151 Total editing time          Word count                  Width
';152 Word count                  Due date                    Vertical resolution
';153 Due date                    End date                    Height
';154 End date                    File count                  Importance
';155 File count                  Filename                    Is attachment
';156 File extension              File version                Is deleted
';157 Filename                    Flag color                  Has flag
';158 File version                Flag status                 Is completed
';159 Flag color                  Space free                  Incomplete
';160 Flag status                 Bit depth                   Read status
';161 Space free                  Horizontal resolution       Shared
';162 Sharing type                Width                       Creator
';163 Bit depth                   Vertical resolution         Date
';164 Horizontal resolution       Height                      Folder name
';165 Width                       Importance                  Folder path
';166 Vertical resolution         Is attachment               Folder
';167 Height                      Is deleted                  Participants
';168 Importance                  Encryption status           Path
';169 Is attachment               Has flag                    Contact names
';170 Is deleted                  Is completed                Entry type
';171 Encryption status           Incomplete                  Language
';172 Has flag                    Read status                 Date visited
';173 Is completed                Shared                      Description
';174 Incomplete                  Creators                    Link status
';175 Read status                 Date                        Link target
';176 Shared                      Folder name                 URL
';177 Creators                    Folder path                 Media created
';178 Date                        Folder                      Date released
';179 Folder name                 Participants                Encoded by
';180 Folder path                 Path                        Producers
';181 Folder                      By location                 Publisher
';182 Participants                Type                        Subtitle
';183 Path                        Contact names               User web URL
';184 By location                 Entry type                  Writers
';185 Type                        Language                    Attachments
';186 Contact names               Date visited                Bcc addresses
';187 Entry type                  Description                 Bcc names
';188 Language                    Link status                 Cc addresses
';189 Date visited                Link target                 Cc names
';190 Description                 URL                         Conversation ID
';191 Link status                 Media created               Date received
';192 Link target                 Date released               Date sent
';193 URL                         Encoded by                  From addresses
';194 Media created               Producers                   From names
';195 Date released               Publisher                   Has attachments
';196 Encoded by                  Subtitle                    Sender address
';197 Episode number              User web URL                Sender name
';198 Producers                   Writers                     Store
';199 Publisher                   Attachments                 To addresses
';200 Season number               Bcc addresses               To do title
';201 Subtitle                    Bcc                         To names
';202 User web URL                Cc addresses                Mileage
';203 Writers                     Cc                          Album artist
';204 Attachments                 Conversation ID             Beats-per-minute
';205 Bcc addresses               Date received               Composers
';206 Bcc                         Date sent                   Initial key
';207 Cc addresses                From addresses              Mood
';208 Cc                          From                        Part of set
';209 Conversation ID             Has attachments             Period
';210 Date received               Sender address              Color
';211 Date sent                   Sender name                 Parental rating
';212 From addresses              Store                       Parental rating reason
';213 From                        To addresses                Space used
';214 Has attachments             To do title                 EXIF version
';215 Sender address              To                          Event
';216 Sender name                 Mileage                     Exposure bias
';217 Store                       Album artist                Exposure program
';218 To addresses                Album ID                    Exposure time
';219 To do title                 Beats-per-minute            F-stop
';220 To                          Composers                   Flash mode
';221 Mileage                     Initial key                 Focal length
';222 Album artist                Part of a compilation       35mm focal length
';223 Sort album artist           Mood                        ISO speed
';224 Album ID                    Part of set                 Lens maker
';225 Sort album                  Period                      Lens model
';226 Sort contributing artists   Color                       Light source
';227 Beats-per-minute            Parental rating             Max aperture
';228 Composers                   Parental rating reason      Metering mode
';229 Sort composer               Space used                  Orientation
';230 Initial key                 EXIF version                Program mode
';231 Part of a compilation       Event                       Saturation
';232 Mood                        Exposure bias               Subject distance
';233 Part of set                 Exposure program            White balance
';234 Period                      Exposure time               Priority
';235 Color                       F-stop                      Project
';236 Parental rating             Flash mode                  Channel number
';237 Parental rating reason      Focal length                Episode name
';238 Space used                  35mm focal length           Closed captioning
';239 EXIF version                ISO speed                   Rerun
';240 Event                       Lens maker                  SAP
';241 Exposure bias               Lens model                  Broadcast date
';242 Exposure program            Light source                Program description
';243 Exposure time               Max aperture                Recording time
';244 F-stop                      Metering mode               Station call sign
';245 Flash mode                  Orientation                 Station name
';246 Focal length                People                      Auto summary
';247 35mm focal length           Program mode                Summary
';248 ISO speed                   Saturation                  Search ranking
';249 Lens maker                  Subject distance            Sensitivity
';250 Lens model                  White balance               Shared with
';251 Light source                Priority                    Product name
';252 Max aperture                Project                     Product version
';253 Metering mode               Channel number              Source
';254 Orientation                 Episode name                Start date
';255 People                      Closed captioning           Billing information
';256 Program mode                Rerun                       Complete
';257 Saturation                  SAP                         Task owner
';258 Subject distance            Broadcast date              Total file size
';259 White balance               Program description         Legal trademarks
';260 Priority                    Recording time              Video compression
';261 Project                     Station call sign           Directors
';262 Channel number              Station name                Data rate
';263 Episode name                Summary                     Frame height
';264 Closed captioning           Snippets                    Frame rate
';265 Rerun                       Auto summary                Frame width
';266 SAP                         Search ranking              Total bitrate
';267 Broadcast date              Sensitivity
';268 Program description         Shared with
';269 Recording time              Sharing status
';270 Station call sign           Product name
';271 Station name                Product version
';272 Summary                     Support link
';273 Snippets                    Source
';274 Auto summary                Start date
';275 Search ranking              Billing information
';276 Sensitivity                 Complete
';277 Shared with                 Task owner
';278 Sharing status              Total file size
';279 Product name                Legal trademarks
';280 Product version             Video compression
';281 Support link                Directors
';282 Source                      Data rate
';283 Start date                  Frame height
';284 Billing information         Frame rate
';285 Complete                    Frame width
';286 Task owner                  Total bitrate
';287 Sort title                  Creator
';288 Total file size             Encryption Level
';289 Legal trademarks            Content Accessibility
';290 Video compression           Document Assembly
';291 Directors                   Changing
';292 Data rate                   Commenting
';293 Frame height                Copying
';294 Frame rate                  Form Filling
';295 Frame width                 Printing
';296 Video orientation           Producer
';297 Total bitrate               PDF Specification
'

Private Function GetFiles(ByVal strPath As String, _
                ByVal dctDict As Object, _
                Optional ByVal blnRecursive As Boolean) As Boolean
'Function GetFiles(strPath As String, _
                dctDict As Dictionary, _
                Optional blnRecursive As Boolean) As Boolean
' This procedure returns all the files in a directory into
' a Dictionary object. If called recursively, it also returns
' all files in subfolders.

    #If FSORef = 0 Then  ' Late binding
        ' Ref: http://msdn.microsoft.com/en-us/library/office/gg278516.aspx
        Dim fsoSysObj As Object
        Dim oFolder As Object
        Dim oSubFolder As Object
        Dim oFile As Object
        Set fsoSysObj = CreateObject("Scripting.FileSystemObject")
        ' <=======
        ' Remove the Object reference if it is present
        On Error Resume Next
        Set Ref = References!Scripting
        If Err.Number = 0 Then
            References.Remove Ref
        ElseIf Err.Number <> 9 Then 'Subscript out of range meaning not reference not found
            MsgBox Err.Description
            Exit Function
        End If
        ' Use your own error handling label here
        On Error GoTo PROC_ERR
        '<=======
    #Else
        ' A reference to the Object Library must be specified
        Dim fsoSysObj As FileSystemObject
        Dim oFolder As Folder
        Dim oSubFolder As Folder
        Dim oFile As File
        ' Return new FileSystemObject.
        Set fsoSysObj = New FileSystemObject
    #End If


    On Error Resume Next
    ' Get folder.
    Set oFolder = fsoSysObj.GetFolder(strPath)
    If Err <> 0 Then
        ' Incorrect path.
        GetFiles = False
        GoTo PROC_EXIT
    End If
    On Error GoTo 0

    ' Loop through Files collection, adding to dictionary.
    For Each oFile In oFolder.Files
        dctDict.Add oFile.Path, oFile.Path
    Next oFile

    ' If Recursive flag is true, call recursively.
    If blnRecursive Then
        For Each oSubFolder In oFolder.SubFolders
            GetFiles oSubFolder.Path, dctDict, True
        Next oSubFolder
    End If

    ' Return True if no error occurred.
    GetFiles = True

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.Description, vbCritical, "GetFiles Error!"
    Exit Function

End Function

Public Sub TestGetFiles()
' Ref: http://msdn.microsoft.com/en-us/library/office/aa164475(v=office.10).aspx

    On Error GoTo 0
    #If FSORef = 0 Then  ' Late binding
        Dim dctDict As Object
        ' Create new dictionary
        Set dctDict = CreateObject("Scripting.Dictionary")
    #Else
        ' A reference to the Object Library must be specified
        Dim dctDict As Dictionary
        Set dctDict = New Dictionary
    #End If

    Dim varItem As Variant
    Dim GetTempDir As String

    GetTempDir = "C:\TEMP"
    ' Call recursively, return files into Dictionary object.
    If GetFiles(GetTempDir, dctDict, True) Then
        ' Print items in dictionary.
        For Each varItem In dctDict
            Debug.Print varItem
        Next
    End If

End Sub