@ECHO OFF
REM /***************************************************************************************/
REM /* FILENAME       :  DECOMPILE.CMD                                                     */
REM /* TYPE           :  Windows NT Command Script                                         */
REM /* DESCRIPTION    :  This module drives the Access DB Decompile process                */
REM /*                                                                                     */
REM /* AUTHOR         :  Michael D Lueck                                                   */
REM /*                   mlueck@lueckdatasystems.com                                       */
REM /*                                                                                     */
REM /* NEEDS          :                                                                    */
REM /*                                                                                     */
REM /* USAGE          : http://www.access-programmers.co.uk/forums/showthread.php?t=219948 */
REM /*                                                                                     */
REM /* REVISION HISTORY                                                                    */
REM /*                                                                                     */
REM /* DATE       REVISED BY DESCRIPTION OF CHANGE                                         */
REM /* ---------- ---------- ------------------------------------------------------------- */
REM /* 10/28/2011 MDL        Initial Creation                                              */
REM /* 11/02/2011 MDL        Updated to parameterize and also display the pre/post size    */
REM /* 12/27/2011 MDL        Updated to parse out the filesize and do the compare          */
REM /* 07/03/2012 MDL        Update to make UserID independent                             */
REM /* 01/21/2013 MDL        Pre request of April15Hater, updated to make safe for DB      */
REM /*                       filenames containing space characters                         */
REM /* 04/03/2014 PFE        Modified for aegit for testing #002                           */
REM /***************************************************************************************/

REM Suggested Decompile Steps:
REM 1) The database should be in the same directory on the C: drive as the decompile.cmd script attached. Update the script as necessary to make correct for your working environment.
REM 2) Run the decompile.cmd script which will start Access, the database, and Access will decompile it.
REM 3) Next close the database – not all of Access.
REM 4) Then reopen the database. (Remember the shift key if you have an autoexec macro!)
REM 5) Compact the database at this point (MS Office icon \ Manage \ Compact and Repair) (Remember the shift key if you have an autoexec macro!)
REM 6) Press Ctrl+G to open the VBA window
REM 7) Click the Debug menu \ Clear All Breakpoints
REM 8) Click the Debug menu \ Compile - ONLY do this step the FIRST time!
REM 9) Then Compact again as in Step 6 (Remember the shift key if you have an autoexec macro!)
REM 10) Completely exit Access (Remember the shift key if you have an autoexec macro!)
REM When the before / after file size are finally the same, the decompile.cmd script will end.

REM Support for multiple database files within the one directory
REM Simply unREM the correct LOC to decompile that database file
SET DBfile=adaept revision control.accdb
SET AccessPath="C:\Program Files\Microsoft Office\Office15\MSACCESS.EXE"

ECHO.
ECHO This script will Decompile the %DBfile% database.
ECHO.
ECHO Do you want to do that?
ECHO.
ECHO If NO, then Ctrl-Break NOW!
ECHO.
ECHO Please remember to hold down the shift key to prevent Access from running
ECHO the autoexec macro if there is one in the database being decompiled.
ECHO.
ECHO DBfile=%DBfile%
ECHO AccessPath=%AccessPath%
PAUSE

:RunDecompile
FOR /F "delims=" %%A IN (' dir  /a-d/b "%DBfile%" ') DO (
  SET DBfilesizepre=%%~zA
  SET DBfilefullyqualified=%%~fA
)
ECHO File Size pre-decompile:  %DBfilesizepre%

%AccessPath% /decompile "%DBfilefullyqualified%"

FOR /F "delims=" %%A IN (' dir  /a-d/b "%DBfile%" ') DO (
  SET DBfilesizepost=%%~zA
)
ECHO File Size post-decompile: %DBfilesizepost%

IF %DBfilesizepre% == %DBfilesizepost% GOTO End

ECHO File size is different, so running decompile again...
ECHO.
GOTO RunDecompile

:End
ECHO Decompile Completed!
ECHO Press any key to finish ...
PAUSE > nul