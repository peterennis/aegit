@ECHO OFF
REM /************************************************************************************/
REM /* FILENAME       :  DECOMPILE.CMD                                                  */
REM /* TYPE           :  Windows NT Command Script                                      */
REM /* DESCRIPTION    :  This module drives the Access DB Decompile process             */
REM /*                                                                                  */
REM /* AUTHOR         :  Michael D Lueck                                                */
REM /*                   mlueck@lueckdatasystems.com                                    */
REM /*                                                                                  */
REM /* NEEDS          :                                                                 */
REM /*                                                                                  */
REM /* USAGE          :                                                                 */
REM /*                                                                                  */
REM /* REVISION HISTORY                                                                 */
REM /*                                                                                  */
REM /* DATE       REVISED BY DESCRIPTION OF CHANGE                                      */
REM /* ---------- ---------- -------------------------------------------------------    */
REM /* 10/28/2011 MDL        Initial Creation                                           */
REM /* 11/02/2011 MDL        Updated to parameterize and also display the pre/post size */
REM /* 12/27/2011 MDL        Updated to parse out the filesize and do the compare       */
REM /* 07/03/2012 MDL        Update to make UserID independent                          */
REM /* 01/21/2013 MDL        Pre request of April15Hater, updated to make safe for DB   */
REM /*                       filenames containing space characters                      */
REM /************************************************************************************/

REM Support for multiple database files within the one directory
REM Simply unREM the correct LOC to decompile that database file
SET DBfile=Fandango_FE_2007.accdb
REM SET DBfile=JDEFandangoReplicate.accdb
REM SET DBfile=SchemaIdeas.accdb
SET DBfile=Schema Ideas.accdb

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
PAUSE

:RunDecompile
FOR /F "delims=" %%A IN (' dir  /a-d/b "%DBfile%" ') DO (
  SET DBfilesizepre=%%~zA
  SET DBfilefullyqualified=%%~fA
)
ECHO File Size pre-decompile:  %DBfilesizepre%

"C:\Program Files\Microsoft Office\Office12\MSACCESS.EXE" /decompile "%DBfilefullyqualified%"

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