Attribute VB_Name = "DLPortIO"
'****************************************************************************
'*  @doc INTERNAL
'*  @module dlportio.bas |
'*
'*  DriverLINX Port I/O Driver Interface
'*  <cp> Copyright 1996 Scientific Software Tools, Inc.<nl>
'*  All Rights Reserved.<nl>
'*  DriverLINX is a registered trademark of Scientific Software Tools, Inc.
'*
'*  Win32 Prototypes for DriverLINX Port I/O
'*
'*  Please report bugs to:
'*  Scientific Software Tools, Inc.
'*  19 East Central Avenue
'*  Paoli, PA 19301
'*  USA
'*  E-mail: support@sstnet.com
'*  Web: www.sstnet.com
'*
'*  @comm
'*  Author: RoyF<nl>
'*  Date:   09/26/96 14:08:58
'*
'*  @group Revision History
'*  @comm
'*  $Revision: 2 $
'*  <nl>
'*  $Log: /XN Alarm.ADO/Dlportio.bas $
rem 
rem 2     4/05/04 9:48a Tkharak
rem 
rem 19    1/27/04 1:39p Tkharak
rem 
rem 16    2/06/03 11:35a Tkharak
rem 
rem 15    1/06/03 2:09p Tkharak
rem 
rem 14    12/16/02 10:11a Tkharak
rem 
rem 13    11/30/01 10:15a Tkharak
rem 
rem 12    8/09/00 10:58a Jslawin
rem This is absoluty my last checkin. Made changes to ignore NULLs instead
rem of converting them to space.
rem 
rem 11    8/08/00 9:58a Jslawin
rem JOseph's last checkin, really. Fix made for line feeds in the simplex
rem inteface
rem 
rem 10    8/07/00 9:59a Jslawin
rem Joseph's last checkin
rem 
rem 9     7/18/00 2:59p Jslawin
rem Simplex Interface modified to work with any alarm system
rem 
rem 8     6/07/00 2:14p Jslawin
rem Converted to use queue btrieve file as an option
rem 
rem 7     6/14/99 2:26p Jslawin
rem Check in before vacation
rem 
rem 6     2/02/99 12:33p Jslawin
rem Statistics correction
rem 
rem 5     12/29/98 9:07a Jslawin
rem Fix made to sendpage of XoQue.bas and XnQue.bas. Passed parameter
rem extension may be overwritten creating unpredictable results for calling
rem procedueres.
rem 
rem 4     12/28/98 11:15a Jslawin
rem Feature added for Supervisor paging, action reminder bug fixed.
rem 
rem 3     12/23/98 11:31a Jslawin
rem Fix made to handle Action reminder correcly
rem 
rem 2     10/23/98 2:51p Jslaw
rem 
rem 1     7/01/98 1:29p Jslaw
rem Xn Alarm Project
rem 
rem 1     3/11/98 1:48p Jslaw
'
' 1     9/27/96 2:03p Royf
' Initial revision.
'*
'****************************************************************************


Public Declare Function DlPortReadPortUchar Lib "dlportio.dll" (ByVal Port As Long) As Byte
Public Declare Function DlPortReadPortUshort Lib "dlportio.dll" (ByVal Port As Long) As Integer
Public Declare Function DlPortReadPortUlong Lib "dlportio.dll" (ByVal Port As Long) As Long

Public Declare Sub DlPortReadPortBufferUchar Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortReadPortBufferUshort Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortReadPortBufferUlong Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)

Public Declare Sub DlPortWritePortUchar Lib "dlportio.dll" (ByVal Port As Long, ByVal Value As Byte)
Public Declare Sub DlPortWritePortUshort Lib "dlportio.dll" (ByVal Port As Long, ByVal Value As Integer)
Public Declare Sub DlPortWritePortUlong Lib "dlportio.dll" (ByVal Port As Long, ByVal Value As Long)

Public Declare Sub DlPortWritePortBufferUchar Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortWritePortBufferUshort Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortWritePortBufferUlong Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)

