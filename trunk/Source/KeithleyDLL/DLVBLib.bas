Attribute VB_Name = "DriverLINXLibrary"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc INTERNAL
 '   @module DLVBLib.bas |
 '
 '  DriverLINX<rtm> Visual Basic Function Call Library<nl>
 '  <cp> Copyright 1997 Scientific Software Tools, Inc.<nl>
 '  All Rights Reserved.<nl>
 '
 '  User Library Functions
 '
 '  @comm
 '  Author: KevinD<nl>
 '  Date:   10/27/97 11:05:00
 '
 '  @group Revision History
 '  @comm
 ' 1     10/27/97 2:30p KevinD
 ' Initial revision.
 '
 ' 2    7/1/99 3:05 PM KevinD
 '  Made changes to the DoesSupportChannelGainList and
 '  IsInDriverLINXAnalogRange to fix bug and to accomodate
 '  DDA8/16 boards.
 '
 
Option Explicit
' @const    Integer |   Foreground  |Defines polled mode as the .Req_mode.
Public Const Foreground As Integer = 0
' @const    Integer |   Background  |Defines DMA or IRQ as the .Req_mode.
' This library will automatically determine the best option.
Public Const Background As Integer = 1
' @const    String  |   PreventOpenDialog   |Defines the string that prevents
' the Open DriverLINX Driver Dialog from opening.
Public Const PreventOpenDialog As String = "$"

 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Opens a DriverLINX driver.
 '
 ' @rdesc   String - returns the name of the open driver or an empty string
 '          if a driver is not opened.
 '
 ' @parm    DriverLINXSR   |   SR          |Name of the control
 ' @parm    String         |DriverName     |Name of driver to open
 ' @parm    Boolean        |NoDialogBox    |Determines whether the open dialog box
 ' is displayed
 '
 ' @comm    <f OpenDriverLINXDriver> This function opens a driver. If the
 '          DriverName argument is "" the function will display a
 '          "Open DriverLINX" dialog box. User can also specify the driver
 '          to open. The NoDialogBox argument prevents the displaying of
 '          the "Open DriverLINX" dialog box.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f CloseDriverLINXDriver>
 '
Public Function OpenDriverLINXDriver(SR As DriverLINXSR, ByVal DriverName As String, _
                                    ByVal NoDialogBox As Boolean) As String
    
    Dim DLDriverName As String
    
    
    If (DriverName = "") Or (DriverName = Null) Then
        ' If the DriverName parameter is a null string, then
        '   DriverLINX/VB will display its "Open DriverLINX" dialog
        '   box.
        DLDriverName = "*.DLL"
    Else
        ' If you specify a DriverName, DriverLINX will try to open
        '   this driver. If your driver fails to open, DriverLINX
        '   will display the "Open DriverLINX" dialog box.
        DLDriverName = DriverName
        If NoDialogBox Then
            ' By appending "PreventOpenDialog" string to the name of the
            '   driver you want to open, you prevent DriverLINX
            '   from displaying the "Open DriverLINX" dialog box.
            '   In this case, if DriverLINX can not open the
            '   requested driver, it will set the Req_DLL_name to
            '   an empty string. Only Works with Newer Drivers.
            DLDriverName = DLDriverName & PreventOpenDialog
        End If
    End If
        
    With SR
        ' Open a DriverLINX driver.
        .Req_DLL_name = DLDriverName
        ' DriverLINX/VB tries to open the specified driver as soon as
        '   you set the Req_DLL_name property. You should not call
        '   the Refresh method.
        
        ' Clicking the "Cancel" button on the "Open Driver" dialog box
        '   will return a null string, i.e. ""
        OpenDriverLINXDriver = .Req_DLL_name
    End With
    
End Function

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Initializes a device.
 '
 ' @rdesc   Boolean - returns result code for operation.
 '
 ' @parm    DriverLINXSR    |   SR          |Name of Service Request control
 ' @parm    Integer         |   Device      |Device to Initialize
 '
 ' @comm    <f InitializeDriverLINXDevice> This function sets the appropriate
 '           fields of the Service Request required to initialize a logical
 '           device. This function executes a Service Request. A device has
 '           to be initialized before it can perform any data acquisition
 '           tasks.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function InitializeDriverLINXDevice(SR As DriverLINXSR, ByVal device As Integer _
                                            ) As Integer

    ' Initialize a Logical Device
    
    ' Setup the Service Request to initialize the desired Logical
    '   Device.
    With SR
        ' ------- Service Request Group -------------
        ' Specify type of Service Request
        AddRequestGroupInitialize SR, device
    
        AddStartEventNullEvent SR       'set these fields to 0 or DL_NULLEVENT
        AddStopEventNullEvent SR        'if not used in the Service Request. This
        AddTimingEventNullEvent SR      'function is synchronous, therefore, events
                                   'and buffers are not needed and should be omitted.
        'Add Select Channel Group
        AddSelectZeroChannels SR
        'Add Select Buffer Group
        AddSelectBuffers SR, 0, 0, 0
        
        ' Execute the service request
        .Refresh
    
        ' Process the results. In this case, just return the error code
        InitializeDriverLINXDevice = .Res_result
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Determines if subsystem exists on board.
 '
 ' @rdesc   Boolean - returns true is subsystem exists
 '
 ' @parm    DriverLINXSR    |   SR          |Name of Service Request control
 ' @parm    DriverLINXLDD   |   LDD         |Name of LDD control
 ' @parm    Integer         |   Subsystem   |Subsystem to check for
 '
 ' @comm    <f HasDriverLINXSubsystem> This function queries the LDD to see
 '           if the board supports the desired subsystem.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
Public Function HasDriverLINXSubsystem(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                        ByVal subsystem As Integer) As Boolean

    Dim DLDriverName As String
    Dim DLSupportedSubsystem As Integer
    
    ' See if the Logical Device supports the subsystem that you want.
    ' This function gets a driver name and Logical device from a
    ' DriverLINXSR control, and uses these values in a DriverLINXLDD
    ' control.
    
    ' First check to see if a driver is open
    DLDriverName = SR.Req_DLL_name
    
    If DLDriverName <> "" Then
        ' If driver is open, then see if it has the desired subsystem
        With LDD
            ' Make sure that the Service Request and LDD controls open
            '   the same DriverLINX driver
            .device = SR.Req_device
            .Req_DLL_name = DLDriverName
            
            ' Convert the subsystem number to a bit-number in the LDD's
            '   device feature map
            subsystem = 2 ^ subsystem
            ' Is the bit that corresponds to the requested subsystem
            '   set in the LDD's Dev_Feature Map property?
            DLSupportedSubsystem = (.Dev_FeatureMap And subsystem)
            ' If bit is set, return True. If bit is not set, return False
            HasDriverLINXSubsystem = (DLSupportedSubsystem = subsystem)
            ' Close the LDD's driver
            .Req_DLL_name = ""
        End With
    Else
        ' Make sure that HasDriverLINXsubsystem returns False if no
        '   driver is open.
        HasDriverLINXSubsystem = False
    End If
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Determines the number of logical channels within a subsystem.
 '
 ' @rdesc   Integer - returns the number of channels supported by the subsystem
 '
 ' @parm    DriverLINXSR   |    SR              |Name of the control
 ' @parm    DriverLINXLDD  |    LDD             |Name of LDD control
 ' @parm    Integer        |    Subsystem       |Subsystem to check for
 '
 ' @comm    <f HowManyDriverLINXLogicalChannels> This function queries the
 '           LDD to see how many channels are supported within the specified
 '           subsystem.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f HasDriverLINXSubsystem>, <f IsHardwareIntel8255>,
 '           <f HowManyExtendedDigitalChannels>,
 '           <f HowManyBitsPerDigitalChannel>,
 '           <f HowManyBytesPerDigitalChannel>
 '
 '
Public Function HowManyDriverLINXLogicalChannels(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                ByVal subsystem As Integer) As Integer

    Dim DLDriverName As String
    Dim DLLogicalChannels As Integer
    Dim i As Integer
    
    ' Get the number of logical channels supported by this subsystem
    ' This function gets a driver name and Logical device from a
    '   DriverLINXSR control, and uses these values in a DriverLINXLDD
    '   control.
    
    ' First check to see if a driver is open
    DLDriverName = SR.Req_DLL_name
    
    If DLDriverName <> "" Then
        ' If driver is open, then see if it has the desired subsystem
        With LDD
            ' Make sure that the Service Request and LDD controls open
            '   the same DriverLINX driver
            .device = SR.Req_device
            .Req_DLL_name = DLDriverName
            
            ' Get the number of channels in this subsystem
            Select Case subsystem
                Case DL_AI
                    DLLogicalChannels = .AI_nChan
                    
                Case DL_AO
                    DLLogicalChannels = .AO_nChan
                    
                Case DL_DI
                    DLLogicalChannels = .DI_nChan
                    For i = 0 To DLLogicalChannels - 1
                        If .DI_Type(i) = 2 Or _
                                        .DI_Type(i) = 3 Then
                            DLLogicalChannels = DLLogicalChannels - 1
                        End If
                    Next i
                           
                Case DL_DO
                    DLLogicalChannels = .DO_nChan
                    
                Case DL_CT
                    DLLogicalChannels = .CT_nChan
                    
                Case Else
                    DLLogicalChannels = 0
                
            End Select
            
            ' Close the LDD's driver
            .Req_DLL_name = ""
        End With
    Else
        ' Make sure that HowManyLINXLogicalChannels returns 0 if no
        '   driver is open.
        DLLogicalChannels = 0
    End If
    
    HowManyDriverLINXLogicalChannels = DLLogicalChannels
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Determines the number of channels in a start stop list.
 '
 ' @rdesc   Integer - returns the number channels to be sampled in a start stop list.
 '
 ' @parm    DriverLINXSR   |   SR          |Name of the control
 ' @parm    DriverLINXLDD  |   LDD         |Name of LDD control
 ' @parm    Integer        |   channels    |Name of 2 element array that
 ' contains the channel start and the channel stop arguments
 ' @parm    Integer        |    Subsystem   |Subsystem that corresponds to the
 ' start stop list
 '
 ' @comm    <f SizeOfStartStopList> This function determines how many channels are
 '          to be sampled in a start stop list.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function SizeOfStartStopList(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                    ByRef channels() As Integer, ByVal subsystem As Integer) As Integer
    'Channels() is a two element array. The first element is the starting channel and the
    'second element is the ending element. Used with Service Request that use a Start/Stop List
    
    Dim MaxChannel As Integer
    MaxChannel = HowManyDriverLINXLogicalChannels(SR, LDD, subsystem) 'get the max number
                                                                      'of Channels
    'The if statement below handles two situations: First if the stop channel is
    'greater than the start channel the number of channels is the difference plus 1.
    'Second if the start channel is greater than stop channel. DriverLINX will
    'include all channels starting from the start channel to the max number of channels
    'plus channel 0 to the stop channel. In this case, DriveLINX wraps around back to
    'the stop channel.
    If channels(0) <= channels(1) Then
        SizeOfStartStopList = channels(1) - channels(0) + 1
    ElseIf channels(0) > channels(1) Then
        SizeOfStartStopList = (MaxChannel - channels(0)) + channels(1) + 1
    End If
        
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Determines the number of logical channels within a subsystem.
 '
 ' @rdesc   Integer - returns the number of channels supported by the subsystem
 '
 ' @parm    DriverLINXSR   |    SR          |Name of the control
 ' @parm    DriverLINXLDD  |    LDD         |Name of LDD control
 ' @parm    Integer        |    Subsystem   |Subsystem to check for
 '
 ' @comm    <f HowManyDriverLINXLogicalChannels> This function queries the
 '           LDD to see how many channels are supported within the specified
 '           subsystem.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f HasDriverLINXSubsystem>, <f IsHardwareIntel8255>,
 '           <f HowManyExtendedDigitalChannels>,
 '           <f HowManyBitsPerDigitalChannel>,
 '           <f HowManyBytesPerDigitalChannel>
 '
 '
Public Function HowManyDriverLINXChannelsSampled(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                        ByVal subsystem As Integer, _
                                        ByVal StartStopChannelGain As Integer) As Integer
    'Function can be used after the service request has been setup to determine how many
    'channels are to be sampled
    Dim MaxChannel As Integer
    Dim channels(1) As Integer  '2 dimensional array
    
    With SR
        If StartStopChannelGain Then    '1= Channel Gain List
            HowManyDriverLINXChannelsSampled = .Sel_chan_N
        Else
            channels(0) = .Sel_chan_start
            channels(1) = .Sel_chan_stop
            HowManyDriverLINXChannelsSampled = SizeOfStartStopList(SR, LDD, _
                                                                    channels(), subsystem)
        End If
    End With
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Read or writes a single value.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   Channel                 |Channel to read or write
' @parm    Single           |   gain                    |Channels gain value
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm    <f SetupDriverLINXSingleValueIO> This function sets up a Service
'           Request that either inputs or outputs one value from/to a
'           subsystem.
'
' @devnote KevinD 10/27/97 11:40:00AM
'
' @xref    <f GetDriverLINXDISingleValue>, <f GetDriverLINXAISingleValue>,
'          <f PutDriverLINXAIBuffer>
'
'
Public Sub SetupDriverLINXSingleValueIO(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                        ByVal device As Integer, _
                                        ByVal subsystem As Integer, _
                                        ByVal Channel As Integer, _
                                        ByVal gain As Single, _
                                        BackGroundForeGround As Integer)
' Setup the Service Request for single-value I/O

        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        ' Single-value I/O doesn't require a timing event
        AddTimingEventNullEvent SR
        
        ' Start immediately on software command
        AddStartEventNullEvent SR               ' or AddStartEventOnCommand
        
        ' Stop as soon as DriverLINX processes the sample
        AddStopEventNullEvent SR               ' or AddStopEventOnTerminalCount
        
        ' ------------ Select Channel Group ----------------
        ' Specify channels, gain and data format
         AddSelectSingleChannel SR, Channel, subsystem, gain    'setup channel Number,
                                                                'and Gain settings
        
        ' ------------ Select Buffers Group ----------------
        ' Single value transfers do not use buffers
        AddSelectBuffers SR, 0, 0, 0
        
         ' ------------ Select Flags -----------------------
        ' Single-value I/O doesn't need ServiceStart or ServiceDone
        '   events
        AddSelectFlags SR, False
        ' NOTE: You do not have to block any events. However,
        '   DriverLINX is somewhat more efficient if you do.
        
        'Note: Your application must call the refresh method to execute this function.
        'Note: For output Service Requests make sure to set the .Res_Sta_ioValue property
        '      appropriately before calling .Refresh
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Reads or writes one or more buffers of data from/to
'          a subsystem using the onboard clock.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   Channel                 |Channel to read or write
' @parm    Single           |   gain                    |Channels gain value
' @parm    Single           |   frequency               |Rate or which to read or write
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXBufferedIO> This function sets up a Service
'           Request that either inputs or outputs one or more buffers of data
'           from/to a subsystem. The Service Requests starts when submitted
'           and stops when the buffer is either full or empty depending
'           whether it is an input or output task. Data is clocked in/out
'           using the boards default clock. Function only will work if board
'           has an onboard clock and if the board supports interrupt or DMA
'           data transfer.
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f PutDriverLINXAIBuffer>,<f GetDriverLINXAIBuffer>,
'           <f SetupDriverLINXContinuousBufferedIO>,
'
'

Public Sub SetupDriverLINXBufferedIO(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                    ByVal device As Integer, _
                                    ByVal subsystem As Integer, _
                                    ByVal Channel As Integer, _
                                    ByVal gain As Single, _
                                    ByRef frequency As Single, _
                                    ByVal SamplesPerChannel As Single, _
                                    ByVal Buffers As Integer, _
                                    ByVal BackGroundForeGround As Integer)
    Dim ChannelsSampled As Integer
    ChannelsSampled = 1 'In this case function is looking for 1 channel
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        'Specify timing event
        AddTimingEventDefault SR, frequency
        
        ' Specify start event
        AddStartEventOnCommand SR 'Start on Software Command
        ' Specify stop event
        AddStopEventOnTerminalCount SR 'Stop on Terminal Count
        
        ' ------------ Select Channel Group ----------------
        ' Specify channels, gain and data format
        AddSelectSingleChannel SR, Channel, subsystem, gain     'setup channel Number,
                                                                'and Gain settings
        
        ' ------------ Select Buffers Group ----------------
        ' Specify the number of buffers and size
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
        
        'Note: Your application must call the refresh method to execute this function.
        'Note: For output Service Requests make sure to set the buffers are appropriately
        '      filled before calling .Refresh
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Reads or writes multiple channels using a single buffer.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   subsystem               |Subsystem to setup
' @parm    Integer          |   channel()               |Two element array containing the
' start channel and the stop channel
' @parm    Single           |   Gains()                 |Two element array containing the
' start channels gain and the stop channel gain
' @parm    Single           |   frequency               |Rate or which to read or write in Hz
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXMultiChannelStartStopList> This function sets
'           up a Service Request that either inputs or outputs one buffer
'           from/to a subsystem. This function makes use of a start stop
'           list. The draw back to this is that the user can on specify
'           consecutive channels with a gain setting for the first channel
'           and a gain setting for all of the other channels that are
'           included in the start stop list.
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f PutDriverLINXAIBuffer>,<f GetDriverLINXAIBuffer>,
'           <f SetupDriverLINXMultiChannelGainList>,
'           <f SetupDriverLINXMultiChannelDigitalStartStopList>,
'           <f SetupDriverLINXMultiChannelBurstMode>,
'           <f SetupDriverLINXSingleScanIO>
'
'
Public Sub SetupDriverLINXMultiChannelStartStopList(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                    ByVal device As Integer, _
                                                    ByVal subsystem As Integer, _
                                                    channels() As Integer, Gains() As Single, _
                                                    ByVal frequency As Single, _
                                                    ByVal SamplesPerChannel As Single, _
                                                    ByVal Buffers As Integer, _
                                                    ByVal BackGroundForeGround As Integer)
        Dim ChannelsSampled As Integer
   
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        'Specify timing event
        AddTimingEventDefault SR, frequency
        
        ' Specify start event
        AddStartEventOnCommand SR 'Start on Software Command
        ' Specify stop event
        AddStopEventOnTerminalCount SR 'Stop on Terminal Count
        
        'Calculate the number of channel to be sampled
        ChannelsSampled = SizeOfStartStopList(SR, LDD, channels(), subsystem)
        
        ' ------------ Select Channel Group ----------------
        ' Specify channels, gain and data format
        AddStartStopList SR, subsystem, channels(), Gains()
        
        ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
        
        'Note: Your application must call the refresh method to execute this function.
        'Note: For output Service Requests make sure to fill the buffers
        '      appropriately before calling .Refresh
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Reads or writes multiple channels using a single buffer.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   channels()              |Two element array containing the
' start channel and the stop channel
' @parm    Single           |   Gains()                 |Two element array containing the
' start channels gain and the stop channel gain
' @parm    Single           |   frequency               |Rate or which to read or write in Hz
' @parm    Single           |   BurstRate               |Burst mode conversion rate in Hz
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXMultiChannelBurstMode> This function sets
'           up a Service Request that either inputs or outputs one buffer
'           from/to a subsystem. This function makes use of a start stop
'           list. The draw back to this is that the user can on specify
'           consecutive channels with a gain setting for the first channel
'           and a gain setting for all of the other channels that are
'           included in the start stop list.
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f PutDriverLINXAIBuffer>,<f GetDriverLINXAIBuffer>,
'           <f SetupDriverLINXMultiChannelGainList>,
'           <f SetupDriverLINXSimultaneousDigitalIO>
'
'
Public Sub SetupDriverLINXMultiChannelBurstMode(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                ByVal device As Integer, _
                                                ByVal subsystem As Integer, _
                                                channels() As Integer, Gains() As Single, _
                                                ByVal frequency As Single, _
                                                ByVal BurstRate As Single, _
                                                ByVal SamplesPerChannel As Single, _
                                                ByVal Buffers As Integer, _
                                                ByVal BackGroundForeGround As Integer)
        Dim ChannelsSampled As Integer
        
        'Calculate the number of channel to be sampled
        ChannelsSampled = SizeOfStartStopList(SR, LDD, channels(), subsystem)
        
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        'Specify timing event
        AddTimingEventBurstMode SR, frequency, BurstRate, ChannelsSampled
        
        ' Specify start event
        AddStartEventOnCommand SR 'Start on Software Command
        ' Specify stop event
        AddStopEventOnTerminalCount SR 'Stop on Terminal Count
        
        ' ------------ Select Channel Group ----------------
        ' Specify channels, gain and data format
        AddStartStopList SR, subsystem, channels(), Gains()
        
        ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
        
        'Note: Your application must call the refresh method to execute this function.
        'Note: For output Service Requests make sure to fill the buffers
        '      appropriately before calling .Refresh
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Reads or writes multiple channels using a single buffer.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   ChannelList             |Array of non-consecutive or
' consecutive channels to read or write
' @parm    Single           |   Gains                   |Array of channel gain value(s)
' @parm    Single           |   frequency               |Rate or which to read or write
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   ChannelsSampled         |Number of Channels in the channel
' gain list
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXMultiChannelGainList> This function sets up a
'           Service Request that either inputs or outputs one buffer from/to
'           a subsystem. This function makes use of a channel gain list which
'           allows the programmer the ability to operate on more than one
'           non consecutive channels that can all have different gain
'           settings.
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f PutDriverLINXAIBuffer>,<f GetDriverLINXAIBuffer>,
'           <f SetupDriverLINXMultiChannelStartStopList>,
'           <f SetupDriverLINXMultiChannelBurstMode>,
'           <f SetupDriverLINXSimultaneousDigitalIO>,
'           <f SetupDriverLINXSingleScanIO>
'
'
Public Sub SetupDriverLINXMultiChannelGainList(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                ByVal device As Integer, _
                                                ByVal subsystem As Integer, _
                                                ChannelList() As Integer, _
                                                Gains() As Single, _
                                                ByVal frequency As Single, _
                                                ByVal SamplesPerChannel As Single, _
                                                ByVal ChannelsSampled As Integer, _
                                                ByVal Buffers As Integer, _
                                                ByVal BackGroundForeGround As Integer)
       'This routine make use of a Channel Gain List
   
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        'Specify timing event
        AddTimingEventDefault SR, frequency
        
        ' Specify start event
        AddStartEventOnCommand SR 'Start on Software Command
        ' Specify stop event
        AddStopEventOnTerminalCount SR 'Stop on Terminal Count
        
        ' ------------ Select Channel Group ----------------
        ' Specify channels, gain and data format
        AddChannelGainList SR, subsystem, ChannelsSampled, ChannelList(), Gains()
        
         ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
        
        'Note: Your application must call the refresh method to execute this function.
        'Note: For output Service Requests make sure to fill the buffers
        '      appropriately before calling .Refresh
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Reads or writes a single channel scan using a single buffer.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   ChannelList             |Array of non-consecutive or
' consecutive channels to read or write
' @parm    Single           |   Gains                   |Array of channel gain value(s)
' @parm    Integer          |   ChannelsSampled         |Number of Channels in the channel
' gain list
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXSingleScanIO> This function sets up a
'           Service Request that either inputs or outputs a single scan of
'           available channels from/to a subsystem. The number of channels
'           does not have to be equal to the available channels supported by
'           the subsystem but the channels sampled should not exceed the total
'           number of channels available to the subsystem. This function makes use
'           of a channel gain list which allows the programmer the ability
'           to operate on more than one non consecutive channels that can
'           all have different gain settings.
'
'
' @devnote  KevinD 4/21/99 2:13:00PM
'
' @xref     <f PutDriverLINXAIBuffer>,<f GetDriverLINXAIBuffer>,
'           <f SetupDriverLINXMultiChannelStartStopList>,
'           <f SetupDriverLINXMultiChannelBurstMode>,
'           <f SetupDriverLINXSimultaneousDigitalIO>,
'           <f SetupDriverLINXMultiChannelGainList>
'
'
Public Sub SetupDriverLINXSingleScanIO(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                ByVal device As Integer, _
                                                ByVal subsystem As Integer, _
                                                ChannelList() As Integer, _
                                                Gains() As Single, _
                                                ByVal ChannelsSampled As Integer, _
                                                ByVal BackGroundForeGround As Integer)
       'This routine make use of a Channel Gain List
       
       Dim Buffers As Integer
       Dim SamplesPerChannel As Integer
       
       Buffers = 1
       SamplesPerChannel = 1
   
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        'Specify timing event
        AddTimingEventNullEvent SR
        ' Specify start event
        AddStartEventOnCommand SR 'Start on Software Command
        ' Specify stop event
        AddStopEventNullEvent SR
        
        ' ------------ Select Channel Group ----------------
        ' Specify channels, gain and data format
        AddChannelGainList SR, subsystem, ChannelsSampled, ChannelList(), Gains()
        
         ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Synchronous tasks do not need DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, False
        
        'Note: Your application must call the refresh method to execute this function.
        'Note: For output Service Requests make sure to fill the buffers
        '      appropriately before calling .Refresh
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc UtilityFunctions
'
' @func     Returns the result from a single analog input read.
'
' @rdesc    Single - returns the analog value scaled in volts.
'
' @parm     DriverLINXSR    |   SR                      |Name of the control
'
' @comm     <f GetDriverLINXAISingleValue> This function returns the result
'           of a single-value, analog input Service Request.
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f GetDriverLINXDISingleValue>, <f GetDriverLINXDIBuffer>,
'           <f GetDriverLINXAIBuffer>, <f PutDriverLINXAIBuffer>,
'           <f WriteDriverLINXDIBuffer>
'
'
Public Function GetDriverLINXAISingleValue(SR As DriverLINXSR _
                                            ) As Single

    Dim ResultInVolts(0 To 0) As Single
    Dim ErrorCode As Long
    With SR
        ' Only call this function after executing a single-value, analog input
        '   Service Request
        If .Req_subsystem <> DL_AI Then
            ErrorHandler "GetDriverLINXAISingleValue", 2, _
                                            "Service Request Subsystem is not DL_AI!"
        Else
            ' Use the VBArrayBufferConvert method to convert this
            '   value from board units to volts
            ErrorCode = .VBArrayBufferConvert(0, 0, 1, ResultInVolts(), DL_tSINGLE, 0, 0)
            If ErrorCode <> 1 Then   'Call Error Handler
                ErrorHandler "GetDriverLINXAISingleValue", 1, "Can't convert data!"
            End If
        End If
    End With
    
    GetDriverLINXAISingleValue = ResultInVolts(0)
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc UtilityFunctions
'
' @func     Fills a VB array with acquired data.
'
' @rdesc    Long - returns the status of the data transfer.
'
' @parm     DriverLINXSR    |   SR                      |Name of the control
' @parm     Integer         |   nBuffer                 |Buffer number to convert
' @parm     Integer         |   VBArray                 |Name VB array to place converted data
'
' @comm     <f GetDriverLINXAIBuffer> This function returns buffer filled by an
'           analog input Service Request. The returned values are scaled in volts.
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f GetDriverLINXDISingleValue>, <f GetDriverLINXDIBuffer>,
'           <f GetDriverLINXAISingleValue>, <f PutDriverLINXAIBuffer>,
'           <f WriteDriverLINXDIBuffer>, <f GetDriverLINXAIPartialBuffer>
'
'

Public Function GetDriverLINXAIBuffer(SR As DriverLINXSR, ByVal nBuffer As Integer, _
                                        VBArray() As Single) As Long

    Dim nSamples As Long
    
    With SR
        nSamples = .Sel_buf_samples
        'Convert raw counts to screen units
        GetDriverLINXAIBuffer = _
                .VBArrayBufferConvert(nBuffer, 0, nSamples, VBArray(), DL_tSINGLE, 0#, 0#)
            If GetDriverLINXAIBuffer <> 1 Then   'Call Error Handler
                ErrorHandler "GetDriverLINXAIBuffer", 1, "Can't convert data!"
            End If
    End With
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc UtilityFunctions
'
' @func     Fills a VB array with acquired data.
'
' @rdesc    Long - returns the status of the data transfer.
'
' @parm     DriverLINXSR    |   SR                      |Name of the control
' @parm     Integer         |   nBuffer                 |Buffer number to convert
' @parm     Long            |   samples                 |Number of samples to convert
' @parm     Integer         |   VBArray                 |Name VB array to place converted data
'
' @comm     <f GetDriverLINXAIPartialBuffer> This function returns buffer filled by an
'           analog input Service Request. The returned values are scaled in volts.
'
'
' @devnote  KevinD 7/28/99 2:20:00PM
'
' @xref     <f GetDriverLINXDISingleValue>, <f GetDriverLINXDIBuffer>,
'           <f GetDriverLINXAISingleValue>, <f PutDriverLINXAIBuffer>,
'           <f WriteDriverLINXDIBuffer>, <f GetDriverLINXAIBuffer>
'
'

Public Function GetDriverLINXAIPartialBuffer(SR As DriverLINXSR, ByVal nBuffer As Integer, _
                                        samples As Long, VBArray() As Single) As Long

    
    With SR
        'Convert raw counts to screen units
        GetDriverLINXAIPartialBuffer = _
                .VBArrayBufferConvert(nBuffer, 0, samples, VBArray(), DL_tSINGLE, 0#, 0#)
            If GetDriverLINXAIPartialBuffer <> 1 Then   'Call Error Handler
                ErrorHandler "GetDriverLINXAIPartialBuffer", 1, "Can't convert data!"
            End If
    End With
    
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc UtilityFunctions
'
' @func     Writes data to the output buffer.
'
' @rdesc    Long - returns the result of the conversion.
'
' @parm     DriverLINXSR    |   SR                      |Name of the control
' @parm     Integer         |   nBuffer                 |Buffer number to write to
' @parm     Integer         |   nSamples                |Number of samples to transfer
' to DriverLINX
' @parm     Integer         |   VBArray                 |Name VB array to transfer
'
' @comm     <f PutDriverLINXAIBuffer> This function converts user input
'           voltage(s) and converts it to D/A units and writes the data to the
'           output buffer.
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f GetDriverLINXDISingleValue>, <f GetDriverLINXDIBuffer>,
'           <f GetDriverLINXAISingleValue>, <f GetDriverLINXAIBuffer>,
'           <f WriteDriverLINXDIBuffer>
'
'
Public Function PutDriverLINXAIBuffer(SR As DriverLINXSR, ByVal nBuffer As Integer, _
                                ByVal nSamples As Integer, VBArray() As Single) As Long
    
    With SR
        'Convert raw counts to screen units
        PutDriverLINXAIBuffer = _
                    .VBArrayBufferConvert(nBuffer, 0, nSamples, VBArray(), DL_tSINGLE, 0#, 0#)
            If PutDriverLINXAIBuffer <> 1 Then   'Call Error Handler
                ErrorHandler "PutDriverLINXAIBuffer", 1, "Can't convert data!"
            End If
    End With
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc UtilityFunctions
'
' @func     Returns the result code and the associated text message.
'
' @rdesc    Integer - returns the result of the operation.
'
' @parm     DriverLINXSR    |   SR                      |Name of the control
' @parm     String          |   Status                  |String to receive the Status
'
' @comm     <f GetDriverLINXStatus> This function returns the result code for
'           the operation and also returns a string containing a text message for
'           the associated result code.
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f ShowDriverLINXStatus>
'
'
Public Function GetDriverLINXStatus(SR As DriverLINXSR, Status As String _
                                    ) As Integer

    With SR
        ' DriverLINX returns result codes in the Res_result property.
        GetDriverLINXStatus = .Res_result
        
        ' DriverLINX returns status messages in the Message property.
        Status = .Message
    End With
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks DriverLINX for errors.
 '
 ' @parm    DriverLINXSR   |   SR          |Name of the control
 '
 ' @comm    <f ShowDriverLINXStatus> This function returns the result code
 '          for the operation. If an error occurred, this function displays
 '          the error in a message box.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f GetDriverLINXStatus>
 '
 '
Public Function ShowDriverLINXStatus(SR As DriverLINXSR _
                                    ) As Integer

    Dim DLResultCode As Integer
    Dim OriginalReq_op As Integer
    
    With SR
        ' DriverLINX returns result codes in the Res_result property.
        DLResultCode = .Res_result
        
        If DLResultCode <> DL_NoErr Then
            ' DriverLINX can display a message box that describes
            '   the status of a Service Request.
            
            ' First save the current Req_op property value
            OriginalReq_op = .Req_op
            
            ' Then change the value of the Req_op property to
            '   DL_MESSAGEBOX, and call the Refresh method.
            .Req_op = DL_MESSAGEBOX
            .Refresh
            
            ' Afterwards, restore Req_op to its original value
            .Req_op = OriginalReq_op
        End If
    End With
       
    ' Return the result code
    ShowDriverLINXStatus = DLResultCode

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Stops an active task.
 '
 ' @parm    DriverLINXSR   |   SR          |Name of the control
 '
 ' @comm    <f StopDriverLINXIO> This function stops an active DriverLINX
 '          task.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddStopEventOnCommand>
 '
Public Function StopDriverLINXIO(SR As DriverLINXSR _
                                                ) As Integer
    Dim OriginalReq_op As Integer
    
    With SR
        ' First save the current Req_op property value
        OriginalReq_op = .Req_op
        
        ' Then change the value of the Req_op property to
        ' DL_STOP, and call the Refresh method.
        .Req_op = DL_STOP
        .Refresh
        
        ' Afterwards, restore Req_op to its original value
        .Req_op = OriginalReq_op
    
        ' Return the result code
        StopDriverLINXIO = .Res_result
    End With

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Closes a DriverLINX driver.
 '
 ' @parm    DriverLINXSR   |   SR          |Name of the control
 '
 ' @comm    <f CloseDriverLINXDriver> This function closes a DriverLINX
 '          driver.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f OpenDriverLINXDriver>
 '
Public Sub CloseDriverLINXDriver(SR As DriverLINXSR)
    'If Service Request allocated buffers they should be deleted.
    If (SR.Sel_buf_N > 0) Then SR.Sel_buf_N = 0
    
    ' To close a DriverLINX driver,
    ' just set the Req_DLL_name property
    ' equal to an empty string.
    SR.Req_DLL_name = ""

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks whether analog value is within the channel's range.
 '
 ' @rdesc   Boolean - returns true if user input is in range.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   Subsystem       |Subsystem to check
 ' @parm    Integer         |userinput          |Value to check if in range
 ' @parm    Single          |Gain               |Gain to check range for
 '
 ' @comm    <f IsInDriverLINXAnalogRange> This function checks the userinput
 '           argument to see if it is within a channel's acceptable voltage
 '           limits. This function first determines if the device supports a
 '           channel gain list. If the device supports a channel gain list
 '           the user input is compared to limits based on the Gain Multiplier
 '           Table. If the channel gain list is not supported the user input
 '           is compared to the values stored in the Min/Max Range Table.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f IsInDriverLINXDigitalRange>,
 '          <f IsInDriverLINXExtendedDigitalRange>,
 '          <f DoesDeviceSupportAnalogChannelGainList>,
 '          <f ConvertVoltsToADUnits>
 '
Public Function IsInDriverLINXAnalogRange(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                            ByVal subsystem As Integer, _
                                            ByVal UserInput As Single, _
                                            gain As Single _
                                            ) As Boolean
                                            
    'Function returns true if value submitted falls between
    'the allowable low and high range analog limits.
    Dim GainList As Boolean
    Dim Min As Single
    Dim Max As Single
    Dim bBipolar As Boolean
    Dim i As Integer
    Dim bMatch As Boolean
    
    bMatch = False  'Initialize variable to false
    Min = 0
    Max = 0
    
    If (gain < 0) Then
        bBipolar = True
    Else
        bBipolar = False
    End If
    
    With LDD
    .device = SR.Req_device
    .Req_DLL_name = SR.Req_DLL_name 'open the LDD
    
    'First check to see if a channel gain list is supported
    GainList = DoesDeviceSupportAnalogChannelGainList(SR, LDD, subsystem)
         
    If subsystem = DL_AO Then
        If GainList Then
           i = 0
           Do
            If (Abs(.AO_GM_mul(i) - Abs(gain)) < 0.1) Then 'We have a match check for bipolar value
                Min = .AO_GM_min(i)
                Max = .AO_GM_max(i)
                'Check to see that we have obtained the correct data from
                'channel gain table.
                Select Case bBipolar
                Case True
                    If Min < 0 Then bMatch = True
                Case False
                    If Min >= 0 Then bMatch = True
                End Select
                
            End If
            i = i + 1
           Loop Until (bMatch = True) Or (i = LDD.AO_GM_n)
           
           'no that we min & max values check userinput
       If bMatch = True Then
            If (UserInput >= Min) And (UserInput <= Max) Then
                 IsInDriverLINXAnalogRange = True
            Else
                 IsInDriverLINXAnalogRange = False
            End If
        Else
                IsInDriverLINXAnalogRange = False
        End If
           
        Else
            If (UserInput >= .AO_MM_min(0)) And (UserInput <= .AO_MM_max(0)) Then
                IsInDriverLINXAnalogRange = True
            Else
                IsInDriverLINXAnalogRange = False
            End If
        End If
    ElseIf subsystem = DL_AI Then
        If GainList Then
           i = 0
           Do
            If (Abs(.AI_GM_mul(i) - Abs(gain)) < 0.1) Then 'We have a match check for bipolar value
                Min = .AI_GM_min(i)
                Max = .AI_GM_max(i)
                'Check to see that we have obtained the correct data from
                'channel gain table.
                Select Case bBipolar
                Case True
                    If Min < 0 Then bMatch = True
                Case False
                    If Min >= 0 Then bMatch = True
                End Select
                
            End If
            i = i + 1
           Loop Until (bMatch = True) Or (i = LDD.AI_GM_n)
           
           'no that we min & max values check userinput
       If bMatch = True Then
            If (UserInput >= Min) And (UserInput <= Max) Then
                 IsInDriverLINXAnalogRange = True
            Else
                 IsInDriverLINXAnalogRange = False
            End If
        Else
                IsInDriverLINXAnalogRange = False
        End If
        Else
            If (UserInput >= .AI_MM_min(0)) And (UserInput <= .AI_MM_max(0)) Then
                IsInDriverLINXAnalogRange = True
            Else
                IsInDriverLINXAnalogRange = False
            End If
        End If
    Else
        IsInDriverLINXAnalogRange = False
    End If
    ' Close the LDD's driver
            .Req_DLL_name = ""
    End With
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks if device supports specified clock mode.
 '
 ' @rdesc   Boolean - returns true if a analog channel gain list is supported.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   ClockMode       |requested clock mode
 ' @parm    Integer         |   ClockChan       |clock channel to check
 '
 ' @comm    <f DoesDeviceSupportRateMode> function determines
 '           if the counter/timer clock channel supports the
 '           requested clock mode.
 '
 ' @devnote RoyF 9/23/99 11:40:00AM
 '
 ' @xref
 '
 '
Public Function DoesDeviceSupportRateMode(SR As DriverLINXSR, _
                                          LDD As DriverLINXLDD, _
                                          ByVal ClockMode As Integer, _
                                          ByVal ClockChan As Integer _
                                         ) As Boolean
    Dim bRateMode As Boolean
    bRateMode = False  'Initially set to false

    With LDD
        .device = SR.Req_device
        .Req_DLL_name = SR.Req_DLL_name 'open the LDD
        If .Req_DLL_name = "" Then     ' unexpected: LDD not found
            DoesDeviceSupportRateMode = False
            Exit Function
        End If

        ' Is clock channel in range?
        If ClockChan < 0 Or ClockChan >= .CT_nChan Then
            DoesDeviceSupportRateMode = False
            Exit Function
        End If

        ' Is requested clock mode supported?
        If .CT_Modes(ClockChan) And (2 ^ ClockMode) <> 0 Then
            bRateMode = True
        End If

    End With
    
    DoesDeviceSupportRateMode = bRateMode
                                                        
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks whether subsystem supports a channel gain list.
 '
 ' @rdesc   Boolean - returns true if a analog channel gain list is supported.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   Subsystem       |Subsystem to check
 '
 ' @comm    <f DoesDeviceSupportAnalogChannelGainList> Function determines
 '           if the analog subsystem supports a channel gain list by querying
 '           the LDD. Function is only valid for analog IO subsystems.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f IsInDriverLINXAnalogRange>,<f ConvertVoltsToADUnits>
 '
 '
Public Function DoesDeviceSupportAnalogChannelGainList(SR As DriverLINXSR, _
                                                        LDD As DriverLINXLDD, _
                                                        subsystem As Integer _
                                                        ) As Boolean
    Dim GainList As Boolean
    With LDD
        .device = SR.Req_device
        .Req_DLL_name = SR.Req_DLL_name 'open the LDD
    
        GainList = False    'Initially set to false
    
        If subsystem = DL_AI Then
            If (.AI_MaxCG > 0) Then GainList = True
        ElseIf subsystem = DL_AO Then
            If (.AO_GM_defined = True) Then
                If (.AO_MaxCG > 0) Then GainList = True
            End If
        Else
            GainList = False    'Invalid subsystem
        End If
        
    End With
    
    DoesDeviceSupportAnalogChannelGainList = GainList
                                                        
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Returns the result from a single digital input read.
 '
 ' @rdesc   Integer - returns the digital value.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 '
 ' @comm    <f GetDriverLINXDISingleValue> This function returns the result
 '           of a single-value, digital input Service Request.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f GetDriverLINXAISingleValue>, <f GetDriverLINXDIBuffer>,
 '          <f GetDriverLINXAIBuffer>, <f PutDriverLINXAIBuffer>,
 '          <f WriteDriverLINXDIBuffer>
 '
 '
Public Function GetDriverLINXDISingleValue(SR As DriverLINXSR _
                                            ) As Integer
    
    With SR
        ' Only call this function after executing a single-value, digital input
         '  Service Request
        If .Req_subsystem <> DL_DI Then
            ErrorHandler "GetDriverLINXDISingleValue", 2, _
                                            "Service Request Subsystem is not DL_DI!"
        Else
            ' Read value out
            GetDriverLINXDISingleValue = .Res_Sta_ioValue
        End If
    End With
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks whether digital value is within the channels range.
 '
 ' @rdesc   Boolean - returns true if user input is in range.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   Subsystem       |Subsystem to check
 ' @parm    Integer         |userinput          |Value to check if in range
 ' @parm    Integer         |Channel            |Channel to check range for
 '
 ' @comm    <f IsInDriverLINXDigitalRange> This function queries the LDD and
 '          compares the user input with the channels acceptable limits.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f IsInDriverLINXAnalogRange>,
 '          <f IsInDriverLINXExtendedDigitalRange>,
 '          <f ConvertVoltsToADUnits>
 '
 '
Public Function IsInDriverLINXDigitalRange(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                            ByVal subsystem As Integer, _
                                            ByVal UserInput As Integer, _
                                            Channel As Integer _
                                            ) As Boolean
    
    'Function returns true if value submitted falls between
    'the allowable low and high range digital limits.
    
    With LDD
    .device = SR.Req_device
    .Req_DLL_name = SR.Req_DLL_name 'open the LDD
        If SR.Req_subsystem <> DL_DO Then
            ErrorHandler "IsInDriverLINXDigitalRange", 2, _
                                        "Service Request Subsystem is not DL_DO!"
        Else
            If (UserInput >= 0) And (UserInput <= .DO_Mask(Channel)) Then
                IsInDriverLINXDigitalRange = True
            Else
                IsInDriverLINXDigitalRange = False
            End If
        End If

    ' Close the LDD's driver
            .Req_DLL_name = ""
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Reads or writes one or more buffers of data from/to
'          a digital subsystem using an external clock.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Channel                 |Channel to Acquire/Write
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXDigitalIOBuffer> This function sets up a
'           Service Request that either inputs or outputs one or more buffers
'           of data from/to a digital subsystem. The Service Requests
'           starts when submitted and stops when the buffer is either full
'           or empty depending whether it is an input or output task.
'           Data is clocked in/out using an external clock. Function only
'           will work if board supports an external clock and if the board
'           supports interrupt or DMA data transfer.
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f SetupDriverLINXContinuousDigitalIO>, <f WriteDriverLINXDIBuffer>,
'           <f GetDriverLINXDIBuffer>
'
'
Public Sub SetupDriverLINXDigitalIOBuffer(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                            ByVal device As Integer, _
                                            ByVal Channel As Integer, _
                                            SamplesPerChannel As Integer, _
                                            subsystem As Integer, Buffers As Integer, _
                                            BackGroundForeGround As Integer)

' Use this procedure to read/write a
' data array to a digital I/O subsystem
Dim gain As Single
Dim ChannelsSampled As Integer
ChannelsSampled = 1     'one channel only
gain = 0    'Gain is 0 for digital I/O

  ' ------- Service Request Group -------------
  AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
  
   ' ------------- Event Group -----------------
  ' Setup external interrupt as timing clock to be used with an external clock
  ' DriverLINX will sample data a the external clock rate
  ' Add Timing event
  AddTimingEventExternalDigitalClock SR
  
  ' Specify start event
  AddStartEventOnCommand SR 'Start on Software Command
  ' Specify stop event
  AddStopEventOnTerminalCount SR 'Stop on Terminal Count
  
  ' ------------ Select Channel Group ----------------
  ' Specify channels, gain and data format
  AddSelectSingleChannel SR, Channel, subsystem, gain
  
  ' ------------ Select Buffer Group ----------------
    AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
  ' ------------ Select Flags ----------------
  ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
  
    'Note: Your application must call the refresh method to execute this function.
    'Note: For output Service Requests make sure to fill the output buffer
    '      appropriately before calling .Refresh
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc UtilityFunctions
'
' @func     Fills a VB array with acquired data.
'
' @rdesc    Long - returns the status of the data transfer.
'
' @parm     DriverLINXSR    |   SR                      |Name of the control
' @parm     Integer         |   nBuffer                 |Buffer number to convert
' @parm     Integer         |   VBArray                 |Name VB array to place converted data
' @parm     Integer         |   nSamples                |Number of Samples to convert
'
' @comm     <f GetDriverLINXDIBuffer> This function returns a buffer filled by a
'           digital input Service Request.
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f GetDriverLINXDISingleValue>, <f GetDriverLINXAIBuffer>,
'           <f GetDriverLINXAISingleValue>, <f PutDriverLINXAIBuffer>,
'           <f WriteDriverLINXDIBuffer>
'
'
Public Function GetDriverLINXDIBuffer(SR As DriverLINXSR, ByVal nBuffer As Integer, _
                                        VBArray() As Byte, nSamples As Integer _
                                        ) As Long

    With SR
          'GetDriverLINXDIBuffer returns the status of the data transfer
           GetDriverLINXDIBuffer = .VBArrayBufferXfer(nBuffer, VBArray, DL_BufferToVBArray)
          If GetDriverLINXDIBuffer <> 0 Then   'Call Error Handler
                ErrorHandler "GetDriverLINXDIBuffer", 1, "Can't convert data!"
          End If
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc UtilityFunctions
'
' @func     Writes data to a digital output buffer.
'
' @rdesc    Long - returns the result of the conversion.
'
' @parm     DriverLINXSR    |   SR                      |Name of the control
' @parm     Integer         |   nBuffer                 |Buffer number to write to
' @parm     Integer         |   VBArray                 |Name VB array to transfer
' @parm     Integer         |   nSamples                |Number of samples to transfer
' to DriverLINX
'
' @comm     <f WriteDriverLINXDIBuffer> This function writes a user buffer to a DriverLINX
'           output buffer.
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f GetDriverLINXDISingleValue>, <f GetDriverLINXDIBuffer>,
'           <f GetDriverLINXAISingleValue>, <f GetDriverLINXAIBuffer>,
'           <f PutDriverLINXAIBuffer>
'
'
Public Function WriteDriverLINXDIBuffer(SR As DriverLINXSR, ByVal nBuffer As Integer, _
                                        VBArray() As Byte, nSamples As Integer) As Long

    With SR
          'WriteDriverLINXDIBuffer returns the status of the data transfer
          WriteDriverLINXDIBuffer = .VBArrayBufferXfer(nBuffer, VBArray, DL_VBArrayToBuffer)
          If WriteDriverLINXDIBuffer <> 0 Then   'Call Error Handler
                ErrorHandler "WriteDriverLINXDIBuffer", 1, "Can't convert data!"
          End If
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks whether subsystem supports DMA mode.
 '
 ' @rdesc   Boolean - returns true if device and subsystem supports DMA.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   Subsystem       |Subsystem to check
 '
 ' @comm    <f DoesDeviceSupportDMA> Function determines if the subsystem
 '          supports DMA by querying the LDD.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f DoesDeviceSupportIRQ>
 '
 '
Public Function DoesDeviceSupportDMA(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                        subsystem As Integer _
                                        ) As Boolean
                                        
    'Function can be used to see whether subsystem supports DMA data transfer
     With LDD
    .device = SR.Req_device
    .Req_DLL_name = SR.Req_DLL_name 'open the LDD
     
   Select Case subsystem        'The following determines if a subsystem supports DMA
   Case DL_AI
    If .AI_DMA0 <> -1 Then
        DoesDeviceSupportDMA = True
    Else
        DoesDeviceSupportDMA = False
    End If
   Case DL_AO
    If .AO_DMA0 <> -1 Then
        DoesDeviceSupportDMA = True
    Else
        DoesDeviceSupportDMA = False
    End If
   Case DL_DI
    If .DI_DMA0 <> -1 Then
        DoesDeviceSupportDMA = True
    Else
        DoesDeviceSupportDMA = False
    End If
   Case DL_DO
    If .DO_DMA0 <> -1 Then
        DoesDeviceSupportDMA = True
    Else
        DoesDeviceSupportDMA = False
    End If
   Case DL_CT
        DoesDeviceSupportDMA = False
   End Select

    ' Close the LDD's driver
            .Req_DLL_name = ""
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks whether subsystem supports IRQ mode.
 '
 ' @rdesc   Boolean - returns true if device and subsystem supports IRQ.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   Subsystem       |Subsystem to check
 '
 ' @comm    <f DoesDeviceSupportIRQ> Function determines if the subsystem
 '          supports IRQ by querying the LDD.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f DoesDeviceSupportDMA>
 '
 '
Public Function DoesDeviceSupportIRQ(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                        subsystem As Integer _
                                        ) As Boolean
    'Function can be used to see whether subsystem supports IRQ data transfer
    With LDD
    .device = SR.Req_device
    .Req_DLL_name = SR.Req_DLL_name 'open the LDD
     
   Select Case subsystem            'The following determines if a subsystem supports IRQ
   Case DL_AI
    If .AI_IRQ <> -1 Then
        DoesDeviceSupportIRQ = True
    Else
        DoesDeviceSupportIRQ = False
    End If
   Case DL_AO
    If .AO_IRQ <> -1 Then
        DoesDeviceSupportIRQ = True
    Else
        DoesDeviceSupportIRQ = False
    End If
   Case DL_DI
    If .DI_IRQ <> -1 Then
        DoesDeviceSupportIRQ = True
    Else
        DoesDeviceSupportIRQ = False
    End If
   Case DL_DO
    If .DO_IRQ <> -1 Then
        DoesDeviceSupportIRQ = True
    Else
        DoesDeviceSupportIRQ = False
    End If
   Case DL_CT
    If .CT_IRQ <> -1 Then
        DoesDeviceSupportIRQ = True
    Else
        DoesDeviceSupportIRQ = False
    End If
   End Select

    ' Close the LDD's driver
            .Req_DLL_name = ""
    End With
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Reads or writes one or more buffers of data from/to
'          a digital subsystem using an external clock.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   ChannelList()           |Array of non-consecutive or
' consecutive channels to read or write
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   ChannelsSampled         |Number of Channels in the channel
' gain list
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXMultiChannelDigitalGainList> This function sets
'           up a Service Request that either inputs or outputs one buffer
'           from/to a subsystem. This function makes use of a channel gain
'           list which allows the programmer the ability to operate on more
'           than one non consecutive channels.
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
'

Public Sub SetupDriverLINXMultiChannelDigitalGainList(SR As DriverLINXSR, _
                                                LDD As DriverLINXLDD, _
                                                ByVal device As Integer, _
                                                ByVal subsystem As Integer, _
                                                ChannelList() As Integer, _
                                                ByVal SamplesPerChannel As Single, _
                                                ByVal ChannelsSampled As Integer, _
                                                ByVal Buffers As Integer, _
                                                ByVal BackGroundForeGround As Integer)
        'This routine make use of a Channel Gain List
        Dim Gains() As Single
        ReDim Gains(ChannelsSampled) As Single
        Dim i As Integer
        For i = 0 To ChannelsSampled - 1
            Gains(i) = 0    'Zero out the gain settings, Digital IO does not use gain
        Next i
        
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        'Specify timing event
        AddTimingEventExternalDigitalClock SR  'Add timing based on external clock pulse
        
        ' Specify start event
        AddStartEventOnCommand SR 'Start on Software Command
        ' Specify stop event
        AddStopEventOnTerminalCount SR 'Stop on Terminal Count
        
        ' ------------ Select Channel Group ----------------
        ' Specify channels, gain and data format
        AddChannelGainList SR, subsystem, ChannelsSampled, ChannelList(), Gains()
        
         ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Reads or writes one or more buffers of data from/to
'          a digital subsystem using an external clock.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   channels()              |Two element array containing the
' start channel and the stop channel
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXMultiChannelDigitalStartStopList> This function sets
'           up a Service Request that either inputs or outputs multiple channels
'           to one or more buffers. The Service Requests uses a start stop list and
'           starts when submitted and stops when the buffer(s) is either full
'           or empty depending whether it is an input or output task.
'           Data is clocked in/out using an external clock. Function only
'           will work if board supports an external clock and if the board
'           supports Interrupt or DMA data transfer.
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
'
Public Sub SetupDriverLINXMultiChannelDigitalStartStopList(SR As DriverLINXSR, _
                                                LDD As DriverLINXLDD, _
                                                ByVal device As Integer, _
                                                ByVal subsystem As Integer, _
                                                channels() As Integer, _
                                                ByVal SamplesPerChannel As Integer, _
                                                ByVal Buffers, _
                                                ByVal BackGroundForeGround As Integer)
    Dim ChannelsSampled, i As Integer
    Dim Gains(1) As Single
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        'Specify timing event
        AddTimingEventExternalDigitalClock SR  'Add timing based on external clock pulse
        
        ' Specify start event
        AddStartEventOnCommand SR 'Start on Software Command
        ' Specify stop event
        AddStopEventOnTerminalCount SR 'Stop on Terminal Count
        
        'Calculate the number of channel to be sampled
        ChannelsSampled = SizeOfStartStopList(SR, LDD, channels(), subsystem)
        
        'For Digital Tasks the Gains() are not passed because they are not used for
        'Digital I/O fill the gains() with zeros
        For i = 0 To 1  'Two element array in a start stop list
            Gains(i) = 0
        Next i
        
        ' ------------ Select Channel Group ----------------
        ' Specify channels, gain and data format
        AddStartStopList SR, subsystem, channels(), Gains()
        
         ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Reads or writes multiple digital channels simultaneously using
'          a single buffer.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   channels()              |Two element array containing the
' start channel and the stop channel
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
' @parm    Boolean          |   Simultaneous            |Select Simultaneous Sampling or One
' Sample per Clock Tic
'
' @comm     <f SetupDriverLINXSimultaneousDigitalIO> This function
'           sets up a Service Request that either inputs or outputs one
'           buffer from/to a subsystem. This function makes use of a start
'           stop list. The draw back to this is that the user can only
'           specify consecutive channels. This function also uses an external
'           trigger to acquire digital data. This function only works with
'           digital I/O boards that support an external clock in conjunction
'           with the Digital I/O channels. This function also allows the user
'           to select whether to sample all channels in the start/stop list
'           simultaneously (or as close together as possible), or at a rate
'           equal to one sample per clock tic.
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f GetDriverLINXDIBuffer>, <f GetDriverLINXDISingleValue>,
'           <f SetupDriverLINXMultiChannelDigitalGainList>,
'           <f SetupDriverLINXMultiChannelStartStopList>,
'           <f SetupDriverLINXMultiChannelDigitalStartStopList>,
'           <f SetupDriverLINXSingleScanIO>
'
'
'
Public Sub SetupDriverLINXSimultaneousDigitalIO(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                ByVal device As Integer, _
                                                ByVal subsystem As Integer, _
                                                channels() As Integer, _
                                                ByVal SamplesPerChannel As Integer, _
                                                ByVal Buffers As Integer, _
                                                ByVal BackGroundForeGround As Integer, _
                                                ByVal Simultaneous As Boolean)
    'Similar to SetupDriverLINXMultiChannelDigitalStartStopList except that upon each external
    'trigger the board will sample all channels specified in the Start/Stop List.
    Dim ChannelsSampled, i As Integer
    Dim Gains(1) As Single
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        'Specify timing event
        AddTimingEventExternalDigitalClock SR  'Add timing based on external clock pulse
        
        ' Specify start event
        AddStartEventOnCommand SR 'Start on Software Command
        ' Specify stop event
        AddStopEventOnTerminalCount SR 'Stop on Terminal Count
        
        'Calculate the number of channel to be sampled
        ChannelsSampled = SizeOfStartStopList(SR, LDD, channels(), subsystem)
        
        'For Digital Tasks the Gains() are not passed because they are not used for
        'Digital I/O fill the gains() with zeros
        For i = 0 To 1  'Two element array in a start stop list
            Gains(i) = 0
        Next i
        
        ' ------------ Select Channel Group ----------------
        ' Specify channels, gain and data format
        AddStartStopList SR, subsystem, channels(), Gains()
        'Specify the channels to be input or output simultaneously
        AddSelectSimultaneous SR, Simultaneous
        
         ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
    
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Continuously reads or writes one or more buffers of data from/to
'          a subsystem using the onboard clock.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   Channel                 |Channel to read or write
' @parm    Single           |   gain                    |Channels gain value
' @parm    Single           |   frequency               |Rate or which to read or write
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXContinuousBufferedIO> This function sets up a
'           Service Request that either inputs or outputs one or more buffers
'           of data from/to a subsystem. The Service Requests starts when
'           submitted and stops when a the Service Request operation field
'           is changed to Stop  Data is clocked in/out using the
'           boards default clock. Function only will work if board has
'           an onboard clock and if the board supports interrupt
'           or DMA data transfer.
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f PutDriverLINXAIBuffer>,<f GetDriverLINXAIBuffer>,
'           <f SetupDriverLINXBufferedIO>,
'           <f SetupDriverLINXContinuousBufferedAIAnalogTrigger>
'
'
Public Sub SetupDriverLINXContinuousBufferedIO(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                ByVal device As Integer, _
                                                ByVal subsystem As Integer, _
                                                ByVal Channel As Integer, _
                                                ByVal gain As Single, _
                                                ByVal frequency As Single, _
                                                ByVal SamplesPerChannel As Integer, _
                                                ByVal Buffers As Integer, _
                                                ByVal BackGroundForeGround As Integer)
'Subroutine is the same as SetupDriverLINXBufferIO except the Stop event is different
    Dim ChannelsSampled As Integer
    ChannelsSampled = 1 'One channel in this case
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        AddTimingEventDefault SR, frequency
        
        ' Specify start event
        AddStartEventOnCommand SR 'Start on Software Command
        
        ' Specify stop event
        AddStopEventOnCommand SR  'use Software Command to stop acquisition
        
        ' ------------ Select Group ----------------
        ' Specify channels, gain and data format
        AddSelectSingleChannel SR, Channel, subsystem, gain
        
        ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Continuously reads or writes one or more buffers of data from/to
'          two different subsystems using the onboard clock.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXSR     |   SR1                     |Name of the control to
' synchronize task with
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   Channel                 |Channel to read or write
' @parm    Single           |   gain                    |Channels gain value
' @parm    Single           |   frequency               |Rate or which to read or write
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXContinuousSynchronizedBufferedIO> This
'           function sets up a Service Request that inputs and outputs data
'           concurrently. This Service Requests uses a timing event that
'           synchronizes on subsystems clock to that of the other. The result
'           is when pacing Service Request starts this service request starts
'           as long as it was previously refreshed. The Service Request
'           will stop when the operation is changed to stop. Data is clocked
'           in/out using the boards default clock. Function only will
'           work if board has an onboard clock and if the board supports
'           interrupt or DMA data transfer. SR1 is a the Service Request
'           whose clock you wish pSR to use for synchronization.
'
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f PutDriverLINXAIBuffer>,<f GetDriverLINXAIBuffer>,
'           <f SetupDriverLINXBufferedIO>
'
'
Public Sub SetupDriverLINXContinuousSynchronizedBufferedIO(SR As DriverLINXSR, _
                                                        SR1 As DriverLINXSR, _
                                                        LDD As DriverLINXLDD, _
                                                        ByVal device As Integer, _
                                                        ByVal subsystem As Integer, _
                                                        ByVal Channel As Integer, _
                                                        ByVal gain As Single, _
                                                        ByVal SamplesPerChannel As Integer, _
                                                        ByVal Buffers As Integer, _
                                                        ByVal BackGroundForeGround As Integer)
'This Service request is similar to SetupDriverLINXContinuousBufferedIO except that the
'timing event is different. I uses a previously defined clock so that analog synchronization
'is possible.
    Dim ChannelsSampled As Integer
    Dim Clock As Integer
    ChannelsSampled = 1 'One channel in this case
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        'Retrieve Analog Input subsytems default clock
        Clock = GetSubSystemsDefaultClock(SR1, LDD, SR1.Req_subsystem)
        
        ' ------------- Event Group -----------------
        AddTimingEventSyncIO SR, Clock
       
        
        ' Specify start event
        AddStartEventOnCommand SR 'Start on Software Command
        
        ' Specify stop event
        AddStopEventOnCommand SR  'use Software Command to stop acquisition
        
        ' ------------ Select Group ----------------
        ' Specify channels, gain and data format
        AddSelectSingleChannel SR, Channel, subsystem, gain
        
        ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Continuously reads or writes one or more buffers of data from/to
'          a subsystem using the onboard clock.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   Channel                 |Channel to read or write
' @parm    Single           |   gain                    |Channels gain value
' @parm    Single           |   frequency               |Rate or which to read or write
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
' @parm    Integer          |   TrigChan                |Analog trigger channel
' @parm    Single           |   TrigGain                |Trigger channels gain value
' @parm    Single           |   UpperThresholdVoltage   |Upper threshold voltage
' @parm    Single           |   LowerThresholdVoltage   |Lower threshold voltage
' @parm    Integer          |   Slope                   |Define the trigger
'
'
' @comm     <f SetupDriverLINXContinuousBufferedAIAnalogTrigger> This
'           function sets up a Service Request that either inputs or
'           outputs one or more buffers of data from/to a subsystem.
'           The Service Requests starts on an the receipt of a analog
'           trigger and stops when a the Service Request operation field
'           is changed to Stop.  Data is clocked in/out using the
'           boards default clock. Function only will work if board has
'           an onboard clock, if the board supports interrupt
'           or DMA data transfer, and if the hardware supports a analog
'           triggers
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f PutDriverLINXAIBuffer>,<f GetDriverLINXAIBuffer>,
'           <f SetupDriverLINXContinuousBufferedIO>,
'           <f SetupDriverLINXContinuousBufferedAIAnalogStopTrigger>
'
'
Public Sub SetupDriverLINXContinuousBufferedAIAnalogTrigger(SR As DriverLINXSR, _
                        LDD As DriverLINXLDD, ByVal device As Integer, _
                        ByVal subsystem As Integer, ByVal Channel As Integer, _
                        ByVal gain As Single, ByVal frequency As Single, _
                        ByVal SamplesPerChannel As Integer, _
                        ByVal Buffers As Integer, _
                        ByVal BackGroundForeGround As Integer, _
                        ByVal TrigChan As Integer, ByVal TrigGain As Integer, _
                        ByVal UpperThresholdVoltage As Single, _
                        ByVal LowerThresholdVoltage As Single, ByVal Slope As Integer)

    Dim ChannelsSampled As Integer
    Dim bResult As Boolean
    ChannelsSampled = 1 'One channel in this case
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        AddTimingEventDefault SR, frequency
        
        ' Specify start event
        bResult = AddStartEventAnalogTrigger(SR, LDD, subsystem, TrigChan, TrigGain, _
                                UpperThresholdVoltage, LowerThresholdVoltage, Slope)
         
        ' Specify stop event
        AddStopEventOnCommand SR  'use Software Command to stop acquisition
        
        ' ------------ Select Group ----------------
        ' Specify channels, gain and data format
        AddSelectSingleChannel SR, Channel, subsystem, gain
        
        ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Continuously reads or writes one or more buffers of data from/to
'          a subsystem using the onboard clock.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   Channel                 |Channel to read or write
' @parm    Single           |   gain                    |Channels gain value
' @parm    Single           |   frequency               |Rate or which to read or write
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
' @parm    Integer          |   TrigChan                |Analog trigger channel
' @parm    Single           |   TrigGain                |Trigger channels gain value
' @parm    Single           |   UpperThresholdVoltage   |Upper threshold voltage
' @parm    Single           |   LowerThresholdVoltage   |Lower threshold voltage
' @parm    Integer          |   Slope                   |Define the trigger
' @parm    Long             |   Delay                   |Amount of samples to acquire after trigger
'
'
' @comm     <f SetupDriverLINXContinuousBufferedAIAnalogStopTrigger> This
'           function sets up a Service Request that either inputs or
'           outputs one or more buffers of data from/to a subsystem.
'           The Service Requests starts when the service request is submitted to
'           DriverLINX and stops upon recieving a analog stop trigger
'           Data is clocked in/out using the boards default clock.
'           Function only will work if board has an onboard clock,
'           if the board supports interrupt or DMA data transfer,
'           and if the hardware supports a analog stop triggers.
'
'
' @devnote  KevinD 7/28/99 1:50:00PM
'
' @xref     <f PutDriverLINXAIBuffer>,<f GetDriverLINXAIBuffer>,
'           <f SetupDriverLINXContinuousBufferedIO>,
'           <f SetupDriverLINXContinuousBufferedAIAnalogTrigger>
'
'
Public Sub SetupDriverLINXContinuousBufferedAIAnalogStopTrigger(SR As DriverLINXSR, _
                        LDD As DriverLINXLDD, ByVal device As Integer, _
                        ByVal subsystem As Integer, ByVal Channel As Integer, _
                        ByVal gain As Single, ByVal frequency As Single, _
                        ByVal SamplesPerChannel As Integer, _
                        ByVal Buffers As Integer, _
                        ByVal BackGroundForeGround As Integer, _
                        ByVal TrigChan As Integer, ByVal TrigGain As Integer, _
                        ByVal UpperThresholdVoltage As Single, _
                        ByVal LowerThresholdVoltage As Single, ByVal Slope As Integer, ByVal Delay As Long)

    Dim ChannelsSampled As Integer
    Dim bResult As Boolean
    ChannelsSampled = 1 'One channel in this case
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        AddTimingEventDefault SR, frequency
        
        ' Specify start event
        AddStartEventOnCommand SR  'use Software Command to stop acquisition
        
        ' Specify stop event
        bResult = AddStopEventAnalogTrigger(SR, LDD, subsystem, TrigChan, TrigGain, _
                                UpperThresholdVoltage, LowerThresholdVoltage, Slope, Delay)
        
        ' ------------ Select Group ----------------
        ' Specify channels, gain and data format
        AddSelectSingleChannel SR, Channel, subsystem, gain
        
        ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Continuously reads or writes one or more buffers of data from/to
'          a subsystem using the onboard clock.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   Channel                 |Channel to read or write
' @parm    Single           |   gain                    |Channels gain value
' @parm    Single           |   frequency               |Rate or which to read or write
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
' @parm    Integer          |   TrigChan                |Digital trigger channel
' @parm    Integer          |   Slope                   |Define the trigger
' @parm    Long             |   Delay                   |Amount of samples to acquire after trigger
'
'
' @comm     <f SetupDriverLINXContinuousBufferedAIDigitalStopTrigger> This
'           function sets up a Service Request that either inputs or
'           outputs one or more buffers of data from/to a subsystem.
'           The Service Requests starts when the service request is submitted to
'           DriverLINX and stops upon recieving a digital stop trigger
'           Data is clocked in/out using the boards default clock.
'           Function only will work if board has an onboard clock,
'           if the board supports interrupt or DMA data transfer,
'           and if the hardware supports a digital stop triggers.
'
'
' @devnote  KevinD 7/28/99 1:50:00PM
'
' @xref     <f PutDriverLINXAIBuffer>,<f GetDriverLINXAIBuffer>,
'           <f SetupDriverLINXContinuousBufferedIO>,
'           <f SetupDriverLINXContinuousBufferedAIAnalogTrigger>,
'           <f SetupDriverLINXAIDigitalStartTrigger>
'
'
Public Sub SetupDriverLINXContinuousBufferedAIDigitalStopTrigger(SR As DriverLINXSR, _
                        LDD As DriverLINXLDD, ByVal device As Integer, _
                        ByVal subsystem As Integer, ByVal Channel As Integer, _
                        ByVal gain As Single, ByVal frequency As Single, _
                        ByVal SamplesPerChannel As Integer, _
                        ByVal Buffers As Integer, _
                        ByVal BackGroundForeGround As Integer, _
                        ByVal TrigChan As Integer, _
                        ByVal Slope As Integer, ByVal Delay As Long)

    Dim ChannelsSampled As Integer
    ChannelsSampled = 1 'One channel in this case
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        AddTimingEventDefault SR, frequency
        
        ' Specify start event
        AddStartEventOnCommand SR  'use Software Command to stop acquisition
        
        ' Specify stop event
        AddStopEventDigitalTrigger SR, TrigChan, Slope, Delay
        
        ' ------------ Select Group ----------------
        ' Specify channels, gain and data format
        AddSelectSingleChannel SR, Channel, subsystem, gain
        
        ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
    
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Reads or writes one or more buffers of data from/to
'          a digital subsystem using an external clock.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Channel                 |Channel to read or write
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXContinuousDigitalIO> This function sets up a
'           Service Request that either inputs or outputs one or more
'           buffers of data from/to a digital subsystem. The Service
'           Requests starts when submitted and stops when the buffer is
'           either full or empty depending whether it is an input or output
'           task. Data is clocked in/out using an external clock. Function
'           only will work if board supports an external clock and if the
'           board supports interrupt or DMA data transfer.
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f SetupDriverLINXDigitalIOBuffer>, <f WriteDriverLINXDIBuffer>,
'           <f GetDriverLINXDIBuffer>
'
'
Public Sub SetupDriverLINXContinuousDigitalIO(SR As Control, LDD As DriverLINXLDD, _
                                                Channel As Integer, _
                                                SamplesPerChannel As Integer, _
                                                subsystem As Integer, _
                                                device As Integer, _
                                                Buffers As Integer, _
                                                BackGroundForeGround As Integer)

' Use this procedure to read/write a data array to a digital I/O subsystem

Dim ChannelsSampled As Integer
Dim gain As Single
ChannelsSampled = 1     'one channel only
gain = 0

  ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
  
  ' Setup external interrupt as timing clock to be used with an external clock
  ' DriverLINX will sample data a the external clock rate
  AddTimingEventExternalDigitalClock SR
  
  AddStartEventOnCommand SR          'start on command
  AddStopEventOnCommand SR           'stop on command
  
 
  ' ------------ Select Group ----------------
  ' Specify channels, gain and data format
    AddSelectSingleChannel SR, Channel, subsystem, gain
  
  ' ------------ Select Buffer Group ----------------
    AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
  ' ------------ Select Flags ----------------
  ' Request DriverLINXSR ServiceStart and ServiceDone events.
    AddSelectFlags SR, True
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Reads or writes one or more buffers of data from/to
'          a subsystem using the onboard clock.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   Channel                 |Channel to read or write
' @parm    Single           |   gain                    |Channels gain value
' @parm    Single           |   frequency               |Rate or which to read or write
' @parm    Single           |   SamplesPerChannel       |Number of samples per channel
' @parm    Integer          |   Buffers                 |Number of buffers to read or write
' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
'
' @comm     <f SetupDriverLINXAIDigitalStartTrigger> This function sets up a
'           Service Request that either inputs or outputs one or more
'           buffers of data from/to a subsystem. The Service Requests starts
'           when submitted and the Digital Start Trigger parameters are
'           satisfied and stops when the buffer is either full or
'           empty depending whether it is an input or output task. Data is
'           clocked in/out using the boards default clock. Function only will
'           work if board has an onboard clock and if the board supports
'           interrupt or DMA data transfer.
'
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
' @xref     <f PutDriverLINXAIBuffer>,<f GetDriverLINXAIBuffer>,
'           <f SetupDriverLINXBufferedIO>,
'           <f SetupDriverLINXContinuousBufferedAIDigitalStopTrigger>
'
Public Sub SetupDriverLINXAIDigitalStartTrigger(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                            ByVal device As Integer, ByVal subsystem As Integer, _
                                            ByVal Channel As Integer, ByVal gain As Single, _
                                            ByVal frequency As Single, _
                                            ByVal SamplesPerChannel As Integer, _
                                            ByVal Buffers, _
                                            ByVal BackGroundForeGround As Integer, _
                                            ByVal TrigChan As Integer, _
                                            ByVal Slope As Integer, ByVal Delay As Long)
    Dim ChannelsSampled As Integer
    
    ChannelsSampled = 1 'One channel only
        ' ------- Service Request Group -------------
        AddRequestGroupStart SR, LDD, device, subsystem, BackGroundForeGround
        
        ' ------------- Event Group -----------------
        'Specify timing event
        AddTimingEventDefault SR, frequency
        
        ' Specify start event
        AddStartEventDigitalTrigger SR, TrigChan, Slope, Delay
                                    
        ' Specify stop event
        AddStopEventOnTerminalCount SR
        
        ' ------------ Select Channel Group ----------------
        ' Specify channels, gain and data format
        AddSelectSingleChannel SR, Channel, subsystem, gain
        
        ' ------------ Select Buffer Group ----------------
        AddSelectBuffers SR, Buffers, SamplesPerChannel, ChannelsSampled
        
        ' ------------ Select Flags ----------------
        ' Request DriverLINXSR ServiceStart and ServiceDone events.
        AddSelectFlags SR, True
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the start type event of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 ' @parm    Integer          |   TrigChan                |Digital trigger channel
 ' @parm    Integer          |   Slope                   |Define the trigger
 ' @parm    Long             |   Delay                   |Number of samples to wait after trigger
 '
 ' @comm    <f AddStartEventDigitalTrigger> sets up the start type event
 '          portion of the Service Request. This functions tells the
 '          Service Request to start task upon receiving the specified
 '          digital input trigger.
 '
 ' @devnote KevinD 8/16/99 1:40:00PM
 '
 ' @xref    <f AddStartEventOnCommand>, <f AddStartEventNullEvent>,
 '          <f AddStartEventAnalogTrigger>, <f AddStopEventAnalogTrigger>
 '
 '
Public Sub AddStartEventDigitalTrigger(SR As DriverLINXSR, _
                                    ByVal TrigChan As Integer, _
                                    ByVal Slope As Integer, _
                                    ByVal Delay As Long _
                                    )
        
    With SR
        .Evt_Str_type = DL_DIEVENT
        .Evt_Str_delay = Delay
        .Evt_Str_diChannel = TrigChan ' DL_DI_EXTTRG   'Specify external trigger
        .Evt_Str_diMask = 1
        .Evt_Str_diMatch = DL_NotEquals
        If Slope >= 0 Then
          .Evt_Str_diPattern = 0 ' != 0 is rising edge
        Else
          .Evt_Str_diPattern = 1 ' != 1 is falling edge
        End If
    End With

    
    
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the start type event of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 '
 ' @comm    <f AddStartEventOnCommand> The function sets up the start type
 '          event portion of the Service Request. This functions tells the
 '          Service Request to start on command.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddStartEventNullEvent>, <f AddStartEventDigitalTrigger>,
 '          <f AddStartEventAnalogTrigger>
 '
 '
Public Sub AddStartEventOnCommand(SR As DriverLINXSR)
    With SR
        .Evt_Str_type = DL_COMMAND
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the start type event of the Service Request.
 '
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 ' @parm    DriverLINXLDD    |   LDD                     |Name of the control
 ' @parm    Integer          |   Subsystem               |Subsystem to setup
 ' @parm    Integer          |   TrigChan                |Analog trigger channel
 ' @parm    Single           |   TrigGain                |Trigger channels gain value
 ' @parm    Single           |   UpperThresholdVoltage   |Upper threshold voltage
 ' @parm    Single           |   UpperThresholdVoltage   |Lower threshold voltage
 ' @parm    Integer          |   Slope                   |Define the trigger
 '
 '
 ' @comm    <f AddStartEventAnalogTrigger> sets up the start type event
 '          portion of the Service Request. This functions tells the
 '          Service Request to start task upon the analog trigger
 '          meeting the criteria defined below with the Slope and Threshold
 '          arguments.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddStartEventOnCommand>, <f AddStartEventNullEvent>,
 '          <f AddStartEventDigitalTrigger>,<f AddStopEventAnalogTrigger>
 '
 '
Public Function AddStartEventAnalogTrigger(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                    ByVal subsystem As Integer, _
                                    ByVal TrigChan As Integer, _
                                    ByVal TrigGain As Single, _
                                    ByRef UpperThresholdVoltage As Single, _
                                    ByRef LowerThresholdVoltage As Single, _
                                    ByVal Slope As Integer _
                                    ) As Boolean
    Dim bUpper As Boolean
    Dim bLower As Boolean
    Dim lUpper As Long
    Dim lLower As Long
    
    With SR
        .Evt_Str_type = DL_AIEVENT
        .Evt_Str_aiChannel = TrigChan
        .Evt_Str_aiGainCode = SR.DLGain2Code(TrigGain)
        .Evt_Str_aiSlope = Slope
        
        'Calculate the threshold value(s) in AD units then fill in the
        'appropriate properties
        bUpper = ConvertVoltsToADUnits(SR, LDD, subsystem, UpperThresholdVoltage, _
                                        TrigChan, TrigGain, lUpper)
        If bUpper Then
            If lUpper <= 32767 Then
                .Evt_Str_aiUpperThreshold = lUpper
            Else
                .Evt_Str_aiUpperThreshold = lUpper - 65536
            End If
        End If
        bLower = ConvertVoltsToADUnits(SR, LDD, subsystem, LowerThresholdVoltage, _
                                        TrigChan, TrigGain, lLower)
        If bLower Then
            If lLower <= 32767 Then
                .Evt_Str_aiLowerThreshold = lLower
            Else
                .Evt_Str_aiLowerThreshold = lLower - 65536
            End If
        End If
        'return True if only both threshold voltages have been converted
        If (bUpper And bLower) Then
            AddStartEventAnalogTrigger = True 'return true
        Else
            AddStartEventAnalogTrigger = False 'return false
        End If
        
    End With
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the start type event of the Service Request.
 '
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 ' @parm    DriverLINXLDD    |   LDD                     |Name of the control
 ' @parm    Integer          |   Subsystem               |Subsystem to setup
 ' @parm    Integer          |   TrigChan                |Analog trigger channel
 ' @parm    Single           |   TrigGain                |Trigger channels gain value
 ' @parm    Single           |   UpperThresholdVoltage   |Upper threshold voltage
 ' @parm    Single           |   UpperThresholdVoltage   |Lower threshold voltage
 ' @parm    Integer          |   Slope                   |Define the trigger
 ' @parm    Long             |   Delay                   |Number of samples to acquire after trigger
 '
 '
 ' @comm    <f AddStopEventAnalogTrigger> sets up the start type event
 '          portion of the Service Request. This functions tells the
 '          Service Request to start task upon the analog trigger
 '          meeting the criteria defined below with the Slope and Threshold
 '          arguments.
 '
 ' @devnote KevinD 7/28/99 11:40:00AM
 '
 ' @xref    <f AddStartEventOnCommand>, <f AddStartEventNullEvent>,
 '          <f AddStartEventDigitalTrigger>,<f AddStartEventAnalogTrigger>
 '
 '
Public Function AddStopEventAnalogTrigger(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                    ByVal subsystem As Integer, _
                                    ByVal TrigChan As Integer, _
                                    ByVal TrigGain As Single, _
                                    ByRef UpperThresholdVoltage As Single, _
                                    ByRef LowerThresholdVoltage As Single, _
                                    ByVal Slope As Integer, _
                                    ByVal Delay As Long _
                                    ) As Boolean
    Dim bUpper As Boolean
    Dim bLower As Boolean
    Dim lUpper As Long
    Dim lLower As Long
    
    With SR
        .Evt_Stp_type = DL_AIEVENT
        .Evt_Stp_aiChannel = TrigChan
        .Evt_Stp_aiGainCode = SR.DLGain2Code(TrigGain)
        .Evt_Stp_aiSlope = Slope
        .Evt_Stp_delay = Delay
        
        'Calculate the threshold value(s) in AD units then fill in the
        'appropriate properties
        bUpper = ConvertVoltsToADUnits(SR, LDD, subsystem, UpperThresholdVoltage, _
                                        TrigChan, TrigGain, lUpper)
        If bUpper Then
            If lUpper <= 32767 Then
                .Evt_Stp_aiUpperThreshold = lUpper
            Else
                .Evt_Stp_aiUpperThreshold = lUpper - 65536
            End If
        End If
        bLower = ConvertVoltsToADUnits(SR, LDD, subsystem, LowerThresholdVoltage, _
                                        TrigChan, TrigGain, lLower)
        If bLower Then
            If lLower <= 32767 Then
                .Evt_Stp_aiLowerThreshold = lLower
            Else
                .Evt_Stp_aiLowerThreshold = lLower - 65536
            End If
        End If
        'return True if only both threshold voltages have been converted
        If (bUpper And bLower) Then
            AddStopEventAnalogTrigger = True 'return true
        Else
            AddStopEventAnalogTrigger = False 'return false
        End If
        
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the start type event of the Service Request.
 '
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 ' @parm    Integer          |   TrigChan                |Digital trigger channel
 ' @parm    Integer          |   Slope                   |Define the trigger
 ' @parm    Long             |   Delay                   |Number of samples to acquire after trigger
 '
 '
 ' @comm    <f AddStopEventDigitalTrigger> sets up the start type event
 '          portion of the Service Request. This functions tells the
 '          Service Request to stop the task upon the digital trigger
 '          meeting the criteria defined below.
 '
 ' @devnote KevinD 7/28/99 11:40:00AM
 '
 ' @xref    <f AddStartEventOnCommand>, <f AddStartEventNullEvent>,
 '          <f AddStartEventDigitalTrigger>, <f AddStartEventAnalogTrigger>,
 '          <f AddStopEventAnalogTrigger>
 '
 '
Public Sub AddStopEventDigitalTrigger(SR As DriverLINXSR, _
                                    ByVal TrigChan As Integer, _
                                    ByVal Slope As Integer, _
                                    ByVal Delay As Long _
                                    )
   
    With SR
        .Evt_Stp_type = DL_DIEVENT
        .Evt_Stp_delay = Delay
        .Evt_Stp_diChannel = TrigChan ' DL_DI_EXTTRG   'Specify external trigger
        .Evt_Stp_diMask = 1
        .Evt_Stp_diMatch = DL_NotEquals
        If Slope >= 0 Then
          .Evt_Stp_diPattern = 0 ' != 0 is rising edge
        Else
          .Evt_Stp_diPattern = 1 ' != 1 is falling edge
        End If
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the start type event of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 '
 ' @comm    <f AddStartEventNullEvent> sets up the start type event portion
 '           of the Service Request. This functions tells the Service
 '           Request that there is no start type event.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddStartEventOnCommand>, <f AddStartEventDigitalTrigger>,
 '          <f AddStartEventAnalogTrigger>
 '
 '
Public Sub AddStartEventNullEvent(SR As DriverLINXSR)
    With SR
        .Evt_Str_type = DL_NULLEVENT    'Same as DL_Command
    End With
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the stop type event of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 '
 ' @comm    <f AddStopEventOnTerminalCount> sets up the stop type event
 '           portion of the Service Request. This functions tells the Service
 '           Request to stop acquiring data when all buffers specified are
 '           filled.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddStopEventNullEvent>, <f AddStopEventOnCommand>,
 '          <f AddStopEventDigitalTrigger>, <f AddStopEventAnalogTrigger>
 '
 '
Public Sub AddStopEventOnTerminalCount(SR As DriverLINXSR)
    With SR
        .Evt_Stp_type = DL_TCEVENT
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the stop type event of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 '
 ' @comm    <f AddStopEventOnCommand> sets up the stop type event portion
 '           of the Service Request. This functions tells the Service
 '           Request to run continuously until told to stop by issuing
 '           a software stop.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddStopEventNullEvent>, <f AddStopEventOnTerminalCount>,
 '          <f AddStopEventDigitalTrigger>, <f AddStopEventAnalogTrigger>
 '
 '
Public Sub AddStopEventOnCommand(SR As DriverLINXSR)
    With SR
        .Evt_Stp_type = DL_COMMAND
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the stop type event of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 '
 ' @comm    <f AddStopEventNullEvent> sets up the stop type event portion
 '          of the Service Request. This functions tells the Service
 '          Request that there is no stop type event.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddStopEventOnCommand>, <f AddStopEventOnTerminalCount>,
 '          <f AddStopEventDigitalTrigger>, <f AddStopEventAnalogTrigger>
 '
 '
Public Sub AddStopEventNullEvent(SR As DriverLINXSR)
    With SR
        .Evt_Stp_type = DL_NULLEVENT    'Same as DL_TCEVENT
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the timing type event of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 '
 ' @comm    <f AddTimingEventNullEvent> sets up the timing type event portion
 '           of the Service Request. This functions tells the Service
 '           Request that there is no timing type event.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddTimingEventDefault>,
 '          <f AddTimingEventExternalDigitalClock>,
 '          <f AddTimingEventBurstMode>,
 '          <f AddTimingEventDIConfigure>,
 '          <f AddTimingEventSyncIO>
 '
 '
Public Sub AddTimingEventNullEvent(SR As DriverLINXSR)
    With SR
        .Evt_Tim_type = DL_NULLEVENT
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the timing type event of the Service Request.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    Single          |   frequency               |Rate at which to sample data
 '
 ' @comm    <f AddTimingEventDefault> sets up the timing type event portion
 '           of the Service Request. This functions sets up the timing event
 '           to use the the board's default clock. It sets up the Service
 '           Request to sample a channel at the user defined frequency.
 '           Function will fail if the frequency is 0 because this
 '           will cause a divide by zero.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddTimingEventNullEvent>,
 '          <f AddTimingEventExternalDigitalClock>,
 '          <f AddTimingEventBurstMode>,
 '          <f AddTimingEventDIConfigure>,
 '          <f AddTimingEventSyncIO>
 '
 '
Public Sub AddTimingEventDefault(SR As DriverLINXSR, frequency As Single)
    'Specify timing event
    With SR
        .Evt_Tim_type = DL_RATEEVENT
        .Evt_Tim_rateChannel = DL_DEFAULTTIMER
        .Evt_Tim_rateMode = DL_RATEGEN
        .Evt_Tim_rateClock = DL_INTERNAL1
        .Evt_Tim_rateGate = DL_DISABLED
        If frequency = 0 Then   'Call Error Handler
            ErrorHandler "AddTimingEventDefault", 0, "Check the Frequency Argument!"
        End If
        'Convert sampling rate, in Hertz, to clock tics
        .Evt_Tim_ratePeriod = .DLSecs2Tics(DL_DEFAULTTIMER, 1 / frequency)
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the timing type event of the Service Request.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    Integer         |   Clock                   |Counter/Timer channel to use
 '
 ' @comm    <f AddTimingEventSyncIO> sets up the timing type event portion
 '           of the Service Request. This functions sets the default clock to
 '           that of another subsystem via the Clock argument. This sets the
 '           timing of one Service Request to that of another. This in turns
 '           allows the programmer to synchronize two different Service
 '           Request(s) to run off of the same clock.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddTimingEventNullEvent>,
 '          <f AddTimingEventExternalDigitalClock>,
 '          <f AddTimingEventBurstMode>,
 '          <f AddTimingEventDIConfigure>,
 '          <f AddTimingEventDefault>
 '
 '
Public Sub AddTimingEventSyncIO(SR As DriverLINXSR, ByVal Clock As Integer)
    'This timing event is used to Synchronize the output clock to the
    'input clock.
    With SR
        .Evt_Tim_type = DL_RATEEVENT
        .Evt_Tim_rateChannel = Clock    'Clock is the default clock of pacing subsystem
        .Evt_Tim_rateMode = DL_RATEGEN
        .Evt_Tim_rateClock = DL_INTERNAL1
        .Evt_Tim_rateGate = DL_DISABLED
        
        'Convert sampling rate, in Hertz, to clock tics
        .Evt_Tim_ratePeriod = 0                         'Set to zero
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func    Setup the request group for polled, interrupt, or DMA operations.
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 ' @parm    DriverLINXLDD    |   LDD                     |Name of LDD control
 ' @parm    Integer          |   Device                  |Number of the device
 ' @parm    Integer          |   Subsystem               |Subsystem to setup
 ' @parm    Integer          |   BackGroundForeGround    |BackGround or ForeGround task
 '
 ' @comm    <f AddRequestGroupStart> sets up the request group portion
 '           of the Service Request for polled, interrupt, or DMA operations.
 '           If task is asynchronous, the function determines the appropriate
 '           mode. If the board supports DMA, it sets the mode to DMA.
 '           Otherwise, it sets the mode to IRQ. If the board does not support
 '           either DMA or IRQ modes, the function will fail.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddRequestGroupConfigure>, <f AddRequestGroupInitialize>
 '
 '
Public Sub AddRequestGroupStart(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                ByVal device As Integer, ByVal subsystem As Integer, _
                                BackGroundForeGround As Integer)
     ' ------- Service Request Group -------------
        ' Specify type of Service Request
        With SR
            .Req_device = device
            .Req_subsystem = subsystem
            If BackGroundForeGround = Foreground Then
                .Req_mode = DL_POLLED
            ElseIf DoesDeviceSupportDMA(SR, LDD, subsystem) Then    'If data-acquistion mode
                   .Req_mode = DL_DMA                               'is background then
            ElseIf DoesDeviceSupportIRQ(SR, LDD, subsystem) Then    'determine the
                .Req_mode = DL_INTERRUPT                            'best mode to use
            Else
                .Req_mode = DL_POLLED 'Return polled if board doesn't support either IRQ or DMA
            End If
            .Req_op = DL_START  'assume operation will always be start
        End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func    Sets up the flags group portion of the Service Request.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    Boolean         |   OnOff                   |Send ServiceStart and
 ' ServiceDone events
 '
 ' @comm    <f AddSelectFlags> sets up the flags group portion of the Service
 '           Request. This function sets the taskflags field.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Sub AddSelectFlags(SR As DriverLINXSR, OnOff As Boolean)
        With SR
        If OnOff Then
            .Sel_taskFlags = CS_NONE    ' send ServiceStart and ServiceDone events.
        Else
            .Sel_taskFlags = NO_SERVICESTART Or NO_SERVICEDONE  'Used in Service Requests
                                                                'that do not use events.
                                                                'i.e. Initialization
        End If
        End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the timing type event of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 '
 ' @comm    <f AddTimingEventExternalDigitalClock> sets up the timing type
 '           event portion of the Service Request. This functions sets up
 '           the timing event to use an external clock source. It sets up
 '           the Service Request to sample a channel at the external
 '           clocks frequency.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddTimingEventDefault>,
 '          <f AddTimingEventNullEvent>,
 '          <f AddTimingEventBurstMode>,
 '          <f AddTimingEventDIConfigure>,
 '          <f AddTimingEventSyncIO>
 '
 '
Public Sub AddTimingEventExternalDigitalClock(SR As DriverLINXSR)
  ' Setup external interrupt as timing clock to be used with an external clock
  ' DriverLINX will sample data at the external clock rate
  ' This function is not supported by all boards. It allows boards that do not
  ' have internal clocks to fill a data buffer.
    With SR
    .Evt_Tim_type = DL_DIEVENT
    .Evt_Tim_diChannel = DL_DI_EXTCLK
    .Evt_Tim_diMask = 1                 'When clock edge goes high interrupt
    .Evt_Tim_diMatch = DL_NotEquals     'occurs and data is logged.
    .Evt_Tim_diPattern = 0
  End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func    Sets up the buffers group portion of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                     |Name of the control
 ' @parm    Integer          |   Buffers                |Number of buffers to read or write
 ' @parm    Single           |   SamplesPerChannel      |Number of samples per channel
 ' @parm    Integer          |   ChannelsSampled        |Number of channels sampled
 '
 ' @comm    <f AddSelectBuffers> sets up the buffers group portion of the
 '           Service Request. This function sets the number of buffers and
 '           the buffer size.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Sub AddSelectBuffers(SR As DriverLINXSR, ByVal Buffers As Integer, _
                            ByVal SamplesPerChannel As Single, _
                            ByVal ChannelsSampled As Integer)
    With SR
        .Sel_buf_samples = SamplesPerChannel * ChannelsSampled
        .Sel_buf_N = Buffers
        If Buffers > 0 Then
            .Sel_buf_notify = DL_NOTIFY
        Else
            .Sel_buf_notify = DL_NOEVENTS
        End If
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func    Sets up the channels group portion of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                     |Name of the control
 ' @parm    Integer          |   Subsystem              |Subsystem to setup
 ' @parm    Integer          |   channels()             |Two element array containing the
 ' start channel and the stop channel
 ' @parm    Single           |   Gains()                |Two element array containing the
 ' start channels gain and the stop channel gain
 '
 '
 ' @comm    <f AddStartStopList> Function sets up the channel group portion
 '           of the Service Request. This function sets up the channel group
 '           for a single channel.
 '
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddSelectZeroChannels>, <f AddChannelGainList>,
 '          <f AddSelectSingleChannel>
 '
 '
Public Sub AddStartStopList(SR As DriverLINXSR, ByVal subsystem As Integer, _
                            ByRef channels() As Integer, ByRef Gains() As Single)
       With SR
            .Sel_chan_format = DL_tNATIVE
            .Sel_chan_N = 2 'This parameter is always 2 when using a Start Stop List regardless
                            'of channel count
            .Sel_chan_start = channels(0)   'Channels(0) is the starting channel
            .Sel_chan_stop = channels(1)    'Channels(1) is the ending Channel
            If (subsystem = DL_AI) Or (subsystem = DL_AO) Then
            ' For analog subsystems, you can specify a gain supported
            '   by your hardware. Use a positive gain factor for
            '   unipolar I/O, or use a negative gain factor for
            '   bipolar I/O.
                .Sel_chan_startGainCode = .DLGain2Code(Gains(0))  'Starting Channels Gain
                .Sel_chan_stopGainCode = .DLGain2Code(Gains(1))   'Second Channel to Stop
                                                                  'Channel Gain setting
             Else
            ' For other subsystems, set the Sel_chan_startGainCode
            ' property to zero
                .Sel_chan_startGainCode = 0
                .Sel_chan_stopGainCode = 0
            End If
         End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func    Sets up the channels group portion of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                     |Name of the control
 ' @parm    Integer          |   Subsystem              |Subsystem to setup
 ' @parm    Single           |   ChannelsSampled        |Number of channels sampled
 ' @parm    Integer          |   ChannelList            |Array of non-consecutive or
 ' consecutive channels to read or write
 ' @parm    Single           |   Gains                  |Array of channel gain value(s)
 '
 '
 ' @comm    <f AddChannelGainList> sets up the channel group portion of
 '           the Service Request. This function sets up the channel group for
 '           a channel gain list. The user can use this function to acquire
 '           a single or multiple channel(s) that can be in any order with
 '           individual gain codes.
 '
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddSelectZeroChannels>, <f AddStartStopList>,
 '          <f AddSelectSingleChannel>
 '
 '
Public Sub AddChannelGainList(SR As DriverLINXSR, ByVal subsystem As Integer, _
                            ByVal ChannelsSampled As Integer, ByRef ChannelList() As Integer, _
                            ByRef Gains() As Single)
       Dim i As Integer
       With SR
            .Sel_chan_format = DL_tNATIVE
            .Sel_chan_N = ChannelsSampled 'This parameter is equivalent to the channel count
            
            .Sel_chan_start = 0         ' Zero out any start stop list information.
            .Sel_chan_stop = 0          ' Make sure that there is no previously entered
            .Sel_chan_startGainCode = 0 ' Service Request Information that may conflict
            .Sel_chan_stopGainCode = 0  ' with the ChannelGainList.
             
             For i = 0 To ChannelsSampled - 1
                .Sel_chan_list(i) = ChannelList(i)  'ChannelList is an array that contains the
                                                    'order of channels that are to be acquired
                If (subsystem = DL_AI) Or (subsystem = DL_AO) Then
                ' For analog subsystems, you can specify a gain supported
                '   by your hardware. Use a positive gain factor for
                '   unipolar I/O, or use a negative gain factor for
                '   bipolar I/O.
                    .Sel_chan_gainCodeList(i) = SR.DLGain2Code(Gains(i))
                 Else
                ' For other subsystems, set the Sel_chan_startGainCode
                ' property to zero
                .Sel_chan_gainCodeList(i) = 0
                End If
            Next i
         End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the timing type event of the Service Request.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    Single          |   frequency               |Rate or which to read or write in Hz
 ' @parm    Single          |   BurstRate               |Burst mode conversion rate in Hz
 ' @parm    Integer         |   ChannelsSampled         |Number of channels sampled
 '
 ' @comm    <f AddTimingEventBurstMode> sets up the timing type event portion
 '           of the Service Request. This functions sets up the timing event
 '           to use the the board's default clock. It sets up the Service
 '           Request to sample data using burst mode data acquisition.
 '           Function will fail if the Frequency or Burst Rate is 0 because
 '           this will cause a divide by zero.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddTimingEventNullEvent>,
 '          <f AddTimingEventExternalDigitalClock>,
 '          <f AddTimingEventDefault>,
 '          <f AddTimingEventDIConfigure>,
 '          <f AddTimingEventSyncIO>
 '
 '
Public Sub AddTimingEventBurstMode(SR As DriverLINXSR, ByVal frequency As Single, _
                                    ByVal BurstRate As Single, _
                                    ByVal ChannelsSampled As Integer)
    With SR
        .Evt_Tim_type = DL_RATEEVENT
        .Evt_Tim_delay = 0 ' (not used)
        .Evt_Tim_rateChannel = DL_DEFAULTTIMER ' or other allowed counter/timer channel
        .Evt_Tim_rateMode = DL_BURSTGEN
        .Evt_Tim_rateClock = DL_INTERNAL1 ' or other selectable internal frequency
        .Evt_Tim_rateGate = DL_DISABLED
        
        If frequency = 0 Then   'Call Error Handler
            ErrorHandler "AddTimingEventBurstMode", 0, "Check the Frequency Argument!"
        End If
        
        .Evt_Tim_ratePeriod = .DLSecs2Tics(.Evt_Tim_rateChannel, (1 / frequency))
        'specify the minor period of the burst generator in HZ
        
        If BurstRate = 0 Then   'Call Error Handler
            ErrorHandler "AddTimingEventBurstMode", 0, "Check the BurstRate Argument!"
        End If
        .Evt_Tim_rateOnCount = .DLSecs2Tics(.Evt_Tim_rateChannel, (1 / BurstRate))
        .Evt_Tim_ratePulses = ChannelsSampled
    End With
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func    Setup the request group for device initialization.
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 ' @parm    Integer          |   Device                  |Number of the device
 '
 ' @comm    <f AddRequestGroupInitialize> sets up the request group portion
 '           of the Service Request for device initialization.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddRequestGroupConfigure>, <f AddRequestGroupStart>
 '
 '
Public Sub AddRequestGroupInitialize(SR As DriverLINXSR, ByVal device As Integer)
    With SR
        .Req_device = device
        .Req_subsystem = DL_DEVICE
        .Req_mode = DL_OTHER
        .Req_op = DL_INITIALIZE
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func    Sets up the channels group portion of the Service Request.
 '
 ' @parm    DriverLINXSR     |   SR                     |Name of the control
 '
 ' @comm    <f AddSelectZeroChannels> sets up the channel group portion of
 '          the Service Request. This function sets up the channels group
 '          for zero channels. This function is typically used when
 '          initializing the device.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddSelectSingleChannel>, <f AddChannelGainList>,
 '          <f AddStartStopList>
 '
 '
Public Sub AddSelectZeroChannels(SR As DriverLINXSR)
    With SR     'Used in Service Requests that do not use events. i.e. Initialization
        .Sel_chan_N = 0
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func    Sets up the channels group portion of the Service Request.
 '
 ' @parm    DriverLINXSR     |  SR                      |Name of the control
 ' @parm    Integer          |  Channel                 |Channel to read or write
 ' @parm    Integer          |  Subsystem               |Subsystem to setup
 ' @parm    Single           |  gain                    |Channels gain value
 '
 ' @comm    <f AddSelectSingleChannel> sets up the channel group portion of
 '          the Service Request. This function sets up the channel group for
 '          a single channel.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddSelectZeroChannels>, <f AddChannelGainList>,
 '          <f AddStartStopList>
 '
 '
Public Sub AddSelectSingleChannel(SR As DriverLINXSR, ByVal Channel As Integer, _
                                ByVal subsystem As Integer, ByVal gain As Single)
    With SR
        .Sel_chan_format = DL_tNATIVE
        .Sel_chan_N = 1
        .Sel_chan_start = Channel
        .Sel_chan_stop = Channel
        If (subsystem = DL_AI) Or (subsystem = DL_AO) Then
            ' For analog subsystems, you can specify a gain supported
            '   by your hardware. Use a positive gain factor for
            '   unipolar I/O, or use a negative gain factor for
            '   bipolar I/O.
            .Sel_chan_startGainCode = .DLGain2Code(gain)
        Else
            ' For other subsystems, set the Sel_chan_startGainCode
            '   property to zero
            .Sel_chan_startGainCode = 0
        End If
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Library error handler routine.
 '
 ' @parm    String          |  Source               |Function that caused the error
 ' @parm    Long            |  ErrorNumber          |Offset to be added
 ' to the base Error Number
 ' @parm    String          |  Tip                  |Clue as to what caused the error
 '
 ' @comm    <f ErrorHandler> Function raises error that might be caused in the library.
 '          The raised error will contain the following info: The source of the error,
 '          the appropriate error number, and a string providing a clue as to what
 '          caused the argument. The Base error number is 65000 which the user can
 '          change to suit their application. This function is used for internal use.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Private Sub ErrorHandler(ByVal Source As String, ByVal ErrorNumber As Long, _
                        ByVal Tip As String)
  Dim ErrorMsg, Msg As String
  Dim ErrBaseValue As Long
  ErrBaseValue = 65000      'User can change value of errors that are generated.
                            'This Library generates errors starting at 65000. If
                            'that conflicts with ones application and errors that
                            'the application might generate. Change the ErrBaseValue
                            'to offset where this library might generate errors.
  
  'Calculate the appropriate error message
  Select Case ErrorNumber
    Case 0
        ErrorMsg = "Divide by Zero "
    Case 1
        ErrorMsg = "VBArrayBufferConvert Failure"
    Case 2
        ErrorMsg = "Incorrect Subsystem"
    Case 3
        ErrorMsg = "Invalid Size Code"
    End Select
   
   Msg = ErrorMsg & " error occurred in the DriverLINX API Library Function " & Source _
            & ". " & Tip
   'Raise the error!
   Err.Raise Number:=ErrorNumber + ErrBaseValue, Source:=App.Title, Description:=Msg
    
    'Terminate the Application


End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks whether device supports analog start triggers.
 '
 ' @rdesc   Boolean - returns true if device supports analog start triggers.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   Subsystem       |Subsystem to check
 '
 ' @comm    <f DoesDeviceSupportAIStartTrigger> Function determines if device
 '          supports analog start triggers by querying the LDD. Use this
 '          only after the Req_mode field has been set in the SR control.
 '
 ' @xref    <f DoesDeviceSupportAIStopTrigger>
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function DoesDeviceSupportAIStartTrigger(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                subsystem As Integer _
                                                ) As Boolean
     With LDD
    .device = SR.Req_device
    .Req_DLL_name = SR.Req_DLL_name 'open the LDD
     
   Select Case subsystem        'Determines if a subsystem supports analog triggers
    Case DL_AI
        If (LDD.AI_Str_Evt(SR.Req_mode) And (1 * 2 ^ DL_AIEVENT)) <> 0 Then
            DoesDeviceSupportAIStartTrigger = True
        Else
            DoesDeviceSupportAIStartTrigger = False
        End If
    Case DL_AO
        If (.AO_Str_Evt(SR.Req_mode) And (1 * 2 ^ DL_AIEVENT)) <> 0 Then
            DoesDeviceSupportAIStartTrigger = True
        Else
            DoesDeviceSupportAIStartTrigger = False
        End If
    End Select
    ' Close the LDD's driver
            .Req_DLL_name = ""
    End With
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks whether device supports analog stop triggers.
 '
 ' @rdesc   Boolean - returns true if device supports analog start triggers.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   Subsystem       |Subsystem to check
 '
 ' @comm    <f DoesDeviceSupportAIStopTrigger> Function determines if device
 '          supports analog stop triggers by querying the LDD. Use this
 '          only after the Req_mode field has been set in the SR control.
 '
 ' @xref    <f DoesDeviceSupportAIStartTrigger>
 '
 ' @devnote KevinD 7/28/99 1:47:00PM
 '
 '
Public Function DoesDeviceSupportAIStopTrigger(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                subsystem As Integer _
                                                ) As Boolean
     With LDD
    .device = SR.Req_device
    .Req_DLL_name = SR.Req_DLL_name 'open the LDD
     
   Select Case subsystem        'Determines if a subsystem supports analog triggers
    Case DL_AI
        If (LDD.AI_Stp_Evt(SR.Req_mode) And (1 * 2 ^ DL_AIEVENT)) <> 0 Then
            DoesDeviceSupportAIStopTrigger = True
        Else
            DoesDeviceSupportAIStopTrigger = False
        End If
    Case DL_AO
        If (.AO_Stp_Evt(SR.Req_mode) And (1 * 2 ^ DL_AIEVENT)) <> 0 Then
            DoesDeviceSupportAIStopTrigger = True
        Else
            DoesDeviceSupportAIStopTrigger = False
        End If
    End Select
    ' Close the LDD's driver
            .Req_DLL_name = ""
    End With
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks whether device supports digital start triggers.
 '
 ' @rdesc   Boolean - returns true if device supports digital start triggers.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   Subsystem       |Subsystem to check
 '
 ' @comm    <f DoesDeviceSupportDigStartTrigger> Function determines if device
 '          supports digital start triggers by querying the LDD. Use this
 '          only after the Req_mode field has been set in the SR control.
 '
 ' @xref    <f DoesDeviceSupportDigStopTrigger>
 '
 ' @devnote KevinD 8/16/99 10:57:00AM
 '
 '
Public Function DoesDeviceSupportDigStartTrigger(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                subsystem As Integer _
                                                ) As Boolean
     With LDD
    .device = SR.Req_device
    .Req_DLL_name = SR.Req_DLL_name 'open the LDD
     
   Select Case subsystem        'Determines if a subsystem supports digital triggers
    Case DL_AI
        If (.AI_Str_Evt(SR.Req_mode) And (1 * 2 ^ DL_DIEVENT)) <> 0 Then
            DoesDeviceSupportDigStartTrigger = True
         Else
            DoesDeviceSupportDigStartTrigger = False
         End If
    
    Case DL_AO
        If (.AO_Str_Evt(SR.Req_mode) And (1 * 2 ^ DL_DIEVENT)) <> 0 Then
            DoesDeviceSupportDigStartTrigger = True
        Else
            DoesDeviceSupportDigStartTrigger = False
        End If
    
    Case DL_DI
        If (.DI_Str_Evt(SR.Req_mode) And (1 * 2 ^ DL_DIEVENT)) <> 0 Then
            DoesDeviceSupportDigStartTrigger = True
         Else
            DoesDeviceSupportDigStartTrigger = False
         End If
    
    Case DL_DO
        If (.DO_Str_Evt(SR.Req_mode) And (1 * 2 ^ DL_DIEVENT)) <> 0 Then
            DoesDeviceSupportDigStartTrigger = True
        Else
            DoesDeviceSupportDigStartTrigger = False
        End If
    
    Case DL_CT
        If (.CT_Str_Evt(SR.Req_mode) And (1 * 2 ^ DL_DIEVENT)) <> 0 Then
            DoesDeviceSupportDigStartTrigger = True
        Else
            DoesDeviceSupportDigStartTrigger = False
        End If
    
    Case Else
        DoesDeviceSupportDigStartTrigger = False
        ErrorHandler "DoesDeviceSupportDigStartTrigger", 2, _
                                        "Subsystem must be AI, AO, DI, DO or CT"
    End Select
    ' Close the LDD's driver
            .Req_DLL_name = ""
    End With
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks whether device supports digital stop triggers.
 '
 ' @rdesc   Boolean - returns true if device supports digital start triggers.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   Subsystem       |Subsystem to check
 '
 ' @comm    <f DoesDeviceSupportDigStopTrigger> Function determines if device
 '          supports digital stop triggers by querying the LDD. Use this
 '          only after the Req_mode field has been set in the SR control.
 '
 ' @xref    <f DoesDeviceSupportDigStartTrigger>
 '
 ' @devnote KevinD 7/28/99 1:47:00PM
 '
 '
Public Function DoesDeviceSupportDigStopTrigger(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                subsystem As Integer _
                                                ) As Boolean
     With LDD
    .device = SR.Req_device
    .Req_DLL_name = SR.Req_DLL_name 'open the LDD
     
   Select Case subsystem        'Determines if a subsystem supports digital triggers
    Case DL_AI
        If (.AI_Stp_Evt(SR.Req_mode) And (1 * 2 ^ DL_DIEVENT)) <> 0 Then
            DoesDeviceSupportDigStopTrigger = True
        Else
            DoesDeviceSupportDigStopTrigger = False
        End If
     
    Case DL_AO
        If (.AO_Stp_Evt(SR.Req_mode) And (1 * 2 ^ DL_DIEVENT)) <> 0 Then
            DoesDeviceSupportDigStopTrigger = True
        Else
            DoesDeviceSupportDigStopTrigger = False
        End If
        
    Case DL_DI
        If (.DI_Stp_Evt(SR.Req_mode) And (1 * 2 ^ DL_DIEVENT)) <> 0 Then
            DoesDeviceSupportDigStopTrigger = True
        Else
            DoesDeviceSupportDigStopTrigger = False
        End If
     
    Case DL_DO
        If (.DO_Stp_Evt(SR.Req_mode) And (1 * 2 ^ DL_DIEVENT)) <> 0 Then
            DoesDeviceSupportDigStopTrigger = True
        Else
            DoesDeviceSupportDigStopTrigger = False
        End If
        
    Case DL_CT
        If (.CT_Stp_Evt(SR.Req_mode) And (1 * 2 ^ DL_DIEVENT)) <> 0 Then
            DoesDeviceSupportDigStopTrigger = True
        Else
            DoesDeviceSupportDigStopTrigger = False
        End If
        
    Case Else
            DoesDeviceSupportDigStopTrigger = False
            ErrorHandler "DoesDeviceSupportDigStopTrigger", 2, _
                                            "Subsystem must be AI, AO, DI, DO or CT"
    End Select
    ' Close the LDD's driver
            .Req_DLL_name = ""
    End With
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Determines the logical counter/timer channel used as the default
 '          timing channel
 '
 ' @rdesc   Integer - returns subsystems default Clock.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   Subsystem       |Subsystem to check
 '
 ' @comm    <f GetSubSystemsDefaultClock> This function queries the LDD and
 '          returns the default counter/timer channel.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function GetSubSystemsDefaultClock(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                          ByVal subsystem As Integer _
                                          ) As Integer
     'Function returns a subsystems default clock
    
    With LDD
    .device = SR.Req_device
    .Req_DLL_name = SR.Req_DLL_name 'open the LDD
    
    Select Case subsystem
        Case DL_AI
            GetSubSystemsDefaultClock = .AI_DefaultCT
        Case DL_AO
            GetSubSystemsDefaultClock = .AO_DefaultCT
        Case DL_DI
            GetSubSystemsDefaultClock = .DI_DefaultCT
        Case DL_DO
            GetSubSystemsDefaultClock = .DO_DefaultCT
        Case DL_CT
            GetSubSystemsDefaultClock = .CT_DefaultCT
    End Select
    
    ' Close the LDD's driver
            .Req_DLL_name = ""
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func    Fills in the Simultaneous field for the channels group portion
 '          of the Service Request.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    Boolean         |   Simultaneous    |Select Simultaneous Sampling
 ' or One Sample per Clock Tic
 '
 ' @comm    <f AddSelectSimultaneous> sets up part of the channel group
 '           portion of the Service Request. This function determine if the
 '           channels in the start stop list or the channels specified in the
 '           channel gain list are to be sampled as close together as possible
 '           or whether the channels will be acquired at a rate equal to one
 '           clock tic per sample. This feature is particularily usefull when
 '           used with digital IO.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Sub AddSelectSimultaneous(SR As DriverLINXSR, ByVal Simultaneous As Boolean)
    With SR
        If Simultaneous Then
            .Sel_chan_simultaneousScan = DL_True
        Else
            .Sel_chan_simultaneousScan = DL_False
        End If
    End With
End Sub
   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks if any digital channels can be re-programmed.
 '
 ' @rdesc   Boolean - returns true any of the digital channels can be
 '          reprogrammed.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   nChannels       |Number of Digital Channels
 '
 ' @comm    <f CheckDigitalProgamming> This function returns true if any
 '           digital channels can be reconfigured. This only indicates that
 '           one of the digital channels can be reprogrammed to some degree.
 '           Not every channel is re-programmable.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function CheckDigitalProgamming(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                        ByVal nChannels As Integer _
                                        ) As Boolean
    Dim i As Integer
    
    CheckDigitalProgamming = False  'Set to false initially
    With LDD
    .device = SR.Req_device
    .Req_DLL_name = SR.Req_DLL_name 'open the LDD
    
        For i = 0 To nChannels - 1
            If Not (.DI_Config(i) = 0) Then CheckDigitalProgamming = True      'Return true
        Next i
        
    ' Close the LDD's driver
            .Req_DLL_name = ""
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @doc ServiceRequests
'
' @func    Configures a digital IO channel as input or output.
'
' @parm    DriverLINXSR     |   SR                      |Name of the control
' @parm    Integer          |   Device                  |Number of the device
' @parm    Integer          |   Subsystem               |Subsystem to setup
' @parm    Integer          |   Channel                 |Channel to configure
'
' @comm     <f SetupDriverLINXInitDIOPort> This function sets up a Service
'           Request that configures a digital channel as either an input
'           or an output channel if the board supports this feature. This
'           Service Requests is aimed at boards such as the Metrabyte PIO
'           series boards that have the Intel 8255 chip.
'
' @devnote  KevinD 10/27/97 11:40:00AM
'
'
Public Sub SetupDriverLINXInitDIOPort(SR As DriverLINXSR, ByVal device As Integer, _
                                        ByVal subsystem As Integer, _
                                        ByVal Channel As Integer)
    'Used for reconfiguring a digital IO channel.
     ' ------- Service Request Group -------------
        AddRequestGroupConfigure SR, device, subsystem
        
        ' ------------- Event Group -----------------
        ' Timing event for DIO configuration
        AddTimingEventDIConfigure SR, Channel
        
        ' Start immediately on software command
        AddStartEventNullEvent SR               ' or AddStartEventOnCommand
        
        ' Stop as soon as DriverLINX processes the sample
        AddStopEventNullEvent SR               ' or AddStopEventOnTerminalCount
        
        ' ------------ Select Channel Group ----------------
        ' Specify channels, gain and data format
         AddSelectZeroChannels SR
        
        ' ------------ Select Buffers Group ----------------
        ' DIO configuration does not use buffers
        AddSelectBuffers SR, 0, 0, 0
        
         ' ------------ Select Flags -----------------------
        ' DIO configuration doesn't need ServiceStart or ServiceDone
        '   events
        AddSelectFlags SR, False
        ' NOTE: You do not have to block any events. However,
        '   DriverLINX is somewhat more efficient if you do.
        
        'Note: Your application must call the refresh method to execute this function.
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func    Setup the request group for configuring digital IO.
 '
 ' @parm    DriverLINXSR     |   SR                      |Name of the control
 ' @parm    Integer          |   Device                  |Number of the device
 ' @parm    Integer          |   Subsystem               |Subsystem to setup
 '
 ' @comm    <f AddRequestGroupConfigure> sets up the request group portion
 '          of the Service Request for digital IO configuration.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddRequestGroupStart>, <f AddRequestGroupInitialize>
 '
 '
Public Sub AddRequestGroupConfigure(SR As DriverLINXSR, ByVal device As Integer, _
                                    ByVal subsystem As Integer)
    With SR
        .Req_device = device
        .Req_subsystem = subsystem
        .Req_mode = DL_OTHER
        .Req_op = DL_CONFIGURE
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc AddFunctions
 '
 ' @func     Sets up the timing type event of the Service Request.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    Integer         |   Channel                 |Digital channel to configure
 '
 ' @comm    <f AddTimingEventDIConfigure> sets up the timing type
 '           event portion of the Service Request. This functions sets up
 '           the timing event for configuring digital IO.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 ' @xref    <f AddTimingEventNullEvent>,
 '          <f AddTimingEventExternalDigitalClock>,
 '          <f AddTimingEventBurstMode>,
 '          <f AddTimingEventDefault>,
 '          <f AddTimingEventSyncIO>
 '
 '
Public Sub AddTimingEventDIConfigure(SR As DriverLINXSR, ByVal Channel As Integer)
    'Used for reconfiguring a digital IO channel.
    With SR
        .Evt_Tim_type = DL_DIOSETUP
        .Evt_Tim_dioChannel = Channel
        .Evt_Tim_dioMode = DL_DIO_BASIC
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc ServiceRequests
 '
 ' @func    Checks if the channel is configured as a digital input channel
 '          or a digital output channel.
 '
 ' @rdesc   Integer - returns 0 if Channel is an output, 1 if Channel is an input.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    Integer         |   Channel                 |Digital channel to check
 '
 ' @comm    <f ReadCurrentDigitalIOConfiguration> This function verifies if
 '          a digital channel is configured as either a input or output
 '          channel. This function assumes that the user has passed a Service
 '          Request that has been set up for single value IO. This function
 '          makes this determination by first reading a channel followed by
 '          a writing the value back to the channel. If successful the
 '          channel is an output channel if operation fails channel is an
 '          input channel.
 '
 ' @xref     <f SetupDriverLINXSingleValueIO>
 '
 ' @devnote  KevinD 8/5/97 11:05:00
 '
Public Function ReadCurrentDigitalIOConfiguration(SR As DriverLINXSR, _
                                                    ByVal Channel As Integer _
                                                    ) As Integer
    Dim i As Integer, DLResultCode As Integer
    Dim ReadValue As Single
    Dim oldOperation As String
    Dim DLMessage As String
    
    With SR
    ' In order not to change the current data of each channel,
    ' which may be written by other programs, read it first then
    ' write it back
    
    'Read channel
        .Sel_chan_start = Channel  'increment the channel being checked
        .Sel_chan_stop = Channel
        .Req_subsystem = DL_DI
        .Refresh
    ' NOTE: Single-value transfers execute synchronously, i.e., the data
    '   is available when the call to the Refresh method returns.
    
    ' Get status or error information
        DLResultCode = GetDriverLINXStatus(SR, DLMessage)
        If DLResultCode = DL_NoErr Then
        ' If no error ocurred, return the acquired data
            ReadValue = GetDriverLINXDISingleValue(SR)
        Else
             oldOperation = .Req_op
            .Req_op = DL_MESSAGEBOX   'Show Error Message
            .Refresh
            .Req_op = oldOperation
        End If
        
    ' Write channel
        .Req_subsystem = DL_DO
        .Res_Sta_ioValue = ReadValue
        .Refresh
    
    'if a channel has an error when executing a write operation
    'then we know this channel is configured as an input channel
    'otherwise it is configured as an output channel
        If .Res_result = DL_NoErr Then
            ReadCurrentDigitalIOConfiguration = 0       '0 = output channel
        Else
            ReadCurrentDigitalIOConfiguration = 1
        'ElseIf .Res_result = DL_InvalidOpErr Then
        '    ReadCurrentDigitalIOConfiguration = 1      '1 = input channel
        'Else
        '    oldOperation = .Req_op
        '    .Req_op = DL_MESSAGEBOX   'Show Error Message
        '    .Refresh
        '    .Req_op = oldOperation
        End If

    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Calculates an extended digital address.
 '
 ' @rdesc   Long - returns the extended address.
 '
 ' @parm    Integer         |   SizeCode            |Channel format.
 ' Ex. 0 = Native, 1= Bit
 ' @parm    Integer         |   Channel             |Name of the control
 '
 ' @comm    <f GetExtendedDigitalAddress>  This function calculates the
 '          extended digital address based on a Size Code and a channel
 '          number.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function GetExtendedDigitalAddress(ByVal SizeCode As Integer, _
                                          ByVal Channel As Integer _
                                          ) As Long
    'Returns the extended digital address based on the channel and the size code.
    Select Case SizeCode
        Case 0
            GetExtendedDigitalAddress = &H0 + Channel       'Native
        Case 1
            GetExtendedDigitalAddress = &H1000 + Channel    'Bit
        Case 2
            GetExtendedDigitalAddress = &H2000 + Channel    'Half Nibble
        Case 3
            GetExtendedDigitalAddress = &H3000 + Channel    'Nibble
        Case 4
            GetExtendedDigitalAddress = &H4000 + Channel    'Byte
        Case 5
            GetExtendedDigitalAddress = &H5000 + Channel    'Word
        Case 6
            GetExtendedDigitalAddress = &H6000 + Channel    'Dword
        Case 7
            GetExtendedDigitalAddress = &H7000 + Channel    'Qword
        Case Else   'No Match - Error!
            ErrorHandler "GetExtendedDigitalAddress", 3, "No Matching Size Code Found"
    End Select
        
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Determines the number of bytes per digital channel.
 '
 ' @rdesc   Integer - returns the number of bytes per digital channel.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    DriverLINXLDD   |   LDD                     |Name of LDD control
 ' @parm    Integer         |   Subsystem               |Subsystem to check
 ' @parm    Integer         |   Channel                 |Channel to check
 '
 ' @comm    <f HowManyBytesPerDigitalChannel>  This function queries the LDD
 '          to see how many bytes are used per digital channel.
 '
 ' @xref    <f IsHardwareIntel8255>, <f HowManyDriverLINXLogicalChannels>,
 '          <f HowManyBitsPerDigitalChannel>,
 '          <f HowManyExtendedDigitalChannels>,
 '          <f HowManyBytesPerAnalogSample>
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function HowManyBytesPerDigitalChannel(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                ByVal subsystem As Integer, _
                                                ByVal Channel As Integer _
                                                ) As Integer
    ' returns the number of bytes per logical channel.
    With LDD
        .Req_DLL_name = SR.Req_DLL_name 'open the LDD
        .device = SR.Req_device
        If subsystem = DL_DI Then
            HowManyBytesPerDigitalChannel = .DI_Bytes(Channel)
        End If
        
        If subsystem = DL_DO Then
            HowManyBytesPerDigitalChannel = .DO_Bytes(Channel)
        End If
        
    End With
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Determines the number of bytes per analog sample.
 '
 ' @rdesc   Integer - returns the number of bytes per analog sample.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    DriverLINXLDD   |   LDD                     |Name of LDD control
 ' @parm    Integer         |   Subsystem               |Subsystem to check
 '
 ' @comm    <f HowManyBytesPerAnalogSample>  This function queries the LDD
 '          to see how many bytes are used per analog sample.
 '
 ' @xref    <f HowManyBytesPerDigitalChannel>
 '
 ' @devnote KevinD 4/22/99 4:10:00PM
 '
 '
Public Function HowManyBytesPerAnalogSample(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                ByVal subsystem As Integer _
                                                ) As Integer
    ' returns the number of bytes per sample.
    With LDD
        .Req_DLL_name = SR.Req_DLL_name 'open the LDD
        .device = SR.Req_device
        If subsystem = DL_AI Then
            HowManyBytesPerAnalogSample = .AI_Bytes
        End If
        
        If subsystem = DL_AO Then
            HowManyBytesPerAnalogSample = .AO_Bytes
        End If
        
    End With
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Determines the number of bits per digital channel.
 '
 ' @rdesc   Integer - returns the number of bits per digital channel.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    DriverLINXLDD   |   LDD                     |Name of LDD control
 ' @parm    Integer         |   Subsystem               |Subsystem to check
 ' @parm    Integer         |   Channel                 |Channel to check
 '
 ' @comm    <f HowManyBitsPerDigitalChannel>  This function queries the LDD
 '          to see how many bits are used per digital channel.
 '
 ' @xref    <f IsHardwareIntel8255>, <f HowManyDriverLINXLogicalChannels>,
 '          <f HowManyBytesPerDigitalChannel>,
 '          <f HowManyExtendedDigitalChannels>
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function HowManyBitsPerDigitalChannel(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                ByVal subsystem As Integer, _
                                                ByVal Channel As Integer _
                                                ) As Integer
    ' returns the number of bits per logical channel.

    Dim Mask As Long
    
    With LDD
        .device = SR.Req_device
        .Req_DLL_name = SR.Req_DLL_name 'open the LDD
        If subsystem = DL_DI Then
            Mask = .DI_Mask(Channel)
        End If
        
        If subsystem = DL_DO Then
            Mask = .DO_Mask(Channel)
        End If
        
        Select Case Mask
        Case 65535
            HowManyBitsPerDigitalChannel = 16
        Case 255
            HowManyBitsPerDigitalChannel = 8
        Case 127
            HowManyBitsPerDigitalChannel = 7
        Case 63
            HowManyBitsPerDigitalChannel = 6
        Case 31
            HowManyBitsPerDigitalChannel = 5
        Case 15
            HowManyBitsPerDigitalChannel = 4
        Case 7
            HowManyBitsPerDigitalChannel = 3
        Case 3
            HowManyBitsPerDigitalChannel = 2
        Case 1
            HowManyBitsPerDigitalChannel = 1
        End Select
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Returns the type of Digital Hardware on the board.
 '
 ' @rdesc   Boolean - returns true if subsystem uses an Intel 8255.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    DriverLINXLDD   |   LDD                     |Name of LDD control
 ' @parm    Integer         |   Subsystem               |Subsystem to check
 ' @parm    Integer         |   Channel                 |Channel to check
 '
 ' @comm    <f IsHardwareIntel8255>  This function queries the LDD and returns
 '          true of false if the channel uses an Intel8255 chip.
 '
 ' @xref    <f HowManyBitsPerDigitalChannel>, <f HowManyDriverLINXLogicalChannels>,
 '          <f HowManyBytesPerDigitalChannel>,
 '          <f HowManyExtendedDigitalChannels>
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function IsHardwareIntel8255(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                    ByVal subsystem As Integer, _
                                    ByVal Channel As Integer _
                                    ) As Boolean
   'determines if the digital hardware uses the Intel 8255 chip.
    With LDD
        .device = SR.Req_device
        .Req_DLL_name = SR.Req_DLL_name 'open the LDD
        
        If subsystem = DL_DI Then
            If .DI_Type(Channel) = 1 Then
                IsHardwareIntel8255 = True
            Else
                IsHardwareIntel8255 = False
            End If
        End If
        
        If subsystem = DL_DO Then
             If .DO_Type(Channel) = 1 Then
                IsHardwareIntel8255 = True
            Else
                IsHardwareIntel8255 = False
            End If
        End If
        
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Determines the number of extended digital channels.
 '
 ' @rdesc   Integer - returns the number of extended digital channels.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    DriverLINXLDD   |   LDD                     |Name of LDD control
 ' @parm    Integer         |   Subsystem               |Subsystem to check
 ' @parm    Integer         |   Channel                 |Channel to check
 ' @parm    Integer         |   SizeCode                |Channel format.
 ' Ex. 0 = Native, 1= Bit

 '
 ' @comm    <f HowManyExtendedDigitalChannels>  This function queries the
 '          LDD to see how many extended channels are supported within
 '          the specified subsystem. This function uses the Sizecode argument
 '          to determine how many extended channels are supported by the
 '          digital subsystem.
 '
 ' @xref    <f HowManyBitsPerDigitalChannel>, <f HowManyDriverLINXLogicalChannels>,
 '          <f HowManyBytesPerDigitalChannel>,
 '          <f IsHardwareIntel8255>
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function HowManyExtendedDigitalChannels(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                ByVal subsystem As Integer, _
                                                ByVal Channel As Integer, _
                                                ByVal SizeCode As Integer _
                                                ) As Integer
    Dim NumberOfBytes As Integer
    Dim BytesPerChannel As Integer
    Dim BitsPerChannel As Integer
    Dim nChannels As Integer        'The number of available channels given size code
    
    BytesPerChannel = HowManyBytesPerDigitalChannel(SR, LDD, subsystem, 0)
    BitsPerChannel = HowManyBitsPerDigitalChannel(SR, LDD, subsystem, 0)
    
    ' Get the number of digital input channels that this device supports
    nChannels = HowManyDriverLINXLogicalChannels(SR, LDD, subsystem)
    
    
    Select Case SizeCode
    Case 0      'Native
        HowManyExtendedDigitalChannels = nChannels
    Case 1      'Bit
    
        HowManyExtendedDigitalChannels = _
            BytesPerChannel * BitsPerChannel * nChannels
    Case 2      'Half Nibble
        HowManyExtendedDigitalChannels = _
            (BytesPerChannel * BitsPerChannel * nChannels) / 2
    Case 3      'Nibble
        HowManyExtendedDigitalChannels = _
            (BytesPerChannel * BitsPerChannel * nChannels) / 4
    Case 4      'Byte
        HowManyExtendedDigitalChannels = _
            (BytesPerChannel * BitsPerChannel * nChannels) / 8
    Case 5      'Word
        HowManyExtendedDigitalChannels = nChannels / 2
    Case 6      'Double Word
        HowManyExtendedDigitalChannels = nChannels / 4
    Case 7      'Quad Word
        HowManyExtendedDigitalChannels = nChannels / 8
    Case Else   'No Match - Error!
        ErrorHandler "HowManyExtendedDigitalChannels", 3, "No Matching Size Code Found"
    End Select
   
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Checks whether digital value is within the extended channels
 '          range.
 '
 ' @rdesc   Boolean - returns true if Value is in range.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    DriverLINXLDD   |   LDD                     |Name of LDD control
 ' @parm    Integer         |   Subsystem               |Subsystem to check
 ' @parm    Integer         |   SizeCode                |Channel format.
 ' Ex. 0 = Native, 1= Bit
 ' @parm    Variant         |   Value                   |Value to check
 '
 ' @comm    <f ISInDriverLINXExtendedDigitalRange>  This function compares
 '          the user input with the channels acceptable limits to determine
 '          if the user input falls within the channels extend digital range.
 '          This function uses the SizeCode argument to determine if the
 '          user input is valid.
 '
 ' @xref    <f IsInDriverLINXAnalogRange>, <f IsInDriverLINXDigitalRange>,
 '          <f ConvertVoltsToADUnits>
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function ISInDriverLINXExtendedDigitalRange(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                                    ByVal subsystem As Integer, _
                                                    ByVal SizeCode As Integer, _
                                                    ByVal Value As Variant _
                                                    ) As Boolean
   
     Select Case SizeCode
        Case 0      ' Native - make use of existing function
                ISInDriverLINXExtendedDigitalRange = IsInDriverLINXDigitalRange(SR, LDD, DL_DO, Value, 0) 'Check Channel Zero
        Case 1      ' Bit
            If (Value >= 0) And (Value <= 1) Then
                ISInDriverLINXExtendedDigitalRange = True
            Else
                ISInDriverLINXExtendedDigitalRange = False
            End If
        Case 2      ' half nibble
            If (Value >= 0) And (Value <= 3) Then
                ISInDriverLINXExtendedDigitalRange = True
            Else
                ISInDriverLINXExtendedDigitalRange = False
            End If
        Case 3      ' Nibble
            If (Value >= 0) And (Value <= 15) Then
                ISInDriverLINXExtendedDigitalRange = True
            Else
                ISInDriverLINXExtendedDigitalRange = False
            End If
        Case 4      ' Byte
            If (Value >= 0) And (Value <= 255) Then
                ISInDriverLINXExtendedDigitalRange = True
            Else
                ISInDriverLINXExtendedDigitalRange = False
            End If
        Case 5      ' Word
            If (Value >= 0) And (Value <= 65535) Then
                ISInDriverLINXExtendedDigitalRange = True
            Else
                ISInDriverLINXExtendedDigitalRange = False
            End If
        Case 6      ' dWord
            If (Value >= 0) And (Value <= 4294967295#) Then
                ISInDriverLINXExtendedDigitalRange = True
            Else
                ISInDriverLINXExtendedDigitalRange = False
            End If
        
        Case Else   'No Match - Error!
            ErrorHandler "ISInDriverLINXExtendedDigitalRange", 3, "No Matching Size Code Found"
        End Select
        
        
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Gets the Model of the installed board.
 '
 ' @rdesc   String - returns name of the installed board.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    DriverLINXLDD   |   LDD                     |Name of LDD control
 '
 ' @comm    <f GetModelName>  This function returns the model description
 '          of the board the driver is using. Should only call this function
 '          after a driver is opened.
 '
 ' @devnote KevinD 10/27/97 11:40:00AM
 '
 '
Public Function GetModelName(SR As DriverLINXSR, LDD As DriverLINXLDD _
                            ) As String
    If (SR.Req_DLL_name <> "") Or (SR.Req_DLL_name <> Null) Then
        With LDD
            ' Make sure that the Service Request and LDD controls open
            '   the same DriverLINX driver
            .device = SR.Req_device
            .Req_DLL_name = SR.Req_DLL_name
            GetModelName = .Dev_Model
           
        End With
    Else
        GetModelName = ""
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Converts a voltage to the equivalent A/D units.
 '
 ' @rdesc   Boolean - returns true if user input is in range.
 '
 ' @parm    DriverLINXSR    |   SR              |Name of the control
 ' @parm    DriverLINXLDD   |   LDD             |Name of LDD control
 ' @parm    Integer         |   Subsystem       |Subsystem to check
 ' @parm    Integer         |userinput          |Value to check if in range
 ' and then convert
 ' @parm    Integer         |Channel            |Channel to check range for
 ' @parm    Single          |Gain               |Gain of channel to select
 ' @parm    Long            |ADUNIT             |Converted Voltage
 '
 ' @comm    <f ConvertVoltsToADUnits> This function converts a voltage to
 '           the equivalent A/D units. This function first determines if the
 '           device supports a channel gain list. If the device supports a
 '           channel gain list the user input is compared to limits
 '           based on the Gain Multiplier Table. If the channel gain list is
 '           not supported the user input is compared to the values stored in
 '           the Min/Max Range Table. Then the user input is converted to the
 '           equivalent A/D units and returned in the lADUNIT argument. If the
 '           userinput value is out of range the function will return false
 '           and will fill the lADUNIT with either the maximum or minimum
 '           A/D unit based on whether the user input is less than the
 '           minimum or greater than the maximum A/D value
 '
 ' @devnote KevinD 4/17/98 10:30:00AM
 '
 ' @xref    <f IsInDriverLINXDigitalRange>,
 '          <f IsInDriverLINXExtendedDigitalRange>,
 '          <f DoesDeviceSupportAnalogChannelGainList>,
 '          <f IsInDriverLINXAnalogRange>
 '
 '
Public Function ConvertVoltsToADUnits(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                            ByVal subsystem As Integer, _
                                            ByRef UserInput As Single, _
                                            Channel As Integer, _
                                            ByVal gain As Single, _
                                            ByRef ADUNIT As Long _
                                            ) As Boolean
    
    'Function returns true if value submitted falls between
    'the allowable low and high range analog limits.
    Dim GainList As Boolean
    Dim TestMin As Single
    Dim TestMax As Single
    Dim MinCode As Long
    Dim MaxCode As Long
    Dim Resolution As Byte  'Analog resolution in bits
    Dim bBipolar As Boolean
    Dim i As Integer
    Dim bMatch As Boolean
    Dim Mask As Long
    
    bMatch = False  'Initialize variable to false
    TestMin = 0
    TestMax = 0
    
    If (gain < 0) Then
        bBipolar = True
    Else
        bBipolar = False
    End If
    
    Resolution = 0 'initialize
    
    With LDD
    .device = SR.Req_device
    .Req_DLL_name = SR.Req_DLL_name 'open the LDD
    
    'First check to see if a channel gain list is supported
    GainList = DoesDeviceSupportAnalogChannelGainList(SR, LDD, subsystem)
         
    If subsystem = DL_AO Then
        If GainList Then
           i = 0
           Do
            If (Abs(.AO_GM_mul(i) - Abs(gain)) < 0.1) Then 'We have a match check for bipolar value
                TestMin = .AO_GM_min(i)
                TestMax = .AO_GM_max(i)
                'Check to see that we have obtained the correct data from
                'channel gain table.
                Select Case bBipolar
                Case True
                    If TestMin < 0 Then bMatch = True
                Case False
                    If TestMin >= 0 Then bMatch = True
                End Select
                
            End If
            i = i + 1
           Loop Until (bMatch = True) Or (i = LDD.AO_GM_n)
        Else
            TestMin = .AO_MM_min(0) 'Assume converter 0
            TestMax = .AO_MM_max(0)
        End If
        'Determine Minimum and Maximum ADUnits that channel supports
        MinCode = .AO_MM_minCode(0)
        MaxCode = .AO_MM_maxCode(0)
        Resolution = .AO_Bits   'Analog resolution
    ElseIf subsystem = DL_AI Then
        If GainList Then
            i = 0
           Do
            If (Abs(.AI_GM_mul(i) - Abs(gain)) < 0.1) Then 'We have a match check for bipolar value
                TestMin = .AI_GM_min(i)
                TestMax = .AI_GM_max(i)
                'Check to see that we have obtained the correct data from
                'channel gain table.
                Select Case bBipolar
                Case True
                    If TestMin < 0 Then bMatch = True
                Case False
                    If TestMax >= 0 Then bMatch = True
                End Select
                
            End If
            i = i + 1
           Loop Until (bMatch = True) Or (i = LDD.AI_GM_n)
        Else
            TestMin = .AI_MM_min(0) 'Assume converter 0
            TestMax = .AI_MM_max(0)
        End If
        'Determine Minimum and Maximum ADUnits that channel supports
        MinCode = .AI_MM_minCode(0)
        MaxCode = .AI_MM_maxCode(0)
        Resolution = .AI_Bits   'Analog resolution
    Else
        'Invalid subsystem
        ConvertVoltsToADUnits = False
        ADUNIT = 0  'return 0
        ErrorHandler "ConvertVoltsToADUnits", 2, "Invalid subsystem passed to function!"
    End If
    
    'Test userinput versus allowable range
        If (UserInput >= TestMin) And (UserInput <= TestMax) Then
            ConvertVoltsToADUnits = True
            'Calculate ADUnits
            ADUNIT = Int((UserInput - TestMin) / (TestMax - TestMin) * (MaxCode - MinCode) _
                    + MinCode + 0.5)
            Mask = .AI_Mask
            While ((Mask And 1) = 0)
                Mask = Mask \ 2
                ADUNIT = ADUNIT * 2
            Wend
            ADUNIT = ADUNIT And .AI_Mask
        Else
            ConvertVoltsToADUnits = False
            If (UserInput > TestMax) Then   'set ADUNIT to closest boundary
                ADUNIT = MaxCode
                UserInput = TestMax 'send caller back the changed voltage
            Else
                ADUNIT = MinCode
                UserInput = TestMin 'send caller back the changed voltage
            End If
        End If
    
    ' Close the LDD's driver
            .Req_DLL_name = ""
    End With
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' @doc UtilityFunctions
 '
 ' @func    Returns whether a channel is an external trigger or clock channel.
 '
 ' @rdesc   Boolean - returns true if channel is an external trigger or clock channel.
 '
 ' @parm    DriverLINXSR    |   SR                      |Name of the control
 ' @parm    DriverLINXLDD   |   LDD                     |Name of LDD control
 ' @parm    Integer         |   Subsystem               |Subsystem to check
 ' @parm    Integer         |   Channel                 |Channel to check
 '
 ' @comm    <f IsExternalDigTrgOrClockChannel>  This function queries the LDD and returns
 '          true of false if the channel is either a External Trigger Channel or a
 '          External Clock Channel.
 '
 ' @xref    <f IsHardwareIntel8255>, <f HowManyDriverLINXLogicalChannels>
 '
 ' @devnote KevinD 9/29/97 11:30:00AM
 '
 '
Public Function IsExternalDigTrgOrClockChannel(SR As DriverLINXSR, LDD As DriverLINXLDD, _
                                    ByVal subsystem As Integer, _
                                    ByVal Channel As Integer _
                                    ) As Boolean
    With LDD
        .device = SR.Req_device
        .Req_DLL_name = SR.Req_DLL_name 'open the LDD
        
        Select Case subsystem
            Case DL_DI
                If .DI_Type(Channel) = 2 Or .DI_Type(Channel) = 3 Then
                    IsExternalDigTrgOrClockChannel = True
                Else
                    IsExternalDigTrgOrClockChannel = False
                End If
            Case DL_DO
                If .DO_Type(Channel) = 2 Or .DO_Type(Channel) = 3 Then
                    IsExternalDigTrgOrClockChannel = True
                Else
                    IsExternalDigTrgOrClockChannel = False
                End If
            Case Else   'Invalid subsystem argument passed in
                ErrorHandler "IsExternalDigTrgOrClockChannel", 2, _
                            "Subsystem must be DI or DO"
                IsExternalDigTrgOrClockChannel = False
        End Select
                
        'Close LDD
        .Req_DLL_name = ""
    End With
End Function
