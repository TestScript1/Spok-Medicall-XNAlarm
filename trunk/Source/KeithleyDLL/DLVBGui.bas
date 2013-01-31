Attribute VB_Name = "DriverLINXGUIInterface"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc INTERNAL
 '  @module DLVBGui.bas |
 '
 '  DriverLINX<rtm> Visual Basic Graphical User Interface Library<nl>
 '  <cp> Copyright 1997 Scientific Software Tools, Inc.<nl>
 '  All Rights Reserved.<nl>
 '
 '  Graphical User Interface Library. This module provides non-DriverLINX
 '  functions that are shared between the Visual Basic examples.
 '
 '  @comm
 '  Author: KevinD<nl>
 '  Date:   10/27/97 11:05:00
 '
 '  @group Revision History
 '  @comm
 '  1     10/27/97 2:30p KevinD
 '  Initial revision.
 '

Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Converts a one dimensional array to a multi-dimensional array.
 '
 '  @parm   Byte        |   TempArray()         |One dimensional array to convert
 '  @parm   Byte        |   DataArray()         |Multi-dimensional array to create
 '  @parm   Integer     |   nSamples            |Number of Samples per Channel
 '  @parm   Integer     |   nChannels           |Number of channel's
 '
 '  @comm   <f ArrayTransfer> This subroutine takes the one dimensional array TempArray
 '          and converts it to a multidimensional array (DataArray).
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '
Public Sub ArrayTransfer(ByRef TempArray() As Byte, ByRef DataArray() As Byte, _
                            nSamples As Integer, nChannels As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    k = 0
    For i = 0 To nSamples - 1
    For j = 0 To nChannels - 1
    DataArray(j, i) = TempArray(k)
    k = k + 1
    Next j
    Next i
    
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Converts each byte in an array and makes 8 arrays of bits.
 '
 '  @parm   Byte        |   ByteArray()     |Byte data to break down to bits
 '  @parm   Byte        |   BitArray()      |Resultant arrays - values = 1 for on, 0 = off
 '  @parm   Integer     |   nSamples        |Number of samples per channel
 '
 '  @comm   <f ByteToBitConversion> This subroutine takes a byte and analyzes the
 '          individual bits. It then makes 8 arrays one for each bit and adds an
 '          off set value.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '
Public Sub ByteToBitConversion(ByteArray() As Byte, BitArray() As Byte, nSamples As Integer)

Dim bit As Integer, samp As Integer
Dim x As Integer
Dim TestValue As Integer 'TestValue will test the value of each bit
                         'The values of TestValue are 1,2,4,8,16,32,64,128

For samp = 0 To nSamples - 1
    TestValue = 1
    For bit = 0 To 7
         If (ByteArray(samp) And TestValue) Then    ' check if bit is on or off
            BitArray(bit, samp) = 1 'on state value
        Else
            BitArray(bit, samp) = 0 'off state value
        End If
        TestValue = TestValue * 2
    Next bit
Next samp
End Sub

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Centers a Form.
 '
 '  @parm   Form        |   frm         |Name of form to center
 '
 '  @comm   <f CenterForm> This subroutine centers a form on the screen.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
Public Sub CenterForm(frm As Form)

    With frm
        ' Center the form on the screen
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
        
        ' Make sure that the form is no bigger than the screen
        If .Left < 0 Then
            .Left = 0
        End If
        
        If .Top < 0 Then
            .Top = 0
        End If
        
        If .Height > Screen.Height Then
            .Height = Screen.Height
        End If
        
        If .Width > Screen.Width Then
            .Width = Screen.Width
        End If
    End With
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Creates a multi-dimensional array of sine waves with a 1 volt offset.
 '
 '  @parm   Single      |   PeakVoltage         |One dimensional array to fill
 '  @parm   Integer     |   nSamples            |Number of Samples per sine wave
 '  @parm   Single      |   DataArray1()        |Multi-dimensional of sine waves to create
 '  @parm   Integer     |   nChannels           |Number of sine wave's to create
 '
 '  @comm   <f CreateSineWaves> This subroutine fills a two dimensional array with data that
 '          corresponds to a sine wave. Each successive sine wave will be offset 1 volt from
 '          the previous sine wave. Each sine wave will be created with a peak value that is
 '          entered via the PeakVoltage argument.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '
Public Sub CreateSineWaves(ByVal PeakVoltage As Single, ByVal nSamples As Integer, _
                            DataArray1() As Single, nChannels As Integer)
' subroutine is used to create an array of sine waves.

    Const PI = 3.141592653
    
    Dim alpha As Single
    Dim i, j As Integer
    
    
    alpha = 2! * PI / nSamples
    
    For j = 0 To nChannels - 1  'loop to create wave for each channel one by one
    For i = 0 To nSamples - 1   'create the individual wave
        DataArray1(j, i) = PeakVoltage * Sin(i * alpha) + (j / nChannels) 'add offset
                                                            'for each sine wave
    Next i
    Next j

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Checks if a file exists.
 '
 '  @rdesc  Boolean - Returns True if file exists.
 '
 '  @parm   String          |   File            |Name of File to check for
 '
 '  @comm   <f DoesFileExist> This function verifies if a file exists or not.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '  @xref   <f WriteToDisk>, <f SizeOfFile>
 '
 '
Public Function DoesFileExist(File As String _
                                ) As Boolean
    Dim FileExist As String
    
    FileExist = Dir(File)
    If FileExist <> "" Or Null Then
        DoesFileExist = True
    Else
        DoesFileExist = False
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func  Fills a array with a constant value.
 '
 '  @parm   Byte        |   Value               |Value to be inserted into array
 '  @parm   Integer     |   nSamples            |Number of Samples to fill in array
 '  @parm   Byte        |   DataArray()         |Result of operation.
 '
 '  @comm   <f FillBuffer> This subroutine fills an array with a constant value.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '
Public Sub FillBuffer(ByRef Value As Byte, ByRef nSamples As Integer, DataArray() As Byte)
    'Routine fill an array with a constant value "Value"
    Dim i As Integer
    
    For i = 0 To nSamples - 1
        DataArray(i) = Value
    Next i

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Fill a one dimensional array.
 '
 '  @parm   Byte        |   TempArray()         |One dimensional array to fill
 '  @parm   Integer     |   nSamples            |Number of Samples per channel
 '  @parm   Integer     |   nChannels           |Number of channel's
 '
 '  @comm   <f FillTempArray> This subroutine fills an array with a constant value.
 '          Array is filled with repeated values that correspond to the channel number
 '          plus 1. The data will repeat for how many samples per channel. For example,
 '          if channels is 3 data in the array will be 1,2,3,4,1,2,3,4, etc..
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '
Public Sub FillTempArray(ByRef TempArray() As Byte, nSamples As Integer, _
                            nChannels As Integer)
    'This subroutine fills array that has (nSamples*nChannels) elements
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
  k = 0
  For i = 0 To nSamples - 1
    For j = 1 To nChannels
     TempArray(k) = j + 1   'output 1 for first channel, 2 for second and so on
    k = k + 1
    Next j
   Next i
   
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Initializes a Scroll bar and a corresponding Text box.
 '
 '  @parm   String      |   DataFileName        |Binary file name used to initialize the
 '  amount of buffers for the Scroll bar
 '  @parm   TextBox     |   editBox             |Name of Text box control
 '  @parm   VScrollBar  |   selector            |Name of Scroll bar control
 '  @parm   Integer     |   SamplesPerBuffer    |Number of Samples per Buffer
 '
 '  @comm   <f InitBufferSelector> This subroutine initializes the min, max values of the
 '          vertical scroll bar which corresponds to the number of buffers that have been
 '          written to a binary file on disk. The textbox holds the value of the scroll bar.
 '          Both are initially set to the last buffer acquired.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '  @xref   <f InitChannelSelector>, <f SizeOfFile>
 '
 '
Public Sub InitBufferSelector(DataFileName As String, editBox As TextBox, _
                                selector As VScrollBar, ByVal SamplesPerBuffer As Integer)
    Dim FileLength, BufferNumber As Long
    'Determine Size of the file
    FileLength = SizeOfFile(DataFileName)
    BufferNumber = FileLength / SamplesPerBuffer / 4    '4 bytes per sample
    
    If BufferNumber > 0 Then
        selector.Min = 1
        selector.Max = BufferNumber
        If BufferNumber > 256 Then
            selector.LargeChange = 256
        ElseIf BufferNumber > 64 Then
            selector.LargeChange = 64
        ElseIf BufferNumber > 16 Then
            selector.LargeChange = 16
        Else
            selector.LargeChange = 1
        End If
        selector.Value = BufferNumber   'Set to the most recent buffer
        selector.Enabled = True
        
        editBox.Text = Str(BufferNumber)     'Set to the most recent buffer
        editBox.Enabled = True
    Else
        selector.Min = 1
        selector.Max = 1
        selector.Value = 1
        selector.Enabled = False
        editBox.Text = Str(0)
        editBox.Enabled = False
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Initializes a Scroll bar and a corresponding Text box.
 '
 '  @parm   TextBox     |   editBox             |Name of Text box control
 '  @parm   VScrollBar  |   selector            |Name of Scroll bar control
 '  @parm   Integer     |   Channel             |Initial value for both the Scroll bar
 '  and the Text box
 '  @parm   Integer     |   NumberOfChannels    |Number of channel's
 '
 '  @comm   <f InitChannelSelector> This function initializes the min, max values of the
 '          vertical scroll bar which corresponds to the number channels a subsystem
 '          supports. The textbox holds the value of the scroll bar. Both are initially
 '          set to the Channel parameter.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '
Public Function InitChannelSelector(editBox As TextBox, selector As VScrollBar, _
                                ByVal Channel As Integer, ByVal NumberOfChannels As Integer)
    Dim i As Integer
    
    If Channel < 0 Then
        Channel = 0
    ElseIf Channel >= NumberOfChannels Then
        Channel = NumberOfChannels - 1
    Else
        Channel = Channel
    End If
      
    If NumberOfChannels > 0 Then
        With selector
            If NumberOfChannels > 256 Then
                .LargeChange = 256
            ElseIf NumberOfChannels > 64 Then
                .LargeChange = 64
            ElseIf NumberOfChannels > 16 Then
                .LargeChange = 16
            Else
                .LargeChange = 1
            End If
            
            .Min = 0
            .Max = NumberOfChannels - 1
            .Value = Channel
            .Enabled = True
        End With
        
        With editBox
            .Text = Channel
            .Enabled = True
        End With

        InitChannelSelector = True
    Else
        With selector
            .LargeChange = 1
            .Min = 0
            .Max = 0
            .Value = 0
            .Enabled = False
        End With
        
        With editBox
            .Text = ""
            .Enabled = False
        End With
        
        InitChannelSelector = False
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Initializes the CRT.
 '
 '  @parm   PictureBox      |   pic             |Name of the picture box to initialize
 '  @parm   Single          |   scaleHorizontal |Scaled max x-axis value
 '  @parm   Single          |   scaleVertical   |Scaled max y-axis value
 '  @parm   Boolean         |   Digital         |Specify True for Digital scaling,
 '  specify false for analog scaling. Digital only displays a positive y-axis
 '  @parm   Integer         |   YAxisOffset     |Decimal value to offset multiple bits or
 '  channels.
 '
 '  @comm   <f InitCRT> This subroutine sets the scale for the picture box.
 '          Digital is used to determine if we will be graphing digital or analog data.
 '          Offset is used only for digital graphing. It sets the y-axis minimum value to
 '          something less than zero so that zero can be view on the screen.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
Public Sub InitCRT(pic As PictureBox, ByVal scaleHorizontal As Single, _
                    ByVal scaleVertical As Single, ByVal Digital As Boolean, _
                    ByVal YAxisOffset As Integer)
    With pic
        'Setup the display coordinate system.
        If Digital Then
            .ScaleTop = scaleVertical       'only create a positive Y-axis
        Else
            .ScaleTop = scaleVertical / 2   'create a positive/negative Y-axis
        End If
        .ScaleLeft = 0
        If Digital Then
            .ScaleHeight = -(scaleVertical + YAxisOffset)
        Else
            .ScaleHeight = -scaleVertical
        End If
        .ScaleWidth = scaleHorizontal
    End With

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Initializes the Size Code Option Buttons.
 '
 '  @parm   Object  |   optSizeCode   |Control Array of Option Buttons
 '  @parm   Integer |   Bits          |Number of Bits Per Digital Channel
 '
 '  @comm   <f InitSizeCodeOptionButtons> This subroutine enables the
 '          appropriate size code option buttons based on the number of
 '          bits per channel.
 '
 '  @devnote    KevinD 3/23/98 02:43:00PM
 '
Public Sub InitSizeCodeOptionButtons(optSizeCode As Object, Bits As Integer)

 'If subsystem is supported only enable the valid "Size Code" options
    optSizeCode(0).Enabled = True   'Always enable the native size
                                    'if subsystem is supported
    Select Case Bits
    Case 1  'bit
        optSizeCode(1).Enabled = True
    Case 2 To 3  'half nibble
        optSizeCode(1).Enabled = True
        optSizeCode(2).Enabled = True
    Case 4 To 7  'nibble
        optSizeCode(1).Enabled = True
        optSizeCode(2).Enabled = True
        optSizeCode(3).Enabled = True
    Case 8 To 15  ' byte
        optSizeCode(1).Enabled = True
        optSizeCode(2).Enabled = True
        optSizeCode(3).Enabled = True
        optSizeCode(4).Enabled = True
    Case 16 To 31 ' Word
        optSizeCode(1).Enabled = True
        optSizeCode(2).Enabled = True
        optSizeCode(3).Enabled = True
        optSizeCode(4).Enabled = True
        optSizeCode(5).Enabled = True
    Case 32       'Double Word
        optSizeCode(1).Enabled = True
        optSizeCode(2).Enabled = True
        optSizeCode(3).Enabled = True
        optSizeCode(4).Enabled = True
        optSizeCode(5).Enabled = True
        optSizeCode(6).Enabled = True
    End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Draws analog data.
 '
 '  @parm   PictureBox      |   pic             |Name of the picture box to initialize
 '  @parm   Integer         |   channels        |Number of channels to be graphed
 '  @parm   Integer         |   samples         |Number of samples per channel
 '  @parm   Single          |   frequency       |Rate at which data was acquired/written
 '  @parm   Single          |   VBArray()       |Array of data to be graphed
 '
 '  @comm   <f ShowAnalogResults> This subroutine draws graphs representing each
 '          input/output channel.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '  @xref   <f ShowDigitalResults>
 '
 '
Public Sub ShowAnalogResults(pic As PictureBox, ByVal channels As Integer, _
                            ByVal samples As Integer, ByVal frequency As Single, _
                            VBArray() As Single)

    Dim x As Single
    Dim period As Single
    Dim sample As Integer
    Dim Channel As Integer
    Dim scaleFull As Single
    
    period = 1 / frequency
    scaleFull = pic.ScaleWidth
    'Plot the data
    pic.Cls
    
    For Channel = 0 To channels - 1
        x = 0
        pic.PSet (0, VBArray(Channel, 0))
        For sample = 1 To samples - 1
            x = x + period
            If x > scaleFull Then
                Exit For
            Else
                pic.Line -(x, VBArray(Channel, sample))
            End If
        Next sample
    Next Channel
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Draws digital data.
 '
 '  @parm   PictureBox      |   pic             |Name of the picture box to initialize
 '  @parm   Integer         |   nChannels       |Number of bits or channels to be graphed
 '  @parm   Integer         |   samples         |Number of samples per channel
 '  @parm   Byte            |   VBArray()       |Array of data to be graphed
 '  @parm   Integer         |   Step            |Offset that can be applied to each channel
 '
 '  @comm   <f ShowDigitalResults> This subroutine draws graph(s) representing each
 '          input/output bit or word of a channel. The graphs are drawn as square waves.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '  @xref   <f ShowAnalogResults>
 '
 '
Public Sub ShowDigitalResults(pic As PictureBox, ByVal nChannels As Integer, _
                                ByVal samples As Integer, _
                                VBArray() As Byte, _
                                ByVal Step As Integer)
    'nChannels can be either the number of bits or channels to be graphed
    Dim j, k As Integer
    Dim offset As Integer
    Dim GraphOffset As Integer  'Allows the first graph to be seen when it is off
    'Plot the data
    pic.Cls
    
    'When graphing the first channel (not bits) it may not be seen if channel value = 0
    'The following shifts graphs up so graph can be seen
    GraphOffset = 0
    If Step >= 255 Then
        If nChannels > 10 Then GraphOffset = 10
        If (nChannels <= 10) And (nChannels >= 5) Then GraphOffset = 5
        If (nChannels <= 5) And (nChannels >= 1) Then GraphOffset = 1
    End If
    
    offset = 0
    For k = 0 To nChannels - 1
    pic.PSet (0, (VBArray(k, 0) + offset + GraphOffset))   ' draw the first point
    For j = 1 To samples - 1        ' draw the rest as square waves
        pic.Line -(j, VBArray(k, j - 1) + offset + GraphOffset)
        pic.Line -(j, VBArray(k, j) + offset + GraphOffset)
    Next j
        offset = offset + Step
   Next k
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Gets length of a file.
 '
 '  @rdesc  Long - Returns number of bytes in the file.
 '
 '  @parm   String          |   File            |Name of File to check for
 '
 '  @comm   <f SizeOfFile> This function returns the length of a file in Bytes.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '  @xref   <f WriteToDisk>, <f DoesFileExist>, <f InitBufferSelector>
 '
 '
Public Function SizeOfFile(FileName As String _
                            ) As Long
    
    If FileName <> "" Then
        SizeOfFile = FileLen(FileName)
    Else
        SizeOfFile = 0
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '  @doc    GUI-Functions
 '
 '  @func   Writes data to a binary file.
 '
 '  @parm   String          |   FileName            |Name of File to write to
 '  @parm   Single          |   VBArray()           |Data to write to file
 '
 '  @comm   <f WriteToDisk> This subroutine writes data to a new binary file or
 '          appends data to an existing data file.
 '
 '  @devnote    KevinD 10/27/97 11:40:00AM
 '
 '  @xref   <f SizeOfFile>, <f DoesFileExist>
 '
 '
Public Sub WriteToDisk(FileName As String, VBArray() As Single)
    Dim FileNumber As Integer
    Dim i As Integer
    Dim FileLength As Long
    
    On Error GoTo FileError
    
    If FileName = "" Or Null Then
        MsgBox "File Name is either Empty or Null!", vbOKOnly, "File Error"
        Exit Sub
    Else
        FileNumber = FreeFile   ' Get unused file
        Open FileName For Binary Access Write As #FileNumber
        
        'Check file size if exists determine length and add to it
        FileLength = SizeOfFile(FileName)
        If FileLength = 0 Then
            Put #FileNumber, 1, VBArray()
        Else
            Put #FileNumber, FileLength + 1, VBArray()
        End If
        
    Close FileNumber
    End If
Exit Sub

FileError:
'MsgBox for file Error
    MsgBox "File Error " & Str$(Err.Number) & "" & Err.Description, vbOKOnly, Error.Source
    'terminate the application


End Sub
