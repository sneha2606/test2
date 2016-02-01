Imports System.Text
Imports System.Data
Imports Microsoft.VisualBasic
Imports System.IO.Ports

Public Class MultipleSensor

    'Option Explicit On
    Public strSndCommand As String                             'command to write to Konica Minolta device
    Public strRcvCommand As String                             'command received from Konica Minolta device

    Public strSendStr As String                                'character set to send to Konica Minolta device
    Public strReceiveStr As String                             'character set received from Konica Minolta device

    Public strSTX_Command As String                            'STX & command
    Public strCommand_ETX As String                            'command & ETX
    Public strCommand_ETX_BCC As String                        'command & ETX & BCC

    Public intErrNO As Integer                                 'Error No
    '0: Normal, 1-7: Error Code, 8: Time Out, 9:BCC
    'Error, 10: Range Changing, 11: Battery Out
    Public intErrflg As Integer

    Public strData As String                                    'measurement data Block containing Ev,x and y values 
    Public strData1 As String                                   'Ev( Lux value)
    Public strData2 As String                                   'x(CCT value)
    Public strData3 As String                                   'y(delta T)

    Public sngData1 As Single                                   'measurement data Ev (Lux value)
    Public sngData2 As Single                                   'measurement data x  (CCT value)
    Public sngData3 As Single                                   'measurement data y  (delta T)

    Public sensorNumber() As String = {"00", "01", "02"}        'sensor numbers 00-Sensor 0,01-Sensor 1,02-Sensor 2
    Public luxcct_Values(2, 1) As String                        'array to store Lux and CCT values for each sensor


    Public n As Integer = 3 'number of sensor
    Public j As Integer 'for LOOP

    Private mySerialPort As New SerialPort


    '*******************************
    '*** Serialport configurations setup(based on Konica Minolta communications specifications) ***
    '*******************************

    Public Function serialPortConfigurations()
        mySerialPort.PortName = "COM5"                             'COM Port for serial port communication
        mySerialPort.Parity = Parity.Even                          'Parity is Even
        mySerialPort.BaudRate = 9600                               'Baud rate for coomunication with Konica Minolta device
        mySerialPort.StopBits = 1                                  'Number of Stop Bits 
        mySerialPort.DataBits = 7                                  'Character length
        mySerialPort.Open()                                        'Open serial port for communication 
    End Function



    '*******************************
    '*** Starting Measurement ***
    '*******************************
    Public Function StartReadingData()
        If mySerialPort.IsOpen Then                                 'If serial port is open then start reading data
            GoTo readData
        Else                                                        'Else open serial port and start reading data
            mySerialPort.Open()
readData:
            intErrflg = 0


            '------------------------------
            'Step 2 PC MODE
            '------------------------------
            j = 0
            strSndCommand = "00541 "                                ' Command to set the Konica Minolta device to PC Connection mode
            Call CmdSendRecv(1)
            Call ErrCheck()                                         ' Check for Errors in data received from Konica Minolta device
            If intErrflg = 1 Then
                Exit Function
            End If

            'Code to wait 500ms
            System.Threading.Thread.Sleep(500)

            '------------------------------
            'Step 3 HOLD ON
            '------------------------------
            strSndCommand = "99551  0"                              ' Command to set the Konica Minolta device to Hold status
            Call CmdSend(0)

            'Code to wait 500ms
            System.Threading.Thread.Sleep(500)


            '------------------------------
            'Step 4 EXT MODE
            '------------------------------

            For j = 0 To n - 1
                strSndCommand = sensorNumber(j) & "4010  "          ' Command to set the Konica Minolta device to External mode
                Call CmdSendRecv(1)
                Call ErrCheck()                                     ' Check for Errors in data received from Konica Minolta device

                If intErrflg = 1 Then
                    Exit Function
                End If

            Next j

            'Code to wait 175ms
            System.Threading.Thread.Sleep(175)


            '------------------------------
            'Step 5 EXT MEASUREMENT
            '------------------------------
            strSndCommand = "994021 "                           'Command to set the Konica Minolta device to take Measurement
            Call CmdSend(0)

            'Code to wait 500ms 
            System.Threading.Thread.Sleep(500)


            '------------------------------
            'Step 6 READ MEASUREMENT DATA
            '------------------------------
            For j = 0 To n - 1
                strSndCommand = sensorNumber(j) & "081200"        'Command to set the Konica Minolta device to start reading measurement data from each sensor 
                Call CmdSendRecv(1)
                Call ErrCheck()                                   'Check for Errors in data received from Konica Minolta device

                If intErrflg = 1 Then
                    Exit Function
                End If

                strData = Microsoft.VisualBasic.Right(strRcvCommand, 18)
                strData1 = Microsoft.VisualBasic.Left(strData, 6)
                strData2 = Microsoft.VisualBasic.Mid(strData, 7, 6)
                strData3 = Microsoft.VisualBasic.Right(strData, 6)

                'Lv,x,y
                sngData1 = Val(Microsoft.VisualBasic.Left(strData1, 5)) * 10 ^ (Val(Microsoft.VisualBasic.Right(strData1, 1)) - 4)
                sngData2 = Val(Microsoft.VisualBasic.Left(strData2, 5)) * 10 ^ (Val(Microsoft.VisualBasic.Right(strData2, 1)) - 4)
                sngData3 = Val(Microsoft.VisualBasic.Left(strData3, 5)) * 10 ^ (Val(Microsoft.VisualBasic.Right(strData3, 1)) - 4)

                'Data obtained from the receptor heads are then used.
                luxcct_Values(j, 0) = sngData1  'Lux value 
                luxcct_Values(j, 1) = sngData2  'CCT value

            Next j
        End If

    End Function


    '************************************************************************************
    '*** Send command to Konica Minolta device & Receive command from Konica Minolta ***
    '************************************************************************************

    Public Sub CmdSendRecv(FlgTimeoutCheck As Integer)
        'Dim sngStartTime As Single
        'Dim sngFinishTime As Single
        'Dim varBuf As String
        intErrNO = 0
        strRcvCommand = ""
        strReceiveStr = ""
        '-------------------------
        'Transmission
        '-------------------------
        Call BCC_Append(strSndCommand)
        strSendStr = Chr(2) & strCommand_ETX_BCC & vbCr & vbLf          'STX + command along with ETX and BCC  + delimiter(CR+LF)


        'Code for sending data to Konica Minolta device
        '-------------------------
        System.Threading.Thread.Sleep(1000)
        mySerialPort.WriteLine(strSendStr)                              'Write data to serial port



        'Reception & TimeOut Check
        '-------------------------
        'Code to handle data receiving within timeout limit here
        '-------------------------
        Dim stopwatch As Stopwatch = Stopwatch.StartNew
        stopwatch = Stopwatch.StartNew
        System.Threading.Thread.Sleep(5000)
        stopwatch.Stop()
        strReceiveStr = mySerialPort.ReadExisting()                      'Read data from serial port



        'BCC Check
        '-------------------------
        strSTX_Command = Microsoft.VisualBasic.Left(strReceiveStr, (InStr(1, strReceiveStr, Chr(3)) - 1))
        strRcvCommand = Mid(strSTX_Command, 2)
        Call BCC_Append(strRcvCommand)
        If (strReceiveStr) <> (Chr(2) & strCommand_ETX_BCC & vbCr & vbLf) Then
            intErrNO = 9 'BCC Error
        Else
            intErrNO = 0
        End If
    End Sub


    '**************************
    '*** BCC Calculation ***
    '**************************
    Public Sub BCC_Append(Command As String)
        Dim intBCC As Long
        Dim strBCC As String
        strCommand_ETX = Command & Chr(3)                                   'Command + ETX
        intBCC = 0
        For i = 1 To Len(strCommand_ETX)
            intBCC = intBCC Xor Asc(Mid(strCommand_ETX, i, 1))
        Next i
        strBCC = (Hex(intBCC))
        If Len(strBCC) = 1 Then
            strBCC = "0" & strBCC
        Else
        End If
        strCommand_ETX_BCC = strCommand_ETX & strBCC                         'Command + ETX + BCC
    End Sub


    '**********************
    '*** Error Check for Data receved from Konica Minolta device***
    '**********************
    Public Sub ErrCheck()
        If Mid(strRcvCommand, 8, 1) = "1" Then
            intErrNO = 11 'Battery Out
            Exit Sub
        ElseIf Mid(strRcvCommand, 7, 1) = "6" Then
            intErrNO = 10 'Changing Range
            Exit Sub
        ElseIf intErrNO = 0 Then
            If Mid(strRcvCommand, 6, 1) = " " Then
                intErrNO = 0
            Else
                intErrNO = Val(Mid(strRcvCommand, 6, 1))
            End If
        End If
        'Select Case intErrNO
        '    Case 0 : Exit Sub
        '    Case 1 : MsgBox("POWER OF SENSOR WAS OFF.(No." & ")") : intErrflg = 1
        '    Case 2 : MsgBox("EE-PROM ERROR(No." &  & ")") : intErrflg = 1
        '    Case 3 : MsgBox("EE-PROM ERROR(No." &  & ")") : intErrflg = 1
        '    Case 4 : MsgBox("EXT ERROR(No." & SensorNo(j) & ")") : intErrflg = 1
        '    Case 5 : Exit Sub
        '    Case 6 : Exit Sub
        '    Case 7 : Exit Sub
        '    Case 8 : MsgBox("TIME OUT(No." & SensorNo(j) & ")") : intErrflg = 1
        '    Case 9 : MsgBox("BCC ERROR(No." & SensorNo(j) & ")") : intErrflg = 1
        '    Case 10 : Exit Sub
        '    Case 11 : MsgBox("BATTERY OUT(No." & SensorNo(j) & ")") : intErrflg = 1
        'End Select
    End Sub


    '**********************************************
    '*** Send command to Konica Minolta device  ***
    '**********************************************
    Public Sub CmdSend(FlgTimeoutCheck As Integer)
        intErrNO = 0
        'strRcvCommand = ""
        'strReceiveStr = ""

        '-------------------------
        'Transmission
        '-------------------------
        Call BCC_Append(strSndCommand)
        strSendStr = Chr(2) & strCommand_ETX_BCC & vbCr & vbLf                 'STX + command along with ETX and BCC  + delimiter(CR+LF)


        'Code for sending data to Konica Minolta device 
        '-------------------------
        System.Threading.Thread.Sleep(1000)
        mySerialPort.WriteLine(strSendStr)


    End Sub


End Class
