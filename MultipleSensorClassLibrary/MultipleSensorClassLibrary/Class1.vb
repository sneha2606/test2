Imports System.Text
Imports System.Data
Imports Microsoft.VisualBasic
Imports System.IO.Ports
Imports System.Timers


'************************************************************************************
'*** User Defined Exceptions for different error conditions ***
'************************************************************************************


''' <summary>
''' This exception occurs when there is no data received from Konica Minolta device
''' </summary>
''' <remarks></remarks>
Public Class SerialPortReadDataErrorException : Inherits ApplicationException
    Public Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub
End Class


''' <summary>
''' This exception occurs in case of any error in data received from Konica device
''' </summary>
''' <remarks></remarks>
Public Class KonicaErrorException : Inherits ApplicationException
    Public Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub
End Class

''' <summary>
''' This class is used to read data from 3 sensors : Sensor 0, Sensor 1 and Sensor 1 connected to the Konica Minolta device
''' </summary>
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
    Public strData1 As String                                   'Ev (Lux value in lx)
    Public strData2 As String                                   'Tcp (CCT value in K)
    Public strData3 As String                                   '(delta UV)

    Public sngData1 As Single                                   'measurement data Ev (Lux value)
    Public sngData2 As Single                                   'measurement data T  (CCT value)
    Public sngData3 As Single                                   'measurement data (delta uv)

    Public sensorNumber() As String = {"00", "01", "02"}        'sensor numbers 00-Sensor 0,01-Sensor 1,02-Sensor 2
    Public luxcct_Values(2, 1) As String                        'array to store Lux and CCT values for each sensor


    Public n As Integer = 3                                     'number of sensor
    Public j As Integer                                         'for LOOP

    Public errMsg As String                                     'to store the error message for different error conditions in Konica Minolta device

    Public mySerialPort As New SerialPort

    Public _timer As Timer
    'Public _continue As Boolean
    Public z As Integer
    Public label_value As String = ""
    Public temp As Integer
    '***********************************************************************************************
    '*** Serialport configurations setup(based on Konica Minolta communications specifications) ***
    '***********************************************************************************************
    ''' <summary>
    ''' This function is used to set the Serial port configurations for Konica Minolta Device
    ''' </summary>
    ''' <returns></returns>
    Public Function serialPortConfigurations()
        'Dim ports As String() = SerialPort.GetPortNames()
        mySerialPort.PortName = "COM5"                             'COM Port for serial port communication
        mySerialPort.Parity = Parity.Even                          'Parity is Even
        mySerialPort.BaudRate = 9600                               'Baud rate for coomunication with Konica Minolta device
        mySerialPort.StopBits = 1                                  'Number of Stop Bits 
        mySerialPort.DataBits = 7                                  'Character length
        'mySerialPort.ReadTimeout = 100                            'Read timeout when reading data 
        'mySerialPort.WriteTimeout = 1000                          'Write timeout when writing data
        Try
            mySerialPort.Open()                                     'Open serial port for communication
        Catch ex As Exception
            Throw (New Exception("Serial Port is not open. Please check if Konica device is connected . " + ex.Message))
        End Try

        '_continue = True
        'If (mySerialPort.Open()) Then                              'Open serial port for communication 
        '    GoTo Here
        'Else
        '    Throw New SerialPortNotOpenException("Serial port is not open")
        'End If
        'Here:
    End Function



    '*******************************
    '*** Starting Measurement ***
    '*******************************
    ''' <summary>
    ''' This function is used to read data from the konica minolta device
    ''' 
    ''' </summary>
    ''' <param name="count">If the value of count is 0 , it indicates that the data is being read from the deVICE for the first time and program is going to send all commands from the PC mode command</param>
    ''' <returns></returns>
    Public Function StartReadingData(count As Integer)
        intErrflg = 0
        If count = 0 Then
Check:
            If mySerialPort.IsOpen Then                                 'If serial port is open then start reading data
                GoTo readData
            Else
                Try
                    mySerialPort.Open()                                        'Open serial port for communication
                    GoTo readData
                Catch ex As Exception
                    Throw (New Exception("Serial Port is not open. Please check if Konica device is connected .  " + ex.Message))
                End Try
                'mySerialPort.Open()
                'Try
                '    {
                '    mySerialPort.Open()
                '    }
                'Catch ex As Exception
                '    {
                '    Throw New Exception(ex.Message)
                '        }
                'End Try
            End If
            'Else open serial port and start reading data
        Else
            If mySerialPort.IsOpen Then                                 'If serial port is open then go to Step 3 Hold mode
                GoTo Step5
            Else
                Try
                    mySerialPort.Open()                                        'Open serial port for communication
                    GoTo readData
                Catch ex As Exception
                    Throw (New Exception("Serial Port is not open. Please check if Konica device is connected .  " + ex.Message))
                End Try

            End If
            ''GoTo Com5



readData:
            '------------------------------
            'Step 2 PC MODE
            '------------------------------
            j = 0
            strSndCommand = "00541  "                                ' Command to set the Konica Minolta device to PC Connection mode
            Call CmdSendRecv(1, 2)
            Call ErrCheck()                                          ' Check for Errors in data received from Konica Minolta device

            If (label_value = "Check") Then
                'temp = temp + 1
                GoTo Check
            End If

            If intErrflg = 1 Then
                Exit Function
            End If

            

            'Code to wait 500ms
            System.Threading.Thread.Sleep(500)



            'Up:
            '------------------------------
            'Step 3 HOLD ON
            '------------------------------
            strSndCommand = "99551  0"                              ' Command to set the Konica Minolta device to Hold status
            Call CmdSendRecv(0, 3)

            'Code to wait 500ms
            System.Threading.Thread.Sleep(500)




            '------------------------------
            'Step 4 EXT MODE
            '------------------------------

            For i = 0 To n - 1
                strSndCommand = sensorNumber(i) & "4010  "          ' Command to set the Konica Minolta device to External mode
                Call CmdSendRecv(1, 4, sensorNumber(i))
                Call ErrCheck()                                     ' Check for Errors in data received from Konica Minolta device

                If (label_value = "Check") Then
                    GoTo Check
                End If

                If intErrflg = 1 Then
                    Exit Function
                End If

            Next i

            'Code to wait 175ms
            System.Threading.Thread.Sleep(175)



Step5:
            '------------------------------
            'Step 5 EXT MEASUREMENT
            '------------------------------
            strSndCommand = "994021  "                           'Command to set the Konica Minolta device to take Measurement
            Call CmdSendRecv(0, 5)

            'Code to wait 500ms 
            System.Threading.Thread.Sleep(500)




            '------------------------------
            'Step 6 READ MEASUREMENT DATA
            '------------------------------
            For j = 0 To n - 1
                strSndCommand = sensorNumber(j) & "081200"        'Command to set the Konica Minolta device to read measurement data from each sensor 
                Call CmdSendRecv(1, 6, sensorNumber(j))
                Call ErrCheck()                                   'Check for Errors in data received from Konica Minolta device
                If (label_value = "Check") Then
                    GoTo Check
                End If

                If intErrNO = 10 Then
                    GoTo Step5
                End If
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
    ''' <summary>
    ''' 'This function is used to send and receive commands from Konica Minolta device
    ''' </summary>
    ''' <param name="FlgTimeoutCheck"></param>
    ''' <param name="mode">Mode is used to indicate for different operations</param>
    ''' <param name="sensorNo"> Optional parameter. For commands which is sent for induvidual sensors , Sensor number is sent as a parameter</param>
    ''' <remarks></remarks>
    Public Sub CmdSendRecv(FlgTimeoutCheck As Integer, mode As Integer, Optional sensorNo As String = Nothing)
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
        System.Threading.Thread.Sleep(500)

        mySerialPort.WriteLine(strSendStr)                          'Write data to serial port

        If ((mode = 2) Or (mode = 4) Or (mode = 6)) Then

            'Reception & TimeOut Check
            '-------------------------
            'Code to handle data receiving within timeout limit here
            '-------------------------

            'Timer()
            'Dim stopwatch As Stopwatch = stopwatch.StartNew
            'stopwatch = stopwatch.StartNew
            'System.Threading.Thread.Sleep(1000)
            'stopwatch.Stop()
            'Dim TSTimer As New System.Timers.Timer(5000) ' 5 secs
            'TSTimer.Start()
            'TSTimer_tick()
            System.Threading.Thread.Sleep(1000)
            ' strReceiveStr = mySerialPort.ReadExisting()

            ' While mySerialPort.ReadTimeout
            'Try

            strReceiveStr = mySerialPort.ReadExisting()
            'Console.WriteLine(message)
            If (Not strReceiveStr = "") Then
                GoTo here

            End If


            If (strReceiveStr = "") Then
                If (mode = 2) Then
                    label_value = "Check"
                    ' Throw New SerialPortReadDataErrorException("Unable to receive command for PC mode")
                ElseIf (mode = 4) Then
                    label_value = "Check"
                    'Throw New SerialPortReadDataErrorException("Sensor Number " & sensorNo & " is not connected. Please reconnect the sensor , switch off and switch on the Konica device.")
                ElseIf (mode = 6) Then
                    label_value = "Check"
                    'Throw New SerialPortReadDataErrorException("No data received from serial port at " & sensorNo & " trying again")
                    ' GoTo readData
                End If
            Else


here:
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
            End If
        Else
            Exit Sub
        End If
    End Sub


    '**************************
    '*** BCC Calculation ***
    '**************************
    ''' <summary>
    ''' This function is used to perform Block Character Check for commands received from the Konica Minolta device
    ''' </summary>
    ''' <param name="Command"></param>
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


    '******************************************************************
    '*** Error Check for Data receved from Konica Minolta device***
    '******************************************************************
    ''' <summary>
    ''' This function is used to check if there are any errors in the command received from the Konica Minolta device
    ''' </summary>
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
        Select Case intErrNO
            Case 0 : Exit Sub
            Case 1 : intErrflg = 1
                errMsg = "POWER OF SENSOR HEAD WAS OFF.(No." & sensorNumber(j) & "). PLEASE TURN ON THE POWER AND THE CL-200A OFF AND THEN BACK ON"
                'Throw (New PowerOfSensorOffException(errMsg))
                Throw (New KonicaErrorException(errMsg))
            Case 2 : intErrflg = 1
                errMsg = "EE-PROM ERROR(No." & "). PLEASE TURN THE CL-200A OFF AND THEN BACK ON"
                ' Throw (New EePromErrorException(errMsg))
                Throw (New KonicaErrorException(errMsg))
            Case 3 : intErrflg = 1
                errMsg = "EE-PROM ERROR(No." & ") . PLEASE TURN THE CL-200A OFF AND THEN BACK ON"
                'Throw (New EePromErrorException(errMsg))
                Throw (New KonicaErrorException(errMsg))
            Case 4 : intErrflg = 1
                errMsg = "EXT ERROR(Sensor No." & sensorNumber(j) & ")"
                'Throw (New ExternalErrorException(errMsg))
                Throw (New KonicaErrorException(errMsg))
            Case 5 : Exit Sub
            Case 6 : Exit Sub
            Case 7 : Exit Sub
            Case 8 : intErrflg = 1
                errMsg = "TIME OUT(No.at " & sensorNumber(j) & ")"
                ' Throw (New TimeOutErrorException(errMsg))
                Throw (New KonicaErrorException(errMsg))
            Case 9 : intErrflg = 1
                errMsg = "BCC ERROR(No.at " & sensorNumber(j) & ")"
                ' Throw (New BCCErrorException(errMsg))
                Throw (New KonicaErrorException(errMsg))
            Case 10 : Exit Sub
            Case 11 : intErrflg = 1
                errMsg = "BATTERY OUT(No." & sensorNumber(j) & ") REPLACE BATTERY OR USE AC ADAPTER"
                ' Throw (New BatteryOutErrorException(errMsg))
                Throw (New KonicaErrorException(errMsg))
        End Select


    End Sub


End Class
