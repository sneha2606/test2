Imports System.Text
Imports System.Data
Imports Microsoft.VisualBasic
Imports System.IO.Ports


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
''' This class is used to read data from 3 sensors : Sensor 0, Sensor 1 and Sensor 2 connected to the Konica Minolta device
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
    Public intErrflg As Integer                                 'Error Flag

    Public strData As String                                    'measurement data Block containing Ev,x and y values 
    Public strData1 As String                                   'Ev (Lux value in lx)
    Public strData2 As String                                   'Tcp (CCT value in K)
    Public strData3 As String                                   '(delta UV)

    Public sngData1 As Double                                   'measurement data Ev (Lux value)
    Public sngData2 As Double                                   'measurement data T  (CCT value)
    Public sngData3 As Double                                   'measurement data (delta uv)

    Public sensorNumber() As String = {"00", "01", "02"}        'sensor numbers 00-Sensor 0,01-Sensor 1,02-Sensor 2
    Public luxcct_Values(2, 1) As String                        'array to store Lux and CCT values for each sensor

    Public n As Integer = 3                                     'number of sensor

    Public errMsg As String                                     'to store the error message for different error conditions in Konica Minolta device

    Public mySerialPort As New SerialPort

    Public label_value As String = ""                            'to store the label value to return to a point
    Public temp As Integer                                       'temporary variable to keep track of number of times no data is received from the Konica device
    Public flag As Integer                                       'temporary variable to keep track of number of times values of 0 0r 0.1 is received from the Konica device
    Public sensor_num(10) As Integer                             'temporary variable to store value 0 or 0.1 when received from the device 

    '***********************************************************************************************
    '*** Serialport configurations setup(based on Konica Minolta communications specifications) ***
    '***********************************************************************************************
    ''' <summary>
    ''' 'This function is used to set the Serial port configurations for Konica Minolta Device
    ''' </summary>
    ''' <param name="comPort"></param>
    ''' <param name="parity"></param>
    ''' <param name="baudRate"></param>
    ''' <param name="stopBits"></param>
    ''' <param name="dataBits"></param>
    ''' <remarks></remarks>
    Public Sub serialPortConfigurations(comPort As String, parity As System.IO.Ports.Parity, baudRate As Integer, stopBits As Integer, dataBits As Integer)

        mySerialPort.PortName = comPort                            'COM Port for serial port communication
        mySerialPort.Parity = parity                               'Parity is Even
        mySerialPort.BaudRate = baudRate                           'Baud rate for coomunication with Konica Minolta device
        mySerialPort.StopBits = stopBits                           'Number of Stop Bits 
        mySerialPort.DataBits = dataBits                           'Character length
        Try
            mySerialPort.Open()                                    'Open serial port for communication
        Catch ex As Exception
            Throw (New Exception("Please check if Konica device is connected."))
        End Try
    End Sub


    ''' <summary>
    ''' 'This function is to close the serial port if it is open
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub closeSerialPort()
        Try
            If (mySerialPort.IsOpen) Then
                mySerialPort.Close()                                'Close serial port 
            End If
        Catch ex As Exception
            Throw (New Exception("Error occured while closing serial port."))
        End Try
    End Sub


    '*******************************
    '*** Starting Measurement ***
    '*******************************
    ''' <summary>
    ''' This function is used to read data from the Konica Minolta device - CL 200A
    ''' </summary>
    ''' <param name="count">If the value of count is 0 , it indicates that the data is being read for the first time from the device and program is going to send all commands from the PC mode command</param>
    Public Sub StartReadingData(count As Integer)
        intErrflg = 0
        temp = 0
        If count = 0 Then
Check:      label_value = ""
            flag = 0
            Array.Clear(sensor_num, 0, sensor_num.Length)
            Try
                mySerialPort.Close()                                   'Close serial port
                mySerialPort.Open()                                    'Open serial port for communication
                GoTo readData
            Catch ex As Exception
                Throw (New Exception("Please check if Konica device is connected."))
            End Try
        Else
            label_value = ""
            flag = 0
            Array.Clear(sensor_num, 0, sensor_num.Length)
            If mySerialPort.IsOpen Then                                 'If serial port is open then go to Step 5
                GoTo Step5
            Else
                Try
                    mySerialPort.Close()                                'Close serial port
                    mySerialPort.Open()                                 'Open serial port for communication
                    GoTo readData
                Catch ex As Exception
                    Throw (New Exception("Please check if Konica device is connected."))
                End Try
            End If
        End If



readData:
        '------------------------------
        'Step 2 PC MODE
        '------------------------------
        strSndCommand = "00541   "                                ' Command to set the Konica Minolta device to PC Connection mode
        Call CmdSendRecv(1, 2)
        Call ErrCheck()                                          ' Check for Errors in data received from Konica Minolta device

        If (label_value = "Check") Then                          ' Go to label Check to recheck if the serialport is open or not
            System.Threading.Thread.Sleep(10000)
            GoTo Check
        End If

        If intErrflg = 1 Then
            Exit Sub
        End If

        'Code to wait 500ms
        System.Threading.Thread.Sleep(500)


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
            Call CmdSendRecv(1, 4)
            Call ErrCheck()                                     ' Check for Errors in data received from Konica Minolta device

            If (label_value = "Check") Then                     ' Go to label Check to recheck if the serialport is open or not
                System.Threading.Thread.Sleep(10000)
                GoTo Check
            End If

            If intErrflg = 1 Then
                Exit Sub
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
            Call CmdSendRecv(1, 6)
            Call ErrCheck()                                   'Check for Errors in data received from Konica Minolta device

            If (label_value = "Check") Then                   'Go to label Check to recheck if the serialport is open or not
                System.Threading.Thread.Sleep(10000)
                GoTo Check
            End If

            If (label_value = "Step5") Then                   'Go to label Step 5 to read data again 
                GoTo Step5
            End If

            If intErrNO = 10 Then
                GoTo Step5
            End If

            If intErrflg = 1 Then
                Exit Sub
            End If

            strData = Microsoft.VisualBasic.Right(strRcvCommand, 18)
            strData1 = Microsoft.VisualBasic.Left(strData, 6)
            strData2 = Microsoft.VisualBasic.Mid(strData, 7, 6)
            strData3 = Microsoft.VisualBasic.Right(strData, 6)

            'Lv,x,y
            sngData1 = Val(Microsoft.VisualBasic.Left(strData1, 5)) * 10 ^ (Val(Microsoft.VisualBasic.Right(strData1, 1)) - 4)
            sngData2 = Val(Microsoft.VisualBasic.Left(strData2, 5)) * 10 ^ (Val(Microsoft.VisualBasic.Right(strData2, 1)) - 4)
            sngData3 = Val(Microsoft.VisualBasic.Left(strData3, 5)) * 10 ^ (Val(Microsoft.VisualBasic.Right(strData3, 1)) - 4)


            'Data obtained from the receptor heads are then used

            'Check if sensor data value is 0 or 0.1
            If sngData1 = 0.1 Or sngData2 = 0.1 Or sngData1 = 0 Or sngData2 = 0 Then
                flag = flag + 1
                sensor_num(flag) = j
                If flag = 2 And sensor_num(flag - 1) = j Then           'If sensor data value is 0 or 0.1 for 2 or more times, retain the values 
                    luxcct_Values(j, 0) = sngData1  'Lux value 
                    luxcct_Values(j, 1) = sngData2  'CCT value
                Else
                    GoTo Step5                                          'Else get value again from device
                End If
            Else
                luxcct_Values(j, 0) = sngData1                          'Lux value 
                luxcct_Values(j, 1) = sngData2                          'CCT value
            End If
        Next j

    End Sub




    '************************************************************************************
    '*** Send command to Konica Minolta device & Receive command from Konica Minolta ***
    '************************************************************************************
    ''' <summary>
    ''' 'This function is used to send and receive commands from Konica Minolta device
    ''' </summary>
    ''' <param name="FlgTimeoutCheck"></param>
    ''' <param name="mode">Mode is used to indicate for different operations</param>
    ''' <remarks></remarks>
    Public Sub CmdSendRecv(FlgTimeoutCheck As Integer, mode As Integer)
        intErrNO = 0
        strRcvCommand = ""
        strReceiveStr = ""

        '-----------------------------------------------
        'Transmission
        '-----------------------------------------------
        Call BCC_Append(strSndCommand)
        strSendStr = Chr(2) & strCommand_ETX_BCC & vbCr & vbLf          'STX + command along with ETX and BCC  + delimiter(CR+LF)

        '------------------------------------------------
        'Code for sending data to Konica Minolta device
        '-----------------------------------------------
        System.Threading.Thread.Sleep(500)

        Try
            mySerialPort.WriteLine(strSendStr)                          'Write data to serial port
        Catch ex As Exception
            Throw (New Exception("Unable to write to Serial Port.Please check if Konica Device is connected properly."))
            Exit Sub
        End Try

        '-----------------------------------------------
        'Reception - Only in Step 2, Step 4 and Step 6
        '-----------------------------------------------
        If ((mode = 2) Or (mode = 4) Or (mode = 6)) Then
            System.Threading.Thread.Sleep(1000)
            Try
                strReceiveStr = mySerialPort.ReadExisting()             'Read data from serial port
            Catch ex As Exception
                Throw (New Exception("Unable to read data from Serial Port. Please check if Konica Device is connected properly."))
                Exit Sub
            End Try

            If (strReceiveStr = "") Or (strReceiveStr = " ") Or (strReceiveStr.Length < 2) Then 'If no data is received from the device
                temp = temp + 1
                label_value = "Check"                                    ' Go to label Check to recheck if the serialport is open or not
                If (temp > 2) Then                                       ' If no data is received for more than 2 times , throw exception 
                    Throw New Exception("No Data received from Konica Device.Please check if device is connected properly")
                    Exit Sub
                End If


            ElseIf (strReceiveStr = "0") And (mode = 6) Then             ' If data received from device is 0 and mode is 6(reading measurement data), get data from device again
                label_value = "Step5"


            ElseIf (strReceiveStr.Length > 1) And (strReceiveStr <> "0") Then
                '-------------------------
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
            intBCC = intBCC Xor Asc(Mid(strCommand_ETX, i, 1))              'XOR data
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
            intErrNO = 11                                                   'Battery Out
            Exit Sub
        ElseIf Mid(strRcvCommand, 7, 1) = "6" Then
            intErrNO = 10                                                   'Changing Range
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
                errMsg = "Power Of Sensor Head Was Off: Please restart Konica device."
                Throw (New KonicaErrorException(errMsg))
            Case 2 : intErrflg = 1
                errMsg = "EE-PROM Error: Please restart Konica device."
                Throw (New KonicaErrorException(errMsg))
            Case 3 : intErrflg = 1
                errMsg = "EE-PROM Error: Please restart Konica device."
                Throw (New KonicaErrorException(errMsg))
            Case 4 : intErrflg = 1
                errMsg = "EXT Error"
                Throw (New KonicaErrorException(errMsg))
            Case 5 : Exit Sub
            Case 6 : Exit Sub
            Case 7 : Exit Sub
            Case 8 : intErrflg = 1
                errMsg = "Time Out Error"
                Throw (New KonicaErrorException(errMsg))
            Case 9 : intErrflg = 1
                errMsg = "BCC Error"
                Throw (New KonicaErrorException(errMsg))
            Case 10 : Exit Sub
            Case 11 : intErrflg = 1
                errMsg = "Battery Out: Please replace the Battery or use AC adapter."
                Throw (New KonicaErrorException(errMsg))
        End Select
    End Sub
End Class
