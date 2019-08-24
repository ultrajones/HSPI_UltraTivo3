Imports System.Threading
Imports System.Text
Imports System.Net.Sockets
Imports System.Net
Imports System.Text.RegularExpressions
Imports HomeSeerAPI.VSVGPairs
Imports HomeSeerAPI

Public Class TiVo

  Protected SendCommandThread As Thread
  Protected WatchdogThread As Thread

  Protected CommandQueue As New Queue

  Protected m_ConnectionUUID As String
  Protected m_ConnectionAddr As String
  Protected m_ConnectionPort As Integer = 31339

  Protected m_Machine As String = String.Empty
  Protected m_Platform As String = String.Empty
  Protected m_SWVersion As String = String.Empty
  Protected m_Services As String = String.Empty

  Protected m_Connected As Boolean = False
  Protected m_Initialized As Boolean = False
  Protected m_WatchdogActive As Boolean = False
  Protected m_WatchdogDisabled As Boolean = False

  Protected m_LastReport As DateTime = DateTime.Now

  Protected m_CmdWait As String = ""
  Protected m_CmdAttempt As Byte = 0

  Protected m_MaxAttempts As Byte = 2
  Protected m_MaxWaitTime As Single = 4.0

#Region "TiVo Object"

  Public ReadOnly Property ConnectionUUID() As String
    Get
      Return Me.m_ConnectionUUID
    End Get
  End Property

  Public ReadOnly Property ConnectionAddr() As String
    Get
      Return Me.m_ConnectionAddr
    End Get
  End Property

  Public ReadOnly Property ConnectionPort() As Integer
    Get
      Return Me.m_ConnectionPort
    End Get
  End Property

  Public ReadOnly Property Platform() As String
    Get
      Return Me.m_Platform
    End Get
  End Property

  Public ReadOnly Property Machine() As String
    Get
      Return Me.m_Machine
    End Get
  End Property

  Public ReadOnly Property Services() As String
    Get
      Return Me.m_Services
    End Get
  End Property

  Public ReadOnly Property ConnectionStatus() As String
    Get
      Select Case m_Connected
        Case True
          Return "Connected"
        Case Else
          Return "Disconnected"
      End Select
    End Get
  End Property

  Public Property LastReport() As DateTime
    Set(ByVal value As DateTime)
      m_LastReport = value
    End Set
    Get
      Return m_LastReport
    End Get
  End Property

  Public Sub New(ByVal strConnectionUUID As String,
                 ByVal strConnectionAddr As String,
                 ByVal strMachine As String,
                 ByVal strPlatform As String,
                 ByVal strSWVersion As String,
                 ByVal strServices As String)

    MyBase.New()

    '
    ' Set the UUID for this object
    '
    Me.m_ConnectionUUID = strConnectionUUID
    Me.m_ConnectionAddr = strConnectionAddr

    Me.m_Machine = strMachine
    Me.m_Platform = strPlatform
    Me.m_SWVersion = strSWVersion
    Me.m_Services = strServices

    '
    ' Update the TiVo Connection Device
    '
    CreateTiVoControlDevice(m_ConnectionUUID, m_Machine)
    CreateTiVoChannelDevice(m_ConnectionUUID, m_Machine)
    CreateTiVoVideoDevice(m_ConnectionUUID, m_Machine)

    '
    ' Start the process command queue thread
    '
    SendCommandThread = New Thread(New ThreadStart(AddressOf ProcessCommandQueue))
    SendCommandThread.Name = "ProcessCommandQueue"
    SendCommandThread.Start()

    Dim strMessage As String = SendCommandThread.Name & " Thread Started"
    WriteMessage(strMessage, MessageType.Debug)

    '
    ' Start the watchdog thread
    '
    WatchdogThread = New Thread(New ThreadStart(AddressOf TiVoWatchdogThread))
    WatchdogThread.Name = "WatchdogThread"
    WatchdogThread.Start()

    strMessage = WatchdogThread.Name & " Thread Started"
    WriteMessage(strMessage, MessageType.Debug)

  End Sub

  Protected Overrides Sub Finalize()

    Try

      '
      ' Abort SendCommandThread
      '
      If SendCommandThread.IsAlive = True Then
        SendCommandThread.Abort()
      End If

      '
      ' Abort WatchdogThread
      '
      If WatchdogThread.IsAlive = True Then
        WatchdogThread.Abort()
      End If

      DisconnectEthernet()

    Catch pEx As Exception

    End Try

    MyBase.Finalize()

  End Sub

#End Region

#Region "HSPI - Watchdog"

  ''' <summary>
  ''' Thread that checks we are still connected to the Pioneer AVR
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub TiVoWatchdogThread()

    Dim bAbortThread As Boolean = False
    Dim strMessage As String = ""
    Dim dblTimerInterval As Single = 1000 * 30
    Dim iSeconds As Long = 0

    Try
      '
      ' Stay in PioneerWatchdogThread for duration of program
      '
      While bAbortThread = False

        Try

          If m_WatchdogDisabled = True Then

            dblTimerInterval = 1000 * 60

            strMessage = String.Format("Watchdog Timer indicates the TiVo DVR '{0}' auto reconnect is disabled.", m_ConnectionAddr)
            WriteMessage(strMessage, MessageType.Debug)

          ElseIf m_Initialized = True Then

            If IsDate(m_LastReport) Then
              iSeconds = DateDiff(DateInterval.Second, m_LastReport, DateTime.Now)
            End If

            If m_Platform = "tcd/Series6" Then
              iSeconds = -1
            End If

            strMessage = String.Format("Watchdog Timer indicates a response from the TiVo DVR '{0}' was received at {1}.", m_ConnectionAddr, m_LastReport.ToString)
            WriteMessage(strMessage, MessageType.Debug)

            '
            ' Test to see if we are connected and that we have received a response within the past 300 seconds
            '
            Call CheckPhysicalConnection()

            If iSeconds > 300 Or m_WatchdogActive = True Or m_Connected = False Then
              '
              ' Action for initial watchdog trigger
              '
              If m_WatchdogActive = False Then
                m_WatchdogActive = True
                dblTimerInterval = 1000 * 30

                Dim strWatchdogReason As String = String.Format("No response response from the TiVo DVR '{0}' for {1} seconds.", m_ConnectionAddr, iSeconds)
                If m_Connected = False Then
                  strWatchdogReason = String.Format("Connection to the TiVo DVR '{0}' was lost.", m_ConnectionAddr)
                End If

                strMessage = String.Format("Watchdog Timer indicates {0}.  Attempting to reconnect ...", strWatchdogReason)
                WriteMessage(strMessage, MessageType.Warning)

                '
                ' Check watchdog trigger
                '

              End If

              '
              ' Ensure everything is closed properly and attempt a reconnect
              '
              Call DisconnectFromTiVo()
              Call ConnectToTiVo()

              If m_Connected = False Then

                WriteMessage("Watchdog Timer reconnect attempt failed.", MessageType.Warning)

                dblTimerInterval *= 2
                If dblTimerInterval > 3600000 Then
                  dblTimerInterval = 3600000
                End If

              Else

                WriteMessage("Watchdog Timer reconnect attempt succeeded.", MessageType.Informational)
                m_WatchdogActive = False
                dblTimerInterval = 1000 * 30

                '
                ' Check watchdog trigger
                '

              End If

            Else
              '
              ' Plug-in is connected to the TiVo
              '
              m_WatchdogActive = False
              dblTimerInterval = 1000 * 30

              strMessage = String.Format("Watchdog Timer indicates a response from the TiVo DVR '{0}' was received {1} seconds ago.", m_ConnectionAddr, iSeconds.ToString)
              WriteMessage(strMessage, MessageType.Debug)

            End If

          End If

          '
          ' Sleep Watchdog Thread
          '
          strMessage = String.Format("Watchdog Timer thread for the TiVo DVR '{0}' sleeping for {1}.", m_ConnectionAddr, dblTimerInterval.ToString)
          WriteMessage(strMessage, MessageType.Debug)

          Thread.Sleep(dblTimerInterval)

        Catch pEx As ThreadInterruptedException
          '
          ' Thread sleep was interrupted
          '
          m_Initialized = True
          strMessage = String.Format("Watchdog Timer thread for the TiVo DVR '{0}' was interrupted.", m_ConnectionAddr, iSeconds.ToString)
          WriteMessage(strMessage, MessageType.Debug)

        Catch pEx As Exception
          '
          ' Process Exception
          '
          Call ProcessError(pEx, "PioneerWatchdogThread()")
        End Try

      End While ' Stay in thread until we get an abort/exit request

    Catch ab As ThreadAbortException
      '
      ' Process Thread Abort Exception
      '
      bAbortThread = True      ' Not actually needed
      Call WriteMessage("Abort requested on PioneerWatchdogThread", MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "PioneerWatchdogThread()")
    Finally

    End Try

  End Sub

  ''' <summary>
  ''' Determine if TiVo connection is active
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub CheckPhysicalConnection()

    Try

      If TcpClient.Connected = False Or TcpClient.Client.Connected = False Then
        m_Connected = False
      Else

        If NetworkStream.CanRead = False Or NetworkStream.CanWrite = False Then
          m_Connected = False
        End If

      End If

    Catch pEx As Exception

    End Try

  End Sub

#End Region

#Region "TiVo Connection Initilization/Shutdown"

  ''' <summary>
  ''' Initialize the connection to Pioneer AVR
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ConnectToTiVo() As String

    Dim strMessage As String = ""
    Dim strPortName As String = ""
    Dim strPortAddr As Integer = 0

    strMessage = "Entered ConnectToPAVR() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      '
      ' Try connection via the Pioneer AVR
      '
      strMessage = String.Format("Initiating TiVo connection to '{0}' ...", m_ConnectionAddr)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Inititalize Ethernet connection
      '
      m_Connected = ConnectToEthernet(m_ConnectionAddr, m_ConnectionPort)

      If m_Initialized = False Then
        m_Initialized = m_Connected
      End If

      If m_Connected = False Then
        Throw New Exception(String.Format("Unable to connect to TiVo at '{0}'.", m_ConnectionAddr))
      End If

      '
      ' We are connected here
      '
      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      m_Connected = False
      Return pEx.ToString
    Finally
      '
      ' Update the Pioneer A/V Receiver connection status
      '
      UpdateTiVoConnectionDevice(m_ConnectionUUID, m_Machine, m_Connected)
    End Try

  End Function

  ''' <summary>
  ''' Disconnect the connection to Pioneer AVR
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DisconnectFromTiVo() As Boolean

    Dim strMessage As String = ""

    strMessage = "Entered DisconnectFromPAVR() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      DisconnectEthernet()

      '
      ' Reset Global Variables
      '
      m_Connected = False

      Return True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "DisconnectFromPAVR()")
      Return False
    Finally
      '
      ' Update the TiVo connection status
      '
      UpdateTiVoConnectionDevice(m_ConnectionUUID, m_Machine, m_Connected)
    End Try

  End Function

  ''' <summary>
  ''' Reconnect the connection to Pioneer AVR
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Reconnect()

    '
    ' Ensure plug-in is disconnected
    '
    DisconnectFromTiVo()

    '
    ' Ensure watchdog is not disabled
    '
    m_WatchdogDisabled = False

    '
    ' Ensure PAVR is marked as initialized
    '
    m_Initialized = True

    '
    ' Interrupt the watchdog thread
    '
    If WatchdogThread.ThreadState = ThreadState.WaitSleepJoin Then
      If m_WatchdogActive = False Then
        WatchdogThread.Interrupt()
      End If
    End If

  End Sub

  ''' <summary>
  ''' isconnect from the Pioneer AVR
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Disconnect()

    '
    ' Ensure the watchdog is disabled
    '
    m_WatchdogDisabled = True

    '
    ' Disconnect from the PAVR
    '
    DisconnectFromTiVo()

  End Sub

#End Region

#Region "TiVo Command Processing"

  Public Sub SetChannel(ByVal strChannel As String)

    Try

      If m_Services.Length = 0 Then

        For Each number As Char In strChannel
          If IsNumeric(number) = True Then
            AddCommand(String.Format("IRCODE NUM{0}", number), True)
          End If
        Next
        AddCommand(String.Format("IRCODE {0}", "ENTER"), True)

      Else

        Dim strCommand As String = String.Format("SETCH {0}", strChannel)
        AddCommand(strCommand, True)
      End If

    Catch pEx As Exception
      ProcessError(pEx, "SetChannel()")
    End Try

  End Sub

  ''' <summary>
  ''' Subroutine to process Pioneer AVR commands
  ''' </summary>
  ''' <param name="buttonKey"></param>
  ''' <remarks></remarks>
  Public Sub SendCommand(ByVal buttonKey As Integer)

    Try

      Dim strButtonName As String = Buttons(buttonKey)
      WriteMessage(String.Format("Processing {0} command {1}.", "TiVo DVR", strButtonName), MessageType.Debug)

      Select Case buttonKey
        Case Is <= 58
          AddCommand(String.Format("IRCODE {0}", strButtonName.ToUpper), True)
        Case 59, 60, 61, 62
          AddCommand(String.Format("TELEPORT {0}", strButtonName.Replace("Teleport_", "").ToUpper), True)
        Case Is <= 63
          AddCommand(String.Format("KEYBOARD {0}", strButtonName.ToUpper), True)
        Case Else
          Throw New Exception(String.Format("Unrecognized command {0}", strButtonName))
      End Select

    Catch pEx As Exception
      ProcessError(pEx, "SendTiVoCommand()")
    End Try

  End Sub

  ''' <summary>
  ''' Adds command to command buffer for processing
  ''' </summary>
  ''' <param name="strCommand"></param>
  ''' <param name="bForce"></param>
  ''' <remarks></remarks>
  Protected Sub AddCommand(ByVal strCommand As String, Optional ByVal bForce As Boolean = False)

    Try
      '
      ' bForce may be used to add a command to repeat a command
      '
      If m_Connected = True Then

        If CommandQueue.Contains(strCommand) = False Or bForce = True Then
          CommandQueue.Enqueue(strCommand)
        End If
      Else
        WriteMessage(String.Format("Ignoring command '{0}' because the TiVo '{1}' is not connected.", strCommand, m_ConnectionAddr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "AddCommand()")
    End Try

  End Sub

  ''' <summary>
  ''' Processes commands and waits for the response
  ''' </summary>
  ''' <remarks></remarks>
  Protected Sub ProcessCommandQueue()

    Dim strCommand As String = ""

    Dim dtStartTime As Date
    Dim etElapsedTime As TimeSpan

    Dim bAbortThread As Boolean = False

    Dim iMillisecondsWaited As Integer
    Dim iCmdAttempt As Integer = 0
    Dim iMaxWaitTime As Single = 0

    Try

      While bAbortThread = False

        '
        ' Process commands in command queue
        '
        While CommandQueue.Count > 0 And m_Connected = True And gIOEnabled = True

          '
          ' Set the command response we are waiting for
          '
          strCommand = CommandQueue.Peek
          iMaxWaitTime = 0

          '
          ' Determine if we need to modify the strCmdWait
          '
          If Regex.IsMatch(strCommand, "^FORCECH") = True Then
            '
            ' Tunes the DVR to the specified channel
            '
            m_CmdWait = "CH_STATUS|CH_FAILED"
            iMaxWaitTime = m_MaxWaitTime

          ElseIf Regex.IsMatch(strCommand, "^KEYBOARD") = True Then
            '
            ' Tunes the DVR to the specified channel
            '
            m_CmdWait = ""
            iMaxWaitTime = m_MaxWaitTime

          ElseIf Regex.IsMatch(strCommand, "^IRCODE") = True Then
            '
            ' Sends a code corresponding to a button on the remote
            '
            m_CmdWait = ""
            iMaxWaitTime = m_MaxWaitTime

          ElseIf Regex.IsMatch(strCommand, "^SETCH") = True Then
            '
            ' Tunes the DVR to the specified channel
            '
            m_CmdWait = "CH_STATUS|CH_FAILED"
            iMaxWaitTime = m_MaxWaitTime

          ElseIf Regex.IsMatch(strCommand, "^TELEPORT LIVETV") = True Then
            '
            ' Tunes the DVR to the specified channel
            '
            m_CmdWait = "LIVETV_READY"
            iMaxWaitTime = m_MaxWaitTime

          ElseIf Regex.IsMatch(strCommand, "^TELEPORT") = True Then
            '
            ' Tunes the DVR to the specified channel
            '
            m_CmdWait = ""
            iMaxWaitTime = m_MaxWaitTime

          Else
            '
            ' No response expected
            '
            m_CmdWait = ""
            iMaxWaitTime = 0
          End If

          WriteMessage(String.Format("Sending command '{0}' to TiVo '{1}', waiting for response '{2}'.", strCommand, m_ConnectionAddr, m_CmdWait), MessageType.Debug)

          '
          ' Increment the counter
          '
          iCmdAttempt += 1

          WriteMessage(String.Format("Sending command '{0}' to TiVo '{1}', attempt #{2}.", strCommand, m_ConnectionAddr, iCmdAttempt), MessageType.Debug)
          SendToTiVo(strCommand)

          '
          ' Determine if we need to wait for a response
          '
          If iMaxWaitTime > 0 And m_CmdWait.Length > 0 Then
            '
            ' A response to our command is expected, so lets wait for it
            '
            WriteMessage(String.Format("Waiting for the TiVo '{0}' to respond with '{1}' for up to {2} seconds...", m_ConnectionAddr, m_CmdWait, iMaxWaitTime), MessageType.Debug)

            '
            ' Keep track of when we started waiting for the response
            '
            dtStartTime = Now

            '
            '  Wait for the proper response to come back, or the maximum wait time
            '
            Do
              '
              ' Sleep this thread for 100ms giving the receive function time to get the response
              '
              Thread.Sleep(100)

              '
              ' Find out how long we have been waiting in total
              '
              etElapsedTime = Now.Subtract(dtStartTime)
              iMillisecondsWaited = etElapsedTime.Milliseconds + (etElapsedTime.Seconds * 1000)

              '
              ' Loop until the expected command was received (strCmdWait is cleared) or we ran past the maximum wait time
              '
            Loop Until m_CmdWait.Length = 0 Or iMillisecondsWaited > iMaxWaitTime * 1000 ' Now abort if the command was recieved or we ran out of time

            WriteMessage(String.Format("Waited {0} milliseconds for the command response '{1}' from TiVo '{2}'.", iMillisecondsWaited, strCommand, m_ConnectionAddr), MessageType.Debug)

            If m_CmdWait.Length > 0 Or iCmdAttempt > m_MaxAttempts Then
              '
              ' Command failed, so lets stop trying to send this commmand
              '
              WriteMessage(String.Format("No response/improper response from TiVo '{0}' to command '{1}'.", m_ConnectionAddr, strCommand), MessageType.Warning)
              m_CmdWait = String.Empty

              '
              ' Only Dequeue the command if we have tried more than MAX_ATTEMPTS times
              '
              If iCmdAttempt >= m_MaxAttempts Then
                CommandQueue.Dequeue()
                m_CmdWait = String.Empty
                iCmdAttempt = 0
              End If
            Else
              CommandQueue.Dequeue()
              iCmdAttempt = 0
            End If

          Else
            '
            ' No response expected, so remove command from queue
            '
            WriteMessage(String.Format("Command {0} does not produce a result.", strCommand), MessageType.Debug)
            CommandQueue.Dequeue()
            m_CmdWait = String.Empty
            iCmdAttempt = 0
          End If

        End While ' Done with all commands in queue

        '
        ' Give up some time to allow the main thread to populate the command queue with more commands
        '
        Thread.Sleep(100)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("ProcessCommandQueue thread received abort request, terminating normally."), MessageType.Debug)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "ProcessCommandQueue()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("ProcessCommandQueue terminated."), MessageType.Debug)
    End Try

  End Sub

#End Region

#Region "TiVo Protocol Processing"

  ''' <summary>
  ''' Sends command to TiVo
  ''' </summary>
  ''' <param name="strDataToSend"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function SendToTiVo(ByVal strDataToSend As String) As Boolean

    Dim strMessage As String = ""

    strMessage = "Entered SendToTiVo() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Format packet
      '
      Dim strPacket As String = FormatDataPacket(strDataToSend)

      '
      ' Set data to Ethernet connection
      '
      If m_Connected = True And gIOEnabled = True Then
        strMessage = String.Format("Sending '{0}' to TiVo '{1}' via Ethernet.", strPacket, m_ConnectionAddr)
        Call WriteMessage(strMessage, MessageType.Debug)
        Return SendMessageToEthernet(strPacket)
      Else
        Return False
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SendToTiVo()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Processes a received data string
  ''' </summary>
  ''' <param name="strDataRec"></param>
  ''' <remarks></remarks>
  Private Sub ProcessReceived(ByVal strDataRec As String)

    Dim strMessage As String = ""

    strMessage = String.Format("Data '{0}' sent by TiVo '{1}'.", strDataRec, m_ConnectionAddr)
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      If Regex.IsMatch(strDataRec, "CH_STATUS|CH_FAILED") Then
        Dim strRegex As String = String.Format("^(?<CMD>({0})) (?<DATA>.+)$", "CH_STATUS|CH_FAILED")
        '
        ' Extract data based on expected response
        '
        Dim colMatches As MatchCollection = Regex.Matches(strDataRec, strRegex)

        If colMatches.Count > 0 Then
          For Each objMatch As Match In colMatches

            Dim strCmd As String = objMatch.Groups("CMD").Value
            Dim strData As String = objMatch.Groups("DATA").Value

            ProcessCommand(strCmd, strData)
          Next

        Else
          '
          ' Buffer failed to match regular expression
          '
          WriteMessage(String.Format("Invalid data '{0}' received from TiVo '{1}'.", strDataRec, m_ConnectionAddr), MessageType.Error)
        End If

      ElseIf Regex.IsMatch(strDataRec, "LIVETV_READY") Then
        ProcessCommand(strDataRec, "")
      ElseIf Regex.IsMatch(strDataRec, "INVALID_KEY") Then
        ProcessCommand(strDataRec, "")
      ElseIf Regex.IsMatch(strDataRec, "INVALID_COMMAND") Then
        ProcessCommand(strDataRec, "")
      Else
        WriteMessage(String.Format("Invalid data '{0}' received from TiVo '{1}'.", strDataRec, m_ConnectionAddr), MessageType.Error)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "ProcessReceived()")
    End Try

  End Sub

  ''' <summary>
  ''' Subroutine to process the data received from the Pioneer AVR
  ''' </summary>
  ''' <param name="strCmd"></param>
  ''' <param name="strData"></param>
  ''' <remarks></remarks>
  Private Sub ProcessCommand(ByVal strCmd As String, ByVal strData As String)

    Dim strMessage As String = "Entered ProcessCommand() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Test to see if this was one of our commands
      '
      If Regex.IsMatch(strCmd, m_CmdWait) = True Then
        WriteMessage(String.Format("Expected command response '{0}' received from TiVo '{1}'.", m_CmdWait, m_ConnectionAddr), MessageType.Debug)
        '
        ' This is the flag that tells the outbound command processor that we in fact received the expected command
        '
        m_CmdWait = String.Empty
      Else
        '
        ' Process the command sent by the TiVo DVR
        '
        strMessage = String.Format("Command response '{0}', data '{1}' received from TiVo '{2}'.", strCmd, strData, m_ConnectionAddr)
        Call WriteMessage(strMessage, MessageType.Debug)
      End If

      '
      ' Update the last report date/time
      '
      m_LastReport = DateTime.Now

      Select Case strCmd
        Case "MISSING_TELEPORT_NAME"
          WriteMessage(String.Format("TiVo [{0}] rejected the command due to '{1}'.", m_ConnectionAddr, strCmd), MessageType.Error)
          m_CmdWait = String.Empty
        Case "INVALID_KEY"
          WriteMessage(String.Format("TiVo [{0}] rejected the command due to '{1}'.", m_ConnectionAddr, strCmd), MessageType.Error)
          m_CmdWait = String.Empty
        Case "INVALID_COMMAND"
          WriteMessage(String.Format("TiVo [{0}] rejected the command due to '{1}'.", m_ConnectionAddr, strCmd), MessageType.Error)
          m_CmdWait = String.Empty
        Case "CH_FAILED"
          WriteMessage(String.Format("TiVo [{0}] rejected the command due to '{1} {2}'.", m_ConnectionAddr, strCmd, strData), MessageType.Error)
          m_CmdWait = String.Empty
        Case "CH_STATUS"
          '
          ' Process Channel Status Response
          '
          Dim strChannelPri As String = Regex.Match(strData, "^(?<channel>\d+)").Groups("channel").ToString()
          Dim strChannelSec As String = Regex.Match(strData, " (?<subchannel>\d+) ").Groups("subchannel").ToString()

          Dim channelPri As Integer = 0
          Int32.TryParse(strChannelPri, channelPri)
          Dim channelSec As Integer = 0
          Int32.TryParse(strChannelSec, channelSec)

          Dim channelNum As Double = Double.Parse(channelPri.ToString & "." & channelSec.ToString)
          Dim dv_addr As String = String.Format("{0}-Chan", m_ConnectionUUID)
          hspi_devices.SetDeviceValue(dv_addr, channelNum.ToString)
          hspi_devices.SetDeviceString(dv_addr, channelNum.ToString)

        Case "LIVETV_READY"
          '
          ' Process LIVETV_READY Response
          '
          m_CmdWait = String.Empty
        Case Else
          '
          ' Unhandled message (message type, sub-message type)
          '
          WriteMessage(String.Format("TiVo '{0}' rejected unhandled command '{1}'.", m_ConnectionAddr, strCmd), MessageType.Warning)
      End Select

    Catch pEx As Exception
      If m_Connected = True Then
        '
        ' Process program exception
        '
        Call ProcessError(pEx, "ProcessCommand()")
      End If
    End Try

  End Sub

  ''' <summary>
  ''' Formats data packet before sending to the Pioneer AVR
  ''' </summary>
  ''' <param name="strData"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function FormatDataPacket(ByVal strData As String) As String

    strData &= vbCr

    Return strData

  End Function

#End Region

#Region "Ethernet Support"

  Dim TcpClient As System.Net.Sockets.TcpClient
  Dim NetworkStream As System.Net.Sockets.NetworkStream
  Dim ReadThread As Threading.Thread

  ''' <summary>
  ''' Establish connection to Ethernet Module
  ''' </summary>
  ''' <param name="Ip"></param>
  ''' <param name="Port"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ConnectToEthernet(ByVal Ip As String, ByVal Port As Integer) As Boolean

    Dim strMessage As String
    Dim IPAddress As String = ResolveAddress(Ip)

    Try

      Try
        '
        ' Create TCPClient
        '
        TcpClient = New TcpClient(IPAddress, Port)

      Catch pEx As SocketException
        '
        ' Process Exception
        '
        strMessage = String.Format("Ethernet connection could not be made to {0} ({1}:{2}) - {3}", _
                                  IPAddress, Ip.ToString, Port.ToString, pEx.Message)
        Call WriteMessage(strMessage, MessageType.Debug)
        Return False
      End Try

      NetworkStream = TcpClient.GetStream()
      ReadThread = New Thread(New ThreadStart(AddressOf EthernetReadThreadProc))
      ReadThread.Name = "EthernetReadThreadProc"
      ReadThread.Start()

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "ConnectToEthernet()")
      Return False
    End Try

    Return True

  End Function

  ''' <summary>
  ''' Disconnection From Ethernet Module
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub DisconnectEthernet()

    Try
      If ReadThread.IsAlive = True Then
        ReadThread.Abort()
      End If
      NetworkStream.Close()
      TcpClient.Close()
    Catch ex As Exception
      '
      ' Ignore Exception
      '
    End Try

  End Sub

  ''' <summary>
  ''' Send Message to connected IP address (first send buffer length and then the buffer holding message)
  ''' </summary>
  ''' <param name="Message"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Protected Function SendMessageToEthernet(ByVal Message As String) As Boolean

    Try

      Dim Buffer() As Byte = Encoding.ASCII.GetBytes(Message.ToCharArray)

      If TcpClient.Connected = True Then
        NetworkStream.Write(Buffer, 0, Buffer.Length)
        Return True
      Else
        Call WriteMessage("Attempted to write to a closed ethernet stream in SendMessageToEthernet()", MessageType.Warning)
        Return False
      End If

    Catch ex As Exception
      Call ProcessError(ex, "SendMessageToEthernet()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Process to Read Data From TCP Client
  ''' </summary>
  ''' <remarks></remarks>
  Protected Sub EthernetReadThreadProc()

    Dim Str As New StringBuilder
    Dim Dat As Char
    Dim By As Integer

    Try
      '
      ' Set initial value of the string builder object
      '
      Str.Length = 0

      '
      ' Stay in EthernetReadThreadProc while client is connected
      '
      Do While TcpClient.Connected = True

        Do While NetworkStream.DataAvailable = True
          By = NetworkStream.ReadByte()
          Dat = Chr(By)
          If Dat = vbCr Then
            If Str.Length > 0 Then
              ProcessReceived(Str.ToString)
            End If
            Str.Length = 0
          Else
            Str.Append(Dat)
          End If
        Loop

        Thread.Sleep(25)
      Loop
    Catch ab As ThreadAbortException
      '
      ' Process Thread Abort Exception
      '
      Call WriteMessage("Abort requested on EthernetReadThreadProc", MessageType.Debug)
      Return
    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "EthernetReadThreadProc()")
    Finally
      '
      ' Indicate we are no longer connected to the Pioneer AVR
      '
      m_Connected = False
    End Try

  End Sub

  ''' <summary>
  ''' Check ip string to be an ip address or if not try to resolve using DNS
  ''' </summary>
  ''' <param name="hostNameOrAddress"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ResolveAddress(ByVal hostNameOrAddress As String) As String

    Try
      '
      ' Attempt to identify fqdn as an IP address
      '
      IPAddress.Parse(hostNameOrAddress)

      '
      ' If this did not throw then it is a valid IP address
      '
      Return hostNameOrAddress
    Catch ex As Exception
      Try
        ' Try to resolve it through DNS if it is not in IP address form
        ' and use the first IP address if defined as round robbin in DNS
        Dim ipAddress As IPAddress = Dns.GetHostEntry(hostNameOrAddress).AddressList(0)

        Return ipAddress.ToString
      Catch pEx As Exception
        Return ""
      End Try

    End Try

  End Function

#End Region

End Class
