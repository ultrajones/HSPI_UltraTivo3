Imports System.Threading
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Net.Sockets
Imports System.Text
Imports System.Net
Imports HomeSeerAPI
Imports Scheduler
Imports System.ComponentModel
Imports System.Data.Common
Imports System.Data.SQLite

Module hspi_plugin

  '
  ' Declare public objects, not required by HomeSeer
  '
  Public TiVos As New List(Of TiVo)
  Public TiVoLock As New Object

  Public Buttons As New SortedList

  Dim actions As New hsCollection
  Dim triggers As New hsCollection
  Dim conditions As New Hashtable
  Const Pagename = "Events"

  Public SmtpProfiles As New ArrayList

  Public Const IFACE_NAME As String = "UltraTivo3"

  Public Const LINK_TARGET As String = "hspi_ultrativo3/hspi_ultrativo3.aspx"
  Public Const LINK_URL As String = "hspi_ultrativo3.html"
  Public Const LINK_TEXT As String = "UltraTiVo3"
  Public Const LINK_PAGE_TITLE As String = "UltraTivo3 HSPI"
  Public Const LINK_HELP As String = "/hspi_ultrativo3/UltraTiVo3_HSPI_Users_Guide.pdf"

  Public gBaseCode As String = ""
  Public gIOEnabled As Boolean = True
  Public gImageDir As String = "/images/hspi_ultrativo3/"
  Public gHSInitialized As Boolean = False
  Public gINIFile As String = "hspi_" & IFACE_NAME.ToLower & ".ini"

  Public HSAppPath As String = ""

#Region "HSPI - Public Routines"

  ''' <summary>
  ''' Initialize the Hash Tables
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub InitializeHashTables()

    ' Navigation Buttons
    Buttons.Add(1, "Up")
    Buttons.Add(2, "Down")
    Buttons.Add(3, "Left")
    Buttons.Add(4, "Right")
    Buttons.Add(5, "Select")
    Buttons.Add(6, "TiVo")
    Buttons.Add(7, "LiveTv")
    Buttons.Add(8, "Guide")
    Buttons.Add(9, "Info")
    Buttons.Add(10, "Exit")

    ' Control Buttons
    Buttons.Add(11, "ThumbsUp")
    Buttons.Add(12, "ThumbsDown")
    Buttons.Add(13, "ChannelUp")
    Buttons.Add(14, "ChannelDown")
    'Buttons.Add(15, "Mute")
    'Buttons.Add(16, "VolumeDown")
    'Buttons.Add(17, "VolumeUp")
    'Buttons.Add(18, "TvInput")

    Buttons.Add(19, "Video_Mode_Fixed_480i")
    Buttons.Add(20, "Video_Mode_Fixed_480p")
    Buttons.Add(21, "Video_Mode_Fixed_720p")
    Buttons.Add(22, "Video_Mode_Fixed_1080i")
    Buttons.Add(23, "Video_Mode_Hybird")
    Buttons.Add(24, "Video_Mode_Hybird_720p")
    Buttons.Add(25, "Video_Mode_Hybird_1080i")
    Buttons.Add(26, "Video_Mode_Native")

    Buttons.Add(27, "CC_On")
    Buttons.Add(28, "CC_Off")
    Buttons.Add(29, "Options")

    Buttons.Add(30, "Aspect_Correction_Full")
    Buttons.Add(31, "Aspect_Correction_Panel")
    Buttons.Add(32, "Aspect_Correction_Zoom")
    Buttons.Add(33, "Aspect_Correction_Wide_Zoom")

    ' TrickPlay Buttons
    Buttons.Add(34, "Play")
    Buttons.Add(35, "Forward")
    Buttons.Add(36, "Reverse")
    Buttons.Add(37, "Pause")
    Buttons.Add(38, "Slow")
    Buttons.Add(39, "Replay")
    Buttons.Add(40, "Advance")
    Buttons.Add(41, "Record")

    ' Numeric Buttons
    Buttons.Add(42, "Num0")
    Buttons.Add(43, "Num1")
    Buttons.Add(45, "Num2")
    Buttons.Add(46, "Num3")
    Buttons.Add(47, "Num4")
    Buttons.Add(48, "Num5")
    Buttons.Add(49, "Num6")
    Buttons.Add(50, "Num7")
    Buttons.Add(51, "Num8")
    Buttons.Add(52, "Num9")
    Buttons.Add(53, "Enter")
    Buttons.Add(54, "Clear")

    ' Shortcut Buttons
    Buttons.Add(55, "Action_A")
    Buttons.Add(56, "Action_B")
    Buttons.Add(57, "Action_C")
    Buttons.Add(58, "Action_D")

    ' Teleport Buttons
    Buttons.Add(59, "Teleport_TiVo")
    Buttons.Add(60, "Teleport_LiveTv")
    Buttons.Add(61, "Teleport_Guide")
    Buttons.Add(62, "Teleport_NowPlaying")

  End Sub

  ''' <summary>
  ''' Discovers Pioneer AVR units installed on the network
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub DiscoveryBeacon()

    Dim subscriber As UdpClient = New UdpClient()
    Dim addr As IPAddress = IPAddress.Parse("239.255.250.250")
    Dim bAbortThread As Boolean = False

    Try
      WriteMessage("The TiVo Connect Discovery Protocol routine has started ...", MessageType.Debug)

      subscriber = New UdpClient()

      Dim IPEndPoint As IPEndPoint = New IPEndPoint(IPAddress.Any, 2190)

      subscriber.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, 1)
      subscriber.Client.Bind(IPEndPoint)

      While bAbortThread = False

        Dim pdata As Byte() = subscriber.Receive(IPEndPoint)
        Dim strIPAddress As String = IPEndPoint.Address.ToString
        Dim resp As String = Encoding.ASCII.GetString(pdata)

        ' tivoconnect=1                   ' Used to serve as an identifying "signature
        ' swversion=20.3.8-01-2-746       ' This value describes the "primary" software running on the TCM
        ' method=broadcast                ' broadcast (for packets sent using UDP) 
        ' identity=74600119071F67D        ' This value should be unique to the originating TCM 
        ' machine=Media Room              ' This value contains human readable text, naming the TCM, suitable for display to the user. 
        ' platform=tcd / Series4          ' tcd (for TiVo DVR beacons) 
        ' services=TiVoMediaServer:80/http

        Try

          WriteMessage(resp, MessageType.Debug)

          If Regex.IsMatch(resp, "tivoconnect=1") = True Then

            Dim regexPattern As String = "platform=(?<platform>.+)"
            Dim platform As String = Regex.Match(resp, regexPattern).Groups("platform").ToString()

            If Regex.IsMatch(platform, "tcd/Series\d", RegexOptions.IgnoreCase) = True Then

              regexPattern = "machine=(?<machine>.+)"
              Dim machine As String = Regex.Match(resp, regexPattern).Groups("machine").ToString()

              regexPattern = "identity=(?<identity>.+)"
              Dim identity As String = Regex.Match(resp, regexPattern).Groups("identity").ToString()

              regexPattern = "swversion=(?<swversion>.+)"
              Dim swversion As String = Regex.Match(resp, regexPattern).Groups("swversion").ToString()

              regexPattern = "services=(?<services>.+)"
              Dim services As String = Regex.Match(resp, regexPattern).Groups("services").ToString()

              WriteMessage(String.Format("TiVo Beacon from {0} {1} {2}", identity, platform, machine), MessageType.Debug)

              SyncLock TiVoLock

                Dim TiVo As TiVo = TiVos.Find(Function(t) t.ConnectionUUID = identity)

                If Not TiVo Is Nothing Then

                  TiVo.LastReport = DateTime.Now

                Else
                  '
                  ' Add TiVo to the database
                  '
                  Dim bSuccess As Boolean = InsertTiVoDevice(identity, machine, swversion, "TiVo", platform, "Ethernet", strIPAddress, services)
                  If bSuccess = True Then
                    WriteMessage(String.Format("DiscoveryBeacon: TiVo {0} {1} {2} successfully added to database.", identity, platform, machine), MessageType.Informational)
                  End If

                  TiVo = New TiVo(identity, strIPAddress, machine, platform, swversion, services)
                  TiVos.Add(TiVo)

                  Dim strResult As String = TiVo.ConnectToTiVo()
                  If strResult.Length = 0 Then
                    WriteMessage(String.Format("DiscoveryBeacon: Connected to TiVo {0} {1} {2}", identity, platform, machine), MessageType.Informational)
                  Else
                    WriteMessage(String.Format("DiscoveryBeacon: Unable to connect to TiVo {0} {1} {2}.  Please enable Network Remote Control on your TiVo DVR.", identity, platform, machine), MessageType.Error)
                  End If

                End If

              End SyncLock

            End If

          End If

        Catch pEx As Exception
          '
          ' Return message
          '
          WriteMessage("An error occured while processing the TiVo Connect Discovery Protocol beacon message.", MessageType.Error)
          ProcessError(pEx, "DiscoveryBeacon()")
        End Try

        '
        ' Give up some time
        '
        Thread.Sleep(50)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("DiscoveryBeacon thread received abort request, terminating normally."), MessageType.Informational)

      subscriber.DropMulticastGroup(addr)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "DiscoveryBeacon()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("DiscoveryBeacon terminated."), MessageType.Debug)

    End Try

  End Sub

  ''' <summary>
  ''' Connects to the TiVo units installed on the network
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub DiscoveryConnection()

    Dim TiVoDeviceList As New SortedList

    Dim bAbortThread As Boolean = False

    Try
      WriteMessage("The TiVo discovery connection routine has started ...", MessageType.Debug)

      While bAbortThread = False

        Try

          TiVoDeviceList = GetTiVoDeviceList()

          For Each strUUID As String In TiVoDeviceList.Keys

            SyncLock TiVoLock

              Dim TiVo As TiVo = TiVos.Find(Function(t) t.ConnectionUUID = strUUID)

              If TiVo Is Nothing Then
                Dim TiVoDevice As Hashtable = TiVoDeviceList(strUUID)

                Dim identity As String = TiVoDevice("device_uuid")
                Dim strIPAddress As String = TiVoDevice("device_addr")
                Dim machine As String = TiVoDevice("device_name")
                Dim platform As String = TiVoDevice("device_model")
                Dim swversion As String = TiVoDevice("device_swversion")
                Dim services As String = TiVoDevice("device_services")

                TiVo = New TiVo(identity, strIPAddress, machine, platform, swversion, services)
                TiVos.Add(TiVo)

                Dim strResult As String = TiVo.ConnectToTiVo()
                If strResult.Length = 0 Then
                  WriteMessage(String.Format("DiscoveryConnection:  Connected to TiVo {0} {1} {2}", identity, platform, machine), MessageType.Informational)
                Else
                  WriteMessage(String.Format("DiscoveryConnection:  Unable to connect to TiVo {0} {1} {2}.  Please enable Network Remote Control on your TiVo DVR.", identity, platform, machine), MessageType.Error)
                End If

              End If

            End SyncLock

          Next

        Catch pEx As Exception
          '
          ' Return message
          '
          ProcessError(pEx, "DiscoveryConnection()")
        End Try

        '
        ' Give up some time
        '
        Thread.Sleep(1000 * 60)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("DiscoveryConnection thread received abort request, terminating normally."), MessageType.Informational)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "DiscoveryConnection()")

    Finally
      '
      ' Notify that we are exiting the thread
      '
      WriteMessage(String.Format("DiscoveryConnection terminated."), MessageType.Debug)

    End Try

  End Sub

  ''' <summary>
  ''' Get the TiVo Device devices
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetTiVoDeviceList() As SortedList

    Dim TiVoDevices As New SortedList

    Try
      '
      ' Define the SQL Query
      '
      Dim strSQL As String = String.Format("SELECT device_id, device_uuid, device_name, device_swversion, device_make, device_model, device_conn, device_addr, device_services FROM tblTiVoDevices WHERE device_conn <> '{0}'", "Disabled")

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the resutls
          '
          While dtrResults.Read()
            Dim TiVoDevice As New Hashtable

            TiVoDevice.Add("device_id", dtrResults("device_id"))
            TiVoDevice.Add("device_uuid", dtrResults("device_uuid"))
            TiVoDevice.Add("device_name", dtrResults("device_name"))
            TiVoDevice.Add("device_swversion", dtrResults("device_swversion"))
            TiVoDevice.Add("device_make", dtrResults("device_make"))
            TiVoDevice.Add("device_model", dtrResults("device_model"))
            TiVoDevice.Add("device_conn", dtrResults("device_conn"))
            TiVoDevice.Add("device_addr", dtrResults("device_addr"))
            TiVoDevice.Add("device_services", dtrResults("device_services"))

            If TiVoDevices.ContainsKey(dtrResults("device_uuid")) = False Then
              TiVoDevices.Add(dtrResults("device_uuid"), TiVoDevice)
            End If
          End While

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetTiVoDeviceList()")
    End Try

    Return TiVoDevices

  End Function

  ''' <summary>
  ''' Gets the TiVo Devices from the underlying database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetTiVoDevices() As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered GetTiVoDevices() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("SELECT * FROM tblTiVoDevices")

      '
      ' Initialize the command object
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      '
      ' Initialize the dataset, then populate it
      '
      Dim MyDS As DataSet = New DataSet

      Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
      MyDA.SelectCommand = MyDbCommand

      SyncLock SyncLockMain
        MyDA.Fill(MyDS)
      End SyncLock

      '
      ' Get our DataTable
      '
      Dim MyDT As DataTable = MyDS.Tables(0)

      '
      ' Get record count
      '
      Dim iRecordCount As Integer = MyDT.Rows.Count

      If iRecordCount > 0 Then
        '
        ' Build field names
        '
        Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
        For iFieldNum As Integer = 0 To iFieldCount
          '
          ' Create the columns
          '
          Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
          Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

          '
          ' Add the columns to the DataTable's Columns collection
          '
          ResultsDT.Columns.Add(MyDataColumn)
        Next

        '
        ' Let's output our records	
        '
        Dim i As Integer = 0
        For i = 0 To iRecordCount - 1
          '
          ' Create the rows
          '
          Dim dr As DataRow
          dr = ResultsDT.NewRow()
          For iFieldNum As Integer = 0 To iFieldCount
            dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
          Next
          ResultsDT.Rows.Add(dr)
        Next

      End If

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "GetTiVoDevices()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Inserts a new TiVo Device into the database
  ''' </summary>
  ''' <param name="device_uuid"></param>
  ''' <param name="device_name"></param>
  ''' <param name="device_swversion"></param>
  ''' <param name="device_make"></param>
  ''' <param name="device_model"></param>
  ''' <param name="device_conn"></param>
  ''' <param name="device_addr"></param>
  ''' <param name="device_services"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InsertTiVoDevice(ByVal device_uuid As String,
                                   ByVal device_name As String,
                                   ByVal device_swversion As String,
                                   ByVal device_make As String,
                                   ByVal device_model As String,
                                   ByVal device_conn As String,
                                   ByVal device_addr As String,
                                   ByVal device_services As String) As Integer

    Dim strMessage As String = ""
    Dim iRecordsAffected As Integer = 0

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If device_uuid.Length = 0 Or device_name.Length = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to insert new TiVo Device into the database.")
      ElseIf device_conn = "Ethernet" And (device_addr.Length = 0) Then
        Throw New Exception("The IP address is required.  Unable to insert new TiVo Device into the database.")
      End If

      '
      ' Try inserting the TiVo Device into one of the 10 available slots
      '
      For device_id As Integer = 1 To 10

        Dim strSQL As String = String.Format("INSERT INTO tblTiVoDevices (" _
                                     & " device_id, device_uuid, device_name, device_swversion, device_make, device_model, device_conn, device_addr, device_services" _
                                     & ") VALUES (" _
                                     & " {0}, '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}'" _
                                     & ")", device_id, device_uuid, device_name, device_swversion, device_make, device_model, device_conn, device_addr, device_services)

        Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

          dbcmd.Connection = DBConnectionMain
          dbcmd.CommandType = CommandType.Text
          dbcmd.CommandText = strSQL

          Try

            SyncLock SyncLockMain
              iRecordsAffected = dbcmd.ExecuteNonQuery()
            End SyncLock

          Catch pEx As Exception
            '
            ' Ignore this error
            '
          Finally
            dbcmd.Dispose()
          End Try

          If iRecordsAffected > 0 Then
            Return device_id
          End If

        End Using

      Next

      Throw New Exception("Unable to insert TiVo Device into the database.  Please ensure you are not attempting to connect more than 10 TiVo Devices to the plug-in.")

    Catch pEx As Exception
      Call ProcessError(pEx, "InsertTiVoDevice()")
      Return 0
    End Try

  End Function

  ''' <summary>
  ''' Updates existing TiVo Device stored in the database
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="device_name"></param>
  ''' <param name="device_uuid"></param>
  ''' <param name="device_make"></param>
  ''' <param name="device_model"></param>
  ''' <param name="device_conn"></param>
  ''' <param name="device_addr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UpdateTiVoDevice(ByVal device_id As Integer,
                                   ByVal device_name As String,
                                   ByVal device_uuid As String,
                                   ByVal device_make As String,
                                   ByVal device_model As String,
                                   ByVal device_conn As String,
                                   ByVal device_addr As String) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If device_uuid.Length = 0 Or device_name.Length = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to insert new TiVo Device into the database.")
      ElseIf device_conn = "Ethernet" And (device_addr.Length = 0) Then
        Throw New Exception("The IP address is required.  Unable to insert new TiVo Device into the database.")
      End If

      Dim strSql As String = String.Format("UPDATE tblTiVoDevices SET " _
                                          & " device_name='{0}', " _
                                          & " device_uuid='{1}', " _
                                          & " device_make='{2}'," _
                                          & " device_model='{3}'," _
                                          & " device_conn='{4}'," _
                                          & " device_addr='{5}' " _
                                          & "WHERE device_id={6}",
                                             device_name,
                                             device_uuid,
                                             device_make,
                                             device_model,
                                             device_conn,
                                             device_addr,
                                             device_id.ToString)

      '
      ' Build the insert/update/delete query
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSql

      Dim iRecordsAffected As Integer = 0
      SyncLock SyncLockMain
        iRecordsAffected = MyDbCommand.ExecuteNonQuery()
      End SyncLock

      strMessage = "UpdateTiVoDevice() updated " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      MyDbCommand.Dispose()

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "UpdateTiVoDevice()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Removes existing TiVo Device stored in the database
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeleteTiVoDevice(ByVal device_id As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      '
      ' Build the insert/update/delete query
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = String.Format("DELETE FROM tblTiVoDevices WHERE device_id={0}", device_id.ToString)

      Dim iRecordsAffected As Integer = 0
      SyncLock SyncLockMain
        iRecordsAffected = MyDbCommand.ExecuteNonQuery()
      End SyncLock

      strMessage = "DeleteTiVoDevice() removed " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      MyDbCommand.Dispose()

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

      Return True

    Catch pEx As Exception
      Call ProcessError(pEx, "DeleteTiVoDevice()")
      Return False
    End Try

  End Function

#End Region

#Region "HSPI - Misc"

  ''' <summary>
  ''' Gets plug-in setting from INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValueDefault"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSetting(ByVal strSection As String,
                             ByVal strKey As String,
                             ByVal strValueDefault As String) As String

    Dim strMessage As String = ""

    Try
      strMessage = "Entered GetSetting() function."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Get the ini settings
      '
      Dim strValue As String = hs.GetINISetting(strSection, strKey, strValueDefault, gINIFile)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to decrypt the data
      '
      If strKey = "UserPass" Then
        strValue = hs.DecryptString(strValue, "&Cul8r#1")
      End If

      Return strValue

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetSetting()")
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' Saves plug-in setting to INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValue"></param>
  ''' <remarks></remarks>
  Public Sub SaveSetting(ByVal strSection As String,
                         ByVal strKey As String,
                         ByVal strValue As String)

    Dim strMessage As String = ""

    Try
      strMessage = "Entered SaveSetting() subroutine."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to encrypt the data
      '
      If strKey = "UserPass" Then
        If strValue.Length = 0 Then Exit Sub
        strValue = hs.EncryptString(strValue, "&Cul8r#1")
      End If

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Save selected settings to global variables
      '
      'If strSection = "Options" And strKey = "MaxDeliveryAttempts" Then
      '  If IsNumeric(strValue) Then
      '    gMaxAttempts = CInt(Val(strValue))
      '  End If
      'End If

      '
      ' Save the settings
      '
      hs.SaveINISetting(strSection, strKey, strValue, gINIFile)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SaveSetting()")
    End Try

  End Sub

#End Region

#Region "UltraTiVo3 Actions/Triggers/Conditions"

#Region "Trigger Proerties"

  ''' <summary>
  ''' Defines the valid triggers for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetTriggers()
    Dim o As Object = Nothing
    If triggers.Count = 0 Then
      'triggers.Add(o, "Email Delivery Status")           ' 1
    End If
  End Sub

  ''' <summary>
  ''' Lets HomeSeer know our plug-in has triggers
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasTriggers() As Boolean
    Get
      SetTriggers()
      Return IIf(triggers.Count > 0, True, False)
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerCount() As Integer
    SetTriggers()
    Return triggers.Count
  End Function

  ''' <summary>
  ''' Returns the subtrigger count
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer
    Get
      Dim trigger As trigger
      If ValidTrig(TriggerNumber) Then
        trigger = triggers(TriggerNumber - 1)
        If Not (trigger Is Nothing) Then
          Return 0
        Else
          Return 0
        End If
      Else
        Return 0
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String
    Get
      If Not ValidTrig(TriggerNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, triggers.Keys(TriggerNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the subtrigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <param name="SubTriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String
    Get
      Dim trigger As trigger
      If ValidSubTrig(TriggerNumber, SubTriggerNumber) Then
        Return ""
      Else
        Return ""
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is valid
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidTrig(ByVal TrigIn As Integer) As Boolean
    SetTriggers()
    If TrigIn > 0 AndAlso TrigIn <= triggers.Count Then
      Return True
    End If
    Return False
  End Function

  ''' <summary>
  ''' Determines if the trigger is a valid subtrigger
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <param name="SubTrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidSubTrig(ByVal TrigIn As Integer, ByVal SubTrigIn As Integer) As Boolean
    Return False
  End Function

  ''' <summary>
  ''' Tell HomeSeer which triggers have conditions
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean
    Get
      Select Case TriggerNumber
        Case 0
          Return True   ' Render trigger as IF / AND IF
        Case Else
          Return False  ' Render trigger as IF / OR IF
      End Select
    End Get
  End Property

  ''' <summary>
  ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  
  ''' Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Set(ByVal value As Boolean)

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      ' TriggerCondition(sKey) = value

    End Set
    Get

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Return False

    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is a condition
  ''' </summary>
  ''' <param name="sKey"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property TriggerCondition(sKey As String) As Boolean
    Get

      If conditions.ContainsKey(sKey) = True Then
        Return conditions(sKey)
      Else
        Return False
      End If

    End Get
    Set(value As Boolean)

      If conditions.ContainsKey(sKey) = False Then
        conditions.Add(sKey, value)
      Else
        conditions(sKey) = value
      End If

    End Set
  End Property

  ''' <summary>
  ''' Called when HomeSeer wants to check if a condition is true
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Return False
  End Function

#End Region

#Region "Trigger Interface"

  ''' <summary>
  ''' Builds the Trigger UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = TrigInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    Else 'new event, so clean out the trigger object
      trigger = New trigger
    End If

    Select Case TrigInfo.TANumber
      Case TiVoTriggers.EmailDeliveryStatus
        Dim triggerName As String = GetEnumName(TiVoTriggers.EmailDeliveryStatus)

        Dim ActionSelected As String = trigger.Item("DeliveryStatus")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "DeliveryStatus", UID, sUnique)

        Dim jqDSN As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqDSN.autoPostBack = True

        jqDSN.AddItem("(Select Delivery Status)", "", (ActionSelected = ""))
        Dim Actions As String() = {"Success", "Deferral", "Failure"}
        For Each strAction As String In Actions
          Dim strOptionValue As String = strAction
          Dim strOptionName As String = strOptionValue
          jqDSN.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqDSN.Build)

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Process changes to the trigger from the HomeSeer events page
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection,
                                       ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As String = TrigInfo.UID.ToString
    Dim TANumber As Integer = TrigInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = TrigInfo.DataIn
    Ret.TrigActInfo = TrigInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    ' DeSerializeObject
    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If
    trigger.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case TiVoTriggers.EmailDeliveryStatus
          Dim triggerName As String = GetEnumName(TiVoTriggers.EmailDeliveryStatus)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, triggerName & "DeliveryStatus_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("DeliveryStatus") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(trigger, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Trigger not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Trigger UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret

  End Function

  ''' <summary>
  ''' Determines if a trigger is properly configured
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Get
      Dim Configured As Boolean = True
      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Select Case TrigInfo.TANumber
        Case TiVoTriggers.EmailDeliveryStatus
          If trigger.Item("DeliveryStatus") = "" Then Configured = False

      End Select

      Return Configured
    End Get
  End Property

  ''' <summary>
  ''' Formats the trigger for display
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim stb As New StringBuilder

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Select Case TrigInfo.TANumber
      Case TiVoTriggers.EmailDeliveryStatus
        If trigger.uid <= 0 Then
          stb.Append("Trigger has not been properly configured.")
        Else
          Dim strTriggerName As String = GetEnumDescription(TiVoTriggers.EmailDeliveryStatus)
          Dim strDeliveryStatus As String = trigger.Item("DeliveryStatus")

          stb.AppendFormat("{0} is <font class='event_Txt_Option'>{1}</font>", strTriggerName, strDeliveryStatus)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Checks to see if trigger should fire
  ''' </summary>
  ''' <param name="Plug_Name"></param>
  ''' <param name="TrigID"></param>
  ''' <param name="SubTrig"></param>
  ''' <param name="strTrigger"></param>
  ''' <remarks></remarks>
  Private Sub CheckTrigger(Plug_Name As String, TrigID As Integer, SubTrig As Integer, strTrigger As String)

    Try
      '
      ' Check HomeSeer Triggers
      '
      If Plug_Name.Contains(":") = False Then Plug_Name &= ":"
      Dim TrigsToCheck() As IAllRemoteAPI.strTrigActInfo = callback.TriggerMatches(Plug_Name, TrigID, SubTrig)

      Try

        For Each TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo In TrigsToCheck
          Dim UID As String = TrigInfo.UID.ToString

          If Not (TrigInfo.DataIn Is Nothing) Then

            Dim trigger As New trigger
            DeSerializeObject(TrigInfo.DataIn, trigger)

            Select Case TrigID

              Case TiVoTriggers.EmailDeliveryStatus
                Dim strTriggerName As String = GetEnumDescription(TiVoTriggers.EmailDeliveryStatus)
                Dim strDeliveryStatus As String = trigger.Item("DeliveryStatus")

                Dim strTriggerCheck As String = String.Format("{0},{1}", strTriggerName, strDeliveryStatus)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

            End Select

          End If

        Next

      Catch pEx As Exception

      End Try

    Catch pEx As Exception

    End Try

  End Sub

#End Region

#Region "Action Properties"

  ''' <summary>
  ''' Defines the valid actions for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetActions()
    Dim o As Object = Nothing
    If actions.Count = 0 Then
      actions.Add(o, "Set Channel")           ' 1
    End If
  End Sub

  ''' <summary>
  ''' Returns the action count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ActionCount() As Integer
    SetActions()
    Return actions.Count
  End Function

  ''' <summary>
  ''' Returns the action name
  ''' </summary>
  ''' <param name="ActionNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String
    Get
      If Not ValidAction(ActionNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, actions.Keys(ActionNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if an action is valid
  ''' </summary>
  ''' <param name="ActionIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidAction(ByVal ActionIn As Integer) As Boolean
    SetActions()
    If ActionIn > 0 AndAlso ActionIn <= actions.Count Then
      Return True
    End If
    Return False
  End Function

#End Region

#Region "Action Interface"

  ''' <summary>
  ''' Builds the Action UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks>This function is called from the HomeSeer event page when an event is in edit mode.</remarks>
  Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim stb As New StringBuilder

    Try
      Dim UID As String = ActInfo.UID.ToString

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      End If

      stb.AppendLine("<table cellspacing='0'>")

      Select Case ActInfo.TANumber
        Case TiVoActions.SetChannel
          Dim actionName As String = GetEnumName(TiVoActions.SetChannel)

          '
          ' Start TiVo Device
          '
          Dim ActionSelected As String = action.Item("TiVoDevice")

          Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "TiVoDevice", UID, sUnique)

          Dim jqTiVoDevice As New clsJQuery.jqDropList(actionId, Pagename, True)

          jqTiVoDevice.AddItem("Select TiVo", "", False)
          For Each TiVo In TiVos
            Dim strOptionValue As String = TiVo.ConnectionUUID
            Dim strOptionName As String = String.Format("{0} {1}", TiVo.Machine, TiVo.Platform)
            jqTiVoDevice.AddItem(strOptionName, strOptionValue, (ActionSelected.Contains(strOptionValue)))
          Next

          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class=""event_Txt_Selection"">TiVo Device:</td>")
          stb.AppendFormat("<td>{0}</td>", jqTiVoDevice.Build)
          stb.AppendLine(" </tr>")

          '
          ' Start TiVo Channel
          '
          ActionSelected = action.Item("TiVoChannel")

          actionId = String.Format("{0}{1}_{2}_{3}", actionName, "TiVoChannel", UID, sUnique)

          Dim tbTiVoChannel As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 5, False)
          tbTiVoChannel.submitForm = True

          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class=""event_Txt_Selection"">TiVo Channel:</td>")
          stb.AppendFormat("<td class=""event_Txt_Option"">{0}</td>", tbTiVoChannel.Build)
          stb.AppendLine(" </tr>")

      End Select

      stb.AppendLine("</table>")

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "ActionBuildUI()")
    End Try

    Return stb.ToString

  End Function

  ''' <summary>
  ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection,
                                      ByVal ActInfo As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As Integer = ActInfo.UID
    Dim TANumber As Integer = ActInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = ActInfo.DataIn
    Ret.TrigActInfo = ActInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    '
    ' DeSerializeObject
    '
    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If
    action.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case TiVoActions.SetChannel
          Dim actionName As String = GetEnumName(TiVoActions.SetChannel)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "TiVoDevice_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("TiVoDevice") = ActionValue

              Case InStr(sKey, actionName & "TiVoChannel_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("TiVoChannel") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(action, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Action not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Action UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret
  End Function

  ''' <summary>
  ''' Determines if our action is proplery configured
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return TRUE if the given action is configured properly</returns>
  ''' <remarks>There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.</remarks>
  Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim Configured As Boolean = True
    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case TiVoActions.SetChannel
        If action.Item("TiVoDevice") = "" Then Configured = False
        If action.Item("TiVoChannel") = "" Then Configured = False

    End Select

    Return Configured

  End Function

  ''' <summary>
  ''' After the action has been configured, this function is called in your plugin to display the configured action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return text that describes the given action.</returns>
  ''' <remarks></remarks>
  Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String
    Dim stb As New StringBuilder

    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case TiVoActions.SetChannel
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(TiVoActions.SetChannel)

          Dim strTiVoMachine As String = "TiVo"
          Dim strTiVoDevice As String = action.Item("TiVoDevice")
          Dim strTiVoChannel As String = action.Item("TiVoChannel")

          Dim TiVo As TiVo = TiVos.Find(Function(t) t.ConnectionUUID = strTiVoDevice)
          If Not TiVo Is Nothing Then
            strTiVoMachine = TiVo.Machine
          End If

          stb.AppendLine("<table cellspacing='0'>")

          stb.AppendLine(" <tr>")
          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class=""event_if_then_text"" colspan=""2"">")
          stb.AppendFormat("{0}: {1}", IFACE_NAME, strActionName)
          stb.AppendLine("  </td>")
          stb.AppendLine(" </tr>")

          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class=""event_Txt_Selection"">TiVo Device:</td>")
          stb.AppendFormat("<td>{0}</td>", strTiVoMachine)
          stb.AppendLine(" </tr>")

          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class=""event_Txt_Selection"">TiVo Channel:</td>")
          stb.AppendFormat("<td>{0}</td>", strTiVoChannel)
          stb.AppendLine(" </tr>")

          stb.AppendLine("</table>")

        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Handles the HomeSeer Event Action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = ActInfo.UID.ToString

    Try

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      Else
        Return False
      End If

      Select Case ActInfo.TANumber
        Case TiVoActions.SetChannel

          Dim strTiVoDevice As String = action.Item("TiVoDevice")
          Dim strTiVoChannel As String = action.Item("TiVoChannel")

          Dim TiVo As TiVo = TiVos.Find(Function(t) t.ConnectionUUID = strTiVoDevice)
          If Not TiVo Is Nothing Then
            TiVo.SetChannel(strTiVoChannel)
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      hs.WriteLog(IFACE_NAME, "Error executing action: " & pEx.Message)
    End Try

    Return True

  End Function

#End Region

#End Region

End Module

Public Enum TiVoActions
  <Description("Set Channel")>
  SetChannel = 1
End Enum

Public Enum TiVoTriggers
  <Description("Email Delivery Status")>
  EmailDeliveryStatus = 1
End Enum
