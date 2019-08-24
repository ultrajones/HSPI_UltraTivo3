Imports System.Data
Imports System.Data.Common
Imports System.Threading
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Configuration
Imports System.Text
Imports System.IO

Module Database

  Public DBConnectionMain As SQLite.SQLiteConnection  ' Our main database connection
  Public DBConnectionTemp As SQLite.SQLiteConnection  ' Our temp database connection

  Public gDBInsertSuccess As ULong = 0            ' Tracks DB insert success
  Public gDBInsertFailure As ULong = 0            ' Tracks DB insert success

  Public bDBInitialized As Boolean = False        ' Indicates if database successfully initialized

  Public SyncLockMain As New Object
  Public SyncLockTemp As New Object

#Region "Database Initilization"

  ''' <summary>
  ''' Initializes the database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitializeMainDatabase() As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered InitializeMainDatabase() function.", MessageType.Debug)

    Try
      '
      ' Close database if it's open
      '
      If Not DBConnectionMain Is Nothing Then
        If CloseDBConn(DBConnectionMain) = False Then
          Throw New Exception("An existing database connection could not be closed.")
        End If
      End If

      '
      ' Create the database directory if it does not exist
      '
      Dim databaseDir As String = FixPath(String.Format("{0}\Data\{1}\", hs.GetAppPath, IFACE_NAME.ToLower))
      If Directory.Exists(databaseDir) = False Then
        Directory.CreateDirectory(databaseDir)
      End If

      '
      ' Determine the database filename
      '
      Dim strDataSource As String = FixPath(String.Format("{0}\Data\{1}\{1}.db3", hs.GetAppPath(), IFACE_NAME.ToLower))

      '
      ' Determine the database provider factory and connection string
      '
      Dim strDbProviderFactory As String = "System.Data.SQLite"
      Dim strConnectionString As String = String.Format("Data Source={0}; Version=3;", strDataSource)

      '
      ' Attempt to open the database connection
      '
      bSuccess = OpenDBConn(DBConnectionMain, strDbProviderFactory, strConnectionString)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      bSuccess = False
      Call ProcessError(pEx, "InitializeDatabase()")
    End Try

    Return bSuccess

  End Function

  '------------------------------------------------------------------------------------
  ''' <summary>
  ''' Opens a connection to the database
  ''' </summary>
  ''' <param name="DBConnectionMain"></param>
  ''' <param name="strDbProviderFactory"></param>
  ''' <param name="strConnectionString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function OpenDBConn(ByRef DBConnectionMain As SQLite.SQLiteConnection, _
                              ByVal strDbProviderFactory As String, _
                              ByVal strConnectionString As String) As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered OpenDBConn() function.", MessageType.Debug)

    Try
      '
      ' Open database connection
      '
      DBConnectionMain = New SQLite.SQLiteConnection()
      DBConnectionMain.ConnectionString = strConnectionString
      DBConnectionMain.Open()

      '
      ' Run database vacuum
      '
      WriteMessage("Running SQLite database vacuum.", MessageType.Debug)
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = "VACUUM"
        MyDbCommand.ExecuteNonQuery()

        MyDbCommand.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "OpenDBConn()")
    End Try

    '
    ' Determine database connection status
    '
    bSuccess = DBConnectionMain.State = ConnectionState.Open

    '
    ' Record connection state to HomeSeer log
    '
    If bSuccess = True Then
      strMessage = "Database initialization complete."
      Call WriteMessage(strMessage, MessageType.Debug)
    Else
      strMessage = "Database initialization failed using [" & strConnectionString & "]."
      Call WriteMessage(strMessage, MessageType.Debug)
    End If

    Return bSuccess

  End Function

  ''' <summary>
  ''' Closes database connection
  ''' </summary>
  ''' <param name="DBConnectionMain"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CloseDBConn(ByRef DBConnectionMain As SQLite.SQLiteConnection) As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered CloseDBConn() function.", MessageType.Debug)

    Try
      '
      ' Attempt to the database
      '
      If DBConnectionMain.State <> ConnectionState.Closed Then
        DBConnectionMain.Close()
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CloseDBConn()")
    End Try

    '
    ' Determine database connection status
    '
    bSuccess = DBConnectionMain.State = ConnectionState.Closed

    '
    ' Record connection state to HomeSeer log
    '
    If bSuccess = True Then
      strMessage = "Database connection closed successfuly."
      Call WriteMessage(strMessage, MessageType.Debug)
    Else
      strMessage = "Unable to close database; Try restarting HomeSeer."
      Call WriteMessage(strMessage, MessageType.Debug)
    End If

    Return bSuccess

  End Function

  ''' <summary>
  ''' Checks to ensure a table exists
  ''' </summary>
  ''' <param name="strTableName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CheckDatabaseTable(ByVal strTableName As String) As Boolean

    Dim strMessage As String = ""
    Dim bSuccess As Boolean = False

    Try
      '
      ' Build SQL delete statement
      '
      If Regex.IsMatch(strTableName, "tblTiVoDevices") = True Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0) !TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          '
          ' Create the table
          '
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Debug)

          Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

            Dim sqlQueue As New Queue

            ' tivoconnect=1                   ' Used to serve as an identifying "signature
            ' swversion=20.3.8-01-2-746       ' This value describes the "primary" software running on the TCM
            ' method=broadcast                ' broadcast (for packets sent using UDP) 
            ' identity=74600119071F67D        ' This value should be unique to the originating TCM 
            ' machine=Media Room              ' This value contains human readable text, naming the TCM, suitable for display to the user. 
            ' platform=tcd / Series4          ' tcd (for TiVo DVR beacons) 
            ' services=TiVoMediaServer:80/http

            sqlQueue.Enqueue("CREATE TABLE tblTiVoDevices(" _
                            & "device_id INTEGER PRIMARY KEY," _
                            & "device_uuid varchar(255) NOT NULL," _
                            & "device_name varchar(255) NOT NULL," _
                            & "device_swversion varchar(255) NOT NULL," _
                            & "device_make varchar(255) NOT NULL," _
                            & "device_model varchar(255) NOT NULL," _
                            & "device_conn varchar(255) NOT NULL," _
                            & "device_addr varchar(12) NOT NULL," _
                            & "device_services varchar(255)" _
                          & ")")

            sqlQueue.Enqueue("CREATE UNIQUE INDEX IF NOT EXISTS idxUUID ON tblTiVoDevices(device_uuid)")
            sqlQueue.Enqueue("CREATE UNIQUE INDEX IF NOT EXISTS idxNAME ON tblTiVoDevices(device_name)")
            sqlQueue.Enqueue("CREATE UNIQUE INDEX IF NOT EXISTS idxADDR ON tblTiVoDevices(device_addr)")

            While sqlQueue.Count > 0
              Dim strSQL As String = sqlQueue.Dequeue

              dbcmd.Connection = DBConnectionMain
              dbcmd.CommandType = CommandType.Text
              dbcmd.CommandText = strSQL

              Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
              If iRecordsAffected <> 1 Then
                'Throw New Exception("Database schemea update failed due to error.")
              End If

            End While

            dbcmd.Dispose()
          End Using

        End If

      Else
        Throw New Exception(strTableName & " not currently supported.")
      End If

      bSuccess = True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CheckDatabaseTables()")
    End Try

    Return bSuccess

  End Function

#End Region

#Region "Database Date Formatting"

  ''' <summary>
  ''' dateTime as DateTime
  ''' </summary>
  ''' <param name="dateTime"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ConvertDateTimeToEpoch(ByVal dateTime As DateTime) As Long

    Dim baseTicks As Long = 621355968000000000
    Dim tickResolution As Long = 10000000

    Return (dateTime.ToUniversalTime.Ticks - baseTicks) / tickResolution

  End Function

  ''' <summary>
  ''' Converts Epoch to datetime
  ''' </summary>
  ''' <param name="epochTicks"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ConvertEpochToDateTime(ByVal epochTicks As Long) As DateTime

    '
    ' Create a new DateTime value based on the Unix Epoch
    '
    Dim converted As New DateTime(1970, 1, 1, 0, 0, 0, 0)

    '
    ' Return the value in string format
    '
    Return converted.AddSeconds(epochTicks).ToLocalTime

  End Function

  ''' <summary>
  ''' Converts date to format recognized by all regions
  ''' </summary>
  ''' <param name="strDate"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function RegionalizeDate(ByVal strDate As String) As String

    Dim ci As CultureInfo = New CultureInfo(Thread.CurrentThread.CurrentUICulture.ToString())
    Dim TheDate As New DateTime

    Try
      '
      ' Try to parse the date provided
      '
      TheDate = Date.Parse(strDate)
    Catch pEx As Exception
      '
      ' Let's just return the current date
      '
      TheDate = Date.Parse(Date.Now)
    End Try

    Return TheDate.ToString("F", ci)

  End Function

#End Region

End Module

