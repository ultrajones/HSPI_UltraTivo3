Imports System.Text
Imports System.Web
Imports Scheduler
Imports HomeSeerAPI
Imports System.Collections.Specialized
Imports System.Web.UI.WebControls

Public Class hspi_webconfig
  Inherits clsPageBuilder

  Public hspiref As HSPI

  Dim TimerEnabled As Boolean

  ''' <summary>
  ''' Initializes new webconfig
  ''' </summary>
  ''' <param name="pagename"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal pagename As String)
    MyBase.New(pagename)
  End Sub

#Region "Page Building"

  ''' <summary>
  ''' Web pages that use the clsPageBuilder class and registered with hs.RegisterLink and hs.RegisterConfigLink will then be called through this function. 
  ''' A complete page needs to be created and returned.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="queryString"></param>
  ''' <param name="instance"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, instance As String) As String

    Try

      Dim stb As New StringBuilder

      '
      ' Called from the start of your page to reset all internal data structures in the clsPageBuilder class, such as menus.
      '
      Me.reset()

      '
      ' Determine if user is authorized to access the web page
      '
      Dim LoggedInUser As String = hs.WEBLoggedInUser()
      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      '
      ' Handle any queries like mode=something
      '
      Dim parts As Collections.Specialized.NameValueCollection = Nothing
      If (queryString <> "") Then
        parts = HttpUtility.ParseQueryString(queryString)
      End If

      Dim Header As New StringBuilder
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrativo3/css/jquery.dataTables.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrativo3/css/editor.dataTables.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrativo3/css/buttons.dataTables.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrativo3/css/select.dataTables.min.css"" rel=""stylesheet"" />")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrativo3/js/jquery.dataTables.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrativo3/js/dataTables.editor.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrativo3/js/dataTables.buttons.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrativo3/js/dataTables.select.min.js""></script>")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrativo3/js/hspi_ultrativo3_devices.js""></script>")
      Me.AddHeader(Header.ToString)

      Dim pageTile As String = String.Format("{0} {1}", pageName, instance).TrimEnd
      stb.Append(hs.GetPageHeader(pageName, pageTile, "", "", False, False))

      '
      ' Start the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

      '
      ' A message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
      '
      stb.Append(clsPageBuilder.DivStart("divErrorMessage", "class='errormessage'"))
      stb.Append(clsPageBuilder.DivEnd)

      Me.RefreshIntervalMilliSeconds = 3000
      stb.Append(Me.AddAjaxHandlerPost("id=timer", pageName))

      If WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED) = False Then
        '
        ' Current user not authorized
        '
        stb.Append(WebUserNotUnauthorized(LoggedInUser))
      Else
        '
        ' Specific page starts here
        '
        stb.Append(BuildContent)
      End If

      '
      ' End the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivEnd)

      '
      ' Add the body html to the page
      '
      Me.AddBody(stb.ToString)

      '
      ' Return the full page
      '
      Return Me.BuildPage()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "GetPagePlugin")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the HTML content
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildContent() As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table border='0' cellpadding='0' cellspacing='0' width='1000'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td width='1000' align='center' style='color:#FF0000; font-size:14pt; height:30px;'><strong><div id='divMessage'>&nbsp;</div></strong></td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", BuildTabs())
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildContent")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the jQuery Tabss
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BuildTabs() As String

    Try

      Dim stb As New StringBuilder
      Dim tabs As clsJQuery.jqTabs = New clsJQuery.jqTabs("oTabs", Me.PageName)
      Dim tab As New clsJQuery.Tab

      tabs.postOnTabClick = True

      tab.tabTitle = "Status"
      tab.tabDIVID = "tabStatus"
      tab.tabContent = "<div id='divStatus'>" & BuildTabStatus() & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Options"
      tab.tabDIVID = "tabOptions"
      tab.tabContent = "<div id='divOptions'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "TiVo Devices"
      tab.tabDIVID = "tabTiVoDevices"
      tab.tabContent = "<div id='divTiVoDevices'></div>"
      tabs.tabs.Add(tab)

      Return tabs.Build

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the TiVo Devices Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabTiVoDevices(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmTiVoDevices", "frmTiVoDevices", "Post"))

      stb.AppendLine("<table width='100%' class='display cell-border' id='table_devices' cellspacing='0'>")

      '
      ' HA7Net Configuration
      '
      stb.AppendLine(" <thead>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>TiVo Name</th>")         ' machine=Media Room
      stb.AppendLine("   <th>TiVo Identity</th>")     ' identity=74600119071F67D
      stb.AppendLine("   <th>TiVo Make</th>")     ' platform=tcd / Series4
      stb.AppendLine("   <th>TiVo Platform</th>")     ' platform=tcd / Series4
      stb.AppendLine("   <th>Connection Type</th>")   ' method=broadcast
      stb.AppendLine("   <th>Connection Address</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </thead>")

      stb.AppendLine(" <tbody>")
      Dim MyDataTable As DataTable = hspi_plugin.GetTiVoDevices()
      For Each row As DataRow In MyDataTable.Rows
        stb.AppendFormat("  <tr id='{0}'>{1}", row("device_id"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_name"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_uuid"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_make"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_model"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_conn"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_addr"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", "", vbCrLf)
        stb.AppendLine("  </tr>")
      Next
      stb.AppendLine(" </tbody>")

      stb.AppendLine("</table")

      Dim strInfo As String = "If the plug-in is unable to connect to your TiVo, please make sure you have enabled networked remote control in the TiVo DVR settings."
      Dim strHint As String = "The plug-in supports TiVo auto discovery.  If your TiVo does not appear in this list, then you may need to manually add it."
      Dim strWarn As String = "When you modify a TiVo Device, a restart of the plug-in may be required."

      stb.AppendLine(" <div>&nbsp;</div>")
      stb.AppendLine(" <p>")
      stb.AppendFormat("<img alt='Info' src='/images/hspi_ultrativo3/ico_info.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strInfo)
      stb.AppendFormat("<img alt='Hint' src='/images/hspi_ultrativo3/ico_hint.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strHint)
      stb.AppendFormat("<img alt='Hint' src='/images/hspi_ultrativo3/ico_warn.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strWarn)
      stb.AppendLine(" </p>")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divTiVoDevices", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabTiVoDevices")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Status Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabStatus(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine(clsPageBuilder.FormStart("frmStatus", "frmStatus", "Post"))

      stb.AppendLine("<div>")
      stb.AppendLine("<table>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Plug-In Status </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Name:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", IFACE_NAME)
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Status:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", "OK")
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Version:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", HSPI.Version)
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table>")
      stb.AppendLine("</div>")

      stb.AppendLine(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divStatus", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabStatus")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabOptions(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      stb.Append(clsPageBuilder.FormStart("frmOptions", "frmOptions", "Post"))

      '
      ' Web Page Access
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Web Page Access</td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Authorized User Roles</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", BuildWebPageAccessCheckBoxes, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Application Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Application Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' Application Logging Level
      '
      Dim selLogLevel As New clsJQuery.jqDropList("selLogLevel", Me.PageName, False)
      selLogLevel.id = "selLogLevel"
      selLogLevel.toolTip = "Specifies the plug-in logging level."

      Dim itemValues As Array = System.Enum.GetValues(GetType(LogLevel))
      Dim itemNames As Array = System.Enum.GetNames(GetType(LogLevel))

      For i As Integer = 0 To itemNames.Length - 1
        Dim itemSelected As Boolean = IIf(gLogLevel = itemValues(i), True, False)
        selLogLevel.AddItem(itemNames(i), itemValues(i), itemSelected)
      Next
      selLogLevel.autoPostBack = True

      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", "Logging&nbsp;Level", vbCrLf)
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selLogLevel.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divOptions", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabOptions")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Web Page Access Checkbox List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildWebPageAccessCheckBoxes()

    Try

      Dim stb As New StringBuilder

      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      Dim cb1 As New clsJQuery.jqCheckBox("chkWebPageAccess_Guest", "Guest", Me.PageName, True, True)
      Dim cb2 As New clsJQuery.jqCheckBox("chkWebPageAccess_Admin", "Admin", Me.PageName, True, True)
      Dim cb3 As New clsJQuery.jqCheckBox("chkWebPageAccess_Normal", "Normal", Me.PageName, True, True)
      Dim cb4 As New clsJQuery.jqCheckBox("chkWebPageAccess_Local", "Local", Me.PageName, True, True)

      cb1.id = "WebPageAccess_Guest"
      cb1.checked = CBool(USER_ROLES_AUTHORIZED And USER_GUEST)

      cb2.id = "WebPageAccess_Admin"
      cb2.checked = CBool(USER_ROLES_AUTHORIZED And USER_ADMIN)
      cb2.enabled = False

      cb3.id = "WebPageAccess_Normal"
      cb3.checked = CBool(USER_ROLES_AUTHORIZED And USER_NORMAL)

      cb4.id = "WebPageAccess_Local"
      cb4.checked = CBool(USER_ROLES_AUTHORIZED And USER_LOCAL)

      stb.Append(clsPageBuilder.FormStart("frmWebPageAccess", "frmWebPageAccess", "Post"))

      stb.Append(cb1.Build())
      stb.Append(cb2.Build())
      stb.Append(cb3.Build())
      stb.Append(cb4.Build())

      stb.Append(clsPageBuilder.FormEnd())

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildWebPageAccessCheckBoxes")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

#Region "Page Processing"

  ''' <summary>
  ''' Post a message to this web page
  ''' </summary>
  ''' <param name="sMessage"></param>
  ''' <remarks></remarks>
  Sub PostMessage(ByVal sMessage As String)

    Try

      Me.divToUpdate.Add("divMessage", sMessage)

      Me.pageCommands.Add("starttimer", "")

      TimerEnabled = True

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' When a user clicks on any controls on one of your web pages, this function is then called with the post data. You can then parse the data and process as needed.
  ''' </summary>
  ''' <param name="page">The name of the page as registered with hs.RegisterLink or hs.RegisterConfigLink</param>
  ''' <param name="data">The post data</param>
  ''' <param name="user">The name of logged in user</param>
  ''' <param name="userRights">The rights of logged in user</param>
  ''' <returns>Any serialized data that needs to be passed back to the web page, generated by the clsPageBuilder class</returns>
  ''' <remarks></remarks>
  Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String

    Try

      WriteMessage("Entered postBackProc() function.", MessageType.Debug)

      Dim postData As NameValueCollection = HttpUtility.ParseQueryString(data)

      '
      ' Write debug to console
      '
      If gLogLevel >= MessageType.Debug Then
        For Each keyName As String In postData.AllKeys
          Console.WriteLine(String.Format("{0}={1}", keyName, postData(keyName)))
        Next
      End If

      Select Case postData("editor_action")
        Case "device-edit"
          Dim fields() As String = {"device_name", "device_uuid", "device_make", "device_model", "device_conn", "device_addr"}

          For Each field As String In fields
            Dim key As String = String.Format("data[{0}]", field)
            If postData.AllKeys.Contains(key) Then
              If postData(key).Trim.Length = 0 Then
                Return DatatableFieldError(field, "This is a required field.")
              End If
            Else
              Return DatatableError("Unable to modify TiVo Device due to an unexpected error.")
            End If
          Next

          Dim device_id As Integer = Integer.Parse(postData("id"))
          Dim device_name As String = postData("data[device_name]").Trim
          Dim device_uuid As String = postData("data[device_uuid]").Trim
          Dim device_make As String = postData("data[device_make]").Trim
          Dim device_model As String = postData("data[device_model]").Trim
          Dim device_conn As String = postData("data[device_conn]").Trim
          Dim device_addr As String = postData("data[device_addr]").Trim

          '
          ' Update TiVo Device
          '
          Dim bSuccess As Boolean = UpdateTiVoDevice(device_id, device_name, device_uuid, device_make, device_model, device_conn, device_addr)
          If bSuccess = False Then
            Return DatatableError("Unable to modify TiVo Device due to an unexpected error.")
          Else
            BuildTabTiVoDevices(True)
            Me.pageCommands.Add("executefunction", "initTiVoDevices()")

            Return DatatableRowDevice(device_id, device_name, device_uuid, device_make, device_model, device_conn, device_addr)
          End If

        Case "device-create"

          Dim fields() As String = {"device_name", "device_uuid", "device_make", "device_model", "device_conn", "device_addr"}

          For Each field As String In fields
            Dim key As String = String.Format("data[{0}]", field)
            If postData.AllKeys.Contains(key) Then
              If postData(key).Trim.Length = 0 Then
                Return DatatableFieldError(field, "This is a required field.")
              End If
            Else
              Return DatatableError("Unable to modify TiVo Device due to an unexpected error.")
            End If
          Next

          Dim device_name As String = postData("data[device_name]").Trim
          Dim device_uuid As String = postData("data[device_uuid]").Trim
          Dim device_make As String = postData("data[device_make]").Trim
          Dim device_model As String = postData("data[device_model]").Trim
          Dim device_conn As String = postData("data[device_conn]").Trim
          Dim device_addr As String = postData("data[device_addr]").Trim

          '
          ' Insert the TiVo Devices into the database
          '
          Dim device_id As Integer = InsertTiVoDevice(device_uuid, device_name, "", device_make, device_model, device_conn, device_addr, "")
          If device_id = False Then

            Return DatatableError("Unable to modify TiVo Device due to an unexpected error.")

          Else
            BuildTabTiVoDevices(True)
            Me.pageCommands.Add("executefunction", "initTiVoDevices()")

            Return DatatableRowDevice(device_id, device_name, device_uuid, device_make, device_model, device_conn, device_addr)
          End If

        Case "device-remove"
          Dim device_id As String = Val(postData("id[]"))

          Dim bSuccess As Boolean = DeleteTiVoDevice(device_id)
          If bSuccess = False Then
            Return DatatableError("Unable to delete TiVo Device due to an unexpected error.")
          Else
            BuildTabTiVoDevices(True)
            Me.pageCommands.Add("executefunction", "initTiVoDevices()")

            Return "{ }"
          End If

      End Select

      '
      ' Process the post data
      '
      Select Case postData("id")
        Case "tabStatus"
          BuildTabStatus(True)

        Case "tabOptions"
          BuildTabOptions(True)

        Case "tabTiVoDevices"
          BuildTabTiVoDevices(True)
          Me.pageCommands.Add("executefunction", "initTiVoDevices()")

        Case "selLogLevel"
          gLogLevel = Int32.Parse(postData("selLogLevel"))
          hs.SaveINISetting("Options", "LogLevel", gLogLevel.ToString, gINIFile)

          PostMessage("The application logging level has been updated.")

        Case "WebPageAccess_Guest"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Guest") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_GUEST
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_GUEST
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Normal"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Normal") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_NORMAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_NORMAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Local"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Local") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_LOCAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_LOCAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "timer" ' This stops the timer and clears the message
          If TimerEnabled Then 'this handles the initial timer post that occurs immediately upon enabling the timer.
            TimerEnabled = False
          Else
            Me.pageCommands.Add("stoptimer", "")
            Me.divToUpdate.Add("divMessage", "&nbsp;")
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "postBackProc")
    End Try

    Return MyBase.postBackProc(page, data, user, userRights)

  End Function

  ''' <summary>
  ''' Returns the Datatable Row JSON
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="device_name"></param>
  ''' <param name="device_uuid"></param>
  ''' <param name="device_make"></param>
  ''' <param name="device_model"></param>
  ''' <param name="device_conn"></param>
  ''' <param name="device_addr"></param>
  ''' <returns></returns>
  Private Function DatatableRowDevice(ByVal device_id As String,
                                      ByVal device_name As String,
                                      ByVal device_uuid As String,
                                      ByVal device_make As String,
                                      ByVal device_model As String,
                                      ByVal device_conn As String,
                                      ByVal device_addr As String) As String

    Try

      Dim sb As New StringBuilder

      sb.AppendLine("{")
      sb.AppendLine(" ""row"": { ")

      sb.AppendFormat(" ""{0}"": {1}, ", "DT_RowId", device_id)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_name", device_name)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_uuid", device_uuid)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_make", device_make)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_model", device_model)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_conn", device_conn)
      sb.AppendFormat(" ""{0}"": ""{1}"" ", "device_addr", device_addr)

      sb.AppendLine(" }")
      sb.AppendLine("}")

      Return sb.ToString

    Catch pEx As Exception
      Return "{ }"
    End Try

  End Function

  ''' <summary>
  ''' Returns the Datatable Error JSON
  ''' </summary>
  ''' <param name="errorString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DatatableError(ByVal errorString As String) As String

    Try
      Return String.Format("{{ ""error"": ""{0}"" }}", errorString)
    Catch pEx As Exception
      Return String.Format("{{ ""error"": ""{0}"" }}", pEx.Message)
    End Try

  End Function

  ''' <summary>
  ''' Returns the Datatable Field Error JSON
  ''' </summary>
  ''' <param name="fieldName"></param>
  ''' <param name="fieldError"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DatatableFieldError(fieldName As String, fieldError As String) As String

    Try
      Return String.Format("{{ ""fieldErrors"": [ {{""name"": ""{0}"",""status"": ""{1}""}} ] }}", fieldName, fieldError)
    Catch pEx As Exception
      Return String.Format("{{ ""fieldErrors"": [ {{""name"": ""{0}"",""status"": ""{1}""}} ] }}", fieldName, pEx.Message)
    End Try

  End Function

#End Region

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the HTML Not Authorized web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function WebUserNotUnauthorized(LoggedInUser As String) As String

    Try

      Dim sb As New StringBuilder

      sb.AppendLine("<table border='0' cellpadding='2' cellspacing='2' width='575px'>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td nowrap>")
      sb.AppendLine("     <h4>The Web Page You Were Trying To Access Is Restricted To Authorized Users ONLY</h4>")
      sb.AppendLine("   </td>")
      sb.AppendLine("  </tr>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td>")
      sb.AppendLine("     <p>This page is displayed if the credentials passed to the web server do not match the ")
      sb.AppendLine("      credentials required to access this web page.</p>")
      sb.AppendFormat("     <p>If you know the <b>{0}</b> user should have access,", LoggedInUser)
      sb.AppendFormat("      then ask your <b>HomeSeer Administrator</b> to check the <b>{0}</b> plug-in options", IFACE_NAME)
      sb.AppendFormat("      page to make sure the roles assigned to the <b>{0}</b> user allow access to this", LoggedInUser)
      sb.AppendLine("        web page.</p>")
      sb.AppendLine("  </td>")
      sb.AppendLine(" </tr>")
      sb.AppendLine(" </table>")

      Return sb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "WebUserNotUnauthorized")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

End Class