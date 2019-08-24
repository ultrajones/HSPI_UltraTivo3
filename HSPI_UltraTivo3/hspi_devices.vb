Imports HomeSeerAPI
Imports Scheduler
Imports HSCF
Imports HSCF.Communication.ScsServices.Service
Imports System.Text.RegularExpressions

Module hspi_devices

  ''' <summary>
  ''' Create the HomeSeer Root Device
  ''' </summary>
  ''' <param name="strRootId"></param>
  ''' <param name="strRootName"></param>
  ''' <param name="dv_ref_child"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateRootDevice(ByVal strRootId As String, _
                                   ByVal strRootName As String, _
                                   ByVal dv_ref_child As Integer) As Integer

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DeviceShowValues As Boolean = False

    Try
      '
      ' Set the local variables
      '
      If strRootId = "Plugin" Then
        dv_name = "UltraTiVo3 Plugin"
        dv_addr = String.Format("{0}-Root", strRootName.Replace(" ", "-"))
        dv_type = dv_name
      Else
        dv_name = strRootName
        dv_addr = String.Format("{0}-Root", strRootId, strRootName.Replace(" ", "-"))
        dv_type = strRootName
      End If

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} root device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} root device.", dv_name), MessageType.Debug)

      End If

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = IIf(strRootId = "Plugin", "Plug-ins", dv_type)
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this a parent root device
      '
      dv.Relationship(hs) = Enums.eRelationship.Parent_Root
      dv.AssociatedDevice_Add(hs, dv_ref_child)

      Dim image As String = "device_root.png"

      Dim VSPair As VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = 0
      VSPair.Status = "Root"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      Dim VGPair As VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.SingleValue
      VGPair.Set_Value = 0
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, image)
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      '
      ' Update the Device Misc Bits
      '
      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

    Catch pEx As Exception

    End Try

    Return dv_ref

  End Function

  ''' <summary>
  ''' Function to create our plug-in connection device used for status
  ''' </summary>
  ''' <param name="tivo_uuid"></param>
  ''' <param name="tivo_name"></param>
  ''' <param name="tivo_connected"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UpdateTiVoConnectionDevice(ByVal tivo_uuid As String,
                                             ByVal tivo_name As String,
                                             ByVal tivo_connected As Boolean) As String

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Try
      '
      ' Set the local variables
      '
      dv_type = "TiVo Connection"
      dv_name = tivo_name
      dv_addr = String.Format("{0}-Conn", tivo_uuid.ToString)
      dv = LocateDeviceByAddr(dv_addr)

      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Storge the UUID for the device
      '
      Dim pdata As New clsPlugExtraData
      pdata.AddNamed("UUID", tivo_uuid)
      dv.PlugExtraData_Set(hs) = pdata

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = dv_type
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this device a child of the root
      '
      dv.AssociatedDevice_ClearAll(hs)
      Dim dvp_ref As Integer = CreateRootDevice(tivo_uuid, dv_name, dv_ref)
      If dvp_ref > 0 Then
        dv.AssociatedDevice_Add(hs, dvp_ref)
      End If
      dv.Relationship(hs) = Enums.eRelationship.Child

      '
      ' Update the last change date
      ' 
      dv.Last_Change(hs) = DateTime.Now

      Dim VSPair As VSPair

      VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = -3
      VSPair.Status = ""
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = -2
      VSPair.Status = "Disconnect"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = -1
      VSPair.Status = "Reconnect"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = 0
      VSPair.Status = "Disconnected"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = 1
      VSPair.Status = "Connected"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      Dim VGPair As VGPair

      '
      ' Add VGPairs
      '
      VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.Range
      VGPair.RangeStart = -3
      VGPair.RangeEnd = 0
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, "tivo_disconnected.png")
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.SingleValue
      VGPair.Set_Value = 1
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, "tivo_connected.png")
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      '
      ' Update the Device Misc Bits
      '
      dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

      '
      ' Update the connection status
      '
      Dim dv_value As Long = IIf(tivo_connected = True, 1, 0)
      hspi_devices.SetDeviceValue(dv_addr, dv_value)

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "UpdatePAVRConnectionDevice()")
      Return pEx.ToString
    End Try

  End Function

  ''' <summary>
  ''' Function to create our plug-in connection device used for status
  ''' </summary>
  ''' <param name="tivo_uuid"></param>
  ''' <param name="tivo_name"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateTiVoControlDevice(ByVal tivo_uuid As String,
                                          ByVal tivo_name As String) As String

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DevicePairs As New ArrayList

    Try
      '
      ' Set the local variables
      '
      dv_type = "TiVo Control"
      dv_name = tivo_name
      dv_addr = String.Format("{0}-Ctrl", tivo_uuid.ToString)
      dv = LocateDeviceByAddr(dv_addr)

      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Storge the UUID for the device
      '
      Dim pdata As New clsPlugExtraData
      pdata.AddNamed("UUID", tivo_uuid)
      dv.PlugExtraData_Set(hs) = pdata

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = dv_type
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this device a child of the root
      '
      dv.AssociatedDevice_ClearAll(hs)
      Dim dvp_ref As Integer = CreateRootDevice(tivo_uuid, dv_name, dv_ref)
      If dvp_ref > 0 Then
        dv.AssociatedDevice_Add(hs, dvp_ref)
      End If
      dv.Relationship(hs) = Enums.eRelationship.Child

      '
      ' Update the last change date
      ' 
      dv.Last_Change(hs) = DateTime.Now

      '
      ' Define the default image
      '
      Dim strIRImage As String = "tivo_control.png"

      DevicePairs.Clear()
      DevicePairs.Add(New hspi_device_pairs(0, "", strIRImage, HomeSeerAPI.ePairStatusControl.Both))

      For Each key In Buttons.Keys
        Dim value As String = Buttons(key)
        Select Case key
          Case 19 To 26, 30 To 33

          Case Else

            DevicePairs.Add(New hspi_device_pairs(key, value, strIRImage, HomeSeerAPI.ePairStatusControl.Control))

        End Select

      Next

      '
      ' Add the Status Graphic Pairs
      '
      For Each Pair As hspi_device_pairs In DevicePairs

        Dim VSPair As VSPair = New VSPair(Pair.Type)
        VSPair.PairType = VSVGPairType.SingleValue
        VSPair.Value = Pair.Value
        VSPair.Status = Pair.Status
        VSPair.Render = Enums.CAPIControlType.Values
        hs.DeviceVSP_AddPair(dv_ref, VSPair)

        Dim VGPair As VGPair = New VGPair()
        VGPair.PairType = VSVGPairType.SingleValue
        VGPair.Set_Value = Pair.Value
        VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
        hs.DeviceVGP_AddPair(dv_ref, VGPair)

      Next

      dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)

      '
      ' Update the Device Misc Bits
      '
      dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

      '
      ' Update the connection status
      '
      Dim dv_value As Long = 0
      hspi_devices.SetDeviceValue(dv_addr, dv_value)
      hspi_devices.SetDeviceString(dv_addr, "Remote")

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CreateTiVoControlDevice()")
      Return pEx.ToString
    End Try

  End Function

  ''' <summary>
  ''' Function to create our plug-in connection device used for status
  ''' </summary>
  ''' <param name="tivo_uuid"></param>
  ''' <param name="tivo_name"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateTiVoVideoDevice(ByVal tivo_uuid As String,
                                        ByVal tivo_name As String) As String

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DevicePairs As New ArrayList

    Try
      '
      ' Set the local variables
      '
      dv_type = "TiVo Video"
      dv_name = tivo_name
      dv_addr = String.Format("{0}-Video", tivo_uuid.ToString)
      dv = LocateDeviceByAddr(dv_addr)

      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Storge the UUID for the device
      '
      Dim pdata As New clsPlugExtraData
      pdata.AddNamed("UUID", tivo_uuid)
      dv.PlugExtraData_Set(hs) = pdata

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = dv_type
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this device a child of the root
      '
      dv.AssociatedDevice_ClearAll(hs)
      Dim dvp_ref As Integer = CreateRootDevice(tivo_uuid, dv_name, dv_ref)
      If dvp_ref > 0 Then
        dv.AssociatedDevice_Add(hs, dvp_ref)
      End If
      dv.Relationship(hs) = Enums.eRelationship.Child

      '
      ' Update the last change date
      ' 
      dv.Last_Change(hs) = DateTime.Now

      '
      ' Define the default image
      '
      Dim strIRImage As String = "tivo_video.png"

      DevicePairs.Clear()
      DevicePairs.Add(New hspi_device_pairs(0, "", strIRImage, HomeSeerAPI.ePairStatusControl.Both))

      For Each key In Buttons.Keys
        Dim value As String = Buttons(key)
        Select Case key
          Case 19 To 26, 30 To 33
            value = Regex.Replace(value, "Video_Mode", "VM")
            value = Regex.Replace(value, "Aspect_Correction", "AC")

            DevicePairs.Add(New hspi_device_pairs(key, value, strIRImage, HomeSeerAPI.ePairStatusControl.Control))
          Case Else

        End Select

      Next

      '
      ' Add the Status Graphic Pairs
      '
      For Each Pair As hspi_device_pairs In DevicePairs

        Dim VSPair As VSPair = New VSPair(Pair.Type)
        VSPair.PairType = VSVGPairType.SingleValue
        VSPair.Value = Pair.Value
        VSPair.Status = Pair.Status
        VSPair.Render = Enums.CAPIControlType.Values
        hs.DeviceVSP_AddPair(dv_ref, VSPair)

        Dim VGPair As VGPair = New VGPair()
        VGPair.PairType = VSVGPairType.SingleValue
        VGPair.Set_Value = Pair.Value
        VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
        hs.DeviceVGP_AddPair(dv_ref, VGPair)

      Next

      dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)

      '
      ' Update the Device Misc Bits
      '
      dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

      '
      ' Update the connection status
      '
      Dim dv_value As Long = 0
      hspi_devices.SetDeviceValue(dv_addr, dv_value)
      hspi_devices.SetDeviceString(dv_addr, "Video")

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CreateTiVoControlDevice()")
      Return pEx.ToString
    End Try

  End Function

  ''' <summary>
  ''' Function to create our plug-in connection device used for status
  ''' </summary>
  ''' <param name="tivo_uuid"></param>
  ''' <param name="tivo_name"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateTiVoChannelDevice(ByVal tivo_uuid As String,
                                          ByVal tivo_name As String) As String

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DevicePairs As New ArrayList

    Try
      '
      ' Set the local variables
      '
      dv_type = "TiVo Channel"
      dv_name = tivo_name
      dv_addr = String.Format("{0}-Chan", tivo_uuid.ToString)
      dv = LocateDeviceByAddr(dv_addr)

      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Storge the UUID for the device
      '
      Dim pdata As New clsPlugExtraData
      pdata.AddNamed("UUID", tivo_uuid)
      dv.PlugExtraData_Set(hs) = pdata

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = dv_type
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this device a child of the root
      '
      dv.AssociatedDevice_ClearAll(hs)
      Dim dvp_ref As Integer = CreateRootDevice(tivo_uuid, dv_name, dv_ref)
      If dvp_ref > 0 Then
        dv.AssociatedDevice_Add(hs, dvp_ref)
      End If
      dv.Relationship(hs) = Enums.eRelationship.Child

      '
      ' Update the last change date
      ' 
      dv.Last_Change(hs) = DateTime.Now

      '
      ' Define the default image
      '
      Dim strIRImage As String = "tivo_channel.png"

      '
      ' Add VGPairs
      '
      Dim Pair As VSPair

      Pair = New VSPair(ePairStatusControl.Control)
      Pair.PairType = VSVGPairType.SingleValue
      Pair.Render = Enums.CAPIControlType.TextBox_String
      hs.DeviceVSP_AddPair(dv_ref, Pair)

      '
      ' Add VGPairs
      '
      Dim VGPair As VGPair

      VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.Range
      VGPair.RangeStart = 0
      VGPair.RangeEnd = 1000
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, strIRImage)
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      '
      ' Update the Device Misc Bits
      '
      dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

      '
      ' Update the connection status
      '
      Dim dv_value As Long = 0
      hspi_devices.SetDeviceValue(dv_addr, dv_value)

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CreateTiVoChannelDevice()")
      Return pEx.ToString
    End Try

  End Function

  ''' <summary>
  ''' Locates device by device by ref
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByRef(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      Long.TryParse(strDeviceAddr, dev_ref)
      If dev_ref > 0 Then
        objDevice = hs.GetDeviceByRef(dev_ref)
        If Not objDevice Is Nothing Then
          Return objDevice
        End If
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByRef")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Locates device by device by code
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByCode(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.DeviceExistsCode(strDeviceAddr)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByCode")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Locates device by device address
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByAddr(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.DeviceExistsAddress(strDeviceAddr, False)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByAddr")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Sets the HomeSeer string and device values
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_value"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceValue(ByVal dv_addr As String, _
                            ByVal dv_value As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_value), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        If IsNumeric(dv_value) Then

          Dim dblDeviceValue As Double = Double.Parse(hs.DeviceValueEx(dv_ref))
          Dim dblSensorValue As Double = Double.Parse(dv_value)

          If dblDeviceValue <> dblSensorValue Then
            hs.SetDeviceValueByRef(dv_ref, dblSensorValue, True)
          End If

        End If

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceValue()")

    End Try

  End Sub

  ''' <summary>
  ''' Sets the HomeSeer device string
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_string"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceString(ByVal dv_addr As String, _
                             ByVal dv_string As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_string), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        hs.SetDeviceString(dv_ref, dv_string, True)

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceString()")

    End Try

  End Sub

End Module
