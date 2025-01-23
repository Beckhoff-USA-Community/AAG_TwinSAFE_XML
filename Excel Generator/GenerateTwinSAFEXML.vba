Sub GenerateTwinSAFEXML()
 
    ' Set reference to the worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Replace "Sheet1" with your sheet name

    ' Create a new XML document
    Set objXML = CreateObject("MSXML2.DOMDocument.6.0")
    objXML.async = False
    objXML.preserveWhiteSpace = True

    ' Create root element
    Set objRoot = objXML.createElement("TwinCATExport")
    objRoot.setAttribute "Version", "0.31"
    objXML.appendChild objRoot

    ' --- General Information ---
    Set objGeneralInfo = objXML.createElement("GeneralInformation")
    objRoot.appendChild objGeneralInfo

    ' Add child elements to GeneralInformation
    With objGeneralInfo
        .appendChild objXML.createElement("ProjectName")
        .LastChild.Text = "TwinSAFE-XML-Demo"

        .appendChild objXML.createElement("Author")
        .LastChild.Text = "J.Doe"

        .appendChild objXML.createElement("InternalProjectName")
        .LastChild.Text = "TwinSAFE-XML-Demo"
    End With

    ' --- Target System Configuration ---
    Set objTargetSys = objXML.createElement("TargetSystemConfiguration")
    objTargetSys.setAttribute "Id", "t1"
    objTargetSys.setAttribute "Type", "Hardware"
    objTargetSys.setAttribute "SubType", "EL6910"
    objTargetSys.setAttribute "IsExternal", "false"
    objRoot.appendChild objTargetSys

    ' Add child elements to TargetSystemConfiguration
    With objTargetSys
        .appendChild objXML.createElement("SafeAddress")
        .LastChild.Text = "2"

        .appendChild objXML.createElement("IoPath")
        .LastChild.Text = "TIID^Device 1 (EtherCAT)^Term 1 (EK1100)^Term 2 (EL6910)"

        .appendChild objXML.createElement("ProductCode")
        .LastChild.Text = "452866130"

        .appendChild objXML.createElement("RevisionNo")
        .LastChild.Text = "1376256"

        .appendChild objXML.createElement("VersionNumber")
        .LastChild.Text = "1"

        Set backupRestoreNode = objXML.createElement("BackupRestore")
        backupRestoreNode.setAttribute "Activated", "false"
        backupRestoreNode.setAttribute "RestoreUserAdministration", "false"
        backupRestoreNode.setAttribute "NumberOfDevicesWithMatchingCRC", "0"
        .appendChild backupRestoreNode

        .appendChild objXML.createElement("MapProjectCRC")
        .LastChild.Text = "false"

        .appendChild objXML.createElement("MapSerialNumber")
        .LastChild.Text = "false"
    End With

    ' --- Application Configuration ---
    Set objAppConfig = objXML.createElement("ApplicationConfiguration")
    objRoot.appendChild objAppConfig
    
    ' Create the TwinSAFEGroups element
    Set objTwinSAFEGroups = objXML.createElement("TwinSAFEGroups")
    objAppConfig.appendChild objTwinSAFEGroups

    ' Create the TwinSAFEGroup element
    Set objTwinSAFEGroup = objXML.createElement("TwinSAFEGroup")
    objTwinSAFEGroup.setAttribute "Id", "g1"
    objTwinSAFEGroup.setAttribute "OrderId", "0"
    objTwinSAFEGroups.appendChild objTwinSAFEGroup
    
    ' Create the Name element
    Set objTwinSAFEGroupName = objXML.createElement("Name")
    objTwinSAFEGroupName.Text = "TwinSafeGroup1"
    objTwinSAFEGroup.appendChild objTwinSAFEGroupName
    
    ' Create the TwinSAFEGroupOptions element
    Set objTwinSafeGroupOptions = objXML.createElement("TwinSAFEGroupOptions")
    objTwinSAFEGroup.appendChild objTwinSafeGroupOptions
    
    ' Add elements to TwinSAFEGroupOptions
    Set objMapDiag = objXML.createElement("MapDiag")
    objMapDiag.Text = "false"
    objTwinSafeGroupOptions.appendChild objMapDiag

    Set objMapState = objXML.createElement("MapState")
    objMapState.Text = "false"
    objTwinSafeGroupOptions.appendChild objMapState

    Set objPassificationAllowed = objXML.createElement("PassificationAllowed")
    objPassificationAllowed.Text = "false"
    objTwinSafeGroupOptions.appendChild objPassificationAllowed

    Set objTemporaryDeactivationAllowed = objXML.createElement("TemporaryDeactivationAllowed")
    objTemporaryDeactivationAllowed.Text = "false"
    objTwinSafeGroupOptions.appendChild objTemporaryDeactivationAllowed

    Set objPermanentDeactivationAllowed = objXML.createElement("PermanentDeactivationAllowed")
    objPermanentDeactivationAllowed.Text = "false"
    objTwinSafeGroupOptions.appendChild objPermanentDeactivationAllowed

    Set objTimeOutPassificationAllowed = objXML.createElement("TimeOutPassificationAllowed")
    objTimeOutPassificationAllowed.Text = "10000"
    objTwinSafeGroupOptions.appendChild objTimeOutPassificationAllowed

    Set objVerifyAnalogFBInputsByStartup = objXML.createElement("VerifyAnalogFBInputsByStartup")
    objVerifyAnalogFBInputsByStartup.Text = "false"
    objTwinSafeGroupOptions.appendChild objVerifyAnalogFBInputsByStartup

    Set objAnalogFBOutputCustomReplacementValues = objXML.createElement("AnalogFBOutputCustomReplacementValues")
    objAnalogFBOutputCustomReplacementValues.Text = "false"
    objTwinSafeGroupOptions.appendChild objAnalogFBOutputCustomReplacementValues
    
    ' Create the GroupOutputs element
    Set objGroupOutputs = objXML.createElement("GroupOutputs")
    objTwinSAFEGroup.appendChild objGroupOutputs
    
    ' Add Element elements to GroupOutputs
    ' Element 1
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_o1"
    objGroupOutputs.appendChild objElement

    Set objPortId = objXML.createElement("PortId")
    objPortId.Text = "33619968"
    objElement.appendChild objPortId

    Set objAlias = objXML.createElement("Alias")
    objAlias.Text = "FbErr"
    objElement.appendChild objAlias

    Set objActive = objXML.createElement("Active")
    objActive.Text = "true"
    objElement.appendChild objActive

    ' Element 2
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_o2"
    objGroupOutputs.appendChild objElement

    Set objPortId = objXML.createElement("PortId")
    objPortId.Text = "33685504"
    objElement.appendChild objPortId

    Set objAlias = objXML.createElement("Alias")
    objAlias.Text = "ComErr"
    objElement.appendChild objAlias

    Set objActive = objXML.createElement("Active")
    objActive.Text = "true"
    objElement.appendChild objActive

    ' Element 3
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_o3"
    objGroupOutputs.appendChild objElement

    Set objPortId = objXML.createElement("PortId")
    objPortId.Text = "33751040"
    objElement.appendChild objPortId

    Set objAlias = objXML.createElement("Alias")
    objAlias.Text = "OutErr"
    objElement.appendChild objAlias

    Set objActive = objXML.createElement("Active")
    objActive.Text = "true"
    objElement.appendChild objActive

    ' Element 4
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_o4"
    objGroupOutputs.appendChild objElement

    Set objPortId = objXML.createElement("PortId")
    objPortId.Text = "33816576"
    objElement.appendChild objPortId

    Set objAlias = objXML.createElement("Alias")
    objAlias.Text = "OtherErr"
    objElement.appendChild objAlias

    Set objActive = objXML.createElement("Active")
    objActive.Text = "true"
    objElement.appendChild objActive

    ' Element 5
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_o5"
    objGroupOutputs.appendChild objElement

    Set objPortId = objXML.createElement("PortId")
    objPortId.Text = "33882112"
    objElement.appendChild objPortId

    Set objAlias = objXML.createElement("Alias")
    objAlias.Text = "ComStartup"
    objElement.appendChild objAlias

    Set objActive = objXML.createElement("Active")
    objActive.Text = "true"
    objElement.appendChild objActive

    ' Element 6
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_o6"
    objGroupOutputs.appendChild objElement

    Set objPortId = objXML.createElement("PortId")
    objPortId.Text = "33947648"
    objElement.appendChild objPortId

    Set objAlias = objXML.createElement("Alias")
    objAlias.Text = "FbDeactive"
    objElement.appendChild objAlias

    Set objActive = objXML.createElement("Active")
    objActive.Text = "true"
    objElement.appendChild objActive

    ' Element 7
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_o7"
    objGroupOutputs.appendChild objElement

    Set objPortId = objXML.createElement("PortId")
    objPortId.Text = "34013184"
    objElement.appendChild objPortId

    Set objAlias = objXML.createElement("Alias")
    objAlias.Text = "FbRun"
    objElement.appendChild objAlias

    Set objActive = objXML.createElement("Active")
    objActive.Text = "true"
    objElement.appendChild objActive

    ' Element 8
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_o8"
    objGroupOutputs.appendChild objElement

    Set objPortId = objXML.createElement("PortId")
    objPortId.Text = "34078720"
    objElement.appendChild objPortId

    Set objAlias = objXML.createElement("Alias")
    objAlias.Text = "InRun"
    objElement.appendChild objAlias

    Set objActive = objXML.createElement("Active")
    objActive.Text = "true"
    objElement.appendChild objActive
    
    ' Create the GroupInputs element
    Set objGroupInputs = objXML.createElement("GroupInputs")
    objTwinSAFEGroup.appendChild objGroupInputs
    
    ' Add Element elements to GroupInputs
    ' Element 1
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_i1"
    objGroupInputs.appendChild objElement

    Set objPortId = objXML.createElement("PortId")
    objPortId.Text = "16842752"
    objElement.appendChild objPortId

    Set objAlias = objXML.createElement("Alias")
    objAlias.Text = "RunStop"
    objElement.appendChild objAlias

    Set objActive = objXML.createElement("Active")
    objActive.Text = "true"
    objElement.appendChild objActive

    ' Element 2
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_i2"
    objGroupInputs.appendChild objElement

    Set objPortId = objXML.createElement("PortId")
    objPortId.Text = "16908288"
    objElement.appendChild objPortId

    Set objAlias = objXML.createElement("Alias")
    objAlias.Text = "ErrAck"
    objElement.appendChild objAlias

    Set objActive = objXML.createElement("Active")
    objActive.Text = "true"
    objElement.appendChild objActive

    ' Element 3
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_i3"
    objGroupInputs.appendChild objElement

    Set objPortId = objXML.createElement("PortId")
    objPortId.Text = "16973824"
    objElement.appendChild objPortId

    Set objAlias = objXML.createElement("Alias")
    objAlias.Text = "ModuleFault"
    objElement.appendChild objAlias

    Set objActive = objXML.createElement("Active")
    objActive.Text = "true"
    objElement.appendChild objActive
    
    ' Create the AliasDevices element
    Set objAliasDevices = objXML.createElement("AliasDevices")
    objTwinSAFEGroup.appendChild objAliasDevices
    
    ' StandardAliasDevice 1
    Set objStandardAliasDevice = objXML.createElement("StandardAliasDevice")
    objStandardAliasDevice.setAttribute "Id", "g1_a1"
    objStandardAliasDevice.setAttribute "OrderId", "1"
    objAliasDevices.appendChild objStandardAliasDevice

    Set objAliasName = objXML.createElement("AliasName")
    objAliasName.Text = "ErrorAcknowledgement"
    objStandardAliasDevice.appendChild objAliasName

    Set objRxPdo = objXML.createElement("RxPdo")
    objRxPdo.setAttribute "BitSize", "1"
    objStandardAliasDevice.appendChild objRxPdo

    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_a1_i1"
    objRxPdo.appendChild objElement

    Set objName = objXML.createElement("Name")
    objName.Text = "In"
    objElement.appendChild objName

    Set objType = objXML.createElement("Type")
    objType.Text = "BIT"
    objElement.appendChild objType

    Set objBitOffset = objXML.createElement("BitOffset")
    objBitOffset.Text = "0"
    objElement.appendChild objBitOffset

    Set objBitSize = objXML.createElement("BitSize")
    objBitSize.Text = "1"
    objElement.appendChild objBitSize

    Set objIsSafeTimer = objXML.createElement("IsSafeTimer")
    objIsSafeTimer.Text = "false"
    objElement.appendChild objIsSafeTimer

    ' StandardAliasDevice 2
    Set objStandardAliasDevice = objXML.createElement("StandardAliasDevice")
    objStandardAliasDevice.setAttribute "Id", "g1_a2"
    objStandardAliasDevice.setAttribute "OrderId", "2"
    objAliasDevices.appendChild objStandardAliasDevice

    Set objAliasName = objXML.createElement("AliasName")
    objAliasName.Text = "Run"
    objStandardAliasDevice.appendChild objAliasName

    Set objRxPdo = objXML.createElement("RxPdo")
    objRxPdo.setAttribute "BitSize", "1"
    objStandardAliasDevice.appendChild objRxPdo

    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_a2_i1"
    objRxPdo.appendChild objElement

    Set objName = objXML.createElement("Name")
    objName.Text = "In"
    objElement.appendChild objName

    Set objType = objXML.createElement("Type")
    objType.Text = "BIT"
    objElement.appendChild objType

    Set objBitOffset = objXML.createElement("BitOffset")
    objBitOffset.Text = "0"
    objElement.appendChild objBitOffset

    Set objBitSize = objXML.createElement("BitSize")
    objBitSize.Text = "1"
    objElement.appendChild objBitSize

    Set objIsSafeTimer = objXML.createElement("IsSafeTimer")
    objIsSafeTimer.Text = "false"
    objElement.appendChild objIsSafeTimer

    ' SafetyAliasDevice 1
    Set objSafetyAliasDevice = objXML.createElement("SafetyAliasDevice")
    objSafetyAliasDevice.setAttribute "Id", "g1_it1"
    objSafetyAliasDevice.setAttribute "IsExternal", "false"
    objAliasDevices.appendChild objSafetyAliasDevice

    Set objVendorId = objXML.createElement("VendorId")
    objVendorId.Text = "2"
    objSafetyAliasDevice.appendChild objVendorId
        
    Set objProductCode = objXML.createElement("ProductCode")
    objProductCode.Text = "124792914"
    objSafetyAliasDevice.appendChild objProductCode

    Set objRevisionNo = objXML.createElement("RevisionNo")
    objRevisionNo.Text = "1245184"
    objSafetyAliasDevice.appendChild objRevisionNo

    Set objType = objXML.createElement("Type")
    objType.Text = "60"
    objSafetyAliasDevice.appendChild objType

    Set objSubType = objXML.createElement("SubType")
    objSubType.Text = "190"
    objSafetyAliasDevice.appendChild objSubType

    Set objConnectionType = objXML.createElement("ConnectionType")
    objConnectionType.Text = "FSoEMaster"
    objSafetyAliasDevice.appendChild objConnectionType

    Set objAliasName = objXML.createElement("AliasName")
    objAliasName.Text = "EL1904, 4 digital inputs_1"
    objSafetyAliasDevice.appendChild objAliasName

    Set objConnectionId = objXML.createElement("ConnectionId")
    objConnectionId.Text = "5"
    objSafetyAliasDevice.appendChild objConnectionId

    Set objSafeAddress = objXML.createElement("SafeAddress")
    objSafeAddress.Text = "5"
    objSafetyAliasDevice.appendChild objSafeAddress

    Set objWatchdog = objXML.createElement("Watchdog")
    objWatchdog.Text = "100"
    objSafetyAliasDevice.appendChild objWatchdog

    Set objSafetyParameters = objXML.createElement("SafetyParameters")
    objSafetyAliasDevice.appendChild objSafetyParameters

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "1"
    objSafetyParameter.setAttribute "Name", "8000:01 FS Operating Mode:Operating Mode"
    objSafetyParameter.Text = "0"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "2"
    objSafetyParameter.setAttribute "Name", "8001:01 FS Sensor Test:Sensor test Channel 1 active"
    objSafetyParameter.Text = "1"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "3"
    objSafetyParameter.setAttribute "Name", "8001:02 FS Sensor Test:Sensor test Channel 2 active"
    objSafetyParameter.Text = "1"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "4"
    objSafetyParameter.setAttribute "Name", "8001:03 FS Sensor Test:Sensor test Channel 3 active"
    objSafetyParameter.Text = "1"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "5"
    objSafetyParameter.setAttribute "Name", "8001:04 FS Sensor Test:Sensor test Channel 4 active"
    objSafetyParameter.Text = "1"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "6"
    objSafetyParameter.setAttribute "Name", "8002:01 FS Logic of Input pairs:Logic of Channel 1 and 2"
    objSafetyParameter.Text = "0"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "7"
    objSafetyParameter.setAttribute "Name", "8002:03 FS Logic of Input pairs:Logic of Channel 3 and 4"
    objSafetyParameter.Text = "0"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "8"
    objSafetyParameter.setAttribute "Name", "10E0:01 Store Code"
    objSafetyParameter.Text = "0"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "9"
    objSafetyParameter.setAttribute "Name", "10E0:02 Project CRC"
    objSafetyParameter.Text = "0"
    objSafetyParameters.appendChild objSafetyParameter

    Set objRxPdo = objXML.createElement("RxPdo")
    objRxPdo.setAttribute "BitSize", "48"
    objSafetyAliasDevice.appendChild objRxPdo

    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_it1_i1"
    objRxPdo.appendChild objElement

    Set objName = objXML.createElement("Name")
    objName.Text = "InputChannel1"
    objElement.appendChild objName

    Set objType = objXML.createElement("Type")
    objType.Text = "BIT"
    objElement.appendChild objType

    Set objBitOffset = objXML.createElement("BitOffset")
    objBitOffset.Text = "8"
    objElement.appendChild objBitOffset

    Set objBitSize = objXML.createElement("BitSize")
    objBitSize.Text = "1"
    objElement.appendChild objBitSize

    Set objIsSafeTimer = objXML.createElement("IsSafeTimer")
    objIsSafeTimer.Text = "false"
    objElement.appendChild objIsSafeTimer

    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_it1_i2"
    objRxPdo.appendChild objElement

    Set objName = objXML.createElement("Name")
    objName.Text = "InputChannel2"
    objElement.appendChild objName

    Set objType = objXML.createElement("Type")
    objType.Text = "BIT"
    objElement.appendChild objType

    Set objBitOffset = objXML.createElement("BitOffset")
    objBitOffset.Text = "9"
    objElement.appendChild objBitOffset

    Set objBitSize = objXML.createElement("BitSize")
    objBitSize.Text = "1"
    objElement.appendChild objBitSize

    Set objIsSafeTimer = objXML.createElement("IsSafeTimer")
    objIsSafeTimer.Text = "false"
    objElement.appendChild objIsSafeTimer
    
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_it1_i3"
    objRxPdo.appendChild objElement

    Set objName = objXML.createElement("Name")
    objName.Text = "InputChannel3"
    objElement.appendChild objName

    Set objType = objXML.createElement("Type")
    objType.Text = "BIT"
    objElement.appendChild objType

    Set objBitOffset = objXML.createElement("BitOffset")
    objBitOffset.Text = "10"
    objElement.appendChild objBitOffset

    Set objBitSize = objXML.createElement("BitSize")
    objBitSize.Text = "1"
    objElement.appendChild objBitSize

    Set objIsSafeTimer = objXML.createElement("IsSafeTimer")
    objIsSafeTimer.Text = "false"
    objElement.appendChild objIsSafeTimer
    
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_it1_i4"
    objRxPdo.appendChild objElement

    Set objName = objXML.createElement("Name")
    objName.Text = "InputChannel4"
    objElement.appendChild objName

    Set objType = objXML.createElement("Type")
    objType.Text = "BIT"
    objElement.appendChild objType

    Set objBitOffset = objXML.createElement("BitOffset")
    objBitOffset.Text = "11"
    objElement.appendChild objBitOffset

    Set objBitSize = objXML.createElement("BitSize")
    objBitSize.Text = "1"
    objElement.appendChild objBitSize

    Set objIsSafeTimer = objXML.createElement("IsSafeTimer")
    objIsSafeTimer.Text = "false"
    objElement.appendChild objIsSafeTimer

    Set objComErrAck = objXML.createElement("ComErrAck")
    objComErrAck.setAttribute "Id", "g1_it1_c"
    objSafetyAliasDevice.appendChild objComErrAck

    Set objModuleFaultAsComErr = objXML.createElement("ModuleFaultAsComErr")
    objModuleFaultAsComErr.Text = "false"
    objSafetyAliasDevice.appendChild objModuleFaultAsComErr

    Set objMapState = objXML.createElement("MapState")
    objMapState.Text = "false"
    objSafetyAliasDevice.appendChild objMapState

    Set objMapDiag = objXML.createElement("MapDiag")
    objMapDiag.Text = "false"
    objSafetyAliasDevice.appendChild objMapDiag

    Set objMapInputs = objXML.createElement("MapInputs")
    objMapInputs.Text = "false"
    objSafetyAliasDevice.appendChild objMapInputs

    Set objMapOutputs = objXML.createElement("MapOutputs")
    objMapOutputs.Text = "false"
    objSafetyAliasDevice.appendChild objMapOutputs

    Set objIncrementalDownloadEnabled = objXML.createElement("IncrementalDownloadEnabled")
    objIncrementalDownloadEnabled.Text = "false"
    objSafetyAliasDevice.appendChild objIncrementalDownloadEnabled
    
    ' SafetyAliasDevice 2
    Set objSafetyAliasDevice = objXML.createElement("SafetyAliasDevice")
    objSafetyAliasDevice.setAttribute "Id", "g1_ot1"
    objSafetyAliasDevice.setAttribute "IsExternal", "false"
    objAliasDevices.appendChild objSafetyAliasDevice

    Set objVendorId = objXML.createElement("VendorId")
    objVendorId.Text = "2"
    objSafetyAliasDevice.appendChild objVendorId

    Set objProductCode = objXML.createElement("ProductCode")
    objProductCode.Text = "190328914"
    objSafetyAliasDevice.appendChild objProductCode

    Set objRevisionNo = objXML.createElement("RevisionNo")
    objRevisionNo.Text = "1245184"
    objSafetyAliasDevice.appendChild objRevisionNo

    Set objType = objXML.createElement("Type")
    objType.Text = "60"
    objSafetyAliasDevice.appendChild objType

    Set objSubType = objXML.createElement("SubType")
    objSubType.Text = "290"
    objSafetyAliasDevice.appendChild objSubType

    Set objConnectionType = objXML.createElement("ConnectionType")
    objConnectionType.Text = "FSoEMaster"
    objSafetyAliasDevice.appendChild objConnectionType

    Set objAliasName = objXML.createElement("AliasName")
    objAliasName.Text = "EL2904, 4 digital outputs_1"
    objSafetyAliasDevice.appendChild objAliasName

    Set objConnectionId = objXML.createElement("ConnectionId")
    objConnectionId.Text = "17"
    objSafetyAliasDevice.appendChild objConnectionId

    Set objSafeAddress = objXML.createElement("SafeAddress")
    objSafeAddress.Text = "17"
    objSafetyAliasDevice.appendChild objSafeAddress

    Set objWatchdog = objXML.createElement("Watchdog")
    objWatchdog.Text = "100"
    objSafetyAliasDevice.appendChild objWatchdog

    Set objSafetyParameters = objXML.createElement("SafetyParameters")
    objSafetyAliasDevice.appendChild objSafetyParameters

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "1"
    objSafetyParameter.setAttribute "Name", "8000:01 FSOE Settings:Standard outputs active"
    objSafetyParameter.Text = "0"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "2"
    objSafetyParameter.setAttribute "Name", "8000:02 FSOE Settings:Current measurement active"
    objSafetyParameter.Text = "1"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "3"
    objSafetyParameter.setAttribute "Name", "8000:03 FSOE Settings:Testing of outputs active"
    objSafetyParameter.Text = "1"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "4"
    objSafetyParameter.setAttribute "Name", "8000:04 FSOE Settings:Error acknowledge active"
    objSafetyParameter.Text = "0"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "5"
    objSafetyParameter.setAttribute "Name", "10E0:01 Store Code"
    objSafetyParameter.Text = "0"
    objSafetyParameters.appendChild objSafetyParameter

    Set objSafetyParameter = objXML.createElement("SafetyParameter")
    objSafetyParameter.setAttribute "OrderId", "6"
    objSafetyParameter.setAttribute "Name", "10E0:02 Project CRC"
    objSafetyParameter.Text = "0"
    objSafetyParameters.appendChild objSafetyParameter

    Set objTxPdo = objXML.createElement("TxPdo")
    objTxPdo.setAttribute "BitSize", "48"
    objSafetyAliasDevice.appendChild objTxPdo

    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_ot1_o1"
    objTxPdo.appendChild objElement

    Set objName = objXML.createElement("Name")
    objName.Text = "OutputChannel1"
    objElement.appendChild objName

    Set objType = objXML.createElement("Type")
    objType.Text = "BIT"
    objElement.appendChild objType

    Set objBitOffset = objXML.createElement("BitOffset")
    objBitOffset.Text = "8"
    objElement.appendChild objBitOffset

    Set objBitSize = objXML.createElement("BitSize")
    objBitSize.Text = "1"
    objElement.appendChild objBitSize

    Set objIsSafeTimer = objXML.createElement("IsSafeTimer")
    objIsSafeTimer.Text = "false"
    objElement.appendChild objIsSafeTimer
    
    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_ot1_o2"
    objTxPdo.appendChild objElement

    Set objName = objXML.createElement("Name")
    objName.Text = "OutputChannel2"
    objElement.appendChild objName

    Set objType = objXML.createElement("Type")
    objType.Text = "BIT"
    objElement.appendChild objType

    Set objBitOffset = objXML.createElement("BitOffset")
    objBitOffset.Text = "9"
    objElement.appendChild objBitOffset

    Set objBitSize = objXML.createElement("BitSize")
    objBitSize.Text = "1"
    objElement.appendChild objBitSize

    Set objIsSafeTimer = objXML.createElement("IsSafeTimer")
    objIsSafeTimer.Text = "false"
    objElement.appendChild objIsSafeTimer

    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_ot1_o3"
    objTxPdo.appendChild objElement

    Set objName = objXML.createElement("Name")
    objName.Text = "OutputChannel3"
    objElement.appendChild objName

    Set objType = objXML.createElement("Type")
    objType.Text = "BIT"
    objElement.appendChild objType

    Set objBitOffset = objXML.createElement("BitOffset")
    objBitOffset.Text = "10"
    objElement.appendChild objBitOffset

    Set objBitSize = objXML.createElement("BitSize")
    objBitSize.Text = "1"
    objElement.appendChild objBitSize

    Set objIsSafeTimer = objXML.createElement("IsSafeTimer")
    objIsSafeTimer.Text = "false"
    objElement.appendChild objIsSafeTimer

    Set objElement = objXML.createElement("Element")
    objElement.setAttribute "Id", "g1_ot1_o4"
    objTxPdo.appendChild objElement

    Set objName = objXML.createElement("Name")
    objName.Text = "OutputChannel4"
    objElement.appendChild objName

    Set objType = objXML.createElement("Type")
    objType.Text = "BIT"
    objElement.appendChild objType

    Set objBitOffset = objXML.createElement("BitOffset")
    objBitOffset.Text = "11"
    objElement.appendChild objBitOffset

    Set objBitSize = objXML.createElement("BitSize")
    objBitSize.Text = "1"
    objElement.appendChild objBitSize

    Set objIsSafeTimer = objXML.createElement("IsSafeTimer")
    objIsSafeTimer.Text = "false"
    objElement.appendChild objIsSafeTimer

    Set objComErrAck = objXML.createElement("ComErrAck")
    objComErrAck.setAttribute "Id", "g1_ot1_c"
    objSafetyAliasDevice.appendChild objComErrAck

    Set objModuleFaultAsComErr = objXML.createElement("ModuleFaultAsComErr")
    objModuleFaultAsComErr.Text = "false"
    objSafetyAliasDevice.appendChild objModuleFaultAsComErr

    Set objMapState = objXML.createElement("MapState")
    objMapState.Text = "false"
    objSafetyAliasDevice.appendChild objMapState

    Set objMapDiag = objXML.createElement("MapDiag")
    objMapDiag.Text = "false"
    objSafetyAliasDevice.appendChild objMapDiag

    Set objMapInputs = objXML.createElement("MapInputs")
    objMapInputs.Text = "false"
    objSafetyAliasDevice.appendChild objMapInputs

    Set objMapOutputs = objXML.createElement("MapOutputs")
    objMapOutputs.Text = "false"
    objSafetyAliasDevice.appendChild objMapOutputs

    Set objIncrementalDownloadEnabled = objXML.createElement("IncrementalDownloadEnabled")
    objIncrementalDownloadEnabled.Text = "false"
    objSafetyAliasDevice.appendChild objIncrementalDownloadEnabled
    
    ' Create the Application element
    Set objApplicationNode = objXML.createElement("Application")
    objTwinSAFEGroup.appendChild objApplicationNode
    
    ' --- Get Actor Names from Excel ---
    Dim actorNames(2) As String
    actorNames(0) = ws.Range("C3").Value
    actorNames(1) = ws.Range("D3").Value
    actorNames(2) = ws.Range("E3").Value
    
    ' --- Create Networks based on Actor Names---

    Dim networkCounter As Integer
    For networkCounter = 0 To 2
    
        ' Check if the actor has any X values in the column
        Dim actorHasMappings As Boolean
        actorHasMappings = False
        For rowCounter = 0 To 2
            If ThisWorkbook.Sheets(1).Cells(rowCounter + 4, networkCounter + 3).Value = "X" Then
                actorHasMappings = True
                Exit For
            End If
        Next rowCounter
        
        ' Skip network creation if the actor has no mappings
        If Not actorHasMappings Then
            GoTo SkipActorNetwork
        End If

        ' Create Network element
        Dim networkElement As Object
        Set networkElement = objXML.createElement("Network")
        networkElement.setAttribute "Id", "g1_n" & networkCounter + 1
        networkElement.setAttribute "OrderId", networkCounter + 1
        objApplicationNode.appendChild networkElement

        ' Create Name element for Network
        Dim networkNameElement As Object
        Set networkNameElement = objXML.createElement("Name")
        networkNameElement.Text = "Network_" & actorNames(networkCounter)
        networkElement.appendChild networkNameElement

        ' Create FunctionBlock element
        Dim functionBlockElement As Object
        Set functionBlockElement = objXML.createElement("FunctionBlock")
        functionBlockElement.setAttribute "Id", "g1_n" & networkCounter + 1 & "_f1"
        functionBlockElement.setAttribute "OrderId", networkCounter + 1
        networkElement.appendChild functionBlockElement

        ' Create elements within FunctionBlock
        Dim typeElement As Object
        Set typeElement = objXML.createElement("Type")
        typeElement.Text = "34"
        functionBlockElement.appendChild typeElement

        Dim nameElement As Object
        Set nameElement = objXML.createElement("Name")
        nameElement.Text = "FBAnd" & networkCounter + 1
        functionBlockElement.appendChild nameElement

        Dim mapStateElement As Object
        Set mapStateElement = objXML.createElement("MapState")
        mapStateElement.Text = "false"
        functionBlockElement.appendChild mapStateElement

        Dim mapDiagElement As Object
        Set mapDiagElement = objXML.createElement("MapDiag")
        mapDiagElement.Text = "false"
        functionBlockElement.appendChild mapDiagElement

        ' Create Outports element
        Dim outportsElement As Object
        Set outportsElement = objXML.createElement("Outports")
        functionBlockElement.appendChild outportsElement

        ' Create Outport Element
        Dim outportElement As Object
        Set outportElement = objXML.createElement("Element")
        outportElement.setAttribute "Id", "g1_n" & networkCounter + 1 & "_f1_o1"
        outportsElement.appendChild outportElement

        Dim portIdElement As Object
        Set portIdElement = objXML.createElement("PortId")
        portIdElement.Text = "33620002"
        outportElement.appendChild portIdElement

        Dim portNameElement As Object
        Set portNameElement = objXML.createElement("PortName")
        portNameElement.Text = "AndOut"
        outportElement.appendChild portNameElement

        Dim aliasElement As Object
        Set aliasElement = objXML.createElement("Alias")
        aliasElement.Text = "Output1"
        outportElement.appendChild aliasElement

        Dim typeElement2 As Object
        Set typeElement2 = objXML.createElement("Type")
        typeElement2.Text = "BIT"
        outportElement.appendChild typeElement2

        ' Create Inports element
        Dim inportsElement As Object
        Set inportsElement = objXML.createElement("Inports")
        functionBlockElement.appendChild inportsElement
        
        Dim inportCounter As Integer
        For inportCounter = 1 To 8

            ' Create Inport Element
            Dim inportElement As Object
            Set inportElement = objXML.createElement("Element")
            inportElement.setAttribute "Id", "g1_n" & networkCounter + 1 & "_f1_i" & inportCounter
            inportsElement.appendChild inportElement
            
            Dim portIdIn As Object
            Set portIdIn = objXML.createElement("PortId")
            portIdIn.Text = 17301538 + (inportCounter - 1) * 65536
            inportElement.appendChild portIdIn
            
            Dim portNameIn As Object
            Set portNameIn = objXML.createElement("PortName")
            portNameIn.Text = "AndIn" & inportCounter
            inportElement.appendChild portNameIn
            
            Dim aliasIn As Object
            Set aliasIn = objXML.createElement("Alias")
            aliasIn.Text = "Input" & inportCounter
            inportElement.appendChild aliasIn
            
            Dim typeIn As Object
            Set typeIn = objXML.createElement("Type")
            typeIn.Text = "BIT"
            inportElement.appendChild typeIn
            
            Dim negatedIn As Object
            Set negatedIn = objXML.createElement("Negated")
            negatedIn.Text = "false" ' Default to false, will be updated later based on Excel
            inportElement.appendChild negatedIn
            
            Dim activeIn As Object
            Set activeIn = objXML.createElement("Active")
            activeIn.Text = "true"
            inportElement.appendChild activeIn

        Next inportCounter
               
SkipActorNetwork:
    Next networkCounter
    
     
    ' Create the Mappings element
    Set objMappings = objXML.createElement("Mappings")
    objAppConfig.appendChild objMappings

    ' --- Create Mappings from Excel ---
    Dim sensorNames(2) As String
    sensorNames(0) = ThisWorkbook.Sheets(1).Range("A4").Value
    sensorNames(1) = ThisWorkbook.Sheets(1).Range("A5").Value
    sensorNames(2) = ThisWorkbook.Sheets(1).Range("A6").Value
    
    'Dim rowCounter As Integer
    For colCounter = 0 To 2
        Dim inputIndex As Integer
        inputIndex = 1 ' Start with Input 1 for each Actor column

        For rowCounter = 0 To 2
            If ThisWorkbook.Sheets(1).Cells(rowCounter + 4, colCounter + 3).Value = "X" Then
                ' Create Mapping, dynamically assigning inputIndex
                Dim mappingElement As Object
                Set mappingElement = objXML.createElement("Mapping")
                objMappings.appendChild mappingElement
                
                Dim sensorName As String
                sensorName = sensorNames(rowCounter)
                
                Dim negated As Boolean
                negated = False

                If InStr(sensorName, "!") > 0 Then
                    negated = True
                    sensorName = Replace(sensorName, "!", "")
                End If

                mappingElement.setAttribute "TargetId", "g1_n" & colCounter + 1 & "_f1_i" & inputIndex
                mappingElement.setAttribute "SourceId", "g1_it1_i" & rowCounter + 1 ' Still use original rowCounter for SourceId
                mappingElement.setAttribute "LocalVarName", sensorName

                ' Update Negated tag if needed
                If negated Then
                    Dim functionBlockElementToUpdate As Object
                    Set functionBlockElementToUpdate = objXML.SelectSingleNode("//FunctionBlock[@Id='g1_n" & colCounter + 1 & "_f1']")
                    Dim inportElementToUpdate As Object
                    Set inportElementToUpdate = functionBlockElementToUpdate.SelectSingleNode("Inports/Element[@Id='g1_n" & colCounter + 1 & "_f1_i" & inputIndex & "']")
                    Dim negatedElementToUpdate As Object
                    Set negatedElementToUpdate = inportElementToUpdate.SelectSingleNode("Negated")
                    negatedElementToUpdate.Text = "true"
                End If

                inputIndex = inputIndex + 1 ' Increment to next available input
            End If
        Next rowCounter
    Next colCounter

    ' Mapping for Group Ports (assuming these are fixed and don't depend on the Excel sheet)
    Set mappingElement = objXML.createElement("Mapping")
    objMappings.appendChild mappingElement
    mappingElement.setAttribute "TargetId", "g1_i2"
    mappingElement.setAttribute "SourceId", "g1_a1_i1"
    mappingElement.setAttribute "LocalVarName", "GroupPort_ErrAck"

    Set mappingElement = objXML.createElement("Mapping")
    objMappings.appendChild mappingElement
    mappingElement.setAttribute "TargetId", "g1_i1"
    mappingElement.setAttribute "SourceId", "g1_a2_i1"
    mappingElement.setAttribute "LocalVarName", "GroupPort_Run"

    ' Mappings for Actors
    For actorCounter = 0 To 2
        actorHasMappings = False
        For rowCounter = 0 To 2
            If ThisWorkbook.Sheets(1).Cells(rowCounter + 4, actorCounter + 3).Value = "X" Then
                actorHasMappings = True
                Exit For
            End If
        Next rowCounter
        
        ' Skip network creation if the actor has no mappings
        If Not actorHasMappings Then
            GoTo SkipActorMapping
        End If
    
    
        Set mappingElement = objXML.createElement("Mapping")
        objMappings.appendChild mappingElement
        mappingElement.setAttribute "TargetId", "g1_ot1_o" & actorCounter + 1
        mappingElement.setAttribute "SourceId", "g1_n" & actorCounter + 1 & "_f1_o1"
        mappingElement.setAttribute "LocalVarName", actorNames(actorCounter)
        
SkipActorMapping:
    Next actorCounter


    ' --- Get save location from user ---
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="TwinSAFE_XML_Export", _
        fileFilter:="TwinSAFE XML (*.xml), *.xml", _
        Title:="Save XML File")

    ' Check if user canceled
    If filePath = False Then
        MsgBox "File save canceled.", vbExclamation
        Exit Sub
    End If

    ' --- Indentation and New Line Formatting ---
    Dim writer As Object
    Set writer = CreateObject("MSXML2.MXXMLWriter.6.0")
    writer.indent = True
    writer.omitXMLDeclaration = False
    writer.Encoding = "UTF-8"  ' Set encoding directly on the writer

    Dim reader As Object
    Set reader = CreateObject("MSXML2.SAXXMLReader.6.0")
    Set reader.contentHandler = writer

    ' Parse the XML through the reader
    reader.Parse objXML

    ' Save the formatted XML to a file using ADODB.Stream
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8" ' Still set charset on stream for safety
    stream.Open
    stream.WriteText writer.Output
    stream.SaveToFile filePath, 2

    stream.Close
    ' Confirmation message
    MsgBox "XML file saved successfully to: " & filePath, vbInformation
End Sub
