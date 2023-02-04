' by http://kevinisms.fason.org'
'Written by Kevin Fason
'This script will collect various hardware information into variables for use by a TS in SCCM.

'open TS Environmnet
set env = CreateObject("Microsoft.SMS.TSEnvironment")

'*******************************************************************************
'declare global variables
'*******************************************************************************
  Dim objWMIService, colSMBIOS, objSMBIOS, assetTag, ComputerName, ComputerModel, gsChassis, oItem

'set the global variables for use later    
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 
  Set colSMBIOS = objWMIService.ExecQuery("Select * from Win32_SystemEnclosure") 

'*******************************************************************************
'* Get the Computers asset tag or number based on WMI query
'*******************************************************************************
For Each objSMBIOS in colSMBIOS
      assetTag = objSMBIOS.SMBIOSAssetTag
      assetTag = Trim(assetTag)
      ComputerName = Trim(objSMBIOS.SMBIOSAssetTag)
  
      For Each oItem in objSMBIOS.ChassisTypes
          gsChassis = oItem
      Next
      
  Next
    
      Select Case gsChassis 
        Case "8","9","10","11","12","13","14"
            gsChassis = "Laptop"
        Case Else
            'default it to desktop since that's more than likely the correct one
            gsChassis = "Desktop"
      End Select
    
  set colSMBIOS = nothing
  Set colSMBIOS = objWMIService.ExecQuery("Select * from Win32_ComputerSystem") 
    
    For Each objSMBIOS in colSMBIOS
       modelNumber = objSMBIOS.Model
       modelNumber = Trim(modelNumber)
       ComputerModel = Trim(objSMBIOS.Model)
    Next
   
   Set objWMIService = Nothing
   Set colSMBIOS = Nothing

'export to TS variables   
   env("OSDComputerModel") = ComputerModel
   env("OSDAssetTag") = ComputerName
   env("OSDChassis") = gsChassis