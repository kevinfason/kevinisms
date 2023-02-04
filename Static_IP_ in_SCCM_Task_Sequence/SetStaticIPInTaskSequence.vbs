Set objWMIService = GetObject ("winmgmts:\\.\root\cimv2")
 WQL = "Select * from Win32_NetworkAdapterConfiguration where IPEnabled = TRUE"
 Set colNetAdapters = objWMIService.ExecQuery (WQL)
 
WScript.Echo "Number of adapters is " & colNetAdapters.Count
 
set objSCCM = CreateObject("Microsoft.SMS.TSEnvironment")
 
if objSCCM("OSDAdapter0IPAddressList") = "" then
 
 For Each objNetAdapter In colNetAdapters
   if objNetAdapter.DHCPEnabled then
    WScript.Echo "DHCP Enabled"
   else
    WScript.Echo "DHCP Disabled"
    objSCCM("OSDAdapter0EnableDHCP") = "false"
    
    if Not IsNull (objNetAdapter.IPAddress) then 
    strIPAddress = objNetAdapter.IPAddress(0)
     WScript.Echo "IP Address:       " & strIPAddress
     objSCCM("OSDAdapter0IPAddressList") = strIPAddress
    end if
       
   if Not IsNull (objNetAdapter.IPSubnet) then
     strIPSubnet = objNetAdapter.IPSubnet(0)
     WScript.Echo "IP Subnet:        " & strIPSubnet
     objSCCM("OSDAdapter0SubnetMask") = strIPSubnet
    end if
       
   if Not IsNull (objNetAdapter.DefaultIPGateway) then
     strIPGateway = objNetAdapter.DefaultIPGateway(0)
     WScript.Echo "IP Gateway:       " & strIPGateway
     objSCCM("OSDAdapter0Gateways") = strIPGateway
    end if
       
   if Not IsNull (objNetAdapter.DNSServerSearchOrder) then
     strDNSServerSearchOrder = objNetAdapter.DNSServerSearchOrder(0)
     WScript.Echo "DNS Server:       " & strDNSServerSearchOrder
     objSCCM("OSDAdapter0DNSServerList") = strDNSServerSearchOrder
    end if
 
   if Not IsNull (objNetAdapter.MACAddress) then
     strMACAddress = objNetAdapter.MACAddress(0)
     WScript.Echo "MAC Address:      " & strMACAddress
    end if
       
   if Not IsNull (objNetAdapter.DNSDomainSuffixSearchOrder) then
     strDNSDomainSuffixSearchOrder = objNetAdapter.DNSDomainSuffixSearchOrder(0)
     WScript.Echo "DNS DOMAIN:       " & strDNSDomainSuffixSearchOrder
    end if
    
    if Not IsNull (objNetAdapter.WINSPrimaryServer) then
     strWins = objNetAdapter.WINSPrimaryServer
     objSCCM("OSDAdapter0EnableWINS") = "true"
     if Not IsNull (objNetAdapter.WINSSecondaryServer) then
      strWins = strWins & "," & objNetAdapter.WINSSecondaryServer
     end if
     WSCript.Echo "WINS Server:      " & strWins
     objSCCM("OSDAdapter0WINSServerList") = strWins
    else
     objSCCM("OSDAdapter0EnableWINS") = "false"
    end if
 
   objSCCM("OSDAdapterCount") = "1"
        
  end if
  Next
 End If
