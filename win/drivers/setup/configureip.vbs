


strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE DHCPEnabled='true' AND IPEnabled='true'") 
For Each objItem in colItems 
    
    wscript.echo "Adapter DHCP Info:"
    Wscript.Echo "DHCPServer: " & objItem.DHCPServer
    Wscript.Echo "DNSDomain: " & objItem.DNSDomain

    If isNull(objItem.DNSDomainSuffixSearchOrder) Then
        Wscript.Echo "DNSDomainSuffixSearchOrder: "
    Else
        Wscript.Echo "DNSDomainSuffixSearchOrder: " & Join(objItem.DNSDomainSuffixSearchOrder, ",")
    End If

    If isNull(objItem.DNSServerSearchOrder) Then
        Wscript.Echo "DNSServerSearchOrder: "
    Else
        Wscript.Echo "DNSServerSearchOrder: " & Join(objItem.DNSServerSearchOrder, ",")
    End If

    Wscript.Echo "InterfaceIndex: " & objItem.InterfaceIndex

    If isNull(objItem.IPAddress) Then
        Wscript.Echo "IPAddress: "
    Else
        Wscript.Echo "IPAddress: " & Join(objItem.IPAddress, ",")
    End If

    If isNull(objItem.IPSubnet) Then
        Wscript.Echo "IPSubnet: "
    Else
        Wscript.Echo "IPSubnet: " & Join(objItem.IPSubnet, ",")
    End If

    Wscript.Echo "WINSPrimaryServer: " & objItem.WINSPrimaryServer
    Wscript.Echo "WINSSecondaryServer: " & objItem.WINSSecondaryServer

'Set Everything
set objShare = objWMIService.Get("Win32_NetworkAdapterConfiguration.Index='" & objItem.Index & "'")
wscript.echo "Release DHCP"
Set objOutParams = objWMIService.ExecMethod("Win32_NetworkAdapterConfiguration.Index='" & objItem.Index & "'", "ReleaseDHCPLease")
Wscript.echo "ReturnValue: " & objOutParams.ReturnValue

'Static IP
wscript.echo "Setting Static IP " & objItem.IPAddress(0)
Set objInParam = objShare.Methods_("EnableStatic").inParameters.SpawnInstance_()
objInParam.Properties_.Item("SubnetMask") =  Array(objItem.IPSubnet(0))
objInParam.Properties_.Item("IPAddress") =  Array(objItem.IPAddress(0))
Set objOutParams = objWMIService.ExecMethod("Win32_NetworkAdapterConfiguration.Index='" & objItem.Index & "'", "EnableStatic", objInParam)
Wscript.echo "ReturnValue: " & objOutParams.ReturnValue

'Gateway
'Only set one gateway and make this the 10. not the 172. (this is specific to our networking you may want to find a better solution)
if colitems.count = 1 then
 strGateway = objItem.DefaultIPGateway(0)
else
for x = 0 to ubound(objItem.DefaultIPGateway)
 arrGateway = split(objItem.DefaultIPGateway(x), ".")
 if arrGateway(0) = "10" then
  strGateway = objItem.DefaultIPGateway(x)
  
  Set objShell = WScript.CreateObject("WScript.Shell")    
  objResult = objShell.Run("route add 169.254.169.0 mask 255.255.255.0 " & strGateway, 0, True)
wscript.echo "Route Added " & objResult

 else
  strGateway = objItem.IPAddress(0)
 end if
next
end if


wscript.echo "Setting Gateway " & strGateway

Set objInParam = objShare.Methods_("SetGateways").inParameters.SpawnInstance_()
objInParam.Properties_.Item("DefaultIPGateway") =  Array(strGateway)
objInParam.Properties_.Item("GatewayCostMetric") =  Array(1)
Set objOutParams = objWMIService.ExecMethod("Win32_NetworkAdapterConfiguration.Index='" & objItem.Index & "'", "SetGateways", objInParam)
Wscript.echo "ReturnValue: " & objOutParams.ReturnValue


wscript.echo "DNS"
'Static DNS
strDNS = objItem.DNSServerSearchOrder(0)
arrDNS = split(strDNS,".")
if arrDNS(0) = "10" then
set objShare = objWMIService.Get("Win32_NetworkAdapterConfiguration")
Set objInParam = objShare.Methods_("EnableDNS").inParameters.SpawnInstance_()
objInParam.Properties_.Item("DNSHostName") =  objItem.DNSHostName
objInParam.Properties_.Item("DNSDomain") =  objItem.DNSDomain
objInParam.Properties_.Item("DNSServerSearchOrder") =  objItem.DNSServerSearchOrder
objInParam.Properties_.Item("DNSDomainSuffixSearchOrder") =  objItem.DNSDomainSuffixSearchOrder
Set objOutParams = objWMIService.ExecMethod("Win32_NetworkAdapterConfiguration", "EnableDNS", objInParam)
Wscript.echo "ReturnValue: " & objOutParams.ReturnValue
end if

Next
