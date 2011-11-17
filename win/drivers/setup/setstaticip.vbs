'==========================================================================
'
' NAME: setstaticip.vbs
'
' AUTHOR: Jordan Rinke (me@jordanrinke.com, jordan@openstack.org) 
' DATE  : 11/16/2011
'
' COMMENT: Reads the DHCP response from nics and sets the cards To
'           a static configuration, also removed the gateway for the
'           private nic and sets a static route on the pub nic.
'           The static route is always needed, the static nic config
'           is due to DNSMasq not taking direct dhcp requests and Not
'           being able to get windows to broadcast only. If you have
'           broadcast only dhcp working with windows 2008 please 
'           let me know
'
'           I don't have any logging built in since I exec/pipe to file
'           because vbs has no try/catch feature unknown errors will
'           will break to stderr and that is the only way to catch them.
'
'           Just exec as: 
'           cscript.exe //NoLogo setstaticip.vbs >> log.txt 2>&1

'
'           Side note: My images disable ipv6 right now so I have no
'           idea if this works with ipv6 enabled, als this in theory
'           works for nova installs with single nic but I have not tested
'
'==========================================================================

Dim objWMIService, colNICs, objNIC, objNICConfig, objOutParams, objInParams
Dim strGateway, arrGateway, intCount, arrIP 
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2") 
Set colNICs = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_NetworkAdapterConfiguration " _
    & "WHERE DHCPEnabled='true' AND IPEnabled='true'") 
For Each objNIC In colNICs
    'List all of the current configuration info for logging to file
    wscript.echo "Adapter DHCP Info For: " & objNIC.Caption
    Wscript.Echo "DHCPServer: " & objNIC.DHCPServer
    Wscript.Echo "DNSDomain: " & objNIC.DNSDomain

    If isNull(objNIC.DNSDomainSuffixSearchOrder) = False Then
        Wscript.Echo "DNSDomainSuffixSearchOrder: " & _
            Join(objNIC.DNSDomainSuffixSearchOrder, ",")
    End If

    If isNull(objNIC.DNSServerSearchOrder)  = False Then
        Wscript.Echo "DNSServerSearchOrder: " & _
            Join(objNIC.DNSServerSearchOrder, ",")
    End If

    Wscript.Echo "InterfaceIndex: " & objNIC.InterfaceIndex

    If isNull(objNIC.IPAddress) = False Then
        Wscript.Echo "IPAddress: " & Join(objNIC.IPAddress, ",")
    End If

    If isNull(objNIC.IPSubnet)  = False Then
        Wscript.Echo "Subnet Mask: " & Join(objNIC.IPSubnet, ",")
    End If

    'Each of these could be a function but I didn't specifically see
    'a reason or need to break it out, left it procedural for now
    Wscript.Echo "Releasing DHCP: " & objNIC.Caption
    Set objNICConfig = objWMIService.Get(_
        "Win32_NetworkAdapterConfiguration.Index='" _
        & objNIC.Index & "'")
    Set objOutParams = objWMIService.ExecMethod( _
        "Win32_NetworkAdapterConfiguration.Index='" & objNIC.Index _
        & "'", "ReleaseDHCPLease")

    Wscript.Echo "Setting Static IP " & objNIC.IPAddress(0) _
        & " and mask " & objNIC.IPSubnet(0)
    Set objInParam = objNICConfig.Methods_(_
        "EnableStatic").inParameters.SpawnInstance_()
    objInParam.Properties_.Item("SubnetMask") =  Array(objNIC.IPSubnet(0))
    objInParam.Properties_.Item("IPAddress") =  Array(objNIC.IPAddress(0))
    Set objOutParams = objWMIService.ExecMethod( _
        "Win32_NetworkAdapterConfiguration.Index='" & objNIC.Index _
        & "'", "EnableStatic", objInParam)
    Wscript.Echo "EnableStatic ReturnValue: " & objOutParams.ReturnValue
    
    If colNICs.count = 1 Then 'specifically left the wmi query async for this
        strGateway = objNIC.DefaultIPGateway(0)
        'if a single nic is detected no logic/route is needed
    Else
        for intCount = 0 to ubound(objNIC.DefaultIPGateway)
            arrGateway = split(objNIC.DefaultIPGateway(intCount), ".")
            if arrGateway(0) = "10" Then
                strGateway = objNIC.DefaultIPGateway(intCount)
                Set objShell = WScript.CreateObject("WScript.Shell")    
                objResult = objShell.Run(_
                    "route -p add 169.254.169.0 mask 255.255.255.0 "_
                    & strGateway, 0, True)
                wscript.echo "Route Added " & objResult
            Else
                strGateway = objNIC.IPAddress(0)
                'There is no way to clear the gateway with WMI
                'this sets it to 0.0.0.0 which is effective
            end If
        Next
    end If


    wscript.echo "Setting Gateway " & strGateway
    Set objInParam = objNICConfig.Methods_( _
        "SetGateways").inParameters.SpawnInstance_()
    objInParam.Properties_.Item("DefaultIPGateway") =  Array(strGateway)
    objInParam.Properties_.Item("GatewayCostMetric") =  Array(1)
    Set objOutParams = objWMIService.ExecMethod(_
        "Win32_NetworkAdapterConfiguration.Index='" & objNIC.Index _
        & "'", "SetGateways", objInParam)
    Wscript.echo "SetGateways ReturnValue: " _
        & objOutParams.ReturnValue

    WScript.Echo "Setting DNS Domain"
    Set objInParam = objNICConfig.Methods_("SetDNSDomain"). _
    inParameters.SpawnInstance_()
    objInParam.Properties_.Item("DNSDomain") = objNIC.DNSHostName
    Set objOutParams = objWMIService.ExecMethod( _
        "Win32_NetworkAdapterConfiguration.Index='" _
         & objNIC.index & "'", "SetDNSDomain", objInParam)
    Wscript.echo "SetDNSDomain ReturnValue: " _
        & objOutParams.ReturnValue

    WScript.Echo "Setting DNS Server Seach Order"
    Set objInParam = objNICConfig.Methods_("SetDNSServerSearchOrder"). _
    inParameters.SpawnInstance_()
    objInParam.Properties_.Item( _
        "DNSServerSearchOrder") =  objNIC.DNSServerSearchOrder
    Set objOutParams = objWMIService.ExecMethod(_
        "Win32_NetworkAdapterConfiguration.Index='" & objNIC.index & "'", _
        "SetDNSServerSearchOrder", objInParam)
    Wscript.echo "SetDNSServerSearchOrder ReturnValue: " _
        & objOutParams.ReturnValue


    
    arrIP = split(objNIC.IPAddress(0),".")
    If arrIP(0) = "10" Then
    'again 10 is always our pub, you may want to do this
    'in a different way
    'Suffix search is global so need to make sure we only take pub nic info
        Set objInParam = objNICConfig.Methods_("SetDNSSuffixSearchOrder"). _
        inParameters.SpawnInstance_()
        objInParam.Properties_.Item( _
            "DNSDomainSuffixSearchOrder") =  join(objNIC.DNSDomainSuffixSearchOrder,",")
        Set objOutParams = objWMIService.ExecMethod( _
            "Win32_NetworkAdapterConfiguration",_
            "SetDNSSuffixSearchOrder", objInParam)
        Wscript.echo "SetDNSSuffixSearchOrder ReturnValue: " _
        & objOutParams.ReturnValue
    End If

Next