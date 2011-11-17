'==========================================================================
'
' NAME: ec2userdata.vbs
'
' AUTHOR: Jordan Rinke (me@jordanrinke.com, jordan@openstack.org) 
' DATE  : 11/16/2011
'
' COMMENT: Simple ec2 userdata interface for setting admin password. 
'			Please feel free to expand on the functionality of this.
'			http://github.com/jordanrinke/openstack/tree/master/win
'
'           I don't have any logging built in since I exec/pipe to file
'           because vbs has no try/catch feature unknown errors will
'           will break to stderr and that is the only way to catch them.
'
'           Just exec as: 
'           cscript.exe //NoLogo ec2userdata.vbs >> log.txt 2>&1
'
'==========================================================================

On Error Resume Next
Dim intMaxTries, intTryDelay, strServer, strPort, strVersion, strLocation
Dim strURL, strPass, intCount, blnDebug 

blnDebug = False
'will output the metadata to console if true (i.e. password etc)

intMaxTries = 15
intTryDelay = 5000 'in ms
strServer = "169.254.169.254" 'ec2 standard IP for metadata
strPort = "8773"
strVersion = "latest"
strLocation = "user-data"
strURL = "http://" & strServer & ":" & strPort & "/" & strVersion _
  & "/" & strLocation

Function fncGetPass(strToUserData)
    Dim objXMLHTTP
    Set objXMLHTTP = CreateObject ("msxml2.xmlhttp.3.0")
    objXMLHTTP.open "get", strToUserData, False
    objXMLHTTP.send
    fncGetPass = objXMLHTTP.responseText
End Function

Sub subSetPassLocal(strPassword)
'There appears to be no way to verify this works/results easily so
'we just have to run it and hope for the best.
    Dim objUser
    Set objUser = GetObject("WinNT://./Administrator, user")
    objUser.SetPassword strPassword
    'Be warned a password not meeting min sec reqs will not work
    objUser.SetInfo
End Sub

For intCount = 1 To intMaxTries
    strPass = fncGetPass(strURL)
    If strPass = "" Then
        WScript.Echo "UserData not found at " & strURL & " - Waiting " & _
            intTryDelay \ 1000 & " seconds (" & intCount & " of " & _
            intMaxTries & ")" 'ugly but oh well
        WScript.Sleep intTryDelay
    Else
        WScript.Echo "UserData found at " & strURL
        If blnDebug = True Then WScript.Echo "UserData: " & strPass 
        subSetPassLocal strPass
        Exit For
    End If
Next    
