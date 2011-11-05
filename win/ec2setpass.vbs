'On Error Resume Next
strServer = "169.254.169.254"
strPort = "8773"
strVersion = "2009-04-04"
strLocation = "user-data"
strFile1 = "c:\windows\panther\unattend\unattend.xml"

strURL = "http://" & strServer & ":" & strPort & "/" & strVersion _
  & "/" & strLocation

Function fncGetPass(strToUserData)
 dim xmlhttp, xmldoc, xmlnode
 set xmlhttp = createobject ("msxml2.xmlhttp.3.0")
 xmlhttp.open "get", strToUserData, false
 xmlhttp.send
 fncGetPass = xmlhttp.responseText
End Function

Sub SetPass(strFileLoc)
For x = 1 to 5
 strPass = fncGetPass(strURL)
 if strPass = "" then
  Set Wscript = CreateObject(Wscript.Shell) 
  wscript.echo "User-Data not found, retrying in 5 seconds"
  Wscript.Sleep 5000
 else
  wscript.echo "User-Data Found"
  set xmldoc = createobject("Microsoft.XMLDOM")
  xmldoc.async = false
  xmldoc.validateOnParse = true
  xmldoc.load strFileLoc
  xmldoc.setProperty "SelectionLanguage","XPath"
  xmldoc.setProperty "SelectionNamespaces", "xmlns:x='urn:schemas-microsoft-com:unattend'"
  set xmlnode = xmldoc.selectSingleNode("/x:unattend/x:settings/x:component/x:UserAccounts/x:AdministratorPassword/x:Value")
  
  if not xmlnode is nothing then
   xmlnode.text = strPass
   wscript.echo "Password saved in " & strFileLoc
  else
   wscript.echo "Unable to update " & strFileLoc
  end if
 
  xmldoc.save strFileLoc
  Exit For
 end if
Next
End Sub

SetPass strFile1
