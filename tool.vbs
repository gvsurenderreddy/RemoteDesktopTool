Dim vDEVICEID
Dim WshShell : set WshShell = CreateObject("WScript.Shell") 

winWidth=788
winHeight=555
window.resizeto winWidth,winHeight
'Center Window on launch ---
'centerX=(screen.width-winWidth)/2
'centerY=(screen.height-winHeight)/2
'window.moveto centerX,centerY
posX=0
posY=0
move=0
Function closeHTA()
  self.close
End Function
Function setPos()
  posX=window.event.screenX
  posY=window.event.ScreenY
  move=1
End Function
Function moving()
  If move=1 Then
    moveX=0
    moveY=0
    moveX=window.event.screenX-posX
    moveY=window.event.screenY-posY
    window.moveto(window.screenLeft+moveX),(window.screenTop+moveY)
    setPos()    
  End if
End Function
Function stopMoving()
  move=0
End Function

Function GetDn(device)
  Dim objTrans, objDomain
  ' Constants for the NameTranslate object.
  Const ADS_NAME_INITTYPE_GC = 3
  Const ADS_NAME_TYPE_NT4 = 3
  Const ADS_NAME_TYPE_1779 = 1
  Set objTrans = CreateObject("NameTranslate")
  Set objDomain = getObject("LDAP://rootDse")
  objTrans.Init ADS_NAME_INITTYPE_GC, ""
  objTrans.Set ADS_NAME_TYPE_NT4, "OHMCNT" & "\" & device & "$"
  GetDN = objTrans.Get(ADS_NAME_TYPE_1779)
  'Set DN to upper Case
  GetDN = UCase(GetDN)
End Function

Function GetEmail(strAccountName, strDomainName)
  Set adoLDAPCon = CreateObject("ADODB.Connection")
  adoLDAPCon.Provider = "ADsDSOObject"
  adoLDAPCon.Open "ADSI"
  strLDAP = "'LDAP://" & strDomainName & "'"
  Set adoLDAPRS = adoLDAPCon.Execute("select mail from " &strLDAP &" WHERE objectClass = 'user'"&" And samAccountName = '" & strAccountName & "'")
  With adoLDAPRS
    If Not .EOF Then
      GetEmail = .Fields("mail")
    Else
      GetEmail = ""
    End If
  End With
  adoLDAPRS.Close
  Set adoLDAPRS = Nothing
  Set adoLDAPCon = Nothing
End Function

'---------------------------------------

'---- Ping Device ----------------------
Sub PingDevice
  deviceInput.value  = UCase(Replace(deviceInput.value," ",""))
  vDEVICEID          = deviceInput.value
  Set objWMIService = GetObject("winmgmts:\\.\root\cimv2") 
  Set objPing = objWMIService.ExecQuery("Select * From Win32_PingStatus Where Address = '" & vDEVICEID & "'" ) 
  For Each objStatus in objPing 
    If IsNull(objStatus.StatusCode) Or objStatus.Statuscode<>0 Then 
      results = MsgBox("Cant Ping "&vDEVICEID&" !",53,"Pinging Device ID: "&vDEVICEID)
      If results = 4 Then
        PingDevice
      End If
      If results = 2 Then
        WshShell.SendKeys "{F5}"
      End If
    Else 
      LoadDevice
    End If 
  Next 
End Sub

'---- Load Device Information --------
Sub LoadDevice
  If vDEVICEID <> "" Then
    GetDeviceInfo
  Else
    WshShell.SendKeys "{F5}"
    deviceInput.value = ""
  End If
End Sub

Sub GetDeviceInfo
  on error resume next
  DT.Style.Display    = "block"
  TC.Style.Display    = "none"
  
  deviceid.innerhtml = vDEVICEID
  load.innerhtml = "<button class='btn btn-success' onClick='PingDevice'>LOAD DEVICE</button>"

  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\"&vDEVICEID&"\root\cimv2") 

  '--- Get IP ----------------------
  Set IPConfigSet = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE") 
  For Each IPConfig in IPConfigSet 
    If Not IsNull(IPConfig.IPAddress) Then  
      For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress) 
        strMacAddr = IpConfig.MACaddress(i) 
        strNetworkCard = Mid(Ipconfig.Caption(i),12,Len(Ipconfig.Caption(i))-11) 
        If Not Instr(IPConfig.IPAddress(i), ":") > 0 Then
          strIPAddr = IpConfig.IPAddress(i)
        End If
      Next 
      If i > 1 Then strIPAddr = strIPaddr 
        If strIPaddr = "" Then
          thinClient
          Exit Sub
        Else 
          IP.Style.Display = "inline"
        End If
      End If 
    Exit For 
  Next
  ipaddress.innerhtml = strIPAddr
  macaddress.innerhtml = strMacAddr
  
  '--- Get User & Email ------------
  Set colItems = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'explorer.exe'") 
  For Each objItem in colItems 
    on error resume next
    objItem.GetOwner strUserName, strUserDomain
    currentUser = strUserDomain & " \ " & strUserName
    If strUserDomain = "OHMCNT" Then
      emailAdress = GetEmail(strUserName, strUserDomain)
    End If
  Next
  userid.innerhtml = ""
  If Not IsEmpty(strUserName) Then 
    If Not IsEmpty(emailAdress) Then 
      email = "<a href=mailto:"&emailAdress&" title="&emailAdress&" class='btn btn-default email-btn btn-xs'>"&_
      "<span class='glyphicon glyphicon-envelope'></span> email</a>"
    Else 
      email = ""
    End If
    userid.innerhtml = currentUser & email
  End if

  '--- Get RAM ---------------------
  Set colItems = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem") 
  For Each objItem in colItems
    ramAmount   = objItem.TotalPhysicalMemory
  Next
  ram.InnerHTML    = Int(ramAmount / 1048576) & " MB"

  '--- Get Dell Info ---------------
  Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct", "WQL", _ 
  wbemFlagReturnImmediately + wbemFlagForwardOnly) 
  For Each objItem In colItems 
    DellModel = objItem.Vendor&"&nbsp;"&objItem.Name 
  Next 
  model.innerhtml = DellModel

  Set colItems = objWMIService.ExecQuery("Select * from Win32_BIOS") 
  For Each objItem In colItems 
    dellServiceTag = objItem.SerialNumber
  Next 
  servicetag.innerhtml = "| ST: "&dellServiceTag

  '--- Get OS Info ---------------
  Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
  For Each objItem in colItems
    osVersion     = objItem.Caption & " ServicePack " & objItem.ServicePackMajorVersion
    strBootYear   = Left( objItem.LastBootUpTime, 4 )
    strBootMonth  = Mid( objItem.LastBootUpTime,  5, 2 )
    strBootDay    = Mid( objItem.LastBootUpTime,  7, 2 )
    strBootDate   = DateValue( strBootMonth & "-" & strBootDay & "-" & strBootYear )
    strBootHour   = Mid( objItem.LastBootUpTime,  9, 2 )
    strBootMins   = Mid( objItem.LastBootUpTime, 11, 2 )
    strBootTime   = strBootHour & ":" & strBootMins
    sysbootDate   = strBootDate & " at " & strBootTime
    strImageYear  = Left( objItem.InstallDate, 4 )
    strImageMonth = Mid( objItem.InstallDate,  5, 2 )
    strImageDay   = Mid( objItem.InstallDate,  7, 2 )
    strImageDate  = DateValue( strImageMonth & "-" & strImageDay & "-" & strImageYear )
    strImageHour  = Mid( objItem.InstallDate,  9, 2 )
    strImageMins  = Mid( objItem.InstallDate, 11, 2 )
    strImageTime  = strImageHour & ":" & strImageMins
    sysimageDate  = strImageDate & " at " & strImageTime
  Next 
  os.innerhtml        = osVersion
  bootDate.InnerHTML  = sysbootDate
  imageDate.innerhtml = sysimageDate

  '--- Get NS Lookup info ---------------
  strOut=""
  cmdarg="%comspec% /c nslookup.exe " & vDEVICEID
  set objExCmd = WshShell.Exec(cmdarg)
  strOut=objExCmd.StdOut.ReadAll
  Set regEx = New RegExp
    regEx.Pattern = "[\f\n\r\v]+"
    regEx.Global = True
    regEx.Multiline = True
  strOut = regEx.Replace(strOut, "<br>")
  ouValu = GetDN(vDEVICEID)
  ouvalu2 = Split(ouValu,",")
  For Each entry In ouvalu2
    ouvalu4 = Split(entry,"=")
    ouvalu5 = ouvalu5 & ouvalu4(1) & "$"
  Next
  ouvalu6 = Split(ouvalu5,"$")
  ouvalu3 = "OU: " & ouvalu6(1) & " | " & ouvalu6(2) & " | " & ouvalu6(3)
  'nsoutput.innerHTML= strOut
  ou.innerHTML = ouvalu3
  
  overviewid.innerhtml = WshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
End Sub

Sub thinClient
  DT.Style.Display = "none"
  TC.Style.Display = "block"
End Sub


'---- Buttons -------------------
Sub checkDevice
  If vDEVICEID = "" Then Exit Sub
End Sub

'---- Remote Desktop
Sub RDP
  checkDevice
  WshShell.Run "%SystemRoot%\system32\mstsc.exe /v:"&vDEVICEID&" /F /console"
End Sub

'---- SCCM
Sub SCCM
  checkDevice
  WshShell.Run ".\sccm\i386\CmRcViewer.exe "&vDEVICEID
End Sub 

'---- Remote Assistance
Sub RA 
  checkDevice
  WshShell.Run "%SystemRoot%\system32\msra.exe /offerra "&vDEVICEID
End Sub

'---- Browse C:
Sub BrowseC 
  checkDevice
  WshShell.Run "explorer \\"& vDEVICEID & "\c$"
End Sub 

'---- Delete Profiles
Sub DelProf 
  checkDevice
  WshShell.Run "cmd /K .\components\DelProf2.exe /c:"&vDEVICEID
End Sub 

'---- System Info
Sub SysInfo 
  checkDevice
  WshShell.Run "msinfo32 /computer " &vDEVICEID
End Sub

'---- Manage
Sub RemoteMMC
  checkDevice
  WshShell.Run "compmgmt.msc /computer=" &vDEVICEID
End Sub

'---- WDM Server
Sub WDMserver
  checkDevice
  WshShell.Run "%SystemRoot%\system32\mstsc.exe /v:ovwdm01 /F /console"
End Sub

'____ Renew IP
Sub RenewIP
  checkDevice
  WshShell.Run "cmd /K .\components\psexec \\"&vDEVICEID&" ipconfig /renew && exit"
End Sub

'____ IP Config
Sub IPconfigAll
  checkDevice
  WshShell.Run "cmd /K .\components\psexec \\"&vDEVICEID&" ipconfig /all"
End Sub

'____ Flush DNS
Sub FlushDNS
  checkDevice
  WshShell.Run "cmd /K .\components\psexec \\"&vDEVICEID&" ipconfig /flushdns && exit"
End Sub

'____ Continuous Ping
Sub ContinuousPing
  checkDevice
  WshShell.Run "cmd /K ping "&vDEVICEID&" -t"
End Sub
    
'---- Power Button
Sub Power(powerLabel, powerOption)
  Warning = MsgBox ("Are You Sure You Want to "&powerLabel&" "&vDEVICEID&" ?", 276, "Warning") 
  if Warning = 6 then 
  Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//"&vDEVICEID&"/root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")
  For Each OpSys In OpSysSet 
    OpSys.Win32Shutdown(powerOption) 
  Next 
  end if 
end Sub 