' VBScript to display computer details with HTML formatting in the default browser
' Features:
'   - An H1 title ("Computer Details")
'   - A clickable link for the URL
'   - Bold labels (items) and unbolded values, all aligned left
'   - Enhanced Network Details (Default Gateway, Subnet, Ethernet Adapter Name)
'   - Three buttons at the bottom:
'       * Copy to Clipboard
'       * Email
'       * Save as Text File
' Created by Stephen Henry | Modified by ChatGPT
' https://github.com/JackInSightsV2/

On Error Resume Next

' Define WMI objects and other variables
Dim getWMI_obj, getCompSettings, getOSSettings, getComputer
Dim dtmBootup, dtmLastBootUpTime, uptimeSeconds, uptimeDays, uptimeHours, uptimeMinutes
Dim IPConfigSet, colSettingsCPU, colSettingsGPU, colSettingsVDU
Dim ipNum, monNum
Dim fso, outputFile, tempFolder, outputPath, WshShell
Dim htmlOutput

' Connect to WMI
Set getWMI_obj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

' Define queries
Set getCompSettings = getWMI_obj.ExecQuery("Select * from Win32_ComputerSystem")
Set getOSSettings   = getWMI_obj.ExecQuery("Select * from Win32_OperatingSystem")
' Updated query to include needed properties for network details
Set IPConfigSet     = getWMI_obj.ExecQuery("Select IPAddress, DefaultIPGateway, IPSubnet, Description from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
Set colSettingsCPU  = getWMI_obj.ExecQuery("Select Name from Win32_Processor")
Set colSettingsGPU  = getWMI_obj.ExecQuery("Select Caption, Description, DeviceName from Win32_DisplayConfiguration")
Set colSettingsVDU  = getWMI_obj.ExecQuery("Select Caption, Description, DeviceID, DisplayType, MonitorManufacturer, MonitorType, Name, PNPDeviceID, ScreenHeight, ScreenWidth from Win32_DesktopMonitor")

' Begin building the HTML output
htmlOutput = "<html><head><title>Computer Details</title></head>"
htmlOutput = htmlOutput & "<body style='font-family: Courier New, monospace;'>"
htmlOutput = htmlOutput & "<h1>Computer Details</h1>"
htmlOutput = htmlOutput & "<p>Created by <a href='https://github.com/JackInSightsV2/'>Stephen Henry</a></p>"

' --- System Information ---
htmlOutput = htmlOutput & "<h2>System Information</h2>"
htmlOutput = htmlOutput & "<table border='0' cellspacing='0' cellpadding='4'>"
For Each objItem In getCompSettings
    htmlOutput = htmlOutput & AddRow("Computer Name:", objItem.Name)
    htmlOutput = htmlOutput & AddRow("User Name:", objItem.UserName)
Next
htmlOutput = htmlOutput & "</table>"

' --- Calculate Uptime ---
For Each objOS In getOSSettings
    dtmBootup = objOS.LastBootUpTime
    dtmLastBootUpTime = WMIDateStringToDate(dtmBootup)
    uptimeSeconds = DateDiff("s", dtmLastBootUpTime, Now)
    uptimeDays = Int(uptimeSeconds / 86400)
    uptimeHours = Int((uptimeSeconds Mod 86400) / 3600)
    uptimeMinutes = Int((uptimeSeconds Mod 3600) / 60)
Next

htmlOutput = htmlOutput & "<h2>Uptime</h2>"
htmlOutput = htmlOutput & "<table border='0' cellspacing='0' cellpadding='4'>"
' The uptime line is updated to show "Days", "Hours", and "Minutes"
htmlOutput = htmlOutput & AddRow("Uptime:", uptimeDays & " Days, " & uptimeHours & " Hours, " & uptimeMinutes & " Minutes")
htmlOutput = htmlOutput & "</table>"

' --- System Details ---
htmlOutput = htmlOutput & "<h2>System Details</h2>"
htmlOutput = htmlOutput & "<table border='0' cellspacing='0' cellpadding='4'>"
For Each getComputer In getCompSettings
    htmlOutput = htmlOutput & AddRow("Manufacturer:", getComputer.Manufacturer)
    htmlOutput = htmlOutput & AddRow("Model:", getComputer.Model)
    htmlOutput = htmlOutput & AddRow("Memory:", Round(getComputer.TotalPhysicalMemory / 1024 / 1024, 0) & " MB")
Next
htmlOutput = htmlOutput & "</table>"

' --- Processor Details ---
htmlOutput = htmlOutput & "<h2>Processor Details</h2>"
htmlOutput = htmlOutput & "<table border='0' cellspacing='0' cellpadding='4'>"
For Each ObjCPU In colSettingsCPU
    htmlOutput = htmlOutput & AddRow("Processor Name:", ObjCPU.Name)
Next
htmlOutput = htmlOutput & "</table>"

' --- Operating System Details ---
htmlOutput = htmlOutput & "<h2>Operating System Details</h2>"
htmlOutput = htmlOutput & "<table border='0' cellspacing='0' cellpadding='4'>"
For Each getComputer In getOSSettings
    htmlOutput = htmlOutput & AddRow("OS Version:", getComputer.Caption & " " & getComputer.CSDVersion)
    htmlOutput = htmlOutput & AddRow("Version:", getComputer.Version)
    htmlOutput = htmlOutput & AddRow("Install Date:", WMIDateStringToDate(getComputer.InstallDate))
    htmlOutput = htmlOutput & AddRow("Windows Folder:", getComputer.WindowsDirectory)
Next
htmlOutput = htmlOutput & "</table>"

' --- Network Details ---
htmlOutput = htmlOutput & "<h2>Network Details</h2>"
ipNum = 1
For Each IPConfig In IPConfigSet
    If Not IsNull(IPConfig.IPAddress) Then
        htmlOutput = htmlOutput & "<h3>Adapter (" & ipNum & ")</h3>"
        htmlOutput = htmlOutput & "<table border='0' cellspacing='0' cellpadding='4'>"
        ' Ethernet Adapter Name (Description property)
        htmlOutput = htmlOutput & AddRow("Ethernet Adapter:", IPConfig.Description)
        
        Dim ipAddress, ipSubnet, defaultGateway
        If IsArray(IPConfig.IPAddress) Then
            ipAddress = Join(IPConfig.IPAddress, ", ")
        Else
            ipAddress = IPConfig.IPAddress
        End If
        If IsArray(IPConfig.IPSubnet) Then
            ipSubnet = Join(IPConfig.IPSubnet, ", ")
        Else
            ipSubnet = IPConfig.IPSubnet
        End If
        If IsArray(IPConfig.DefaultIPGateway) Then
            defaultGateway = Join(IPConfig.DefaultIPGateway, ", ")
        Else
            defaultGateway = IPConfig.DefaultIPGateway
        End If
        
        htmlOutput = htmlOutput & AddRow("IP Address:", ipAddress)
        htmlOutput = htmlOutput & AddRow("Subnet:", ipSubnet)
        htmlOutput = htmlOutput & AddRow("Default Gateway:", defaultGateway)
        
        htmlOutput = htmlOutput & "</table>"
        ipNum = ipNum + 1
    End If
Next

' --- Graphics Card Details ---
htmlOutput = htmlOutput & "<h2>Graphics Card Details</h2>"
htmlOutput = htmlOutput & "<table border='0' cellspacing='0' cellpadding='4'>"
For Each ObjGPU In colSettingsGPU
    htmlOutput = htmlOutput & AddRow("Graphics Card:", ObjGPU.Description)
Next
htmlOutput = htmlOutput & "</table>"

' --- Monitor Details ---
htmlOutput = htmlOutput & "<h2>Monitor Details</h2>"
monNum = 1
For Each ObjVDU In colSettingsVDU
    If ObjVDU.MonitorManufacturer <> "" Then
        htmlOutput = htmlOutput & "<h3>Monitor (" & monNum & ")</h3>"
        htmlOutput = htmlOutput & "<table border='0' cellspacing='0' cellpadding='4'>"
        htmlOutput = htmlOutput & AddRow("Make:", ObjVDU.MonitorManufacturer)
        htmlOutput = htmlOutput & AddRow("Description:", ObjVDU.Description)
        htmlOutput = htmlOutput & AddRow("Resolution:", ObjVDU.ScreenWidth & " x " & ObjVDU.ScreenHeight)
        htmlOutput = htmlOutput & "</table>"
        monNum = monNum + 1
    End If
Next

' --- Add Buttons at the Bottom ---
htmlOutput = htmlOutput & "<div style='margin-top:20px;'>"
htmlOutput = htmlOutput & "<button onclick='copyToClipboard()'>Copy to Clipboard</button> "
htmlOutput = htmlOutput & "<button onclick='emailContent()'>Email</button> "
htmlOutput = htmlOutput & "<button onclick='saveAsTextFile()'>Save as Text File</button>"
htmlOutput = htmlOutput & "</div>"

' --- Add JavaScript Functions ---
htmlOutput = htmlOutput & "<script>"
htmlOutput = htmlOutput & "function copyToClipboard() {"
htmlOutput = htmlOutput & "  var text = document.body.innerText;"
htmlOutput = htmlOutput & "  if(navigator.clipboard && window.isSecureContext) {"
htmlOutput = htmlOutput & "    navigator.clipboard.writeText(text).then(function() { alert('Copied to clipboard!'); }, function(err) { alert('Failed to copy: ' + err); });"
htmlOutput = htmlOutput & "  } else {"
htmlOutput = htmlOutput & "    var textArea = document.createElement('textarea');"
htmlOutput = htmlOutput & "    textArea.value = text;"
htmlOutput = htmlOutput & "    textArea.style.position = 'fixed'; textArea.style.top = 0; textArea.style.left = 0; textArea.style.width = '2em'; textArea.style.height = '2em';"
htmlOutput = htmlOutput & "    textArea.style.padding = 0; textArea.style.border = 'none'; textArea.style.outline = 'none';"
htmlOutput = htmlOutput & "    textArea.style.boxShadow = 'none'; textArea.style.background = 'transparent';"
htmlOutput = htmlOutput & "    document.body.appendChild(textArea);"
htmlOutput = htmlOutput & "    textArea.focus(); textArea.select();"
htmlOutput = htmlOutput & "    try {"
htmlOutput = htmlOutput & "      var successful = document.execCommand('copy');"
htmlOutput = htmlOutput & "      alert(successful ? 'Copied to clipboard!' : 'Unable to copy');"
htmlOutput = htmlOutput & "    } catch (err) { alert('Unable to copy'); }"
htmlOutput = htmlOutput & "    document.body.removeChild(textArea);"
htmlOutput = htmlOutput & "  }"
htmlOutput = htmlOutput & "}"
htmlOutput = htmlOutput & "function emailContent() {"
htmlOutput = htmlOutput & "  var subject = 'Computer Details';"
htmlOutput = htmlOutput & "  var body = document.body.innerText;"
htmlOutput = htmlOutput & "  window.location.href = 'mailto:?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(body);"
htmlOutput = htmlOutput & "}"
htmlOutput = htmlOutput & "function saveAsTextFile() {"
htmlOutput = htmlOutput & "  var text = document.body.innerText;"
htmlOutput = htmlOutput & "  var filename = 'ComputerDetails.txt';"
htmlOutput = htmlOutput & "  var blob = new Blob([text], {type: 'text/plain;charset=utf-8'});"
htmlOutput = htmlOutput & "  if (window.navigator.msSaveOrOpenBlob) {"
htmlOutput = htmlOutput & "    window.navigator.msSaveOrOpenBlob(blob, filename);"
htmlOutput = htmlOutput & "  } else {"
htmlOutput = htmlOutput & "    var a = document.createElement('a');"
htmlOutput = htmlOutput & "    var url = URL.createObjectURL(blob);"
htmlOutput = htmlOutput & "    a.href = url; a.download = filename;"
htmlOutput = htmlOutput & "    document.body.appendChild(a);"
htmlOutput = htmlOutput & "    a.click();"
htmlOutput = htmlOutput & "    setTimeout(function() { document.body.removeChild(a); window.URL.revokeObjectURL(url); }, 0);"
htmlOutput = htmlOutput & "  }"
htmlOutput = htmlOutput & "}"
htmlOutput = htmlOutput & "</script>"

htmlOutput = htmlOutput & "</body></html>"

' Write the HTML output to a temporary file and open it in the default browser
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
tempFolder = WshShell.ExpandEnvironmentStrings("%TEMP%")
outputPath = tempFolder & "\ComputerDetails.html"
Set outputFile = fso.CreateTextFile(outputPath, True)
outputFile.WriteLine htmlOutput
outputFile.Close

' Launch the HTML file in the default browser
WshShell.Run "cmd /c start """" """ & outputPath & """", 1, False

' --- Function to Add a Table Row with Bold Label, aligned left ---
Function AddRow(label, value)
    ' Returns a table row with the label in bold and both label and value left-aligned.
    AddRow = "<tr><td style='text-align:left; padding-right:10px;'><b>" & label & "</b></td>" & _
             "<td style='text-align:left;'>" & value & "</td></tr>"
End Function

' --- Function to Convert a WMI Date String to a Readable Date ---
Function WMIDateStringToDate(dtm)
    WMIDateStringToDate = DateSerial(CInt(Left(dtm, 4)), CInt(Mid(dtm, 5, 2)), CInt(Mid(dtm, 7, 2))) + _
                           TimeSerial(CInt(Mid(dtm, 9, 2)), CInt(Mid(dtm, 11, 2)), CInt(Mid(dtm, 13, 2)))
End Function
