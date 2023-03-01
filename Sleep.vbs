Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colOS = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOS in colOS
  msg = "Are you sure you want to put your computer to sleep?"
  title = "Sleep Computer"
  style = vbYesNo + vbQuestion
  response = MsgBox(msg, style, title)
  If response = vbYes Then
    psCmd = "powershell.exe -WindowStyle Hidden -Command ""Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Application]::SetSuspendState('Suspend', $false, $true)"""
    Set obiShell = CreateObject("WScript.Shell")
    obiShell.Run psCmd, 0, True
  End If
Next