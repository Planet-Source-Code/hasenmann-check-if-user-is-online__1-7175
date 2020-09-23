<div align="center">

## Check if user is Online


</div>

### Description

The code checks if the User is online or not. If he's online it will prompt a MsgBox "You are connected to the net.") or it will Show: "You are NOT connected to the net.". You can also insert your code. This code is useful when your application uses the internet. So if the user is not connected the net, the program will simply unload.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hasenmann](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hasenmann.md)
**Level**          |Intermediate
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hasenmann-check-if-user-is-online__1-7175/archive/master.zip)





### Source Code

```
Private Type RASCONN
  dwSize As Long
  hRasConn As Long
  szEntryName(256) As Byte
  szDeviceType(16) As Byte
  szDeviceName(128) As Byte
End Type
Private Declare Function RasEnumConnectionsA& Lib "RasApi32.DLL" (lprasconn As Any, lpcb&, lpcConnections&)
Private Sub Command1_Click()
Dim Verbindung As RASCONN
Dim size, Anz As Long
 Verbindung.dwSize = 412
 size = Verbindung.dwSize
 If RasEnumConnectionsA(Verbindung, size, Anz) = 0 Then
  If Anz = 0 Then
  MsgBox ("You are NOT connected to the net.")
  Else
  MsgBox ("You are connected to the net.")
  End If
 End If
End Sub
```

