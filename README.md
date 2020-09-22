<div align="center">

## Print Screen \(take a picture of the screen\!\)


</div>

### Description

Take a picture of the screen
 
### More Info
 
Found off the web


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brent w\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brent-w.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brent-w-print-screen-take-a-picture-of-the-screen__1-1333/archive/master.zip)

### API Declarations

```
Declare Sub keybd_event Lib "user32" _
(ByVal bVk As Byte, ByVal bScan As Byte, _
ByVal Flags As Long, ByVal ExtraInfo As Long)
```


### Source Code

```
Sub ScreenToClipboard()
Const VK_SNAPSHOT = &H2C
Call keybd_event(VK_SNAPSHOT, 1, 0&, 0&)
End Sub
```

