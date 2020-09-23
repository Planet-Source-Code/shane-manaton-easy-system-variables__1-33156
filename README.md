<div align="center">

## Easy System Variables


</div>

### Description

Easy way to return system variables
 
### More Info
 
none, none ,none

your coding becomes a lot quicker!!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Shane Manaton](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shane-manaton.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shane-manaton-easy-system-variables__1-33156/archive/master.zip)

### API Declarations

None!!!


### Source Code

```
Private Sub Form_Load()
For i = 1 To 30
MsgBox Environ(i), vbOKOnly, i
Next i
End Sub
I have seen loads of complicated code that does the same
Try this code, every system variable you need in one line of code
Usage:
MsgBox Environ("ComputerName")
MsgBox Environ("Path")
Etc
```

