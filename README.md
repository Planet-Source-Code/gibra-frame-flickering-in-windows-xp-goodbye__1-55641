<div align="center">

## Frame 'flickering' in Windows XP Goodbye\!


</div>

### Description

This function change the CommandButton style to a BS_GROUPBOX to transform it like a Frame control! This solve many problems (bad drawning controls, flickering, etc.).
 
### More Info
 


When you use a Frame control on Windows XP where Themes is active (XP Style), there is some problems:

1) OptionButton is draw black

2) CommandButton is draw with black border

3) all controls contained on the frame 'flickering' when you move mouse over them.

This code show how to change a CommandButton style and use it like a Frame control.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[gibra](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gibra.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gibra-frame-flickering-in-windows-xp-goodbye__1-55641/archive/master.zip)

### API Declarations

```
'-----------------------------------------------
'Module BAS declarations:
'-----------------------------------------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
 ByVal hWnd As Long, _
 ByVal wMsg As Long, _
 ByVal wParam As Long, _
 lParam As Any) As Long
 Private Const BM_SETSTYLE As Long = &HF4
 Private Const BS_GROUPBOX As Long = &H7&
```


### Source Code

```
--------------------------------
Open a new project,
on Form1 add this controls:
--------------------------------
- 1 CommandButton (Command1)
- 1 OptionButton (Option1)
- 1 TextBox (Text1)
- Set Form1.ClipControls = False
--------------------------------
Add this code to Form1:
--------------------------------
Private Sub Form_Load()
 With Command1
 .Caption = "Pseudo Frame"
 .Left = 300
 .Top = 300
 .ZOrder 1
 End With
 With Text1
 .Height = 330
 .Left = 510
 .Top = 600
 End With
 With Option1
 .Height = 330
 .Left = 510
 .Top = 1050
 End With
 ChangeButtonStyle Command1, Me, 300, 300, 1800, 1800
End Sub
--------------------------------
Add to Module this code
--------------------------------
Public Sub ChangeButtonStyle(ByRef cmd As CommandButton, _
 ByVal Parent As Object, _
 Optional Left As Long = 0, _
 Optional Top As Long = 0, _
 Optional Width As Long = 0, _
 Optional Height As Long = 0)
 '/ Show a CommandButton like a Frame control.
 '/ Also, set the backcolor text to the
 '/ background color
 On Error Resume Next
 '/ Change the CommandButton style
 SendMessage cmd.hWnd, BM_SETSTYLE, BS_GROUPBOX, 0
 '/ Set the backcolor text to emulate the
 '/ transparent background
 cmd.BackColor = cmd.Container.BackColor
 '/ IMPORTANT: Set the TabStop property to false
 '/ otherwise when lost the focus pressing the TAB key
 '/ the style is changed to CheckBox
 '/ Also, the focus state don't need
 '/ with this pseudo-frame.
 cmd.TabStop = False
 '/ Optionally, you can move and size the
 '/ commandbutton (i.e. if you use a
 '/ PictureBox as Container:
 cmd.Move Left, Top, IIf(Width = 0, Parent.Width, Width), _
 IIf(Height = 0, Parent.Height, Height)
 '/ IMPORTANT: This property MUST to be
 '/ set on Design-Time, otherwise
 '/ has not effect!
 '/ --------------------------------------
 '/ Parent.ClipControls = False
 '/ --------------------------------------
End Sub
```

