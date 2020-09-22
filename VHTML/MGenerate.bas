Attribute VB_Name = "MGenerate"
'ce module genere le code...

Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const MousePress = &HA1
Public Const SizeN = 12
Public Const SizeS = 15
Public Const SizeW = 10
Public Const SizeE = 11
Public Const SizeNW = 13
Public Const SizeSW = 16
Public Const SizeNE = 14
Public Const SizeSE = 17



Dim I As Integer

Dim xName As String
Dim xValue As String
Dim xBgcol As String
Dim xFgcol As String
Dim xTitle As String
Dim xAlt As String
Dim xRo As String
Dim xBorder As String
Dim xPath As String
Dim xClick As String
Dim xDblclick As String
Dim xOver As String
Dim xOut As String
Dim xDown As String
Dim xMove As String
Dim xLoad As String
Dim xUnload As String
Dim xKeydown As String
Dim xKeyup As String
Dim xKeypress As String
Dim xSelect As String
Dim xFocus As String
Dim xChange As String
Dim xBlur As String
Dim xError  As String
Dim xAbord As String


Public Function CopyControl(Control As Variant, Visible As Boolean, Top As Integer, Left As Integer, Width As Integer, Height As Integer)
    Dim NewIndex As Integer
    NewIndex = Control.Count + 1
    Load Control(NewIndex)
    With Control(NewIndex)
        .Visible = Visible
        .Top = Top
        .Left = Left
        .Width = Width
        .Height = Height
    End With
End Function

'Function CopyControlWithResize, optional API
Public Function CopyControlWithResize(Control As Variant, Visible As Boolean, Resize As Boolean, Handles As Variant, Top As Integer, Left As Integer, Width As Integer, Height As Integer)
    Dim NewIndex As Integer
    NewIndex = Control.Count + 1
    Load Control(NewIndex)
    With Control(NewIndex)
        .Visible = Visible
        .Top = Top
        .Left = Left
        .Width = Width
        .Height = Height
    End With
    If Resize = True Then
    X = 0
    Do Until X = 8
    Handles(X).Visible = True
    X = X + 1
    Loop
    HandlesMove Control(NewIndex), Handles
    End If
End Function

'Function ControlResize, API
Public Function ControlResize(ControlWithAPIHandle As Control, Handles As Variant, Index As Variant)
On Error Resume Next
    ReleaseCapture
    Select Case Index
        Case 0
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeNW, 0
        Case 1
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeN, 0
        Case 2
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeNE, 0
        Case 7
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeE, 0
        Case 3
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeSE, 0
        Case 6
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeS, 0
        Case 5
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeSW, 0
        Case 4
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeW, 0
    End Select
    HandlesMove ControlWithAPIHandle, Handles
End Function

'Function HandlesMove, API
Public Function HandlesMove(ByVal Control As Control, Handles As Variant)
    Form1.rect(0).Left = Control.Left - Form1.rect(0).Width
    Form1.rect(0).Top = Control.Top - Form1.rect(0).Height
    Form1.rect(1).Left = (Control.Width - Form1.rect(7).Width) / 2 + Control.Left
    Form1.rect(1).Top = Control.Top - Form1.rect(1).Height
    Form1.rect(2).Left = Control.Left + Control.Width
    Form1.rect(2).Top = Control.Top - Form1.rect(0).Height
    Form1.rect(7).Left = Control.Left + Control.Width
    Form1.rect(7).Top = (Control.Height - Form1.rect(7).Height) / 2 + Control.Top
    Form1.rect(3).Left = Control.Left + Control.Width
    Form1.rect(3).Top = Control.Top + Control.Height
    Form1.rect(6).Left = (Control.Width - Form1.rect(6).Width) / 2 + Control.Left
    Form1.rect(6).Top = Control.Top + Control.Height
    Form1.rect(5).Left = Control.Left - Form1.rect(5).Width
    Form1.rect(5).Top = Control.Top + Control.Height
    Form1.rect(4).Left = Control.Left - Form1.rect(4).Width
    Form1.rect(4).Top = (Control.Height - Form1.rect(4).Height) / 2 + Control.Top
End Function


Public Sub Generate()
SetVariables "Page", "N/A"
Form1.RTCode.Text = ""
Form1.RTCode.Text = "<html>" & Form1.space.Text & _
"<head>" & "<title>" & xTitle & "</title>" & Form1.space.Text & _
"<script language='javascript' src='vhtml.js'></SCRIPT></head>" & "<body" & xFgcol & xBgcol & xClick & xDblclick & xDown & xOut & xOver & xKeypress & xKeydown & xKeyup & xLoad & xUnload & ">"
For I = 1 To Form1.cb.Count - 1
Add2Coding Form1.cb(I).Left, Form1.cb(I).Top, "cb", I, Form1.cb(I).Width, Form1.cb(I).Height
Next I
For I = 1 To Form1.cc.Count - 1
Add2Coding Form1.cc(I).Left, Form1.cc(I).Top, "cc", I, Form1.cc(I).Width, Form1.cc(I).Height
Next I
For I = 1 To Form1.ci.Count - 1
Add2Coding Form1.ci(I).Left, Form1.ci(I).Top, "ci", I, Form1.ci(I).Width, Form1.ci(I).Height
Next I
For I = 1 To Form1.cli.Count - 1
Add2Coding Form1.cli(I).Left, Form1.cli(I).Top, "cli", I, Form1.cli(I).Width, Form1.cli(I).Height
Next I
For I = 1 To Form1.clist.Count - 1
Add2Coding Form1.clist(I).Left, Form1.clist(I).Top, "clist", I, Form1.clist(I).Width, Form1.clist(I).Height
Next I
For I = 1 To Form1.ccombo.Count - 1
Add2Coding Form1.ccombo(I).Left, Form1.ccombo(I).Top, "ccombo", I, Form1.ccombo(I).Width, Form1.ccombo(I).Height
Next I
For I = 1 To Form1.ct.Count - 1
Add2Coding Form1.ct(I).Left, Form1.ct(I).Top, "ct", I, Form1.ct(I).Width, Form1.ct(I).Height
Next I
For I = 1 To Form1.cta.Count - 1
Add2Coding Form1.cta(I).Left, Form1.cta(I).Top, "cta", I, Form1.cta(I).Width, Form1.cta(I).Height
Next I
For I = 1 To Form1.cp.Count - 1
Add2Coding Form1.cp(I).Left, Form1.cp(I).Top, "cp", I, Form1.cp(I).Width, Form1.cp(I).Height
Next I
For I = 1 To Form1.ch.Count - 1
Add2Coding Form1.ch(I).Left, Form1.ch(I).Top, "ch", I, Form1.ch(I).Width, Form1.ch(I).Height
Next I
For I = 1 To Form1.cl.Count - 1
Add2Coding Form1.cl(I).X1, Form1.cl(I).Y1, "cl", I, (Form1.cl(I).X2 - Form1.cl(I).X1), 0
Next I
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"</body></html>"

Open "c:\preview.html" For Output As #1
Print #1, Form1.RTCode.Text
Close #1
Open "c:\vhtml.js" For Output As #1
Print #1, Form1.js.Text
Close #1

Form1.RTCode.Colorize
End Sub

Public Sub Add2Coding(X As Integer, Y As Integer, Wtype As String, Index As Integer, Width As Integer, Height As Integer)
SetVariables Wtype, Index
Select Case Wtype
Case "cb"
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<input type='button'" & xValue & xName & _
"' style='position:absolute;width:" & Width & ";height:" & Height & ";left:" & X & ";top:" & Y & _
";'" & xOver & xOut & xClick & xDblclick & xDown & ">"
Case "cc"
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<input type='CheckBox'" & xValue & xName & _
"' style='position:absolute;width:" & Width & ";height:" & Height & ";left:" & X & ";top:" & Y & ";'" & xOver & xOut & xClick & xDblclick & xDown & ">"
Case "ci"
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<img" & xPath & _
xName & _
xBorder & _
xAlt & _
" style='position:absolute;width:" & Width & ";height:" & Height & ";left:" & X & ";top:" & Y & ";'" & xOver & xOut & xClick & xDblclick & xDown & xAbord & xError & ">"
Case "cli"
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & Form1.space.Text & _
"<div style='position:absolute;left:" & X & ";top:" & Y & ";'" & xOver & xOut & xClick & xDblclick & xDown & ">" & Form1.space.Text & _
"<a" & xPath & _
xTitle & ">" & _
xValue & "</a>" & Form1.space.Text & "</div>"
Case "clist"
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<select name='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\clist" & Index, "name") & _
"' size='2' style='position:absolute;width:" & Width & ";height:" & Height & ";left:" & X & _
";top:" & Y & ";'" & xOver & xOut & xClick & xDblclick & xDown & ">" & Form1.space.Text & _
"<option value='1'>" & xValue & _
Form1.space.Text & "</select>"
Case "ccombo"
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<select" & xName & _
"' style='position:absolute;width:" & Width & ";left:" & X & ";top:" & Y & ";" & xOver & xOut & xClick & xDblclick & xDown & ">" & _
Form1.space.Text & "<option>" & _
xValue & _
"</select>"
Case "ct"
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<input type='text'" & xName & _
" style='position:absolute;width:" & Width & ";left:" & X & ";top:" & Y & ";'" & xValue & _
 xOver & xOut & xClick & xDblclick & xDown & xSelect & xChange & xFocus & xBlur & xRo & ">"
Case "cta"
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<textarea" & xName & _
" style='position:absolute;width:" & Width & ";height:" & Height & ";left:" & X & ";top:" & Y & ";'" & _
xOver & xOut & xClick & xDblclick & xDown & xSelect & xChange & xFocus & xBlur & xRo & ">" & Form1.cta(Index).Text & "</textarea>"
Case "cp"
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<input type='password'" & xName & _
"' style='position:absolute;width:" & Width & ";left:" & X & ";top:" & Y & ";'" & xValue & "' xOver & xOut & xClick & xDblclick & xDown & xSelect & xChange & xFocus & xBlur & xro &" > ""
Case "ch"
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<input type='hidden'" & xName & _
"' style='position:absolute;width:" & Width & ";left:" & X & ";top:" & Y & ";" & xValue & "'>"
Case "cl"
Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<hr style='position:absolute;width:" & Width & ";left:" & X & ";top:" & Y & ";'>"
Case "csub"

End Select
End Sub


Public Sub SetVariables(Wtype As String, Index)
If pro.nam.Text <> "" Then
xName = " name='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "name") & "'"
End If

If pro.valu.Text <> "" Then
    If Wtype = "cli" Or wlist = "clist" Then
    xValue = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "value")
    Else
    xValue = " value='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "value") & "'"
    End If
End If

If pro.bgcol.Text <> "" Then
xBgcol = " bgcolor='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "bgcolor") & "'"
End If

If pro.fgcol.Text <> "" Then
xFgcol = " text='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "fgcolor") & "'"
End If

If pro.title.Text <> "" Then
    If Wtype = "Page" Then
    xTitle = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "title")
    Else
    xTitle = " title='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "title") & "'"
    End If
End If

If pro.alt.Text <> "" Then
    xAlt = " alt='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "alt") & "'"
End If

If pro.RO.Value = 1 Or pro.RO.Value = 2 Then
xRo = " READONLY" 'getstring(HKEY_CURRENT_USER, "Software\VHTML\" & wType & index, "ro")
ElseIf pro.RO.Value = 0 Then
xRo = ""
End If

If pro.border.Text <> "" Then
xBorder = " border='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "border") & "'"
End If

If pro.path.Text <> "" Then
    If Wtype = "cli" Then
        xPath = " href='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "path") & "'"
    Else
        xPath = " src='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "path") & "'"
    End If
End If

If pro.onclick.Text <> "" Then
xClick = " onClick='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "click") & "'"
End If

If pro.ondblclick.Text <> "" Then
xDblclick = " onDblClick='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "dblclick") & "'"
End If

If pro.over.Text <> "" Then
xOver = " onMouseOver='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "over") & "'"
End If

If pro.out.Text <> "" Then
xOut = " onMouseOut='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "out") & "'"
End If

If pro.down.Text <> "" Then
xDown = " onMouseDown='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "down") & "'"
End If

If pro.mouve.Text <> "" Then
xMove = " onMouseMove='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "move") & "'"
End If

If pro.onload.Text <> "" Then
xLoad = " OnLoad='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "load") & "'"
End If

If pro.onunload.Text <> "" Then
xUnload = " onUnload='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "unload") & "'"
End If

If pro.onkeydown.Text <> "" Then
xKeydown = " onKeyDown='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "keydown") & "'"
End If

If pro.onkeyup.Text <> "" Then
xKeyup = " onKeyUp='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "keyup") & "'"
End If

If pro.onkeypress.Text <> "" Then
xKeypress = " onKeyPress='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "keypress") & "'"
End If

If pro.onselect.Text <> "" Then
xSelect = " onSelect='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "select") & "'"
End If

If pro.onfocus.Text <> "" Then
xFocus = " onFocus='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "focus") & "'"
End If

If pro.onchange.Text <> "" Then
xChange = " onChange='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "change") & "'"
End If

If pro.onblur.Text <> "" Then
xBlur = " onBlur='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "name") & "'"
End If

If pro.onerror.Text <> "" Then
xError = " onError='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "error") & "'"
End If

If pro.onabord.Text <> "" Then
xAbord = " onAbord='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "abord") & "'"
End If
End Sub



Public Sub SetVariablesBody()
If pro.bgcol.Text <> "" Then
xbgcolor = " bgcolor='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "bgcolor") & "'"
End If

If pro.fgcol.Text <> "" Then
xfgcolor = " text='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "fgcolor") & "'"
End If
End Sub
