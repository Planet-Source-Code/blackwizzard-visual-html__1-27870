VERSION 5.00
Begin VB.MDIForm pro 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Visual HTML Dev"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14430
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5400
      Top             =   7800
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H0080FFFF&
      Height          =   10440
      Left            =   0
      ScaleHeight     =   10380
      ScaleWidth      =   2685
      TabIndex        =   3
      Top             =   0
      Width           =   2745
      Begin VB.PictureBox Picture4 
         BackColor       =   &H0080FFFF&
         Height          =   7920
         Left            =   0
         ScaleHeight     =   7860
         ScaleWidth      =   2700
         TabIndex        =   4
         Top             =   0
         Width           =   2760
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   68
            Text            =   "Top"
            Top             =   6720
            Width           =   1215
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   67
            Text            =   "Left"
            Top             =   6960
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   66
            Text            =   "Width"
            Top             =   6240
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   65
            Text            =   "Height"
            Top             =   6480
            Width           =   1215
         End
         Begin VB.TextBox zonabord 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   64
            Text            =   "onAbord"
            Top             =   6000
            Width           =   1215
         End
         Begin VB.TextBox zonerror 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   63
            Text            =   "onError"
            Top             =   5760
            Width           =   1215
         End
         Begin VB.TextBox zonblur 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   62
            Text            =   "onBlur"
            Top             =   5520
            Width           =   1215
         End
         Begin VB.TextBox zonchange 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   61
            Text            =   "onChange"
            Top             =   5280
            Width           =   1215
         End
         Begin VB.TextBox zonfocus 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   60
            Text            =   "onFocus"
            Top             =   5040
            Width           =   1215
         End
         Begin VB.TextBox zonselect 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   59
            Text            =   "onSelect"
            Top             =   4800
            Width           =   1215
         End
         Begin VB.TextBox zonkeypress 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   58
            Text            =   "onKeypress"
            Top             =   4560
            Width           =   1215
         End
         Begin VB.TextBox zonkeyup 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   57
            Text            =   "onKeyup"
            Top             =   4320
            Width           =   1215
         End
         Begin VB.TextBox zonkeydown 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "onKeydown"
            Top             =   4080
            Width           =   1215
         End
         Begin VB.TextBox zonunload 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   55
            Text            =   "onUnload"
            Top             =   3840
            Width           =   1215
         End
         Begin VB.TextBox zonload 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "onLoad"
            Top             =   3600
            Width           =   1215
         End
         Begin VB.TextBox zmouve 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   53
            Text            =   "onMouseMove"
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox zover 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "onMouseOver"
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox zout 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "onMouseOut"
            Top             =   2880
            Width           =   1215
         End
         Begin VB.TextBox zdown 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "onMouseDown"
            Top             =   3120
            Width           =   1215
         End
         Begin VB.TextBox zondblclick 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "onDblClick"
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox zclick 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "onClick"
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox zpath 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   47
            Text            =   "Path"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CheckBox RO 
            BackColor       =   &H0080FFFF&
            Height          =   200
            Left            =   1800
            TabIndex        =   46
            Top             =   1490
            Width           =   255
         End
         Begin VB.TextBox zborder 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "Border"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox zreadonly 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "ReadOnly"
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox zalt 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "Alt"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox ztitle 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "Title"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox zfgcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "FGColor"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox ztype 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "Type"
            Top             =   7320
            Width           =   1215
         End
         Begin VB.TextBox zitem 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "Item"
            Top             =   7560
            Width           =   1215
         End
         Begin VB.TextBox zbgcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "BGColor"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox zvalue 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "Value"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox zname 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   "Name"
            Top             =   0
            Width           =   1215
         End
         Begin VB.TextBox Cwidth 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   35
            Top             =   6240
            Width           =   1455
         End
         Begin VB.TextBox Cheight 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   34
            Top             =   6480
            Width           =   1455
         End
         Begin VB.TextBox Ctop 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   33
            Top             =   6720
            Width           =   1455
         End
         Begin VB.TextBox Cleft 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   32
            Top             =   6960
            Width           =   1455
         End
         Begin VB.ComboBox onabord 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   31
            Top             =   6000
            Width           =   1455
         End
         Begin VB.ComboBox onerror 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   30
            Top             =   5760
            Width           =   1455
         End
         Begin VB.ComboBox onblur 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   29
            Top             =   5520
            Width           =   1455
         End
         Begin VB.ComboBox onchange 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   28
            Top             =   5280
            Width           =   1455
         End
         Begin VB.ComboBox onfocus 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   27
            Top             =   5040
            Width           =   1455
         End
         Begin VB.ComboBox onselect 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   26
            Top             =   4800
            Width           =   1455
         End
         Begin VB.ComboBox onkeypress 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   25
            Top             =   4560
            Width           =   1455
         End
         Begin VB.ComboBox onkeyup 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   24
            Top             =   4320
            Width           =   1455
         End
         Begin VB.ComboBox onkeydown 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   23
            Top             =   4080
            Width           =   1455
         End
         Begin VB.ComboBox onunload 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   22
            Top             =   3840
            Width           =   1455
         End
         Begin VB.ComboBox onload 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   21
            Top             =   3600
            Width           =   1455
         End
         Begin VB.ComboBox mouve 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   20
            Top             =   3360
            Width           =   1455
         End
         Begin VB.ComboBox down 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   19
            Top             =   3120
            Width           =   1455
         End
         Begin VB.ComboBox out 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   18
            Top             =   2880
            Width           =   1455
         End
         Begin VB.ComboBox over 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   17
            Top             =   2640
            Width           =   1455
         End
         Begin VB.ComboBox ondblclick 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   16
            Top             =   2400
            Width           =   1455
         End
         Begin VB.ComboBox onclick 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   15
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox path 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   14
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox border 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   13
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox alt 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   12
            Top             =   1185
            Width           =   1455
         End
         Begin VB.TextBox title 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox fgcol 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   10
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox bgcol 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   9
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox valu 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox nam 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   7
            Top             =   0
            Width           =   1455
         End
         Begin VB.TextBox Wtype 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   7320
            Width           =   1455
         End
         Begin VB.TextBox Witem 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   7560
            Width           =   1455
         End
         Begin VB.Line Line2 
            X1              =   1800
            X2              =   2160
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line1 
            X1              =   2640
            X2              =   2640
            Y1              =   1320
            Y2              =   1800
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10440
      Left            =   2745
      ScaleHeight     =   10440
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   0
      Width           =   240
      Begin VB.VScrollBar scr 
         Height          =   8415
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FF8080&
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   14370
      TabIndex        =   0
      Top             =   10440
      Width           =   14430
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visual HTML Dev Par BlackWizzard"
         Height          =   195
         Left            =   14160
         TabIndex        =   69
         Top             =   20
         Width           =   2565
      End
   End
End
Attribute VB_Name = "pro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub alt_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "alt", alt.Text)
Call ApplyVal
End Sub

Private Sub bgcol_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "bgcolor", bgcol.Text)
Call ApplyVal
End Sub

Private Sub border_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "border", border.Text)
Call ApplyVal
End Sub

Private Sub Cheight_Change()
On Error Resume Next
Select Case Wtype.Text
Case "cb"
Form1.cb(Witem.Text).Height = Cheight.Text
Case "cc"
'non modifiable...
Case "ct"
Form1.ct(Witem.Text).Height = Cheight.Text
Case "ci"
Form1.ci(Witem.Text).Height = Cheight.Text
Case "cp"
Form1.cp(Witem.Text).Height = Cheight.Text
Case "cta"
Form1.cta(Witem.Text).Height = Cheight.Text
Case "ch"
Form1.ch(Witem.Text).Height = Cheight.Text
Case "clist"
Form1.clist(Witem.Text).Height = Cheight.Text
Case "ccombo"
Form1.ccombo(Witem.Text).Height = Cheight.Text
Case "cli"
Form1.cli(Witem.Text).Height = Cheight.Text
Case "Page"
'Form1.zone.height = Cheight.Text
End Select
End Sub

Private Sub Cleft_Change()
On Error Resume Next
Select Case Wtype.Text
Case "cb"
Form1.cb(Witem.Text).Left = Cleft.Text
Case "cc"
'non modifiable...
Case "ct"
Form1.ct(Witem.Text).Left = Cleft.Text
Case "ci"
Form1.ci(Witem.Text).Left = Cleft.Text
Case "cp"
Form1.cp(Witem.Text).Left = Cleft.Text
Case "cta"
Form1.cta(Witem.Text).Left = Cleft.Text
Case "ch"
Form1.ch(Witem.Text).Left = Cleft.Text
Case "clist"
Form1.clist(Witem.Text).Left = Cleft.Text
Case "ccombo"
Form1.ccombo(Witem.Text).Left = Cleft.Text
Case "cli"
Form1.cli(Witem.Text).Left = Cleft.Text
Case "Page"
'Form1.zone.Left = Cleft.Text
End Select
End Sub

Private Sub Ctop_Change()
On Error Resume Next
Select Case Wtype.Text
Case "cb"
Form1.cb(Witem.Text).Top = Ctop.Text
Case "cc"
'non modifiable...
Case "ct"
Form1.ct(Witem.Text).Top = Ctop.Text
Case "ci"
Form1.ci(Witem.Text).Top = Ctop.Text
Case "cp"
Form1.cp(Witem.Text).Top = Ctop.Text
Case "cta"
Form1.cta(Witem.Text).Top = Ctop.Text
Case "ch"
Form1.ch(Witem.Text).Top = Ctop.Text
Case "clist"
Form1.clist(Witem.Text).Top = Ctop.Text
Case "ccombo"
Form1.ccombo(Witem.Text).Top = Ctop.Text
Case "cli"
Form1.cli(Witem.Text).Top = Ctop.Text
Case "Page"
'Form1.zone.Top = Ctop.Text
End Select
End Sub

Private Sub Cwidth_Change()
On Error Resume Next
Select Case Wtype.Text
Case "cb"
Form1.cb(Witem.Text).Width = Cwidth.Text
Case "cc"
'non modifiable...
Case "ct"
Form1.ct(Witem.Text).Width = Cwidth.Text
Case "ci"
Form1.ci(Witem.Text).Width = Cwidth.Text
Case "cp"
Form1.cp(Witem.Text).Width = Cwidth.Text
Case "cta"
Form1.cta(Witem.Text).Width = Cwidth.Text
Case "ch"
Form1.ch(Witem.Text).Width = Cwidth.Text
Case "clist"
Form1.clist(Witem.Text).Width = Cwidth.Text
Case "ccombo"
Form1.ccombo(Witem.Text).Width = Cwidth.Text
Case "cli"
Form1.cli(Witem.Text).Width = Cwidth.Text
Case "Page"
'Form1.zone.width = Cwidth.Text
End Select
End Sub

Private Sub down_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "down", down.Text)
Call ApplyVal
End Sub

Private Sub fgcol_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "fgcolor", fgcol.Text)
Call ApplyVal
End Sub

Private Sub Form_Load()
forward Me
End Sub

Private Sub MDIForm_Resize()
scr.Height = Picture2.Top
End Sub

Private Sub mouve_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "move", mouve.Text)
Call ApplyVal
End Sub

Private Sub nam_Change()
On Error Resume Next
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "name", nam.Text)
Call ApplyVal
End Sub

Private Sub onabord_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "abord", onabord.Text)
Call ApplyVal
End Sub

Private Sub onblur_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "blur", onblur.Text)
Call ApplyVal
End Sub

Private Sub onchange_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "change", onchange.Text)
Call ApplyVal
End Sub

Private Sub onclick_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "click", onclick.Text)
Call ApplyVal
End Sub

Private Sub ondblclick_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "dblclick", ondblclick.Text)
Call ApplyVal
End Sub

Private Sub onerror_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "error", onerror.Text)
Call ApplyVal
End Sub

Private Sub onfocus_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "focus", onfocus.Text)
Call ApplyVal
End Sub

Private Sub onkeydown_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "keydown", onkeydown.Text)
Call ApplyVal
End Sub

Private Sub onkeypress_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "keypress", onkeypress.Text)
Call ApplyVal
End Sub

Private Sub onkeyup_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "keyup", onkeyup.Text)
Call ApplyVal
End Sub

Private Sub onload_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "load", onload.Text)
Call ApplyVal
End Sub

Private Sub onselect_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "select", onselect.Text)
Call ApplyVal
End Sub

Private Sub onunload_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "unload", onunload.Text)
Call ApplyVal
End Sub

Private Sub out_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "out", out.Text)
Call ApplyVal
End Sub

Private Sub over_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "over", over.Text)
Call ApplyVal
End Sub

Private Sub path_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "path", path.Text)
Call ApplyVal
End Sub

Private Sub RO_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "ro", RO.Value)
Call ApplyVal
End Sub


Private Sub scr_Change()
scr.Height = Picture2.Top
scr.Max = Picture4.Height - Picture2.Top
Picture4.Top = 0 - scr.Value
End Sub

Private Sub Timer1_Timer()
If Label1.Left < 0 - Label1.Width Then
Label1.Left = Me.Width
Else
Label1.Left = Label1.Left - 50
End If
End Sub

Private Sub title_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "title", title.Text)
Call ApplyVal
End Sub

Private Sub valu_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "value", valu.Text)
Call ApplyVal
End Sub

Public Sub ApplyVal()
Dim fileimg
Select Case Wtype
Case "cb"
Form1.cb(Witem.Text).Caption = valu.Text
Form1.cb(Witem.Text).ToolTipText = nam.Text
Case "cc"
Form1.cc(Witem.Text).ToolTipText = nam.Text
Case "ci"
Form1.ci(Witem.Text).ToolTipText = alt.Text & " - (" & nam.Text & ")"
If ExistFile(Form1.cd1.FileName & "\" & path.Text) = True Then
fileimg = path.Text
Form1.ci(Witem.Text).Picture = fileimg
End If
Case "cli"
Form1.cli(Witem.Text).Caption = valu.Text
Form1.cli(Witem.Text).ToolTipText = "URL: " & path.Text
Case "clist"
Form1.clist(Witem.Text).ToolTipText = nam.Text
Form1.clist(Witem.Text).Text = valu.Text
Case "ccombo"
Form1.ccombo(Witem.Text).ToolTipText = nam.Text
Form1.ccombo(Witem.Text).Text = valu.Text
Case "ct"
Form1.ct(Witem.Text).ToolTipText = nam.Text
Form1.ct(Witem.Text).Text = valu.Text
Case "cta"
Form1.cta(Witem.Text).ToolTipText = nam.Text
Form1.cta(Witem.Text).Text = valu.Text
Case "cp"
Form1.cp(Witem.Text).ToolTipText = nam.Text
Form1.cp(Witem.Text).Text = valu.Text
Case "ch"
Form1.ct(Witem.Text).ToolTipText = nam.Text
Form1.ct(Witem.Text).Text = valu.Text
Case "cl"

Case "csub"

Case "Page"
If Len(bgcol.Text) = 7 Then
Form1.zone.BackColor = InvHex(Right(UCase(bgcol.Text), 6))
End If
End Select
End Sub

Public Function InvHex(ValHex As String) As String
  InvHex = val("&H00" & Right(ValHex, 2) & Mid(ValHex, 3, 2) & Left(ValHex, 2) & "&")
End Function



