Attribute VB_Name = "OnTop"
Option Explicit

'API nécessaire pour le mode "toujours visible"
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
'toujours visible
Public Function forward(who As Form) 'who correspond au nom de la form  | exemple: form1
Dim Resultat As Long
Const Flags = &H2 Or &H1 Or &H40 Or &H10
Resultat = SetWindowPos(who.hWnd, -1, 0, 0, 0, 0, Flags)
End Function

'annuler toujours visible
Public Function backward(who As Form)
Dim Resultat As Long
Const Flags = &H2 Or &H1 Or &H40 Or &H10
Resultat = SetWindowPos(who.hWnd, -2, 0, 0, 0, 0, Flags)
End Function

'execution d'un lien...
Public Function WeB(WebPage As String, actualfrmHWND As String)
On Error Resume Next
Dim cod
cod = ShellExecute(actualfrmHWND, vbNullString, WebPage, "", vbNullString, 1)
End Function

'restart...
Public Sub reload(frm As Form)
Unload frm
Load frm
frm.Show
End Sub

'verifier l'existence d'un fichier...
Public Function ExistFile(strPath As String) As Boolean
  Dim fs As Object
  Dim blnFExiste As Boolean
  Set fs = CreateObject("Scripting.FileSystemObject")
  If Not (fs.FileExists(strPath)) Then
    blnFExiste = False
  Else
    blnFExiste = True
  End If
  ExistFile = blnFExiste
End Function
 


