Attribute VB_Name = "modTransparence"
Option Explicit

' Module original : http://www.vbfrance.com/code.aspx?ID=24602


'''''D�claration des constantes en tant que globales
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const GWL_EXSTYLE = (-20)

'''''Apis n�cessaires pour la transparence
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'

' Utilisation dans une forme : le Slide = � de transparence
'
'Private Sub Command1_Click()
'    Transparence "ON", Me, HScroll1.Value
'End Sub
'
'Private Sub Command2_Click()
'    Transparence "OFF", Me
'End Sub
'
'Private Sub hscroll1_Scroll()
'    Transparence "ON", Me, HScroll1.Value   ' 0 � 255
'End Sub


Public Sub Transparence(State As String, Fen�tre As Form, Optional ByVal Alpha As Byte = 255)
    Dim Reference As Long
    Reference = GetWindowLong(Fen�tre.hWnd, GWL_EXSTYLE)
    
    Select Case UCase(State)
        Case "ON"
                SetWindowLong Fen�tre.hWnd, GWL_EXSTYLE, Reference Or WS_EX_LAYERED
                SetLayeredWindowAttributes Fen�tre.hWnd, 0, Alpha, LWA_ALPHA
        Case "OFF"
                SetWindowLong Fen�tre.hWnd, GWL_EXSTYLE, Reference - WS_EX_LAYERED
    End Select
End Sub
