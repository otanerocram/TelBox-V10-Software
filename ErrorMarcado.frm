VERSION 5.00
Begin VB.Form frmErrorMarcado 
   BackColor       =   &H000000FF&
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3450
   ControlBox      =   0   'False
   Icon            =   "ErrorMarcado.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   3450
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   1320
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   120
      Picture         =   "ErrorMarcado.frx":000C
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   120
      Picture         =   "ErrorMarcado.frx":0D76
      Stretch         =   -1  'True
      Top             =   360
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   960
      Picture         =   "ErrorMarcado.frx":198E
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmErrorMarcado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" _
                (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub Form_Load()
    Label1.Caption = "La serie " & NumeroNoPermitido & " marcada en la cabina " _
                    & cabId & " no está registrada en el tarifario"
    Image1.Visible = True
    Image3.Visible = False
    Timer1.Enabled = True
    Timer1.Interval = 500
    PlaySound
End Sub

Private Sub Timer1_Timer()
    If Image1.Visible = True Then
        Image1.Visible = False
        Image3.Visible = True
    Else
        Image1.Visible = True
        Image3.Visible = False
    End If
End Sub

Private Sub Image2_Click()
    Unload Me
End Sub

Private Sub PlaySound()
    Dim destination As String
    Dim playWave As Variant
    
    destination = App.Path & "\warning.wav"
    'play sound file
    playWave = sndPlaySound(ByVal CStr(destination), 1)
End Sub
