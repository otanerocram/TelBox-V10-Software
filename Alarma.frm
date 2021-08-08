VERSION 5.00
Begin VB.Form frmAlarma 
   BackColor       =   &H000000FF&
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3450
   ControlBox      =   0   'False
   Icon            =   "Alarma.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   3450
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   2880
      Top             =   1320
   End
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
      Picture         =   "Alarma.frx":000C
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   120
      Picture         =   "Alarma.frx":0D76
      Stretch         =   -1  'True
      Top             =   360
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   960
      Picture         =   "Alarma.frx":198E
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmAlarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" _
                (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub Form_Load()
    Label1.Caption = "Uno de los Visores ha sido desconectado!"
    Image1.Visible = True
    Image3.Visible = False
    Timer1.Enabled = True
    Timer1.Interval = 500
    Timer2.Enabled = True
    Timer2.Interval = 1000
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

Private Sub Timer2_Timer()
    PlaySound
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
