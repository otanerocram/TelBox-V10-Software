VERSION 5.00
Begin VB.Form frmMensaje 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   945
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMensaje.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      Caption         =   "Mensaje"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
    lblMensaje.Caption = Me.Tag
End Sub

