VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmReporteErrores 
   Caption         =   "Reporte de Errores"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   Icon            =   "frmReporteErrores.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   10470
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7095
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   9615
      ExtentX         =   16960
      ExtentY         =   12515
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmReporteErrores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRefresh_Click()
    WebBrowser1.Navigate (App.Path & "\ErrorLlamadas.htm")
End Sub

Private Sub Form_Load()
    cmdRefresh_Click
End Sub
