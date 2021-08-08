VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmRepair 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verificar/Optimizar archivos de datos"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   HelpContextID   =   900
   Icon            =   "frmRepair.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtResultados 
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   5535
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdVerificar 
      Caption         =   "Verificar"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      _Version        =   327680
      Center          =   -1  'True
      FullWidth       =   57
      FullHeight      =   57
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmRepair.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Resultados"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   $"frmRepair.frx":0316
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmRepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdVerificar_Click()
    Animation1.Visible = True
    Animation1.Open App.Path & "\Findfile.avi"
    Animation1.Play
    txtResultados.Text = ""
    DoEvents
    cmdVerificar.Enabled = False
    cmdCancelar.Enabled = False
    RepairDB "Llamadas.mdb"
    CompactDB "Llamadas.mdb"
    RepairDB "visualzziber.mdb"
    CompactDB "visualzziber.mdb"
    cmdVerificar.Enabled = True
    cmdCancelar.Enabled = True
    If txtResultados.Text = "" Then
        txtResultados.Text = "Proceso terminado OK"
    End If
    Animation1.Stop
    Animation1.Close
    Animation1.Visible = False
    cmdCancelar.SetFocus
End Sub

Sub RepairDB(db As String)

    Dim errBucle As Error
    On Error GoTo Err_Reparar
    DBEngine.RepairDatabase App.Path & "\" & db
    On Error GoTo 0
    'MsgBox "¡Fin del procedimiento reparar!"

Exit Sub

Err_Reparar:
    For Each errBucle In DBEngine.Errors
        txtResultados.Text = txtResultados.Text _
            & "Archivo: " & db & Chr(13) & Chr(10) _
            & "Error " & errBucle.Number & ": " & errBucle.Description & Chr(13) & Chr(10)
        txtResultados.SelStart = Len(txtResultados.Text)
        DoEvents
    Next errBucle

End Sub


Sub CompactDB(db As String)

    Dim pwd As String
    Dim errBucle As Error
    Dim oldDB As String
    pwd = "Enya"
    On Error GoTo errorHandler
    oldDB = Mid(db, 1, Len(db) - 3) & "old"
    If Dir(App.Path & "\" & oldDB) <> "" Then Kill App.Path & "\" & oldDB
    DBEngine.CompactDatabase App.Path & "\" & db, App.Path & "\" & oldDB, , , ";Pwd=" & pwd
    Kill App.Path & "\" & db
    FileCopy App.Path & "\" & oldDB, App.Path & "\" & db
    On Error GoTo 0

Exit Sub

errorHandler:
    For Each errBucle In DBEngine.Errors
        txtResultados.Text = txtResultados.Text _
            & "Archivo :" & db & Chr(13) & Chr(10) _
            & "Error " & errBucle.Number & ": " & errBucle.Description & Chr(13) & Chr(10)
        txtResultados.SelStart = Len(txtResultados.Text)
            DoEvents
    Next errBucle
    
End Sub

