VERSION 5.00
Begin VB.Form frmMasterClaveProg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zziber Visual - Clave de Programación"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   Icon            =   "frmMasterClaveProg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   3975
      Begin VB.TextBox txtClave 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Nueva clave   :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ubicación del programa Zziber Visual"
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3975
      Begin VB.DirListBox dirBuscar 
         Height          =   1890
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
      Begin VB.DriveListBox drvBuscar 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Grabar nueva clave"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   4320
      Width           =   1695
   End
End
Attribute VB_Name = "frmMasterClaveProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currDrive As String
Dim encontrado As Boolean

Private Sub cmdAceptar_Click()
    Dim wrkZziber As Workspace
    Dim dbZziber As Database
    Dim rsZziber As Recordset
    Dim pwd As String
    pwd = "Enya"
    Set wrkZziber = DBEngine.Workspaces(0)
    Set dbZziber = OpenDatabase(App.Path & "\Visualzziber.mdb", False, False, ";Pwd=" & pwd)
    Set rsZziber = dbZziber.OpenRecordset("Claves", dbOpenDynaset)
    With rsZziber
        .Edit
        .Fields("Clave Programacion") = Trim(txtClave.Text)
        .Update
    End With
    rsZziber.Close
    dbZziber.Close
    wrkZziber.Close
    MsgBox "Se actualizó la clave de programación", vbInformation + vbOKOnly, "Nueva clave"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub dirBuscar_Change()
    Dim slash As String
    slash = IIf(Right(dirBuscar.Path, 1) = "\", "", "\")
    Debug.Print Dir(dirBuscar.Path & slash & "visualzziber.mdb")
    If Dir(dirBuscar.Path & slash & "visualzziber.mdb") = "" Then
        'no encuentra el archivo
        encontrado = False
    Else
        encontrado = True
    End If
    Validar
End Sub

Private Sub drvBuscar_Change()
    On Error GoTo errorDrive
    dirBuscar.Path = drvBuscar.Drive '& "\"
    currDrive = drvBuscar.Drive
Exit Sub

errorDrive:
    Select Case Err.Number
    Case 0
        'nada
    Case Else
        If MsgBox("Error " & Err.Number & ": " & Err.Description, vbOKCancel + vbCritical, "Error") <> vbCancel Then
            Resume
        End If
        drvBuscar.Drive = currDrive
    End Select
End Sub

Private Sub Form_Load()
    currDrive = drvBuscar.Drive
    dirBuscar_Change
End Sub

Private Sub Validar()
    If encontrado And Len(txtClave.Text) > 3 Then
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
End Sub

Private Sub txtClave_Change()
    Validar
End Sub
