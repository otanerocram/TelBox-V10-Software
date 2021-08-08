VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmOficinas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Oficina"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmOficinas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmdBuscaArchivo 
         Caption         =   "Buscar..."
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   1400
         Width           =   855
      End
      Begin VB.TextBox txtArchivo 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Nueva..."
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   800
         Width           =   855
      End
      Begin VB.ComboBox cboOficinas 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Programación"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Oficina"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Seleccione la oficina/sucursal y el archivo de programación:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "frmOficinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    
    Dim dbLlamadas As Database
    Dim rsOficinas As Recordset
    Dim rs
    Dim pwd As String
    Dim sqlStr As String
    'Dim oficina_Id As Integer
    
    pwd = "Enya"
    Set dbLlamadas = OpenDatabase(App.Path & "\Llamadas.mdb", False, False, ";Pwd=" & pwd)
    sqlStr = "SELECT * FROM Oficinas WHERE [Nombre]='" & cboOficinas.Text & "'"
    Set rsOficinas = dbLlamadas.OpenRecordset(sqlStr, dbOpenDynaset)
    oficina_Id = rsOficinas.Fields("Oficina_Id")
    
    rsOficinas.Close
    dbLlamadas.Close
    
    Hide
    
End Sub

Private Sub cmdBuscaArchivo_Click()
    
    Dim archivo As String
    Dim archivoRuta As String
    Dim oldtag As String
    
    On Error GoTo HacerNada
    
    With CommonDialog1
        .DialogTitle = "Guardar programación en archivo"
        .CancelError = True
        .Filter = "Archivos de programación Zziber|*.zpr"
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt
        .ShowSave
    End With
    
    'si no se produce error
    archivoRuta = CommonDialog1.filename
    archivo = CommonDialog1.FileTitle
    oldtag = Me.Tag
    txtArchivo.Text = archivo
        
Exit Sub

HacerNada:
    If Err.Number = cdlCancel Then
        'nada, se presionó Cancelar
        CommonDialog1.filename = ""
    Else
        MsgBox "No se pudo abrir el archivo de programación." _
                + vbCr + "Se produjo el error " + Err.Number + ": " + Err.Description, _
                vbCritical + vbOKOnly, "No se pudo abrir el archivo"
    End If

End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim dbLlamadas As Database
    Dim rsOficinas As Recordset
    Dim pwd As String
    Dim sqlStr As String
    
    pwd = "Enya"
    Set dbLlamadas = OpenDatabase(App.Path & "\Llamadas.mdb", False, False, ";Pwd=" & pwd)
    sqlStr = "SELECT * FROM Oficinas ORDER BY Nombre"
    Set rsOficinas = dbLlamadas.OpenRecordset(sqlStr, dbOpenDynaset)
    
    cboOficinas.Clear
    rsOficinas.MoveFirst
    Do While Not rsOficinas.EOF
        cboOficinas.AddItem rsOficinas.Fields("Nombre")
        rsOficinas.MoveNext
    Loop
    cboOficinas.ListIndex = 0
    
    rsOficinas.Close
    dbLlamadas.Close
        
End Sub
