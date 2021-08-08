VERSION 5.00
Begin VB.Form frmPuertoComm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Error de lectura/transmisión de datos"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmPuertoComm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
      Begin VB.OptionButton optCommPort 
         Caption         =   "COM 4"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optCommPort 
         Caption         =   "COM 3"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optCommPort 
         Caption         =   "COM 2"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optCommPort 
         Caption         =   "COM 1"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   240
         Picture         =   "frmPuertoComm.frx":000C
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblAccion 
      Caption         =   "Puede"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label lblError 
      Caption         =   "El"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmPuertoComm.frx":0418
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPuertoComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim puerto As Integer

Private Sub cmdAceptar_Click()
    Dim Control As Variant
    For Each Control In optCommPort
        If Control.Value = True Then
            puerto = Control.Index + 1
            Exit For
        End If
    Next
    If puerto = 0 Then puerto = 2
'actualiza mdb
    Dim dbClave As Database
    Dim rsClave As Recordset
    Dim pwd As String
    pwd = "Enya"
    Set dbClave = OpenDatabase(App.Path & "\VisualZziber.mdb", True, False, ";Pwd=" & pwd)
    Set rsClave = dbClave.OpenRecordset("Claves", dbOpenDynaset)
    With rsClave
        .Edit
        .Fields("CommPort") = puerto
'        .Fields("ModemPort") = modemPort
'        .Fields("CommMode") = modoComm
'        .Fields("GetLinea") = txtAnexo.Text
        .Update
    End With
    rsClave.Close
    dbClave.Close
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim dbClave As Database
    Dim rsClave As Recordset
    Dim pwd As String
    pwd = "Enya"
    Set dbClave = OpenDatabase(App.Path & "\VisualZziber.mdb", True, False, ";Pwd=" & pwd)
    Set rsClave = dbClave.OpenRecordset("Claves", dbOpenDynaset)
    With rsClave
        puerto = .Fields("CommPort")
'        modemPort = .Fields("ModemPort")
'        modoComm = .Fields("CommMode")
'        getLinea = .Fields("GetLinea")
    End With
    rsClave.Close
    dbClave.Close
    
    If puerto = 0 Then puerto = 2

    optCommPort(puerto - 1).Value = True
    Select Case globalPass
    Case vbObjectError + 1050   'custom
        lblError.Caption = "Error en la comunicación."
        lblAccion.Caption = "Verifique que el puerto selecionado sea el correcto y" _
                           & vbCr & "que el cable de conexión entre el Zziber y la PC esté seguro."
    Case 8005   'puerto ya está abierto
        lblError.Caption = "El puerto COM" & puerto & " se encuentra en uso."
        lblAccion.Caption = "Cierre el programa que usa el puerto o " _
                            & vbCr & "seleccione un puerto que esté libre."
    Case 8002   'puerto no válido
        lblError.Caption = "El puerto COM" & puerto & " no es válido."
        lblAccion.Caption = "Seleccione un puerto de comunicaciones válido y" _
                            & vbCr & "que esté libre."
    End Select
    Beep
End Sub

Private Sub Form_Unload(Cancel As Integer)
    globalPass = False      'error
End Sub
