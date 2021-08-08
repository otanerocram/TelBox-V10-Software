VERSION 5.00
Begin VB.Form frmModifClaveIngreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar Clave de Ingreso"
   ClientHeight    =   2175
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5535
   HelpContextID   =   620
   Icon            =   "frmModifClaveIngreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1285.062
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   5197.065
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3600
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   1755
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2880
      TabIndex        =   4
      Top             =   1680
      Width           =   1185
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   3600
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmModifClaveIngreso.frx":000C
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Deje en blanco los casilleros si no desea clave para ingresar:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Confirme la nueva clave:"
      Height          =   270
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nueva clave (hasta 15 caracteres):"
      Height          =   270
      Index           =   1
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   2520
   End
End
Attribute VB_Name = "frmModifClaveIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public LoginSucceeded As Boolean
Private cuentaerror As Integer

Private Sub cmdCancel_Click()
    'establece la variable global a false
    'para indicar un fallo en el inicio de sesión
    'LoginSucceeded = False
    'Me.Hide
    Unload Me
End Sub

Private Sub cmdOK_Click()

    If txtPassword(0).Text = txtPassword(1).Text Then
        Dim dbClave As Database
        Dim rsClave As Recordset
        Dim pwd As String
        pwd = "Enya"
        Set dbClave = OpenDatabase(App.Path & "\VisualZziber.mdb", True, False, ";Pwd=" & pwd)
        Set rsClave = dbClave.OpenRecordset("Claves", dbOpenDynaset)
        rsClave.Edit
        rsClave.Fields("Clave Entrada") = txtPassword(1).Text
        rsClave.Update
        rsClave.Close
        dbClave.Close
        Unload Me
        Opciones.Tag = "ClaveOK"    'para committrans
        MsgBox "La nueva clave está vigente", vbInformation, "Nueva clave OK"
    Else
        cuentaerror = cuentaerror + 1
        If cuentaerror = 3 Then
            MsgBox "Alcanzó el número máximo de intentos permitidos." + Chr(10) + "                         Verifique su Clave", 16, "Acceso denegado"
            cuentaerror = 0
            Unload Me
        Else
            MsgBox "Debe confirmar la nueva clave", 48, "Error en la clave"
            txtPassword(1).SetFocus
            SendKeys "{Home}+{End}"
        End If
    End If
End Sub
