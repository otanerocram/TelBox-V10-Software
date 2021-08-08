VERSION 5.00
Begin VB.Form frmModifClaveProg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar Clave de Programación"
   ClientHeight    =   2070
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5535
   HelpContextID   =   620
   Icon            =   "frmModifClaveProg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1223.024
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   5197.065
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3480
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   390
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2880
      TabIndex        =   4
      Top             =   1560
      Width           =   1185
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   3480
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   1845
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmModifClaveProg.frx":000C
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Caption         =   "Confirme la nueva clave:"
      Height          =   270
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nueva clave (4 a 15 caracteres):"
      Height          =   270
      Index           =   1
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2400
   End
End
Attribute VB_Name = "frmModifClaveProg"
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
        rsClave.Fields("Clave Programacion") = txtPassword(1).Text
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

Private Sub txtPassword_Change(Index As Integer)
    If Len(Trim(txtPassword(Index).Text)) < 4 Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub
