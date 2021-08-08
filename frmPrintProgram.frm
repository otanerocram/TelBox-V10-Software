VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrintProgram 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir programación"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmPrintProgram.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1440
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "Configrurar..."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imprimir"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.CheckBox chkUsuarios 
         Caption         =   "Lista de usuarios y bloqueos"
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkTelefonos 
         Caption         =   "Lista de teléfonos"
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   360
         Picture         =   "frmPrintProgram.frx":000C
         Top             =   480
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmPrintProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isPrinting As Boolean

Private Sub chkTelefonos_Click()
    ValidateChkBox
End Sub

Private Sub chkUsuarios_Click()
    ValidateChkBox
End Sub
Private Sub ValidateChkBox()
    If chkTelefonos.Value = 0 And chkUsuarios.Value = 0 Then
        cmdImprimir.Enabled = False
    Else
        cmdImprimir.Enabled = True
    End If
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdConfig_Click()
    MDIMainForm.mnuImpresora_Click
End Sub

Private Sub cmdImprimir_Click()
    isPrinting = True
    cmdImprimir.Enabled = False
    cmdCancelar.Enabled = False
    With CrystalReport1
        .PrinterDriver = Printer.DriverName
        .PrinterName = Printer.DeviceName
        .PrinterPort = Printer.Port
        .Destination = crptToPrinter
        '.Destination = crptToWindow
        If chkTelefonos.Value = 1 Then
            .ReportFileName = App.Path() + "\" + "rptListatelefonos.rpt"
            '.Password = "Enya" & Chr(10) & "Enya"       'session & database-level pwd
            .Password = Chr(10) & "Enya"
            DoEvents
            Debug.Print "antes "; Timer
            .Action = 1
            Debug.Print "después "; Timer
        End If
        DoEvents
        If chkUsuarios.Value = 1 Then
            .ReportFileName = App.Path() + "\" + "rptlistausuarios.rpt"
            .Password = "Enya" & Chr(10) & "Enya"
            DoEvents
            .Action = 1
        End If
        isPrinting = False
        Unload Me
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If isPrinting Then Cancel = 1   'no interrumpe la impresión
End Sub
