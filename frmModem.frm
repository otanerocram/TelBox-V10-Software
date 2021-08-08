VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmModem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conectar vía modem"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "frmModem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdContinuar 
      Caption         =   "Continuar"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtRespuesta 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3000
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdLlamar 
      Caption         =   "Marcar número"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtNumeroTelefono 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtPuertoModem 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Puerto"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmModem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tempComm As MSComm
Dim xForm As Form


Private Sub cmdContinuar_Click()
    tempComm.Tag = "DATA"
    xForm.lstResultados.AddItem "Modem Conectado..."
    'frmRecibirMulti.Show vbModal
    'frmRecibirMulti.SetFocus
    Unload Me
End Sub

Private Sub cmdLlamar_Click()
    Dim salida As String
    Dim respuesta As String
    
    If IsFormLoaded("frmTransmitirMulti") Then
        Set tempComm = frmTransmitirMulti.MSComm1
    ElseIf IsFormLoaded("frmRecibirMulti") Then
        Set tempComm = frmRecibirMulti.MSComm1
        Set xForm = frmRecibirMulti
    Else
        MsgBox "Error: no hay formulario Tx/Rx"
        Exit Sub
    End If
    
    With tempComm
        If .PortOpen Then .PortOpen = False
        '.Settings = "9600,n,8,1"
        .Settings = ConfigVariable("BaudRate") & ",n,8,1"
        '.CommPort = Val(txtPuertoModem.Text)
        .CommPort = ConfigParameter("CommPort")
        .DTREnable = True
        .RThreshold = 1
        .RTSEnable = False
        '.InBufferSize = 16384
        '.InputMode = comInputModeBinary
        .PortOpen = True
        .Tag = ""
    End With
    
    salida = "ATZ" & vbCr
    tempComm.InputLen = 0
    tempComm.Output = salida
    espera 3
    'respuesta = tempComm.Input
    tempComm.Tag = "ATD"

    
    salida = "ATDT" & Trim(txtNumeroTelefono.Text) & vbCr
    tempComm.Output = salida
    tempComm.InputLen = 1
    espera 1
    
End Sub

Private Sub tempComm_OnComm()
    Dim inputChar As String
        
    Select Case tempComm.CommEvent
    ' Controlar cada evento o error escribiendo
    ' código en cada instrucción Case

    ' Errores
        'Case comBreak   ' Se ha recibido una interrupción.
        'Case comEventCDTO   ' Tiempo de espera CD (RLSD).
        'Case comEvent0CTSTO  ' Tiempo de espera CTS.
        'Case comEventDSRTO  ' Tiempo de espera DSR.
        'Case comEventFrame  ' Error de trama
        'Case comEventOverrun    ' Datos perdidos.
        'Case comEventRxOver ' Desbordamiento del búfer
                            ' de recepción.

        'Case comEventRxParity   ' Error de paridad.
        'Case comEventTxFull ' Búfer de transmisión lleno.
        'Case comEventDCB    ' Error inesperado al recibir DCB]

    ' Eventos
        'Case comEvCD    ' Cambio en la línea CD.
        'Case comEvCTS   ' Cambio en la línea CTS.
        'Case comEvDSR   ' Cambio en la línea DSR.
        'Case comEvRing  ' Cambio en el indicador de
                            ' llamadas.
        Case comEvReceive   ' Recibido nº RThreshold de
                                ' caracteres.

            inputChar = tempComm.Input
            'Debug.Print inputChar, Asc(inputChar), Hex(Asc(inputChar)), Cont
            If tempComm.Tag = "ATD" Then txtRespuesta.Text = txtRespuesta.Text & inputChar
            
        'Case comEvSend  ' Hay un número SThreshold de
                            ' caracteres en el búfer
                            ' de transmisión.
        'Case comEvEOF   ' Se ha encontrado un carácter
                            ' EOF en la entrada
    End Select

End Sub

