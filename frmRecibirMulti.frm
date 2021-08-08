VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRecibirMulti 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Leer datos desde el Zziber"
   ClientHeight    =   5820
   ClientLeft      =   705
   ClientTop       =   885
   ClientWidth     =   5805
   FillColor       =   &H0000C000&
   HelpContextID   =   200
   Icon            =   "frmRecibirMulti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkModem 
      Caption         =   "Usar Modem"
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CheckBox Troncales_all 
      Caption         =   "Seleccionar Todo"
      Height          =   195
      Left            =   3240
      TabIndex        =   19
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Calendario Zziber"
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   2520
      Width           =   5055
      Begin VB.CommandButton Command1 
         Caption         =   "Cambiar..."
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblHora 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOficina 
      Caption         =   "..."
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   1600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtOficina 
      Height          =   285
      Left            =   3600
      TabIndex        =   17
      Text            =   "Principal"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox chkOficina 
      Caption         =   "Usar archivo / Oficina:"
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ListBox lstResultados 
      Height          =   1425
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   5535
   End
   Begin VB.Frame fraOperacion 
      Caption         =   "Leer"
      Height          =   1095
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optLeerDatos 
         Caption         =   "Programación y llamadas"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optLeerDatos 
         Caption         =   "Sólo programación"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton optLeerDatos 
         Caption         =   "Sólo llamadas"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.ListBox lstTroncales 
      Height          =   1860
      Left            =   840
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   720
   End
   Begin VB.CommandButton cmdOpciones 
      Caption         =   "Opciones..."
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   180
      Picture         =   "frmRecibirMulti.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   240
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   1200
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4560
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton CmdIniciar 
      Caption         =   "Iniciar"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      InBufferSize    =   2048
      BaudRate        =   2400
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Resultados"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Troncal(es)"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmRecibirMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim k As Long
Dim tiempo As Single
Dim Maxim As Single
Dim puerto As Integer
Dim dbLlamadas As Database
Dim rsAgenda As Recordset
Dim rsLlamadas As Recordset
Dim rsTroncales As Recordset
Dim rsUsuarios As Recordset
Dim pwd As String
Dim modoComm As String
Dim modemport As Integer
Dim getLinea As String
Dim operStatus As Variant

Private Sub chkModem_Click()
    If chkModem.value = 1 Then
        frmModem.Show vbModal ', Me
    ElseIf IsFormLoaded("frmModem") Then
        Unload frmModem
    Else
    
    End If
End Sub

Private Sub MSComm1_OnComm()
    If MSComm1.Tag = "DATA" Then Exit Sub
    Dim inputChar As String
        
    Select Case MSComm1.CommEvent
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

            inputChar = MSComm1.Input
            'Debug.Print inputChar, Asc(inputChar), Hex(Asc(inputChar)), Cont
            If MSComm1.Tag = "ATD" Then frmModem.txtRespuesta.Text = frmModem.txtRespuesta.Text & inputChar
            
        'Case comEvSend  ' Hay un número SThreshold de
                            ' caracteres en el búfer
                            ' de transmisión.
        'Case comEvEOF   ' Se ha encontrado un carácter
                            ' EOF en la entrada
    End Select
    
End Sub

Private Sub Troncales_all_Click()
If Troncales_all.value = 1 Then
    Dim numerodetroncales, nt
    
    numerodetroncales = lstTroncales.ListCount
    For nt = 0 To numerodetroncales - 1
        lstTroncales.ListIndex = nt
        lstTroncales.selected(lstTroncales.ListIndex) = True
    Next nt
    lstTroncales.ListIndex = 0
Else
    numerodetroncales = lstTroncales.ListCount
    For nt = 0 To numerodetroncales - 1
        lstTroncales.ListIndex = nt
        lstTroncales.selected(lstTroncales.ListIndex) = False
    Next nt
    lstTroncales.ListIndex = 0
End If
End Sub

Private Sub chkOficina_Click()
    If chkOficina.value = 1 Then
        txtOficina.Visible = True
        cmdOficina.Visible = True
    Else
        txtOficina.Visible = False
        cmdOficina.Visible = False
    End If
End Sub
    
Private Sub BuscaArchivo()
    Dim archivo As String
    Dim archivoRuta As String
    Dim oldtag As String
    
    On Error GoTo HacerNada
    
    With CommonDialog1
        .DialogTitle = "Leer datos usando archivo"
        .CancelError = True
        .Filter = "Archivos de programación Zziber|*.zpr"
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .ShowOpen
    End With
    
    'si no se produce error
    archivoRuta = CommonDialog1.filename
    archivo = CommonDialog1.FileTitle
    txtOficina.Text = archivo
    txtOficina.Tag = archivoRuta
        
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

Private Sub CmdIniciar_Click()
'    recepcionOK = False     'variable pública
    Dim oldtag As String
    Dim archivoExt As Boolean

    lstResultados.Clear
    Me.Height = 6225 '5895
    DoEvents
    leePuerto
'    On Error GoTo errComm
    
    If chkOficina.value = 1 And txtOficina.Text <> "Principal" Then
        If MDIMainForm.GetProgramFile(txtOficina.Tag) = False Then GoTo cancelComm
        archivoExt = True
        oldtag = MDIMainForm.Tag
        MDIMainForm.Tag = "LeerDatosExt;" & txtOficina.Text
    End If
    
    Dim i As Integer
    Dim indice As Integer
    For i = 0 To 2
        If optLeerDatos(i).value = True Then
            indice = i + 1
            Exit For
        End If
    Next
    Select Case indice
        Case 1
            Me.Tag = "RecibeTodo"   'usado por Zziber_Recibir
        Case 2
            Me.Tag = "RecibeProg"
        Case 3
            Me.Tag = "RecibeLlamadas"
    End Select
    operStatus = "receiving"
    Dim first As Boolean
    first = True
    For i = 0 To lstTroncales.ListCount - 1
        lstTroncales.ListIndex = i
        If lstTroncales.selected(i) = True Then
            If first Then Me.Move 30, 1260
            first = False
            lstResultados.AddItem "Troncal " & i + 1 & ": Conmutando..."
            lstResultados.ListIndex = lstResultados.ListCount - 1
            DoEvents
            If operStatus = "truncated" Then GoTo endComm
            IniciaConmutador i
            'globalPass = True
            If globalPass <> True Then
                GoTo endComm
            Else
                lstResultados.List(lstResultados.ListCount - 1) = "Troncal " & i + 1 & ": en proceso..."
            End If
            DoEvents
            If operStatus = "truncated" Then GoTo cancelComm
            ZZiber_Recibir.Show vbModal
            DoEvents
            Select Case globalPass
            Case True
                lstResultados.List(lstResultados.ListCount - 1) = _
                    "Troncal " & i + 1 & ": Lectura de datos OK."
            Case "truncated"
                lstResultados.List(lstResultados.ListCount - 1) = _
                    "Troncal " & i + 1 & ": Proceso cancelado por el usuario."
            Case Else
                lstResultados.List(lstResultados.ListCount - 1) = _
                    "Troncal " & i + 1 & ": ERROR de comunicación/lectura de datos. Proceso cancelado."
            End Select
            DoEvents
            If operStatus = "truncated" Then GoTo cancelComm
        End If
    Next
    IniciaConmutador -1
    If operStatus = "truncated" Then GoTo cancelComm
    lstResultados.AddItem "Terminado."
    lstResultados.ListIndex = lstResultados.ListCount - 1
    operStatus = "OK"
    GoTo endComm

Exit Sub

cancelComm:
    lstResultados.AddItem "Cerrando..."
    lstResultados.ListIndex = lstResultados.ListCount - 1
    DoEvents
    IniciaConmutador -1
    TerminaComm
Exit Sub

endComm:
    If archivoExt Then MDIMainForm.Tag = oldtag
    TerminaComm
Exit Sub

'errComm:
'    MsgBox "Eureka"
End Sub

Private Sub IniciaConmutador(troncal As Integer)
    CmdIniciar.Enabled = False
    CmdCancelar.Enabled = False
    cmdOpciones.Enabled = False
    globalPass = False
    Dim tiempo As Single
    Dim respuesta As String
    Dim salida As String
    Dim Contador As Single
    On Error GoTo errorComm
    Dim i As Integer
    'Label1.Caption = "Buscando el Conmutador..."
    With MSComm1
        If chkModem.value = 0 Then
            .CommPort = puerto
            If .PortOpen = True Then .PortOpen = False
            .PortOpen = True
            .Settings = "1200,N,8,1"
            .RTSEnable = False
        End If
        .InputLen = 0
        respuesta = .Input
        Select Case True
        Case troncal = 255
            salida = Chr(255) & Chr(255) & Chr(255) & Chr(255) _
                     & Chr(255) & Chr(255) & Chr(255) & Chr(255)
        Case troncal < 8
            salida = Chr(2 ^ (troncal)) & Chr(0) & Chr(0) & Chr(0) _
                     & Chr(0) & Chr(0) & Chr(0) & Chr(0)
        Case troncal >= 8 And troncal < 16
            salida = Chr(0) & Chr(2 ^ (troncal - 8)) & Chr(0) & Chr(0) _
                     & Chr(0) & Chr(0) & Chr(0) & Chr(0)
        Case troncal >= 16 And troncal < 24
            salida = Chr(0) & Chr(0) & Chr(2 ^ (troncal - 16)) & Chr(0) _
                     & Chr(0) & Chr(0) & Chr(0) & Chr(0)
        Case troncal >= 24 And troncal < 32
            salida = Chr(0) & Chr(0) & Chr(0) & Chr(2 ^ (troncal - 24)) _
                     & Chr(0) & Chr(0) & Chr(0) & Chr(0)
        Case troncal >= 32 And troncal < 40
            salida = Chr(0) & Chr(0) & Chr(0) & Chr(0) _
                     & Chr(2 ^ (troncal - 32)) & Chr(0) & Chr(0) & Chr(0)
        Case troncal >= 40 And troncal < 48
            salida = Chr(0) & Chr(0) & Chr(0) & Chr(0) _
                     & Chr(0) & Chr(2 ^ (troncal - 40)) & Chr(0) & Chr(0)
        Case troncal >= 48 And troncal < 56
            salida = Chr(0) & Chr(0) & Chr(0) & Chr(0) _
                     & Chr(0) & Chr(0) & Chr(2 ^ (troncal - 48)) & Chr(0)
        Case Else
            salida = Chr(0) & Chr(0) & Chr(0) & Chr(0) _
                     & Chr(0) & Chr(0) & Chr(0) & Chr(2 ^ (troncal - 56))
        End Select
        'salida = IIf(troncal < 8, Chr(2 ^ (troncal)) & Chr(0), Chr(0) & Chr(2 ^ (troncal)))
        'salida = Chr(0) & Chr(0)
        salida = "ZZIBER" & salida
        For i = 1 To Len(salida)
            .Output = Mid(salida, i, 1)
            'Debug.Print Mid(salida, i, 1)
            tiempo = Timer + 0.01
            Do While Timer < tiempo
            Loop
        Next
        .PortOpen = False
        tiempo = Timer() + 1.5
        Do While Timer < tiempo
'            If .InBufferCount > 0 Then
                globalPass = True
'                Exit Do
'            End If
            DoEvents
        Loop
        If globalPass = False Then Err.Raise vbObjectError + 1050
    End With


Exit Sub

errorComm:
    'Beep
    lstResultados.List(lstResultados.ListCount - 1) = lstResultados.List(lstResultados.ListCount - 1) & " ERROR"
    DoEvents
    Select Case Err.Number
    Case vbObjectError + 1050   'custom
        MsgBox "No se encuentra el conmutador en COM" & puerto & "." _
        & vbCr & "Verifique que el conmutador esté conectado al puerto COM" & puerto & " o" _
        & vbCr & "indique en qué otro puerto se encuentra conectado el conmutador," _
        & vbCr & "o seleccione discado manual o por módem en opciones de comunicación." _
        , vbExclamation + vbOKOnly, "No se encuentra el conmutador"
        DoEvents
        If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Case 8005   'puerto ya está abierto
        MsgBox "El puerto de comunicaciones COM" & puerto & " está en uso por otro programa." _
        & vbCr & "Verifique que el conmutador esté en COM" & puerto & " y desactive el otro programa," _
        & vbCr & "especifique en qué otro puerto se encuentra conectado el conmutador," _
        & vbCr & "o seleccione discado manual o por módem en opciones de comunicación." _
        , vbExclamation + vbOKOnly, "El puerto del conmutador está ocupado"
    Case 8002   'puerto no válido
        MsgBox "El puerto de comunicaciones COM" & puerto & " no es válido." _
        & vbCr & "Verifique que el conmutador esté conectado a un puerto serial válido," _
        & vbCr & "o seleccione discado manual o por módem en opciones de comunicación." _
        , vbExclamation + vbOKOnly, "El puerto del conmutador no es válido"
    Case Else
        MsgBox "Error " & Err.Number & ":" & Err.Description, vbCritical + vbOKOnly, "Conmutador - Error"
    End Select
    DoEvents
End Sub

Private Sub TerminaComm()
    If operStatus = "truncated" Then
        Unload Me
        Exit Sub
    End If
    If globalPass <> True Then recepcionOK = False
    CmdIniciar.Enabled = True
    CmdCancelar.Enabled = True
    cmdOpciones.Enabled = True
    DoEvents
End Sub

Private Sub cmdOficina_Click()
    BuscaArchivo
End Sub

Private Sub cmdOpciones_Click()
    Opciones.Show vbModal
End Sub

Private Sub Command1_Click()
    Shell "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl,,0"
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Timer2_Timer
'llena lista de troncales
    lstTroncales.Clear
    'Dim i As Integer
    'lstTroncales.AddItem "Todas"
    'For i = 1 To 16
    '    lstTroncales.AddItem i
    'Next
    'lstTroncales.AddItem "Ninguna"
    Dim dbTroncal As Database
    Dim rsTroncal As Recordset
    Dim pwd As String
    Dim elemento As Integer
    Dim textoTroncal As String
    Dim strSQL As String
    
    pwd = "Enya"
    strSQL = "SELECT TOP " & ConfigVariable("numTroncales") & " * FROM Troncales ORDER BY Troncal_Id "
    Set dbTroncal = OpenDatabase(App.Path & "\Llamadas.mdb", False, False, ";Pwd=" & pwd)
    Set rsTroncal = dbTroncal.OpenRecordset(strSQL, dbOpenDynaset)
    rsTroncal.MoveFirst
    Do While Not rsTroncal.EOF
        textoTroncal = rsTroncal.Fields("troncal_id")
        textoTroncal = textoTroncal + IIf(Len(rsTroncal.Fields("troncal")) > 0, _
                                          " (" & rsTroncal.Fields("Troncal") & ")", "")
        lstTroncales.AddItem textoTroncal
        If Len(rsTroncal.Fields("troncal")) > 0 Then lstTroncales.selected(elemento) = True
        elemento = elemento + 1
        rsTroncal.MoveNext
    Loop
    rsTroncal.Close
    dbTroncal.Close
    lstTroncales_Click
    Screen.MousePointer = vbDefault
End Sub
Private Sub leePuerto()
    
    Dim dbClave As Database
    Dim rsClave As Recordset
'    Dim pwd As String
    pwd = "Enya"
    Set dbClave = OpenDatabase(App.Path & "\VisualZziber.mdb", True, False, ";Pwd=" & pwd)
    Set rsClave = dbClave.OpenRecordset("Claves", dbOpenDynaset)
    With rsClave
        puerto = .Fields("CommPort")
        modemport = .Fields("ModemPort")
        modoComm = .Fields("CommMode")
        getLinea = .Fields("GetLinea")
    End With
    rsClave.Close
    dbClave.Close
    
Exit Sub
    Open App.Path & "\zzibervisual.ini" For Input As #1
    Dim linea As String
    Do While Not EOF(1)
        Line Input #1, linea
        linea = LTrim(linea)
        If UCase(Mid(linea, 1, 9)) = "COMMPORT=" Then
            puerto = Val(Mid(linea, 10, 5))
            Exit Do
        End If
    Loop
    Close #1
    If puerto = 0 Then puerto = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If operStatus = "receiving" Then
        operStatus = "truncated"
        Cancel = 1
        Exit Sub
    End If
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Dim i As Integer
    For i = 0 To Forms.Count - 1
        Debug.Print Forms(i).Name
    Next
    'If Zziber_Transmitir.Tag = "AutoRead" Then Zziber_Transmitir.Tag = ""
End Sub

Private Sub lstTroncales_Click()
    'Static click As Integer
    'click = click + 1
    'Debug.Print "click"; click
    Dim selected As Boolean
    Dim i As Integer
    For i = 0 To lstTroncales.ListCount - 1
        If lstTroncales.selected(i) = True Then
            selected = True
            Exit For
        End If
    Next
    If selected Then
        CmdIniciar.Enabled = True
    Else
        CmdIniciar.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    CmdIniciar_Click
    'If Zziber_Transmitir.Tag = "AutoRead" Then CmdIniciar_Click
End Sub

Private Sub Timer2_Timer()
    lblFecha.Caption = Format(now, "dd/mm/yy")
    lblHora.Caption = Format(now, "hh:mm")
End Sub
