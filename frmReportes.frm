VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmReportes 
   Caption         =   "Reportes"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame fraParam 
      Height          =   975
      Left            =   960
      TabIndex        =   1
      Top             =   5520
      Width           =   10815
      Begin VB.CommandButton cmdCerrar 
         Height          =   615
         Left            =   9960
         Picture         =   "frmReportes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cerrar"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdNuevoReporte 
         Height          =   615
         Left            =   240
         Picture         =   "frmReportes.frx":1042
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Nuevo Reporte."
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdExportar 
         Height          =   615
         Left            =   1920
         Picture         =   "frmReportes.frx":2084
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " Exportar reporte."
         Top             =   240
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6720
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   615
         Left            =   1080
         Picture         =   "frmReportes.frx":30C6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir reporte."
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Data datReporte 
      Caption         =   "Reporte"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin TrueDBGrid60.TDBGrid TDBGridReporte 
      Bindings        =   "frmReportes.frx":4108
      Height          =   4455
      Left            =   240
      OleObjectBlob   =   "frmReportes.frx":4121
      TabIndex        =   0
      Top             =   960
      Width           =   11535
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2295
      Left            =   2400
      TabIndex        =   13
      Top             =   2040
      Width           =   5175
      ExtentX         =   9128
      ExtentY         =   4048
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
      Location        =   ""
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmReportes.frx":6829
      Top             =   5640
      Width           =   480
   End
   Begin VB.Label lblSerie 
      Caption         =   "Serie"
      Height          =   255
      Left            =   7920
      TabIndex        =   11
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblDestino 
      Caption         =   "Destino"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblServicio 
      Caption         =   "Servicio"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label lblTurnoUsuario 
      Caption         =   "TurnoUsuario"
      Height          =   255
      Left            =   7920
      TabIndex        =   6
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblCabina 
      Caption         =   "Cabina"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblFechaInicio 
      Caption         =   "FechaInicio"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblTipoReporte 
      Caption         =   "TipoReporte"
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
      TabIndex        =   3
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'estas variables se toman de frmReportesParam
Dim refFechaInicio As String
Dim refFechaFin As String
Dim refTurnoUsuario As String
Dim refCabina As String
Dim refServicio As String
Dim refDestino As String
Dim refSerie As String
Dim refDetalle As Boolean

Private Sub GeneraRecordset()
    
    Select Case Me.Tag
    Case "DetalleTurno"
        RptDetalleTurno
    Case "ResumenDiario"
        RptResumenDiario
    Case "Personalizado"
        RptPersonalizado
    End Select
    
End Sub

Private Sub cmdCerrar_Click()
  Unload Me
End Sub

Private Sub cmdExportar_Click()
    cmdPrint_Click
    ExportaArchivoReporte
End Sub

Private Sub cmdNuevoReporte_Click()
    frmReportesParam.Show
End Sub

Private Sub cmdPrint_Click()
    
    Me.MousePointer = vbHourglass
        
    On Error Resume Next
    Dim espera As Single
    
    'genera archivo
    Select Case Me.Tag
    Case "DetalleTurno"
        ExportHTMLDetalleTurno
    Case "ResumenDiario"
        ExportHTMLResumenDiario
    Case "Personalizado"
        ExportHTMLPersonalizado
    End Select
    
    'carga archivo y presenta vista preliminar
    espera = Timer + 3
    With WebBrowser1
        '.Navigate (archivoExport)
        .Navigate (App.Path & "\rptCabina.htm")
        Do While Timer < espera
            DoEvents
        Loop
        '.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
        ''.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
        '.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0
        .ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
    End With
    Me.MousePointer = vbDefault




End Sub

Private Sub Form_Load()

    WebBrowser1.ZOrder 1
    
    With datReporte
        .DatabaseName = dbMain()
        .Connect = ";Pwd=" & dbPwd
    End With
    
    Me.Tag = frmReportesParam.Tag
    refFechaInicio = frmReportesParam.txtFechaInicio.Text
    refFechaFin = frmReportesParam.txtFechaFin.Text
    refTurnoUsuario = frmReportesParam.cboTurnoUsuario.Text
    refCabina = frmReportesParam.cboCabina.Text
    refServicio = frmReportesParam.cboServicio.Text
    refDestino = frmReportesParam.cboDestino.Text
    refSerie = Trim(frmReportesParam.txtSerie.Text)
    refDetalle = frmReportesParam.optDetalle(0).Value
    
    GeneraRecordset
    Unload frmReportesParam
    
End Sub

Private Sub RptDetalleTurno()
    Dim strSQL1 As String
    Dim StrSQL2 As String
    Dim strSQL3 As String
    Dim strSQL As String
    Dim strTurno As String
    Dim strFecha As String

    If refTurnoUsuario = "Todos" Then
        strFecha = "#" & Format(frmReportesParam.txtFechaInicio, "MM/dd/yyyy") & "#"
        strTurno = "SELECT Id FROM Turnos " _
                & "WHERE Turnos.FechaHoraInicio >= " & strFecha & " " _
                & "AND Turnos.FechaHoraInicio < " & strFecha & " + 1 "
    Else
        strFecha = Mid(refTurnoUsuario, 1, InStr(refTurnoUsuario, "-") - 1)
        strFecha = "#" & Format(strFecha, "MM/dd/yyyy HH:nn:ss") & "#"
        strTurno = "SELECT Id FROM Turnos " _
                & "WHERE Turnos.FechaHoraInicio = " & strFecha & " "
    End If
    
    strSQL1 = "SELECT Turnos.UsuarioLogin AS Turno, " _
            & "Format(FechaHora, 'dd/mm HH:nn') AS Inicio, " _
            & "TroncalId AS Cabina, NumTelefono, " _
            & "NombreDestino AS Destino, DuracionTexto AS Duracion,  " _
            & "Format(Costo, '#,##0.0') AS Precio, Servicio, 1 AS Orden " _
            & "FROM LLamadas, Turnos " _
            & "WHERE Llamadas.TurnoId = Turnos.Id " _
            & "AND Llamadas.TurnoId IN (" & strTurno & ") " _
            & "ORDER BY FechaHora "

    StrSQL2 = "SELECT 'Subtotal ' & Servicio AS Turno, " _
            & "COUNT(FechaHora) & ' llamadas' AS Inicio, " _
            & "' -' AS Cabina, ' -' AS NumTelefono, ' -' AS Destino, " _
            & "Format(SUM(DuracionSeg)/(3600 * 24), 'hh:nn:ss') AS Duracion, " _
            & "Format(SUM(Costo), '#,##0.0') AS Precio, " _
            & "Servicio, 2 AS Orden " _
            & "FROM Llamadas " _
            & "WHERE Llamadas.TurnoId IN (" & strTurno & ") " _
            & "GROUP BY Servicio "
            
    strSQL3 = "SELECT 'Total Reporte' AS Turno, " _
            & "COUNT(FechaHora) & ' llamadas' AS Inicio, " _
            & "' -' AS Cabina, ' -' AS NumTelefono, ' -' AS Destino, " _
            & "Format(SUM(DuracionSeg)/(3600 * 24), 'hh:nn:ss') AS Duracion, " _
            & "Format(SUM(Costo), '#,##0.0') AS Precio, " _
            & "' -' AS Servicio, 3 AS Orden " _
            & "FROM Llamadas " _
            & "WHERE Llamadas.TurnoId IN (" & strTurno & ") "
            
    strSQL = strSQL1 & "UNION (" & StrSQL2 & ") " & "UNION (" & strSQL3 & ") " _
            & "ORDER BY Orden, Inicio, Servicio "

    With datReporte
        .RecordSource = strSQL
        .Refresh
    End With
    With TDBGridReporte
        .Columns("Orden").Visible = False
        .Columns("Precio").Alignment = dbgRight
        .Columns("Precio").HeadAlignment = dbgCenter
        .Columns("Precio").Width = .Columns("Precio").Width * 0.5
        .Columns("Cabina").Width = .Columns("Cabina").Width * 0.5
        .Columns("Duracion").Width = .Columns("Duracion").Width * 0.5
        .Columns("Duracion").Alignment = dbgRight
        .Columns("Duracion").HeadAlignment = dbgCenter
        .Columns("Duracion").Caption = "Duración"
        .Refresh
    End With
    
    lblTipoReporte.Caption = "Detalle de Llamadas por Turno"
    lblFechaInicio.Visible = True
    lblFechaInicio.Caption = "Día: " & refFechaInicio
    lblTurnoUsuario.Visible = True
    lblTurnoUsuario.Caption = "Turno: " & refTurnoUsuario
    lblCabina.Visible = False
    lblServicio.Visible = False
    lblDestino.Visible = False
    lblSerie.Visible = False

End Sub

Private Sub RptResumenDiario()
    Dim strSQL1 As String
    Dim StrSQL2 As String
    Dim strSQL3 As String
    Dim strSQL As String
    Dim strTurno As String
    Dim strFecha As String
    Dim numTroncales As Integer
    Dim i As Integer
    
    numTroncales = Val(ConfigVariable("numTroncales"))

    strFecha = "#" & Format(frmReportesParam.txtFechaInicio, "MM/dd/yyyy") & "#"
    strTurno = "SELECT Id FROM Turnos " _
            & "WHERE Turnos.FechaHoraInicio >= " & strFecha & " " _
            & "AND Turnos.FechaHoraInicio < " & strFecha & " + 1 "
    
    
    'Número de llamadas
    strSQL1 = "SELECT 'Número de Llamadas' AS Servicio, "
    For i = 1 To numTroncales
        strSQL1 = strSQL1 _
                & "'' AS Cab_" & i & ", "
    Next
    strSQL1 = strSQL1 & "'' AS Total, 0.5 AS Orden "
    strSQL1 = strSQL1 _
            & "FROM Llamadas " _
            & "WHERE Llamadas.TurnoId IN (" & strTurno & ") "
    
    strSQL1 = strSQL1 & "UNION (SELECT Servicio, "
    For i = 1 To numTroncales
        strSQL1 = strSQL1 _
                & "SUM(IIF(TroncalId = " & i & ", 1, 0)) AS Cab_" & i & ", "
    Next
    strSQL1 = strSQL1 & "COUNT(TroncalId) AS Total, 1 AS Orden "
    strSQL1 = strSQL1 _
            & "FROM Llamadas " _
            & "WHERE Llamadas.TurnoId IN (" & strTurno & ") " _
            & "GROUP BY Servicio) "
    
    strSQL1 = strSQL1 & "UNION (" _
            & "SELECT 'Total Llamadas' AS Servicio, "
    For i = 1 To numTroncales
        strSQL1 = strSQL1 _
                & "SUM(IIF(TroncalId = " & i & ", 1, 0)) AS Cab_" & i & ", "
    Next
    strSQL1 = strSQL1 & "COUNT(TroncalId) AS Total, 1.5 AS Orden "
    strSQL1 = strSQL1 _
            & "FROM Llamadas " _
            & "WHERE Llamadas.TurnoId IN (" & strTurno & ") " _
            & ") "

    'Duración de las llamadas
    StrSQL2 = "SELECT 'Duración de Llamadas' AS Servicio, "
    For i = 1 To numTroncales
        StrSQL2 = StrSQL2 & "'' AS Cab_" & i & ", "
    Next
    StrSQL2 = StrSQL2 & "'' AS Total, " _
            & "2 AS Orden "
    StrSQL2 = StrSQL2 _
            & "FROM Llamadas " _
            & "WHERE Llamadas.TurnoId IN (" & strTurno & ") "
    
    StrSQL2 = StrSQL2 & "UNION (SELECT Servicio, "
    For i = 1 To numTroncales
        StrSQL2 = StrSQL2 & "Format(" _
                & "SUM(IIF(TroncalId = " & i & ", DuracionSeg, 0))/(3600*24)" _
                & ", 'hh:nn:ss') AS Cab_" & i & ", "
    Next
    StrSQL2 = StrSQL2 & "Format(SUM(DuracionSeg)/(3600*24), 'hh:nn:ss') AS Total, " _
            & "2.1 AS Orden "
    StrSQL2 = StrSQL2 _
            & "FROM Llamadas " _
            & "WHERE Llamadas.TurnoId IN (" & strTurno & ") " _
            & "GROUP BY Servicio) "
    
    StrSQL2 = StrSQL2 & "UNION (" _
            & "SELECT 'Total Duración' AS Servicio, "
    For i = 1 To numTroncales
        StrSQL2 = StrSQL2 & "Format(" _
                & "SUM(IIF(TroncalId = " & i & ", DuracionSeg, 0))/(3600*24)" _
                & ", 'hh:nn:ss') AS Cab_" & i & ", "
    Next
    StrSQL2 = StrSQL2 & "Format(SUM(DuracionSeg)/(3600*24), 'hh:nn:ss') AS Total, " _
            & "2.5 AS Orden "
    StrSQL2 = StrSQL2 _
            & "FROM Llamadas " _
            & "WHERE Llamadas.TurnoId IN (" & strTurno & ") " _
            & ") "


    'Costo
    strSQL3 = "SELECT 'Precio de Llamadas' AS Servicio, "
    For i = 1 To numTroncales
        strSQL3 = strSQL3 & "'' AS Cab_" & i & ", "
    Next
    strSQL3 = strSQL3 & "'' AS Total, 3 AS Orden "
    strSQL3 = strSQL3 _
            & "FROM Llamadas " _
            & "WHERE Llamadas.TurnoId IN (" & strTurno & ") "

    strSQL3 = strSQL3 & "UNION (SELECT Servicio, "
    For i = 1 To numTroncales
        strSQL3 = strSQL3 & "Format(" _
                & "SUM(IIF(TroncalId = " & i & ", Costo, 0))" _
                & ", '#,##0.0') AS Cab_" & i & ", "
    Next
    strSQL3 = strSQL3 & "Format(SUM(Costo), '#,##0.0') AS Total, 3.1 AS Orden "
    strSQL3 = strSQL3 _
            & "FROM Llamadas " _
            & "WHERE Llamadas.TurnoId IN (" & strTurno & ") " _
            & "GROUP BY Servicio) "

    strSQL3 = strSQL3 & "UNION (" _
            & "SELECT 'Total Precio' AS Servicio, "
    For i = 1 To numTroncales
        strSQL3 = strSQL3 & "Format(" _
                & "SUM(IIF(TroncalId = " & i & ", Costo, 0))" _
                & ", '#,##0.0') AS Cab_" & i & ", "
    Next
    strSQL3 = strSQL3 & "Format(SUM(Costo), '#,##0.0') AS Total, 3.5 AS Orden "
    strSQL3 = strSQL3 _
            & "FROM Llamadas " _
            & "WHERE Llamadas.TurnoId IN (" & strTurno & ") " _
            & ") "
            
            
    strSQL = strSQL1 & "UNION (" & StrSQL2 & ") " & "UNION (" & strSQL3 & ") " _
            & "ORDER BY Orden, Servicio "

    'Debug.Print strSQL
    
    With datReporte
        .RecordSource = strSQL
        .Refresh
    End With
    With TDBGridReporte
        .Columns("Orden").Visible = False
        .Columns("Servicio").Width = .Columns("Servicio").Width * 1.2
        For i = 1 To .Columns.Count - 1
            .Columns(i).Alignment = dbgRight
            .Columns(i).HeadAlignment = dbgCenter
        Next
        .Refresh
    End With
    
    lblTipoReporte.Caption = "Resumen por Cabina y Servicio"
    lblFechaInicio.Visible = True
    lblFechaInicio.Caption = "Día: " & refFechaInicio
    lblTurnoUsuario.Visible = False
    'lblTurnoUsuario.Caption = "Turno: " & refTurnoUsuario
    lblCabina.Visible = False
    lblServicio.Visible = False
    lblDestino.Visible = False
    lblSerie.Visible = False

End Sub


Private Sub RptPersonalizado()
    Dim strSQL1 As String
    Dim StrSQL2 As String
    Dim strSQL3 As String
    Dim strSQL As String
    Dim strUsuario As String
    Dim strFecha As String
    Dim strCabina As String
    Dim strServicio As String
    Dim strDestino As String
    Dim strSerie As String

    'strFecha = "#" & Format(refFechaInicio, "MM/dd/yyyy") & "#"
    strFecha = " AND Llamadas.FechaHora >= " _
            & "#" & Format(refFechaInicio, "MM/dd/yyyy") & "#" & " " _
            & "AND Llamadas.FechaHora < " _
            & "#" & Format(refFechaFin, "MM/dd/yyyy") & "#" & " + 1 "
    
    'Usuario
    If refTurnoUsuario <> "Todos" Then
        strUsuario = " AND Usuarios.Nombre = '" & refTurnoUsuario & "' "
    End If
    
    'Cabina
    If refCabina <> "Todas" Then
        strCabina = " AND Llamadas.TroncalId = " & refCabina & " "
    End If
    
    'Servicio
    If refServicio <> "Todos" Then
        strServicio = " AND Llamadas.Servicio = '" & refServicio & "' "
    End If
    
    'Destino
    If refDestino <> "Todos" And refDestino <> "" Then
        strDestino = " AND Llamadas.NombreDestino LIKE '" & refDestino & "' "
    End If
    
    'Serie / Número
    If refSerie <> "Todos" And refSerie <> "" Then
        strDestino = " AND Llamadas.NumTelefono LIKE '" & refSerie & "*' "
    End If
    
    If refDetalle Then _
    strSQL1 = "SELECT Turnos.UsuarioLogin AS Turno, " _
            & "Format(FechaHora, 'dd/mm/yy HH:nn') AS Inicio, " _
            & "TroncalId AS Cabina, NumTelefono, " _
            & "NombreDestino AS Destino, DuracionTexto AS Duracion,  " _
            & "Format(Costo, '#,##0.0') AS Precio, Servicio, 1 AS Orden " _
            & "FROM LLamadas, Turnos, Usuarios " _
            & "WHERE Llamadas.TurnoId = Turnos.Id " _
            & "AND Usuarios.Login = Turnos.UsuarioLogin " _
            & strFecha & strUsuario & strCabina _
            & strServicio & strDestino & strSerie _
            & "ORDER BY FechaHora "

    StrSQL2 = "SELECT 'Subtotal ' & Servicio AS Turno, " _
            & "COUNT(FechaHora) & ' llamadas' AS Inicio, " _
            & "' -' AS Cabina, ' -' AS NumTelefono, ' -' AS Destino, " _
            & "Format(SUM(DuracionSeg)/(3600 * 24), 'hh:nn:ss') AS Duracion, " _
            & "Format(SUM(Costo), '#,##0.0') AS Precio, " _
            & "Servicio, 2 AS Orden " _
            & "FROM Llamadas, Turnos, Usuarios " _
            & "WHERE Llamadas.TurnoId = Turnos.Id " _
            & "AND Usuarios.Login = Turnos.UsuarioLogin " _
            & strFecha & strUsuario & strCabina _
            & strServicio & strDestino & strSerie _
            & "GROUP BY Servicio "
            
    strSQL3 = "SELECT 'Total Reporte' AS Turno, " _
            & "COUNT(FechaHora) & ' llamadas' AS Inicio, " _
            & "' -' AS Cabina, ' -' AS NumTelefono, ' -' AS Destino, " _
            & "Format(SUM(DuracionSeg)/(3600 * 24), 'hh:nn:ss') AS Duracion, " _
            & "Format(SUM(Costo), '#,##0.0') AS Precio, " _
            & "' -' AS Servicio, 3 AS Orden " _
            & "FROM Llamadas, Turnos, Usuarios " _
            & "WHERE Llamadas.TurnoId = Turnos.Id " _
            & "AND Usuarios.Login = Turnos.UsuarioLogin " _
            & strFecha & strUsuario & strCabina _
            & strServicio & strDestino & strSerie

    If refDetalle Then
        strSQL = strSQL1 & "UNION (" & StrSQL2 & ") " & "UNION (" & strSQL3 & ") " _
                & "ORDER BY Orden, Inicio, Servicio "
    Else
        strSQL = StrSQL2 & " " & "UNION (" & strSQL3 & ") " _
                & "ORDER BY Orden, Inicio, Servicio "
    End If
    
    With datReporte
        .RecordSource = strSQL
        .Refresh
    End With
    With TDBGridReporte
        .Columns("Turno").Caption = "Usuario"
        .Columns("Orden").Visible = False
        .Columns("Precio").Alignment = dbgRight
        .Columns("Precio").HeadAlignment = dbgCenter
        .Columns("Precio").Width = .Columns("Precio").Width * 0.5
        .Columns("Cabina").Width = .Columns("Cabina").Width * 0.5
        .Columns("Duracion").Width = .Columns("Duracion").Width * 0.5
        .Columns("Duracion").Alignment = dbgRight
        .Columns("Duracion").HeadAlignment = dbgCenter
        .Columns("Duracion").Caption = "Duración"
        .Refresh
    End With
    
    lblTipoReporte.Caption = "Reporte de Llamadas"
    lblFechaInicio.Visible = True
    lblFechaInicio.Caption = "Periodo: " & refFechaInicio & " - " & refFechaFin
    lblTurnoUsuario.Visible = True
    lblTurnoUsuario.Caption = "Usuario: " & refTurnoUsuario
    lblCabina.Visible = True
    lblCabina.Caption = "Cabina: " & refCabina
    lblServicio.Visible = True
    lblServicio.Caption = "Servicio: " & refServicio
    lblDestino.Visible = True
    lblDestino.Caption = "Destino: " & refDestino
    lblSerie.Visible = True
    lblSerie.Caption = "Serie / Número: " & refSerie

End Sub


Private Sub ExportHTMLDetalleTurno()
    Dim archivoHTM As String
    
    'archivoHTM = App.Path & "\DetalleTurno.htm"
    archivoHTM = App.Path & "\rptCabina.htm"

    'crea nuevo archivo
    GoSub abrirHTM
    'exporta datos al archivo
    TDBGridReporte.ExportToFile archivoHTM, True
    '
    GoSub cerrarHTM
    
Exit Sub

abrirHTM:
    'abre archivos htm para reporte
    Open archivoHTM For Output As #1
    
    Print #1, "<html>"
    Print #1, "<head>"
    Print #1, "<title>" & nombreProducto & "</title>"
    Print #1, "<style>table {font-size:10pt}</style>"
    Print #1, "</head>"
    Print #1, "<body>"
    Print #1, "<font face=arial size=2>"
    Print #1, "<h3>Detalle de LLamadas por Turno</h3>"
    'Print #1, "<br>"
    Print #1, "<table>"
    Print #1, "<tr>"
    Print #1, "<td>Fecha: " & Format(refFechaInicio, "dd/MM/yyyy") & "&nbsp;</td>"
    Print #1, "<td>" & "&nbsp;</td>"
    Print #1, "<td>Turno: " & refTurnoUsuario & "&nbsp;</td>"
    Print #1, "</tr>"
    Print #1, "</table>"
    Print #1, "<br>"
    
    Close #1
    Return
    
cerrarHTM:
    'abre archivos htm para reporte
    Open archivoHTM For Append As #1
    
    'Print #1, "<br>Total: " & datLlamadasGrid.Recordset.RecordCount & " registro(s) <br>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close #1
    Return
    
End Sub


Private Sub ExportHTMLResumenDiario()
    Dim archivoHTM As String
    
    'archivoHTM = App.Path & "\DetalleTurno.htm"
    archivoHTM = App.Path & "\rptCabina.htm"

    'crea nuevo archivo
    GoSub abrirHTM
    'exporta datos al archivo
    TDBGridReporte.ExportToFile archivoHTM, True
    '
    GoSub cerrarHTM
    
Exit Sub

abrirHTM:
    'abre archivos htm para reporte
    Open archivoHTM For Output As #1
    
    Print #1, "<html>"
    Print #1, "<head>"
    Print #1, "<title>" & nombreProducto & "</title>"
    Print #1, "<style>table {font-size:10pt}</style>"
    Print #1, "</head>"
    Print #1, "<body>"
    Print #1, "<font face=arial size=2>"
    Print #1, "<h3>Resumen por Cabina y Servicio</h3>"
    'Print #1, "<br>"
    Print #1, "<table>"
    Print #1, "<tr>"
    Print #1, "<td>Fecha: " & Format(refFechaInicio, "dd/MM/yyyy") & "&nbsp;</td>"
    Print #1, "<td>" & "&nbsp;</td>"
    Print #1, "<td>" & "&nbsp;</td>"
    Print #1, "</tr>"
    Print #1, "</table>"
    Print #1, "<br>"
    
    Close #1
    Return
    
cerrarHTM:
    'abre archivos htm para reporte
    Open archivoHTM For Append As #1
    
    'Print #1, "<br>Total: " & datLlamadasGrid.Recordset.RecordCount & " registro(s) <br>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close #1
    Return
    
End Sub

Private Sub ExportHTMLPersonalizado()
    Dim archivoHTM As String
    
    'archivoHTM = App.Path & "\DetalleTurno.htm"
    archivoHTM = App.Path & "\rptCabina.htm"

    'crea nuevo archivo
    GoSub abrirHTM
    'exporta datos al archivo
    TDBGridReporte.ExportToFile archivoHTM, True
    '
    GoSub cerrarHTM
    
Exit Sub

abrirHTM:
    'abre archivos htm para reporte
    Open archivoHTM For Output As #1
    
    Print #1, "<html>"
    Print #1, "<head>"
    Print #1, "<title>" & nombreProducto & "</title>"
    Print #1, "<style>table {font-size:8pt}</style>"
    Print #1, "</head>"
    Print #1, "<body>"
    Print #1, "<font face=arial size=2>"
    Print #1, "<h3>Reporte de LLamadas</h3>"
    'Print #1, "<br>"
    Print #1, "<table>"
    Print #1, "<tr>"
    Print #1, "<td>Periodo: " & Format(refFechaInicio, "dd/MM/yyyy") & " - " & Format(refFechaFin, "dd/MM/yyyy") & "&nbsp;</td>"
    Print #1, "<td>" & "&nbsp;</td>"
    Print #1, "<td>Usuario: " & refTurnoUsuario & "&nbsp;</td>"
    Print #1, "<td>" & "&nbsp;</td>"
    Print #1, "<td>Cabina: " & refCabina & "&nbsp;</td>"
    Print #1, "</tr>"
    
    Print #1, "<tr>"
    Print #1, "<td>Servicio: " & refServicio & "&nbsp;</td>"
    Print #1, "<td>" & "&nbsp;</td>"
    Print #1, "<td>Destino: " & refDestino & "&nbsp;</td>"
    Print #1, "<td>" & "&nbsp;</td>"
    Print #1, "<td>Serie/Número: " & refSerie & "&nbsp;</td>"
    Print #1, "</tr>"
    
    Print #1, "</table>"
    Print #1, "<br>"
    
    Close #1
    Return
    
cerrarHTM:
    'abre archivos htm para reporte
    Open archivoHTM For Append As #1
    
    'Print #1, "<br>Total: " & datLlamadasGrid.Recordset.RecordCount & " registro(s) <br>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close #1
    Return
    
End Sub


Private Sub ExportaArchivoReporte()
    Dim archivo As String
    Dim archivoRuta As String
    Dim oldtag As String
    
    On Error GoTo HacerNada
    
    With CommonDialog1
        .DialogTitle = "Exportar reporte"
        .CancelError = True
        .DefaultExt = "htm"
        .Filter = "Archivos htm|*.htm"
        .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
        .ShowSave
    End With
    
    'si no se produce error
    'reemplaza archivo de configuración
    FileCopy App.Path & "\rptCabina.htm", CommonDialog1.FileName
        
Exit Sub

HacerNada:
    If Err.Number = cdlCancel Then
        'nada, se presionó Cancelar
        CommonDialog1.FileName = ""
    Else
        MsgBox "No se pudo seleccionar el archivo de reporte." _
                + vbCr + "Se produjo el error " + Err.Number + ": " + Err.Description, _
                vbCritical + vbOKOnly, "No se pudo abrir el archivo"
    End If

End Sub


