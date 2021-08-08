VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmReporteCabina 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Total de Llamadas"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmReporteCabina.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   280
      Width           =   975
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "Ver"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Data datTotal 
      Caption         =   "Total"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin TrueDBGrid60.TDBGrid TDBGridReporteCabina 
      Bindings        =   "frmReporteCabina.frx":000C
      Height          =   3375
      Left            =   480
      OleObjectBlob   =   "frmReporteCabina.frx":0023
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1575
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Width           =   2295
      ExtentX         =   4048
      ExtentY         =   2778
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
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "Ver llamadas del día"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmReporteCabina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Me.MousePointer = vbHourglass
    
    On Error Resume Next
    Dim archivoExport As String
    Dim espera As Single
    

    archivoExport = App.Path & "\ReporteCabina.htm"
    'genera archivo html
    ExportHTMLReporte
        
    
    espera = Timer + 3
    With WebBrowser1
        .Navigate (archivoExport)
        Do While Timer < espera
            DoEvents
        Loop
        '.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
        .ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
        '.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0
        '.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
    End With
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdReporte_Click()

    Dim strSQL As String
    
    strSQL = "SELECT Troncal_Id AS Cabina, " _
            & "COUNT(Troncal_Id) AS Num_Llamadas, " _
            & "SUM(Costo) AS Costo_Total " _
            & "FROM Llamadas " _
            & "WHERE Fecha = #" & txtFecha.Text & "# " _
            & "GROUP BY Troncal_id " _
            & "UNION " _
            & "SELECT 'Total' AS Cabina, " _
            & "COUNT(Troncal_Id) AS Num_Llamadas, " _
            & "SUM(Costo) AS Costo_Total " _
            & "FROM Llamadas " _
            & "WHERE Fecha = #" & txtFecha.Text & "# "

    With datTotal
        .RecordSource = strSQL
        .Refresh
    End With
    Debug.Print datTotal.Recordset.EOF
    
    
    TDBGridReporteCabina.Refresh
    cmdPrint.Enabled = True
    
End Sub


Private Sub Form_Load()
    Dim pwd As String
    Dim i As Integer
    
    WebBrowser1.ZOrder 1
    
    pwd = "Enya"
    
    With datTotal
        .DatabaseName = dbMain()
        .Connect = ";Pwd=" & pwd
    End With

    With TDBGridReporteCabina
        .Columns("Cabina").Alignment = dbgCenter
        .Columns("Num_Llamadas").Alignment = dbgCenter
        .Columns("Costo_Total").Alignment = dbgCenter
    End With

    txtFecha.Text = Format(Date, "MM/dd/yyyy")

End Sub

Private Sub txtFecha_Change()
    If IsDate(txtFecha.Text) Then
        cmdReporte.Enabled = True
    Else
        cmdReporte.Enabled = False
    End If
    TDBGridReporteCabina.Close
    cmdPrint.Enabled = False
End Sub

Private Sub ExportHTMLReporte()
    Dim archivoHTM As String
    
    archivoHTM = App.Path & "\ReporteCabina.htm"

    'crea nuevo archivo
    GoSub abrirHTM
    'exporta datos al archivo
    TDBGridReporteCabina.ExportToFile archivoHTM, True
    '
    GoSub cerrarHTM
    
    
    
Exit Sub

abrirHTM:
    'abre archivos htm para reporte
    Open archivoHTM For Output As #1
    
    Print #1, "<html>"
    Print #1, "<head>"
    Print #1, "<title>Zziber Visual</title>"
    Print #1, "<style>table {font-size:10pt}</style>"
    Print #1, "</head>"
    Print #1, "<body>"
    Print #1, "<font face=arial size=2>"
    Print #1, "<h3>Reporte de Total de LLamadas del día " & txtFecha.Text & "</h3>"
    'Print #1, "<br>"
    Print #1, "<table>"
    Print #1, "<tr>"
    Print #1, "<td>Fecha: " & Format(Date, "dd/MM/yyyy") & "&nbsp;</td>"
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
    
    Print #1, "<br>Total: " & datTotal.Recordset.RecordCount - 1 & " cabina(s) <br>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close #1
    Return
    
End Sub

