VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDReporte As String
Public Tipo As String
Public Nombre As String
Public Titulo As String
Public MostrarEnVisor As Boolean
Public Personal As Boolean
Public PrinterSetupBeforeShow As Boolean

Private mRecordset As ADODB.Recordset

Public Parametros As Collection

Private Report As CRAXDRT.Report

Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = mRecordset
End Property

Public Property Set Recordset(ByVal Value As ADODB.Recordset)
    Set mRecordset = Value.Clone
End Property

'////////////////////////////////////////////////////////////////////

Public Function Load() As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim ReporteParametro As ReporteParametro

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Reporte_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDReporte_FILTER", adChar, adParamInput, 50, IDReporte)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenKeyset, adLockOptimistic
    Set cmdData = Nothing
        
    Tipo = recData("Tipo").Value
    Nombre = recData("Nombre").Value
    Titulo = recData("Titulo").Value & ""
    Personal = recData("Personal").Value
    
    recData.Close
    Set recData = Nothing
    
    '//////////////////////////////////////////////////
    'PARAMETERS
    Set Parametros = New Collection
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ReporteParametro_List"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDReporte_FILTER", adChar, adParamInput, 50, IDReporte)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenKeyset, adLockOptimistic
    Set cmdData = Nothing
    
    Do While Not recData.EOF
        Set ReporteParametro = New ReporteParametro
        
        ReporteParametro.IDParametro = RTrim(recData("IDParametro").Value)
        ReporteParametro.Nombre = recData("Nombre").Value
        ReporteParametro.Tipo = RTrim(recData("Tipo").Value)
        ReporteParametro.Requerido = recData("Requerido").Value
        ReporteParametro.RequeridoLeyenda = recData("RequeridoLeyenda").Value & ""
        Parametros.Add ReporteParametro, ReporteParametro.IDParametro
        
        recData.MoveNext
    Loop
    
    recData.Close
    Set recData = Nothing
        
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Reporte.Load", "Error al obtener los datos del Reporte." & vbCr & vbCr & "IDReporte: " & IDReporte
End Function

Public Function OpenReport() As Boolean
    Dim errorMessage As String
    Dim ReporteParametro As ReporteParametro
    
    Dim DatabaseTable As CRAXDRT.DatabaseTable
    Dim Section As CRAXDRT.Section
    Dim ReportObject As Object
    Dim SubReportObject As CRAXDRT.SubReportObject
    Dim ParameterFieldDefinition As CRAXDRT.ParameterFieldDefinition
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    'Start the Application if its the first time
    If pCRAXDRTApplication Is Nothing Then
        errorMessage = "Error al iniciar una sesión de Crystal Report."
        Set pCRAXDRTApplication = New CRAXDRT.Application
    End If
    
    errorMessage = "Error al abrir el Reporte."
    Set Report = pCRAXDRTApplication.OpenReport(pParametro.Report_Path & IDReporte & ".rpt", crOpenReportByDefault)
    
    If Not Recordset Is Nothing Then
        Call Report.Database.SetDataSource(mRecordset)
    Else
        For Each DatabaseTable In Report.Database.Tables
            DatabaseTable.DllName = REPORT_DATABASE_DLL_NAME
            'DatabaseTable.ConnectionProperties("Provider").Value = pParametro.Report_Provider
            DatabaseTable.ConnectionProperties("Data Source").Value = pDatabase.DataSource
            DatabaseTable.ConnectionProperties("Initial Catalog").Value = pDatabase.Database
            DatabaseTable.ConnectionProperties("Connect Timeout").Value = pDatabase.ConnectionTimeout
            DatabaseTable.ConnectionProperties("General Timeout").Value = pDatabase.CommandTimeout
            DatabaseTable.ConnectionProperties("User ID").Value = pDatabase.UserID
            DatabaseTable.ConnectionProperties("Password").Value = pDatabase.Password
            DatabaseTable.Location = pDatabase.Database & ".dbo." & DatabaseTable.Location
        Next DatabaseTable
    
        For Each Section In Report.Sections
            For Each ReportObject In Section.ReportObjects
                If ReportObject.Kind = crSubreportObject Then
                    Set SubReportObject = ReportObject
                    For Each DatabaseTable In SubReportObject.OpenSubreport.Database.Tables
                        'DatabaseTable.ConnectionProperties("Provider").Value = pParametro.Report_Provider
                        DatabaseTable.ConnectionProperties("Data Source").Value = pDatabase.DataSource
                        DatabaseTable.ConnectionProperties("Initial Catalog").Value = pDatabase.Database
                        DatabaseTable.ConnectionProperties("Connect Timeout").Value = pDatabase.ConnectionTimeout
                        DatabaseTable.ConnectionProperties("General Timeout").Value = pDatabase.CommandTimeout
                        DatabaseTable.ConnectionProperties("User ID").Value = pDatabase.UserID
                        DatabaseTable.ConnectionProperties("Password").Value = pDatabase.Password
                        DatabaseTable.Location = pDatabase.Database & ".dbo." & DatabaseTable.Location
                    Next DatabaseTable
                End If
            Next ReportObject
        Next Section
    End If

    With Report
        .ApplicationName = App.Title
        If PrinterSetupBeforeShow Then
            .PrinterSetup frmMDI.hWnd
        End If
        .SetDialogParentWindow frmMDI.hWnd
    End With
    
    errorMessage = "Error al abrir el Reporte."
        
    For Each ReporteParametro In Parametros
        For Each ParameterFieldDefinition In Report.ParameterFields
            With ParameterFieldDefinition
                If IIf(.ParameterType = crStoreProcedureParameter, "@", "") & ReporteParametro.IDParametro = .ParameterFieldName Then
                    Select Case ReporteParametro.Tipo
                        Case REPORTE_PARAMETRO_TIPO_COMPANY
                            .AddCurrentValue pParametro.CompanyName
                        Case REPORTE_PARAMETRO_TIPO_TITLE
                            .AddCurrentValue Titulo
                        Case REPORTE_PARAMETRO_TIPO_CONDITION_TEXT
                            .AddCurrentValue GetConditionText()
                        Case REPORTE_PARAMETRO_TIPO_PERSONAL
                            .AddCurrentValue pPersonal
                        Case Else
                            If IsEmpty(ReporteParametro.Valor) Then
                                .EnableNullValue = True
                                .AddCurrentValue Empty
                            Else
                                Select Case .ValueType
                                    Case crCurrencyField, crNumberField
                                        .AddCurrentValue CDbl(ReporteParametro.Valor)
                                    Case crDateTimeField, crDateField, crTimeField
                                        Select Case ReporteParametro.Tipo
                                            Case REPORTE_PARAMETRO_TIPO_DATE_TIME
                                                .AddCurrentValue CDate(ReporteParametro.Valor)
                                            Case REPORTE_PARAMETRO_TIPO_DATE
                                                .AddCurrentValue CDate(ReporteParametro.Valor)
                                            Case REPORTE_PARAMETRO_TIPO_TIME
                                                .AddCurrentValue CDate("1900/01/01 " & ReporteParametro.Valor)
                                            Case REPORTE_PARAMETRO_TIPO_YEAR_MONTH_FROM, REPORTE_PARAMETRO_TIPO_YEAR_MONTH_TO
                                                .AddCurrentValue CDate(ReporteParametro.Valor)
                                        End Select
                                    Case crStringField
                                        .AddCurrentValue CStr(ReporteParametro.Valor)
                                    Case crBooleanField
                                        .AddCurrentValue CBool(ReporteParametro.Valor)
                                End Select
                            End If
                    End Select
                    Exit For
                End If
            End With
        Next ParameterFieldDefinition
    Next ReporteParametro
    
    Report.EnableParameterPrompting = False
    
    Screen.MousePointer = vbDefault
    OpenReport = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Reporte.OpenReport", errorMessage & vbCr & vbCr & "Filename: " & pParametro.Report_Path & IDReporte & ".rpt"
End Function

Public Function PrintReport(ByVal Preview As Boolean) As Boolean
    Dim ReportViewer As frmReportViewer
    Dim FormLoaded As Boolean
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    If Preview Then
        Set ReportViewer = New frmReportViewer
        FormLoaded = True
        With ReportViewer
            .CRViewer.ReportSource = Report
            .Caption = "REPORTE: " & Replace(Titulo, vbCr, " | ")
            .CRViewer.ViewReport
            .Show
        End With
    Else
        Report.PrintOut
    End If
    Screen.MousePointer = vbDefault
    PrintReport = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Reporte.PrintReport", "Error al abrir una nueva sesión del Visor de Reportes."
    If FormLoaded Then
        Unload ReportViewer
        Set ReportViewer = Nothing
    End If
End Function

Public Property Get ExportOptions() As CRAXDRT.ExportOptions
    Set ExportOptions = Report.ExportOptions
End Property

Public Function ExportReport(Optional ByVal PromptUser As Boolean = False) As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    Report.Export PromptUser

    Screen.MousePointer = vbDefault
    ExportReport = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Reporte.ExportReport", "Error al Exportar el Reporte."
End Function

Private Sub Class_Initialize()
    Set Parametros = New Collection
End Sub

Private Sub Class_Terminate()
    Set Parametros = Nothing
    If Not mRecordset Is Nothing Then
        If mRecordset.State = adStateOpen Then
            If Not (mRecordset.BOF Or mRecordset.EOF) Then
                If mRecordset.EditMode <> adEditNone Then
                    mRecordset.CancelUpdate
                End If
            End If
            mRecordset.Close
        End If
        Set mRecordset = Nothing
    End If
End Sub

Private Function GetConditionText() As String
    Dim ReporteParametro As ReporteParametro
    
    For Each ReporteParametro In Parametros
        If Not IsEmpty(ReporteParametro.Valor) Then
            GetConditionText = GetConditionText & IIf(GetConditionText = "", "", vbCr) & ReporteParametro.Nombre & ": " & ReporteParametro.ValorLeyenda
        End If
    Next ReporteParametro
    
    Set ReporteParametro = Nothing
End Function


'    For Each Table In Report.Database.Tables
'        'Table.DllName = REPORT_DATABASE_DLL_NAME
'        With Table.ConnectionProperties
'            .DeleteAll
'            '.TestConnectivity
'
''            .ConnectionProperties("Provider").Value = pParametro.Database_Provider
''            .ConnectionProperties("Data Source").Value = pParametro.Database_DataSource
''            .ConnectionProperties("Initial Catalog").Value = pParametro.Database_Database
''            .ConnectionProperties("User ID").Value = pParametro.Database_UserID
''            .ConnectionProperties("Password").Value = pParametro.Database_Password
''            .ConnectionProperties("Use DSN Default Properties").Value = True
''            .ConnectionProperties("Connect Timeout").Value = pParametro.Database_ConnectionTimeout
''            .ConnectionProperties("General Timeout").Value = pParametro.Database_ConnectionTimeout
'
'            .Add "Provider", pParametro.Database_Provider
'            .Add "Data Source", pParametro.Database_DataSource
'            .Add "Initial Catalog", pParametro.Database_Database
'            .Add "User ID", pParametro.Database_UserID
'            .Add "Password", pParametro.Database_Password
'            .Add "Integrated Security", False
'            .Add "Use DSN Default Properties", False
'            .Add "Asynchronous Processing", 0
'            .Add "Locale Identifier", 11274
'            .Add "Connect Timeout", pParametro.Database_ConnectionTimeout
'            .Add "General Timeout", pParametro.Database_ConnectionTimeout
'            .Add "OLE DB Services", -5
'            .Add "Current Language", Empty
'            .Add "Initial File Name", Empty
'            .Add "Use Encryption for Data", 0
'            .Add "Replication server name connect option", Empty
'            .Add "Tag with column collation when possible", 0
'            .Add "MARS Connection", 0
'            .Add "Failover Partner", Empty
'            .Add "Old Password", Empty
'            .Add "DataTypeCompatibility", 0
'            .Add "Trust Server Certificate", 0
'        End With
'    Next Table
'
