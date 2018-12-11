Attribute VB_Name = "Startup"
Option Explicit

'///////////////////////////////////////////////////////////////////
'RUNTIME
Public pIsCompiled As Boolean
Public pTrapErrors As Boolean

'///////////////////////////////////////////////////////////////////
'CONFIGURATION
Public pParametro As Parametro
Public pCSC_Parameter As CSC_Parameter
Public pRegionalSettings As CSC_RegionalSettings
Public pSpecialFolders As CSC_SpecialFolders
Public pSucursal As Sucursal

'///////////////////////////////////////////////////////////////////
'DATABASE
Public pDatabase As CSC_Database_ADO_SQL

'///////////////////////////////////////////////////////////////////
'REFRESH FLAGS
Public pCSemaforoGeneral As CSemaforoGeneral
Public precSemaforoGeneral As ADODB.Recordset
Public precSemaforoLlamada As ADODB.Recordset
Public precSemaforoUsuario As ADODB.Recordset

'///////////////////////////////////////////////////////////////////
'SESSION DATA
Public pUsuario As Usuario
Public pCPermiso As CPermiso
Public pPersonal As Boolean
Public pMessengerEnabled As Boolean
Public pMessengerBlinking As Boolean

'///////////////////////////////////////////////////////////////////
'REPORTS
Public pCRAXDRTApplication As CRAXDRT.Application

'///////////////////////////////////////////////////////////////////
'TELEPHONY
Public pTelephony As Telephony

Private Sub Main()
    Dim StartTime As Date
    Dim cmdData As ADODB.command
    
    If App.PrevInstance Then
        Call CSM_Instance.ActivatePrevious
    End If
    
    App.OleServerBusyMsgText = "Sistema Ocupado - Aguarde unos instantes."
    App.OleServerBusyMsgTitle = App.Title
    
    pIsCompiled = CSM_Instance.IsCompiled()
    'pIsCompiled = True
    pTrapErrors = pIsCompiled
        
    Screen.MousePointer = vbHourglass
    
    '//////////////////////////////////////////////////////////////////
    'OBTENGO LOS PARAMETROS DE LA ESTACION DE TRABAJO
    Set pParametro = New Parametro
    If Not pParametro.LoadWorkstationParameters() Then
        TerminateApplication
        Exit Sub
    End If
    
    Set pCSC_Parameter = New CSC_Parameter
    
    'PARSE COMMAND-LINE ARGUMENTS AND IF CONFIG IS SPECIFIED, LOAD WORKSTATION PARAMETERS AGAIN
    Call ParseCommandLineArguments
    
    If pParametro.System_MaintenanceMode Then
        TerminateApplication
        Exit Sub
    End If
    
    'INIT LOGGING
    'Verifico que exista el path especificado en el INI para el log
    If FileSystem.Dir(pParametro.Logs_Path, vbDirectory) = "" Then
        MsgBox "La carpeta especificada para los archivos de Log no existe.", vbCritical, App.Title
        TerminateApplication
        Exit Sub
    End If
    Call CSM_ApplicationLog.InitLogging(pParametro.Logs_Path, pParametro.Logs_FileNameTemplate, pParametro.Logs_MonthsToKeep)
    WriteLogEvent "*** Application Starts ***", vbLogEventTypeInformation
    
    frmSplash.MousePointer = vbHourglass
    frmSplash.Show
    DoEvents
    
    Load frmMDI
    
    '//////////////////////////////////////////////////////////////////
    'REALIZO LA CONEXION A LA BASE DE DATOS (ADO)
    Set pDatabase = New CSC_Database_ADO_SQL
    With pDatabase
        .Provider = pParametro.Database_Provider
        .ConnectionTimeout = pParametro.Database_ConnectionTimeout
        .CommandTimeout = pParametro.Database_CommandTimeout
        .PacketSize = pParametro.Database_PacketSize
        .DataTypeCompatibility = pParametro.Database_DataTypeCompatibility
        .CursorLocationServer = (pParametro.Database_CursorLocation = adUseServer)
        .DataSource = pParametro.Database_DataSource
        .FailoverPartner = pParametro.Database_FailoverPartner
        .UserID = pParametro.Database_UserID
        .Password = pParametro.Database_Password
        .Database = pParametro.Database_Database
        If Not .Connect() Then
            Unload frmMDI
            Exit Sub
        End If
    End With
    
    '//////////////////////////////////////////////////////////////////
    'OBTENGO LOS PARAMETROS DEL SISTEMA
    If Not pParametro.LoadSystemParameters() Then
        Unload frmMDI
        Exit Sub
    End If
    
    'PREPARO EL MDI DE ACUERDO AL TIPO DE EMPRESA
    App.Title = App.Title & " == " & pParametro.CompanyName
    frmMDI.Caption = App.Title
    frmMDI.BackColor = pParametro.Interface_MDIBackgroundColor
    frmMDI.tlbMain.Buttons("COMISION").Visible = pParametro.Comision_Habilitar
    
    Set pSucursal = New Sucursal
    pSucursal.IDSucursal = pParametro.IDSucursal
    pSucursal.NoMatchRaiseError = False
    If Not pSucursal.Load() Then
        Unload frmMDI
        Exit Sub
    End If
    If pSucursal.NoMatch Then
        'MsgBox "No está especificada la Sucursal a la cual pertenece esta Estación de Trabajo.", vbInformation, App.Title
    End If
    
    'LOGO DEL SPLASH
    On Error Resume Next
    If pParametro.Interface_CompanyLogo = "" Or FileSystem.Dir(pParametro.Interface_CompanyLogo) = "" Then
        frmSplash.imgCompany.Visible = False
        frmSplash.lblLicense.Visible = True
        frmSplash.lblCompanyName.Caption = pParametro.CompanyName
        frmSplash.lblCompanyName.Visible = True
    Else
        frmSplash.imgCompany.Picture = LoadPicture(pParametro.Interface_CompanyLogo)
        frmSplash.imgCompany.Left = (frmSplash.ScaleWidth - frmSplash.imgCompany.Width) / 2
        frmSplash.imgCompany.Visible = True
        frmSplash.lblLicense.Visible = False
        frmSplash.lblCompanyName.Visible = False
    End If
    DoEvents
    
    'LOGO DEL STATUS BAR
    On Error Resume Next
    If pParametro.Interface_CompanyLogoMini = "" Or FileSystem.Dir(pParametro.Interface_CompanyLogoMini) = "" Then
        frmMDI.stbMain.Panels("COMPANY_NAME").Text = " " & pParametro.CompanyName & " "
    Else
        frmMDI.stbMain.Panels("COMPANY_NAME").Picture = LoadPicture(pParametro.Interface_CompanyLogoMini)
    End If
    On Error GoTo 0
    
    '/////////////////////////////////////////////////////////////////
    'TELEFONIA
    Set pTelephony = New Telephony
    pTelephony.TelephonyType = pParametro.Telephony_Type
    pTelephony.Initialize
    
    '/////////////////////////////////////////////////////////////////
    'TIMER MULTIPLE CADA 1 SEGUNDO
    frmMDI.tmrMulti.Enabled = pIsCompiled
    'frmMDI.tmrMulti.Enabled = True
    
    '/////////////////////////////////////////////////////////////////
    'SEMAFORO GENERAL
    On Error Resume Next
    Err.Clear
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_SemaforoGeneral_Data"
    cmdData.CommandType = adCmdStoredProc
    Set precSemaforoGeneral = New ADODB.Recordset
    precSemaforoGeneral.Open cmdData, , adOpenKeyset, adLockOptimistic
    Set cmdData = Nothing
    If Err.Number <> 0 Then
        ShowErrorMessage "Modules.Startup.Main", "Error al Abrir la Tabla de Semáforos Generales."
        Unload frmMDI
        Exit Sub
    End If
    If pParametro.Database_CursorLocation = adUseClient Then
        precSemaforoGeneral.Properties("Update Criteria").Value = adCriteriaKey
    End If
    Set pCSemaforoGeneral = New CSemaforoGeneral
    
    '//////////////////////////////////////////////////////////////////
    'VERIFICO LOS PARAMETROS DE PERSONAL
    If pParametro.Personal_Status_Global Then
        RefreshList_CheckForPersonal
    ElseIf pParametro.Personal_Status_Persistent Then
        pPersonal = CSM_Registry.GetValue_FromApplication_CurrentUser("Interface", "Personal", pParametro.Personal_Status_OnStartup, csrdtBoolean)
    Else
        pPersonal = pParametro.Personal_Status_OnStartup
    End If
    
    StartTime = Now
    
    '/////////////////////////////////////////////////////////////////
    'SETTINGS
    Set pRegionalSettings = New CSC_RegionalSettings
    Set pSpecialFolders = New CSC_SpecialFolders
    
    If pIsCompiled Then
        Load frmLogin
    End If
    
    If pIsCompiled Then
        Do While DateDiff("s", StartTime, Now) < 4
            DoEvents
        Loop
    End If
        
    frmMDI.Show
    
    Unload frmSplash
    Set frmSplash = Nothing
    
    '/////////////////////////////////////////////////////////////////
    'USUARIO
    Screen.MousePointer = vbDefault
    Set pUsuario = New Usuario
    If pIsCompiled Then
        frmLogin.txtIDUsuario.Text = CSM_Registry.GetValue_FromApplication_CurrentUser("", "LastUserID", "", csrdtString)
        If frmLogin.txtIDUsuario.Text <> "" Then
            frmLogin.lblIDUsuario.tabIndex = 7
            frmLogin.txtIDUsuario.tabIndex = 7
        End If
        frmLogin.Show vbModal, frmMDI
        If pUsuario.IDUsuario = 0 Then
            Unload frmMDI
            Exit Sub
        End If
        If Not pUsuario.LogIn() Then
            WriteLogEvent "Usuario Not Logged In - Exiting", vbLogEventTypeInformation
            Unload frmMDI
            Exit Sub
        End If
    Else
        pUsuario.IDUsuario = USUARIO_ID_ADMINISTRATOR
        If Not pUsuario.Load() Then
            Unload frmMDI
            Exit Sub
        End If
        If Not pUsuario.LogIn() Then
            Unload frmMDI
            Exit Sub
        End If
    End If
        
    WriteLogEvent "Application Loading, Startup and Login: DONE", vbLogEventTypeInformation
End Sub


' *****************************************************************************
' Purpose:  Unload and Cleanup all Objects and Forms
'
' Method:
'
' Inputs:
'       None
'
' Outputs:
'       None
'
' Errors:
'       This Function no raise Errors.
'
' Asserts:
'
' Developer                 Date            Comments
' ---------                 ----            --------
' Tomas A. Cardoner         23-Jan-2002     Initial creation.
' *****************************************************************************
Public Sub TerminateApplication()
    Static Running As Boolean
    
    If Running Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    Running = True
    
    WriteLogEvent "Starting Application Terminate Routine", vbLogEventTypeInformation
    
    'Unload All Forms in memory
    WriteLogEvent "Unloading All Forms from Memory", vbLogEventTypeInformation
    CSM_Forms.UnloadAll
    
    '///////////////////////////////////////////////////////////////////
    'DATABASE CONNECTION
    pDatabase.Disconnect
    Set pDatabase = Nothing
    
    WriteLogEvent "Cleaning Public Objects References", vbLogEventTypeInformation
    '///////////////////////////////////////////////////////////////////
    'REFRESH FLAGS
    Set pCSemaforoGeneral = Nothing
    
    '///////////////////////////////////////////////////////////////////
    'TELEPHONY
    Set pTelephony = Nothing
    
    '///////////////////////////////////////////////////////////////////
    'CRYSTAL REPORTS APPLICATION
    If Not pCRAXDRTApplication Is Nothing Then
        pCRAXDRTApplication.LogOffServer REPORT_DATABASE_DLL_NAME, pParametro.Database_DataSource, pParametro.Database_Database, pParametro.Database_UserID, pParametro.Database_Password
        Set pCRAXDRTApplication = Nothing
    End If
    
    '///////////////////////////////////////////////////////////////////
    'SESSION DATA
    Set pUsuario = Nothing
    Set pCPermiso = Nothing
    
    '///////////////////////////////////////////////////////////////////
    'CONFIGURATION
    Set pSucursal = Nothing
    Set pParametro = Nothing
    Set pCSC_Parameter = Nothing
    Set pRegionalSettings = Nothing
    Set pSpecialFolders = Nothing
    
    WriteLogEvent "*** Application Terminate ***", vbLogEventTypeInformation
    Running = False
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub ParseCommandLineArguments()
    Dim aArguments() As String
    Dim ArgumentIndex As Integer
    Dim Argument As String
    Dim ArgumentEqualPosition As Integer
    Dim ArgumentName As String
    Dim ArgumentValue As String
    Dim DES As CSC_Encryption_DES
    
    If command$ <> "" Then
        aArguments = Split(command$, " ")
        For ArgumentIndex = 0 To UBound(aArguments)
            Argument = aArguments(ArgumentIndex)
            If Len(Argument) > 0 Then
                ArgumentEqualPosition = InStr(1, Argument, "=")
                If ArgumentEqualPosition > 1 And ArgumentEqualPosition < Len(Argument) Then
                    ArgumentName = UCase(Left(Argument, ArgumentEqualPosition - 1))
                    ArgumentValue = Mid(Argument, ArgumentEqualPosition + 1)
                    Select Case ArgumentName
                        Case "DATASOURCE"
                            pParametro.Database_DataSource = ArgumentValue
                        Case "CONNECTIONTIMEOUT"
                            pParametro.Database_ConnectionTimeout = Val(ArgumentValue)
                        Case "USERID"
                            pParametro.Database_UserID = ArgumentValue
                        Case "PASSWORD"
                            Set DES = New CSC_Encryption_DES
                            pParametro.Database_Password = DES.DecryptString(ArgumentValue, PASSWORD_ENCRYPTION_KEY)
                            Set DES = Nothing
                        Case "DATABASE"
                            pParametro.Database_Database = ArgumentValue
                        Case "REPORTSPATH"
                            pParametro.Report_Path = ArgumentValue
                        Case "CONFIG"
                            pParametro.Config_Name = ArgumentValue
                            pParametro.Config_Type = ""
                            Call pParametro.LoadWorkstationParameters
                    End Select
                End If
            End If
        Next ArgumentIndex
    End If
End Sub
