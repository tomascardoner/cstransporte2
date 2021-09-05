VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/////////////////////////////////////////////////////////////
'SYSTEM PARAMETERS
'/////////////////////////////////////////////////////////////

Public Empresa_ID_Predeterminada As Long
Public Provincia_ID_Predeterminada As String
Public CompanyName As String

Public RefreshList_Fastest_CheckInterval_Seconds As Long
Public RefreshList_Slowest_CheckInterval_Seconds As Long

Public LogAccion_Enabled As Boolean
Public System_SyncDateTimeWithDBServer_Enabled As Boolean
Public Workstation_SyncDateTimeWithDBServer_Enabled As Boolean
Public Messenger_Enabled As Boolean

'PERSONAS
Public Persona_DatoIncompleto_AvisoTipo As Byte     '0 = Ignorar / 1 = Avisar / 2 = Exigir
Public Persona_DatoIncompleto_RegistroTodos As Boolean
Public Persona_DatoIncompleto_CampoDocumento As Boolean
Public Persona_DatoIncompleto_CampoDomicilio As Boolean

Public Persona_Apellido_Busqueda_Delay_Milliseconds As Long

Public Viaje_Actual_Rango_Minutos As Long
Public Viaje_Permite_2_Conductores As Boolean
Public Viaje_Permite_RutaConexion As Boolean
Public Viaje_RutaConexion_TiempoEsperaMaximo_Minutos As Long

Public Permitir_Reservas_Condicionales As Boolean

Public Viaje_Estado_Activo_Color As Long
Public Viaje_Estado_Activo_Bold As Boolean
Public Viaje_Estado_EnProgreso_Color As Long
Public Viaje_Estado_EnProgreso_Bold As Boolean
Public Viaje_Estado_Finalizado_Color As Long
Public Viaje_Estado_Finalizado_Bold As Boolean
Public Viaje_Estado_Cancelado_Color As Long
Public Viaje_Estado_Cancelado_Bold As Boolean

Public Viaje_Normal_Estado_Vencido_Horas As Long
Public Viaje_Especial_Estado_Vencido_Horas As Long

Public Viaje_Especial_MostrarEnTodasLasRutas As Boolean
Public Viaje_Especial_Color As Long
Public Viaje_Especial_Bold As Boolean

Public ViajeDetalle_CreadoEnProgreso_Color As Long
Public ViajeDetalle_ModificadoEnProgreso_Color As Long

Public ViajeDetalle_Transferencia_RangoMaximoMinutos As Long

Public ViajeDetalle_MedioPago_UtilizaOperacion_ObligaFacturaNumero As Boolean

Public CuentaCorrienteGrupo_ID_ViajeDebito As Long
Public CuentaCorrienteGrupo_ID_ViajeCredito As Long
Public CuentaCorrienteGrupo_ID_Transferencia As Long
Public CuentaCorrienteGrupo_ID_Sueldo As Long
Public CuentaCorrienteCaja_ID_ViajeDebito As Long
Public CuentaCorriente_MovimientoAnterior_Minutos As Long

Public MedioPago_Predeterminado_ID As Long

'//////////////////////////////////////
'PRECIOS
Public ListaPrecio_ID_Predeterminada As Long
Public ViajeDetalle_Pasajero_Precio_PermitirModificar As Boolean
Public ViajeDetalle_Comision_Precio_PermitirModificar As Boolean
Public Persona_Prepago_AvisoCantidadRestante As Integer

'//////////////////////////////////////
'COMISIONES
Public Comision_Habilitar As Boolean
Public Comision_Seguro_Habilitar As Boolean
Public Comision_Seguro_ValorDeclaradoMinimo As Currency
Public Comision_Seguro_ValorDeclaradoMinimo_Preasignar As Boolean
Public Comision_Seguro_ValorDeclaradoMaximo As Currency
Public Comision_Seguro_Porcentaje As Single
Public Comision_Seguro_PorcentajeAplicarSobreTotal As Boolean
Public Comision_Seguro_RedondeoDecimales As Byte

Public UsuarioGrupo_ID_Predeterminado As Long

Public LugarGrupo_ID_Otro As Long
Public Lugar_ID_Otro As Long
Public Ruta_ID_Otra As String

Public Ruta_Paquete_ID As String
Public ViajeDetalle_Paquete_Permite_Multiples_Pagos As Boolean

Public TelefonoTipo_ID_Otro As Long

Public Vehiculo_TelefonoTipo_ID As Long

Public Recordset_MaxRecords As Long

Public Pasajero_AvisoInasistencia_Veces As Long
Public Pasajero_AvisoInasistencia_Dias As Long

'INTERNET DATA
Public InternetDataEmail_SenderDisplayName As String
Public InternetDataEmail_SenderAddress As String
Public InternetDataEmail_SMTPHost As String
Public InternetDataEmail_SMTPUserName As String
Public InternetDataEmail_SMTPPassword As String

'PLANILLA VIAJE EMAIL
Public PlanillaViajeEmail_SendReportFormat As Integer
Public PlanillaViajeEmail_SendExcel As Boolean
Public PlanillaViajeEmail_SMTPHost As String
Public PlanillaViajeEmail_SMTPUserName As String
Public PlanillaViajeEmail_SMTPPassword As String
Public PlanillaViajeEmail_Margin_Top As Single
Public PlanillaViajeEmail_Margin_Left As Single
Public PlanillaViajeEmail_Margin_Right As Single
Public PlanillaViajeEmail_Margin_Bottom As Single
Public PlanillaViajeEmail_ColumnWidth_A As Single
Public PlanillaViajeEmail_ColumnWidth_B As Single
Public PlanillaViajeEmail_ColumnWidth_C As Single
Public PlanillaViajeEmail_ColumnWidth_D As Single
Public PlanillaViajeEmail_ColumnWidth_E As Single
Public PlanillaViajeEmail_ColumnWidth_F As Single
Public PlanillaViajeEmail_ColumnWidth_G As Single
Public PlanillaViajeEmail_ColumnWidth_H As Single
Public PlanillaViajeEmail_ColumnWidth_I As Single
Public PlanillaViajeEmail_ColumnWidth_J As Single
Public PlanillaViajeEmail_ColumnWidth_K As Single
Public PlanillaViajeEmail_RowHeight_2 As Single
Public PlanillaViajeEmail_RowHeight_4 As Single
Public PlanillaViajeEmail_RowHeight_6 As Single
Public PlanillaViajeEmail_RowHeight_8 As Single
Public PlanillaViajeEmail_RowHeight_35 As Single
Public PlanillaViajeEmail_RowHeight_37 As Single

Public PlanillaViajeEmail_RowHeight_ColumnHeader As Single
Public PlanillaViajeEmail_RowHeight_Detalle As Single

Public Personal_Status_Global As Boolean
Public Personal_Status_Global_CheckInterval_Seconds As Long
Public Personal_Status_Persistent As Boolean
Public Personal_Status_OnStartup As Boolean

'SISTEMA WEB ON LINE
Public WebSite_HyperLink_Main As String
Public WebSite_HyperLink_Login As String
Public WebSite_HyperLink_ChangePassword As String
Public WebSite_Persona_EmailBienvenida_Asunto As String
Public WebSite_Persona_EmailBienvenida_Texto As String

'AGENT
Public StartupDelay_Seconds As Byte
Public InternetData_Send_Interval_Seconds As Long
Public PersonaAdministrativo_Sueldo_Interval_Seconds As Long
Public PlanillaViajeEmail_Send_Interval_Seconds As Long
Public Viaje_AutoPersonal_Interval_Seconds As Long
Public Viaje_AutoPersonal_AfterNumberOfDays As Long
Public Email_Delete_Interval_Seconds As Long
Public Email_Delete_AfterNumberOfDays As Long

'LISTA DE PASAJEROS
Public ListaPasajero_ArchivoNombre As String
Public ListaPasajero_RowStart As Integer
Public ListaPasajero_ColumnPasajero As String
Public ListaPasajero_ColumnDocumento As String

'/////////////////////////////////////////////////////////////
'WORKSTATION PARAMETERS
'/////////////////////////////////////////////////////////////
Public Config_Type As String
Public Config_Name As String
Public Config_INIFile_FullPathAndName As String

'CURRENT
Public Database_Provider As String
Public Database_ConnectionTimeout As Long
Public Database_CommandTimeout As Long
Public Database_PacketSize As Long
Public Database_DataTypeCompatibility As Integer
Public Database_DataSource As String
Public Database_FailoverPartner As String
Public Database_UserID As String
Public Database_Password  As String
Public Database_Database As String
Public Database_DatabaseHistory As String

Public Database_CursorLocation As ADODB.CursorLocationEnum
Public Database_CursorType As ADODB.CursorTypeEnum

Public Logs_Path As String
Public Logs_FileNameTemplate As String
Public Logs_MonthsToKeep As Byte

Public Report_Provider As String
Public Report_Path As String

Public DateTime_SyncWithServer_Primary As Boolean
Public DateTime_SyncWithServer_Secondary As Boolean
Public DateTime_SyncWithServer_IntervalSeconds As Integer

Public Interface_MDIBackgroundColor As Long
Public Interface_CompanyLogo As String
Public Interface_CompanyLogoMini As String

Public System_MaintenanceMode As Boolean

Public IDSucursal As String

Public RefreshList_Enabled As Boolean

'TELEPHONY
Public Telephony_Type As String 'TAPI / COMM / NONE
Public RegistroLlamada_Save As Boolean

'/////////////////////////////////////////////////////////////
'USER PARAMETERS
'/////////////////////////////////////////////////////////////
Public ListView_GridLines As Boolean

Public Usuario_Messenger_Enabled As Boolean

Public ViajeDetalle_SeparateRowsByType As Boolean
Public ViajeDetalle_SeparateRowsByStatus As Boolean

Public VehiculoMantenimiento_Aviso As Boolean
Public PersonaAlarma_Aviso As Boolean
Public Alarma_Aviso As Boolean

'VIAJE ESTADO
Public Viaje_EstadoVencido_Check As Boolean
Public Viaje_EstadoVencido_CheckIntervalSeconds As Long

Public Report_Preview As Boolean

Public GridEvenRowBackColor As Long
Public GridEvenRowForeColor As Long

Public GridOddRowBackColor As Long
Public GridOddRowForeColor As Long

'/////////////////////////////////////////////////////////////
'LOCAL VARIABLES AND OBJECTS
'/////////////////////////////////////////////////////////////
Private mrecUsuarioParametro As ADODB.Recordset

Public Function LoadSystemParameters() As Boolean
    Dim cmdData As ADODB.command
    Dim recParametro As ADODB.Recordset
    Dim DES As CSC_Encryption_DES
    
    WriteLogEvent "Load System Parameters: INIT", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_Parametro_AllData"
    cmdData.CommandType = adCmdStoredProc
    Set recParametro = New ADODB.Recordset
    recParametro.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    Empresa_ID_Predeterminada = LoadSystemParameterNumberInteger(recParametro, "Empresa_ID_Predeterminada", 1)
    LogAccion_Enabled = LoadSystemParameterBoolean(recParametro, "LogAccion_Enabled", False)
    System_SyncDateTimeWithDBServer_Enabled = LoadSystemParameterBoolean(recParametro, "SyncDateTimeWithDBServer_Enabled", False)
    Messenger_Enabled = LoadSystemParameterBoolean(recParametro, "Messenger_Enabled", True)
    
    '////////////////////////////////////////////////////////
    'COMPAÑIA
    CompanyName = LoadSystemParameterText(recParametro, "CompanyName", "")
    If CompanyName <> "" Then
        Set DES = New CSC_Encryption_DES
        CompanyName = DES.DecryptString(CompanyName, PASSWORD_ENCRYPTION_KEY)
        Set DES = Nothing
    End If
    
    '////////////////////////////////////////////////////////
    'PERSONA
    Persona_DatoIncompleto_AvisoTipo = LoadSystemParameterNumberInteger(recParametro, "Persona_DatoIncompleto_AvisoTipo", 0)
    Persona_DatoIncompleto_RegistroTodos = LoadSystemParameterBoolean(recParametro, "Persona_DatoIncompleto_RegistroTodos", True)
    Persona_DatoIncompleto_CampoDocumento = LoadSystemParameterBoolean(recParametro, "Persona_DatoIncompleto_CampoDocumento", True)
    Persona_DatoIncompleto_CampoDomicilio = LoadSystemParameterBoolean(recParametro, "Persona_DatoIncompleto_CampoDomicilio", True)
    '////////////////////////////////////////////////////////
    
    '////////////////////////////////////////////////////////
    'VIAJE ESTADO
    Viaje_Estado_Activo_Color = LoadSystemParameterNumberInteger(recParametro, "Viaje_Estado_Activo_Color", vbBlack)
    Viaje_Estado_Activo_Bold = LoadSystemParameterBoolean(recParametro, "Viaje_Estado_Activo_Bold", False)
    Viaje_Estado_EnProgreso_Color = LoadSystemParameterNumberInteger(recParametro, "Viaje_Estado_EnProgreso_Color", 49152)
    Viaje_Estado_EnProgreso_Bold = LoadSystemParameterBoolean(recParametro, "Viaje_Estado_EnProgreso_Bold", True)
    Viaje_Estado_Finalizado_Color = LoadSystemParameterNumberInteger(recParametro, "Viaje_Estado_Finalizado_Color", 12582912)
    Viaje_Estado_Finalizado_Bold = LoadSystemParameterBoolean(recParametro, "Viaje_Estado_Finalizado_Bold", True)
    Viaje_Estado_Cancelado_Color = LoadSystemParameterNumberInteger(recParametro, "Viaje_Estado_Cancelado_Color", vbRed)
    Viaje_Estado_Cancelado_Bold = LoadSystemParameterBoolean(recParametro, "Viaje_Estado_Cancelado_Bold", True)
    Viaje_Normal_Estado_Vencido_Horas = LoadSystemParameterNumberInteger(recParametro, "Viaje_Normal_Estado_Vencido_Horas", 7)
    Viaje_Especial_Estado_Vencido_Horas = LoadSystemParameterNumberInteger(recParametro, "Viaje_Especial_Estado_Vencido_Horas", 168)
    Viaje_Especial_MostrarEnTodasLasRutas = LoadSystemParameterBoolean(recParametro, "Viaje_Especial_MostrarEnTodasLasRutas", False)
    Viaje_Especial_Color = LoadSystemParameterNumberInteger(recParametro, "Viaje_Especial_Color", vbBlack)
    Viaje_Especial_Bold = LoadSystemParameterBoolean(recParametro, "Viaje_Especial_Bold", True)
    ViajeDetalle_CreadoEnProgreso_Color = LoadSystemParameterNumberInteger(recParametro, "ViajeDetalle_CreadoEnProgreso_Color", vbBlue)
    ViajeDetalle_ModificadoEnProgreso_Color = LoadSystemParameterNumberInteger(recParametro, "ViajeDetalle_ModificadoEnProgreso_Color", vbRed)
    
    ViajeDetalle_Transferencia_RangoMaximoMinutos = LoadSystemParameterNumberInteger(recParametro, "ViajeDetalle_Transferencia_RangoMaximoMinutos", 10)
    
    ViajeDetalle_MedioPago_UtilizaOperacion_ObligaFacturaNumero = LoadSystemParameterBoolean(recParametro, "ViajeDetalle_MedioPago_UtilizaOperacion_ObligaFacturaNumero", True)
    
    '////////////////////////////////////////////////////////
    'CUENTA CORRIENTE
    CuentaCorrienteGrupo_ID_ViajeDebito = LoadSystemParameterNumberInteger(recParametro, "CuentaCorrienteGrupo_ID_ViajeDebito", 0)
    CuentaCorrienteGrupo_ID_ViajeCredito = LoadSystemParameterNumberInteger(recParametro, "CuentaCorrienteGrupo_ID_ViajeCredito", 0)
    CuentaCorrienteGrupo_ID_Transferencia = LoadSystemParameterNumberInteger(recParametro, "CuentaCorrienteGrupo_ID_Transferencia", 0)
    CuentaCorrienteGrupo_ID_Sueldo = LoadSystemParameterNumberInteger(recParametro, "CuentaCorrienteGrupo_ID_Sueldo", 0)
    CuentaCorrienteCaja_ID_ViajeDebito = LoadSystemParameterNumberInteger(recParametro, "CuentaCorrienteCaja_ID_ViajeDebito", 0)
    CuentaCorriente_MovimientoAnterior_Minutos = LoadSystemParameterNumberInteger(recParametro, "CuentaCorriente_MovimientoAnterior_Minutos", 600)
    MedioPago_Predeterminado_ID = LoadSystemParameterNumberInteger(recParametro, "MedioPago_Predeterminado_ID", 1)
    
    '////////////////////////////////////////////////////////
    'PRECIOS
    ListaPrecio_ID_Predeterminada = LoadSystemParameterNumberInteger(recParametro, "ListaPrecio_ID_Predeterminada", 0)
    ViajeDetalle_Pasajero_Precio_PermitirModificar = LoadSystemParameterBoolean(recParametro, "ViajeDetalle_Pasajero_Precio_PermitirModificar", False)
    ViajeDetalle_Comision_Precio_PermitirModificar = LoadSystemParameterBoolean(recParametro, "ViajeDetalle_Comision_Precio_PermitirModificar", True)
    Persona_Prepago_AvisoCantidadRestante = LoadSystemParameterNumberInteger(recParametro, "Persona_Prepago_AvisoCantidadRestante", 3)
    
    '////////////////////////////////////////////////////////
    'VARIOS
    Empresa_ID_Predeterminada = LoadSystemParameterNumberInteger(recParametro, "Empresa_ID_Predeterminada", 0)
    UsuarioGrupo_ID_Predeterminado = LoadSystemParameterNumberInteger(recParametro, "UsuarioGrupo_ID_Predeterminado", 0)
    LugarGrupo_ID_Otro = LoadSystemParameterNumberInteger(recParametro, "LugarGrupo_ID_Otro", 0)
    Lugar_ID_Otro = LoadSystemParameterNumberInteger(recParametro, "Lugar_ID_Otro", 0)
    Ruta_ID_Otra = LoadSystemParameterText(recParametro, "Ruta_ID_Otra", "ESPECIAL")
    Ruta_Paquete_ID = LoadSystemParameterText(recParametro, "Ruta_Paquete_ID", "PAQUETE")
    ViajeDetalle_Paquete_Permite_Multiples_Pagos = LoadSystemParameterBoolean(recParametro, "ViajeDetalle_Paquete_Permite_Multiples_Pagos", False)
    TelefonoTipo_ID_Otro = LoadSystemParameterNumberInteger(recParametro, "TelefonoTipo_ID_Otro", 0)
    Vehiculo_TelefonoTipo_ID = LoadSystemParameterNumberInteger(recParametro, "Vehiculo_TelefonoTipo_ID", 0)
    Recordset_MaxRecords = LoadSystemParameterNumberInteger(recParametro, "Recordset_MaxRecords", 1000)
    Pasajero_AvisoInasistencia_Veces = LoadSystemParameterNumberInteger(recParametro, "Pasajero_AvisoInasistencia_Veces", 4)
    Pasajero_AvisoInasistencia_Dias = LoadSystemParameterNumberInteger(recParametro, "Pasajero_AvisoInasistencia_Dias", 180)
    Provincia_ID_Predeterminada = LoadSystemParameterText(recParametro, "Provincia_ID_Predeterminada", "")
    Viaje_Actual_Rango_Minutos = LoadSystemParameterNumberInteger(recParametro, "Viaje_Actual_Rango_Minutos", 30)
    Viaje_Permite_2_Conductores = LoadSystemParameterBoolean(recParametro, "Viaje_Permite_2_Conductores", False)
    Viaje_Permite_RutaConexion = LoadSystemParameterBoolean(recParametro, "Viaje_Permite_RutaConexion", False)
    Viaje_RutaConexion_TiempoEsperaMaximo_Minutos = LoadSystemParameterNumberInteger(recParametro, "Viaje_RutaConexion_TiempoEsperaMaximo_Minutos", 30)
    Permitir_Reservas_Condicionales = LoadSystemParameterBoolean(recParametro, "Permitir_Reservas_Condicionales", False)
    
    '////////////////////////////////////////////////////////
    'COMISIONES
    Comision_Habilitar = LoadSystemParameterBoolean(recParametro, "Comision_Habilitar", True)
    Comision_Seguro_Habilitar = LoadSystemParameterBoolean(recParametro, "Comision_Seguro_Habilitar", False)
    Comision_Seguro_ValorDeclaradoMinimo = LoadSystemParameterCurrency(recParametro, "Comision_Seguro_ValorDeclaradoMinimo", 300)
    Comision_Seguro_ValorDeclaradoMaximo = LoadSystemParameterCurrency(recParametro, "Comision_Seguro_ValorDeclaradoMaximo", 3000)
    Comision_Seguro_ValorDeclaradoMinimo_Preasignar = LoadSystemParameterBoolean(recParametro, "Comision_Seguro_ValorDeclaradoMinimo_Preasignar", True)
    Comision_Seguro_Porcentaje = LoadSystemParameterNumberDecimal(recParametro, "Comision_Seguro_Porcentaje", 0.06)
    Comision_Seguro_PorcentajeAplicarSobreTotal = LoadSystemParameterBoolean(recParametro, "Comision_Seguro_PorcentajeAplicarSobreTotal", False)
    Comision_Seguro_RedondeoDecimales = LoadSystemParameterNumberInteger(recParametro, "Comision_Seguro_RedondeoDecimales", 1)
    
    '////////////////////////////////////////////////////////
    'INTERNET DATA
    InternetDataEmail_SenderDisplayName = LoadSystemParameterText(recParametro, "InternetDataEmail_SenderDisplayName", "")
    InternetDataEmail_SenderAddress = LoadSystemParameterText(recParametro, "InternetDataEmail_SenderAddress", "")
    InternetDataEmail_SMTPHost = LoadSystemParameterText(recParametro, "InternetDataEmail_SMTPHost", "")
    InternetDataEmail_SMTPUserName = LoadSystemParameterText(recParametro, "InternetDataEmail_SMTPUserName", "")
    InternetDataEmail_SMTPPassword = LoadSystemParameterText(recParametro, "InternetDataEmail_SMTPPassword", "")
    If InternetDataEmail_SMTPPassword <> "" Then
        Set DES = New CSC_Encryption_DES
        InternetDataEmail_SMTPPassword = DES.DecryptString(InternetDataEmail_SMTPPassword, PASSWORD_ENCRYPTION_KEY)
        Set DES = Nothing
    End If
    
    '////////////////////////////////////////////////////////
    'PLANILLA POR EMAIL
    PlanillaViajeEmail_SendReportFormat = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_SendReportFormat", crEFTPortableDocFormat)
    PlanillaViajeEmail_SendExcel = LoadSystemParameterBoolean(recParametro, "PlanillaViajeEmail_SendExcel", True)
    PlanillaViajeEmail_SMTPHost = LoadSystemParameterText(recParametro, "PlanillaViajeEmail_SMTPHost", "")
    PlanillaViajeEmail_SMTPUserName = LoadSystemParameterText(recParametro, "PlanillaViajeEmail_SMTPUserName", "")
    PlanillaViajeEmail_SMTPPassword = LoadSystemParameterText(recParametro, "PlanillaViajeEmail_SMTPPassword", "")
    If PlanillaViajeEmail_SMTPPassword <> "" Then
        Set DES = New CSC_Encryption_DES
        PlanillaViajeEmail_SMTPPassword = DES.DecryptString(PlanillaViajeEmail_SMTPPassword, PASSWORD_ENCRYPTION_KEY)
        Set DES = Nothing
    End If
    PlanillaViajeEmail_Margin_Top = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_Margin_Top", 0)
    PlanillaViajeEmail_Margin_Left = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_Margin_Left", 0)
    PlanillaViajeEmail_Margin_Right = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_Margin_Right", 0)
    PlanillaViajeEmail_Margin_Bottom = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_Margin_Bottom", 0)
    PlanillaViajeEmail_ColumnWidth_A = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_ColumnWidth_A", 3)
    PlanillaViajeEmail_ColumnWidth_B = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_ColumnWidth_B", 11)
    PlanillaViajeEmail_ColumnWidth_C = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_ColumnWidth_C", 11)
    PlanillaViajeEmail_ColumnWidth_D = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_ColumnWidth_D", 13.71)
    PlanillaViajeEmail_ColumnWidth_E = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_ColumnWidth_E", 13.71)
    PlanillaViajeEmail_ColumnWidth_F = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_ColumnWidth_F", 7.14)
    PlanillaViajeEmail_ColumnWidth_G = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_ColumnWidth_G", 7.14)
    PlanillaViajeEmail_ColumnWidth_H = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_ColumnWidth_H", 7.14)
    PlanillaViajeEmail_ColumnWidth_I = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_ColumnWidth_I", 2)
    PlanillaViajeEmail_ColumnWidth_J = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_ColumnWidth_J", 20)
    PlanillaViajeEmail_ColumnWidth_K = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_ColumnWidth_K", 0.08)
    PlanillaViajeEmail_RowHeight_2 = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_RowHeight_2", 12.75)
    PlanillaViajeEmail_RowHeight_4 = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_RowHeight_4", 12.75)
    PlanillaViajeEmail_RowHeight_6 = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_RowHeight_6", 12.75)
    PlanillaViajeEmail_RowHeight_8 = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_RowHeight_8", 12.75)
    PlanillaViajeEmail_RowHeight_35 = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_RowHeight_35", 12.75)
    PlanillaViajeEmail_RowHeight_37 = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_RowHeight_37", 12.75)
    PlanillaViajeEmail_RowHeight_ColumnHeader = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_RowHeight_ColumnHeader", 17.25)
    PlanillaViajeEmail_RowHeight_Detalle = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_RowHeight_Detalle", 23.25)
    
    '////////////////////////////////////////////////////////
    'LISTA DE PASAJEROS
    ListaPasajero_ArchivoNombre = LoadSystemParameterText(recParametro, "ListaPasajero_ArchivoNombre", "")
    ListaPasajero_RowStart = LoadSystemParameterNumberInteger(recParametro, "ListaPasajero_RowStart", 0)
    ListaPasajero_ColumnPasajero = LoadSystemParameterText(recParametro, "ListaPasajero_ColumnPasajero", "")
    ListaPasajero_ColumnDocumento = LoadSystemParameterText(recParametro, "ListaPasajero_ColumnDocumento", "")
    
    '////////////////////////////////////////////////////////
    'PERSONAL
    Personal_Status_Global = LoadSystemParameterBoolean(recParametro, "Personal_Status_Global", False)
    Personal_Status_Global_CheckInterval_Seconds = LoadSystemParameterNumberInteger(recParametro, "Personal_Status_Global_CheckInterval_Seconds", 15)
    Personal_Status_Persistent = LoadSystemParameterBoolean(recParametro, "Personal_Status_Persistent", True)
    Personal_Status_OnStartup = LoadSystemParameterBoolean(recParametro, "Personal_Status_OnStartup", False)
    
    '////////////////////////////////////////////////////////
    'SISTEMA ON LINE
    WebSite_HyperLink_Main = LoadSystemParameterText(recParametro, "WebSite_HyperLink_Main", "")
    WebSite_HyperLink_Login = LoadSystemParameterText(recParametro, "WebSite_HyperLink_Login", "")
    WebSite_HyperLink_ChangePassword = LoadSystemParameterText(recParametro, "WebSite_HyperLink_ChangePassword", "")
    WebSite_Persona_EmailBienvenida_Asunto = LoadSystemParameterText(recParametro, "WebSite_Persona_EmailBienvenida_Asunto", "")
    WebSite_Persona_EmailBienvenida_Texto = LoadSystemParameterText(recParametro, "WebSite_Persona_EmailBienvenida_Texto", "")
    
    '////////////////////////////////////////////////////////
    'AGENT
    InternetData_Send_Interval_Seconds = LoadSystemParameterNumberInteger(recParametro, "InternetData_Send_Interval_Seconds", 120)
    PersonaAdministrativo_Sueldo_Interval_Seconds = LoadSystemParameterNumberInteger(recParametro, "PersonaAdministrativo_Sueldo_Interval_Seconds", 21600)
    PlanillaViajeEmail_Send_Interval_Seconds = LoadSystemParameterNumberInteger(recParametro, "PlanillaViajeEmail_Send_Interval_Seconds", 30)
    Viaje_AutoPersonal_Interval_Seconds = LoadSystemParameterNumberInteger(recParametro, "Viaje_AutoPersonal_Interval_Seconds", 3600)
    Viaje_AutoPersonal_AfterNumberOfDays = LoadSystemParameterNumberInteger(recParametro, "Viaje_AutoPersonal_AfterNumberOfDays", 1)
    Email_Delete_Interval_Seconds = LoadSystemParameterNumberInteger(recParametro, "Email_Delete_Interval_Seconds", 43200)
    Email_Delete_AfterNumberOfDays = LoadSystemParameterNumberInteger(recParametro, "Email_Delete_AfterNumberOfDays", 30)
    
    recParametro.Close
    Set recParametro = Nothing
    
    LoadSystemParameters = True
    WriteLogEvent "Load System Parameters: DONE", vbLogEventTypeInformation, Logs_MonthsToKeep
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.Parametro.LoadSystemParameters", "Error al obtener los Parámetros del Sistema."
End Function

Public Function LoadWorkstationParameters() As Boolean
    Dim DES As CSC_Encryption_DES
    
    'WriteLogEvent "Load Workstation Parameters: INIT", vbLogEventTypeInformation
    
    If Config_Type = "" Then
        'PRIMERA VEZ, DETERMINAR EL TIPO DE CONFIGURACION
        If Config_Name = "" Then
            'NO SE ESPECIFICO NOMBRE, BUSCO UN INI CON EL MISMO NOMBRE QUE EL EXE
            If FileSystem.Dir(App.Path & "\" & App.EXEName & ".ini") = "" Then
                Config_Type = "REG"
            Else
                Config_Type = "INI"
                Config_INIFile_FullPathAndName = App.Path & "\" & App.EXEName & ".ini"
            End If
        Else
            'SE ESPECIFICÓ NOMBRE, BUSCO ESE INI
            If FileSystem.Dir(App.Path & "\" & Config_Name & ".ini") = "" Then
                Config_Type = "REG"
            Else
                Config_Type = "INI"
                Config_INIFile_FullPathAndName = App.Path & "\" & Config_Name & ".ini"
            End If
        End If
    End If
    If Config_Type = "REG" Then
        'GET SETTINGS FROM REGISTRY
        If Config_Name = "" Then
            Config_Name = CSM_Registry.GetValue_FromApplication_LocalMachine("Config", "Current", "", csrdtString)
        End If
    
        Database_Provider = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "Provider", "", csrdtString)
        Database_ConnectionTimeout = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "ConnectionTimeout", 10, csrdtNumberInteger)
        Database_CommandTimeout = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "CommandTimeout", 10, csrdtNumberInteger)
        Database_PacketSize = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "PacketSize", 4096, csrdtNumberInteger)
        Database_DataTypeCompatibility = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "DataTypeCompatibility", 80, csrdtNumberInteger)
        Database_DataSource = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "DataSource", "", csrdtString)
        Database_FailoverPartner = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "FailoverPartner", "", csrdtString)
        Database_UserID = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "UserID", "", csrdtString)
        Database_Password = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "Password", "", csrdtString)
        If Database_Password <> "" Then
            Set DES = New CSC_Encryption_DES
            Database_Password = DES.DecryptString(Database_Password, PASSWORD_ENCRYPTION_KEY)
            Set DES = Nothing
        End If
        Database_Database = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "Database", "", csrdtString)
        Database_DatabaseHistory = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "DatabaseHistory", Database_Database & "_History", csrdtString)
        
        Database_CursorLocation = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "CursorLocation", adUseClient, csrdtNumberInteger)
        Database_CursorType = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "CursorType", adOpenKeyset, csrdtNumberInteger)
        
        Logs_Path = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "LogsPath", App.Path & "\", csrdtString)
        Logs_FileNameTemplate = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "LogsFileNameTemplate", "@USERNAME@_", csrdtString)
        Logs_MonthsToKeep = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "LogsMonthsToKeep", 3, csrdtNumberInteger)
        
        'Report_Provider = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "ReportsProvider", Database_Provider, csrdtString)
        Report_Path = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "ReportsPath", "", csrdtString)
        If Right(Report_Path, 1) <> "\" Then
            Report_Path = Report_Path & "\"
        End If
        
        DateTime_SyncWithServer_Primary = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "DateTime_SyncWithServer_Primary", True, csrdtBoolean)
        DateTime_SyncWithServer_Secondary = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "DateTime_SyncWithServer_Secondary", True, csrdtBoolean)
        DateTime_SyncWithServer_IntervalSeconds = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "DateTime_SyncWithServer_IntervalSeconds", 3600, csrdtNumberInteger)
    
        Interface_MDIBackgroundColor = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "MDIBackgroundColor", vbApplicationWorkspace, csrdtNumberInteger)
        Interface_CompanyLogo = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "CompanyLogo", "", csrdtString)
        Interface_CompanyLogoMini = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "CompanyLogoMini", "", csrdtString)
        
        System_MaintenanceMode = CSM_Registry.GetValue_FromApplication_LocalMachine("Config\" & Config_Name, "System_MaintenanceMode", False, csrdtBoolean)
    Else
        'GET SETTINGS FROM INI FILE
        Database_Provider = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "Provider", "", csidtString)
        Database_ConnectionTimeout = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "ConnectionTimeout", 10, csidtNumberInteger)
        Database_CommandTimeout = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "CommandTimeout", 10, csidtNumberInteger)
        Database_PacketSize = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "PacketSize", 4096, csidtNumberInteger)
        Database_DataTypeCompatibility = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "DataTypeCompatibility", 80, csidtNumberInteger)
        Database_DataSource = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "DataSource", "", csidtString)
        Database_FailoverPartner = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "FailoverPartner", "", csidtString)
        Database_UserID = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "UserID", "", csidtString)
        Database_Password = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "Password", "", csidtString)
        If Database_Password <> "" Then
            Set DES = New CSC_Encryption_DES
            Database_Password = DES.DecryptString(Database_Password, PASSWORD_ENCRYPTION_KEY)
            Set DES = Nothing
        End If
        Database_Database = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "Database", "", csidtString)
        Database_DatabaseHistory = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "DatabaseHistory", Database_Database & "_History", csidtString)
        
        Database_CursorLocation = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "CursorLocation", adUseClient, csidtNumberInteger)
        Database_CursorType = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Database", "CursorType", adOpenKeyset, csidtNumberInteger)
        
        Logs_Path = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Logs", "Path", App.Path & "\", csidtString)
        Logs_FileNameTemplate = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Logs", "FileNameTemplate", "@USERNAME@_", csidtString)
        Logs_MonthsToKeep = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Logs", "MonthsToKeep", 3, csidtNumberInteger)
        
        'Report_Provider = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Reports", "Provider", Database_Provider, csidtString)
        Report_Path = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Reports", "Path", "", csidtString)
        If Right(Report_Path, 1) <> "\" Then
            Report_Path = Report_Path & "\"
        End If
        
        DateTime_SyncWithServer_Primary = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "DateTime", "SyncWithServer_Primary", True, csidtBoolean)
        DateTime_SyncWithServer_Secondary = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "DateTime", "SyncWithServer_Secondary", True, csidtBoolean)
        DateTime_SyncWithServer_IntervalSeconds = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "DateTime", "SyncWithServer_IntervalSeconds", 3600, csidtNumberInteger)
        
        Interface_MDIBackgroundColor = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Interface", "MDIBackgroundColor", vbApplicationWorkspace, csidtNumberInteger)
        Interface_CompanyLogo = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Interface", "CompanyLogo", "", csidtString)
        Interface_CompanyLogoMini = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "Interface", "CompanyLogoMini", "", csidtString)
        
        System_MaintenanceMode = CSM_INI.GetValue(Config_INIFile_FullPathAndName, "System", "MaintenanceMode", False, csidtBoolean)
    End If
    
    Telephony_Type = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "Type", "NONE", csrdtString)
    
    RegistroLlamada_Save = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "CallerIDSave", False, csrdtBoolean)
    
    IDSucursal = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Misc", "IDSucursal", "", csrdtString)
    
    RefreshList_Enabled = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\RefreshLists", "Enabled", True, csrdtBoolean)
    Workstation_SyncDateTimeWithDBServer_Enabled = CSM_Registry.GetValue_FromApplication_LocalMachine("Options", "SyncDateTimeWithDBServer", False, csrdtBoolean)
    
    RefreshList_Fastest_CheckInterval_Seconds = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\RefreshLists", "Fastest_CheckIntervalSeconds", 3, csrdtNumberInteger)
    RefreshList_Slowest_CheckInterval_Seconds = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\RefreshLists", "Slowest_CheckIntervalSeconds", 60, csrdtNumberInteger)
    
    'AGENT
    StartupDelay_Seconds = CSM_Registry.GetValue_FromApplication_LocalMachine("Options\Agent", "StartupDelay_Seconds", 5, csrdtNumberInteger)
    
    LoadWorkstationParameters = True
    'WriteLogEvent "Load Workstation Parameters: DONE", vbLogEventTypeInformation
End Function

Public Function InitializeUserParameters() As Boolean
    Screen.MousePointer = vbHourglass
    
    WriteLogEvent "Initialize User Parameters: INIT", vbLogEventTypeInformation, pParametro.LogAccion_Enabled

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Set mrecUsuarioParametro = New ADODB.Recordset
    mrecUsuarioParametro.Source = "SELECT * FROM UsuarioParametro WHERE IDUsuario = " & pUsuario.IDUsuario
    mrecUsuarioParametro.Open , pDatabase.Connection, pParametro.Database_CursorType, adLockOptimistic

    Screen.MousePointer = vbDefault
    InitializeUserParameters = True
    
    WriteLogEvent "Initialize User Parameters: DONE", vbLogEventTypeInformation, LogAccion_Enabled
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Parametro.Initialize", "Error al Abrir la Tabla de Parámetros de los Usuarios."
End Function

Public Function LoadUserParameters() As Boolean
    WriteLogEvent "Load User Parameters: INIT", vbLogEventTypeInformation, LogAccion_Enabled

    ListView_GridLines = Usuario_LeerSiNo("ListView_GridLines", True)
    Usuario_Messenger_Enabled = Usuario_LeerSiNo("Messenger_Enabled", True)
    
    ViajeDetalle_SeparateRowsByType = Usuario_LeerSiNo("ViajeDetalle_SeparateRowsByType", True)
    ViajeDetalle_SeparateRowsByStatus = Usuario_LeerSiNo("ViajeDetalle_SeparateRowsByStatus", True)
    VehiculoMantenimiento_Aviso = Usuario_LeerSiNo("VehiculoMantenimiento_Aviso", False)
    PersonaAlarma_Aviso = Usuario_LeerSiNo("PersonaAlarma_Aviso", False)
    Alarma_Aviso = Usuario_LeerSiNo("Alarma_Aviso", False)
    Viaje_EstadoVencido_Check = Usuario_LeerSiNo("Viaje_EstadoVencido_Check", False)
    Viaje_EstadoVencido_CheckIntervalSeconds = Usuario_LeerNumero("Viaje_EstadoVencido_CheckIntervalSeconds", 1800)
    Persona_Apellido_Busqueda_Delay_Milliseconds = Usuario_LeerNumero("Persona_Apellido_Busqueda_Delay_Milliseconds", 1500)
    
    Report_Preview = Usuario_LeerSiNo("Report_Preview", True)
    
    ' Grid Colors
    GridEvenRowBackColor = Usuario_LeerNumero("GridEvenRowBackColor", vbButtonFace)
    GridEvenRowForeColor = Usuario_LeerNumero("GridEvenRowForeColor", vbWindowText)
    GridOddRowBackColor = Usuario_LeerNumero("GridOddRowBackColor", vbWindowBackground)
    GridOddRowForeColor = Usuario_LeerNumero("GridOddRowForeColor", vbWindowText)
    
    LoadUserParameters = True
    WriteLogEvent "Load User Parameters: DONE", vbLogEventTypeInformation, LogAccion_Enabled
End Function

Public Sub LoadParameters()
    LoadSystemParameters
    LoadWorkstationParameters
    LoadUserParameters
End Sub

Public Sub SaveSystemParameterText(ByVal IDParametro As String, ByVal value As String)
    Dim recParametro As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recParametro = New ADODB.Recordset
    recParametro.Source = "SELECT * FROM Parametro WHERE IDParametro = '" & IDParametro & "'"
    recParametro.Open , pDatabase.Connection, adOpenKeyset, adLockOptimistic, adCmdText
    
    If recParametro.EOF Then
        recParametro.AddNew
        recParametro("IDParametro").value = IDParametro
    End If
    recParametro("Texto").value = IIf(Trim(value) = "", Null, value)
    recParametro.Update
    
    recParametro.Close
    Set recParametro = Nothing
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Parametro.SaveSystemParameterText", "Error al grabar el Parámetro del Sistema." & vbCr & vbCr & "IDParametro: " & IDParametro
End Sub

Public Sub SaveSystemParameterBoolean(ByVal IDParametro As String, ByVal value As Boolean)
    Dim recParametro As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recParametro = New ADODB.Recordset
    recParametro.Source = "SELECT * FROM Parametro WHERE IDParametro = '" & IDParametro & "'"
    recParametro.Open , pDatabase.Connection, adOpenKeyset, adLockOptimistic, adCmdText
    
    If recParametro.EOF Then
        recParametro.AddNew
        recParametro("IDParametro").value = IDParametro
    End If
    recParametro("SiNo").value = value
    recParametro.Update
    
    recParametro.Close
    Set recParametro = Nothing
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Parametro.SaveSystemParameterBoolean", "Error al grabar el Parámetro del Sistema." & vbCr & vbCr & "IDParametro: " & IDParametro
End Sub

Public Sub SaveSystemParameterNumberInteger(ByVal IDParametro As String, ByVal value As Long)
    Dim recParametro As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recParametro = New ADODB.Recordset
    recParametro.Source = "SELECT * FROM Parametro WHERE IDParametro = '" & IDParametro & "'"
    recParametro.Open , pDatabase.Connection, adOpenKeyset, adLockOptimistic, adCmdText
    
    If recParametro.EOF Then
        recParametro.AddNew
        recParametro("IDParametro").value = IDParametro
    End If
    recParametro("NumeroEntero").value = value
    recParametro.Update
    
    recParametro.Close
    Set recParametro = Nothing
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Parametro.SaveSystemParameterNumberInteger", "Error al grabar el Parámetro del Sistema." & vbCr & vbCr & "IDParametro: " & IDParametro
End Sub

Public Sub SaveSystemParameterNumberDecimal(ByVal IDParametro As String, ByVal value As Single)
    Dim recParametro As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recParametro = New ADODB.Recordset
    recParametro.Source = "SELECT * FROM Parametro WHERE IDParametro = '" & IDParametro & "'"
    recParametro.Open , pDatabase.Connection, adOpenKeyset, adLockOptimistic, adCmdText
    
    If recParametro.EOF Then
        recParametro.AddNew
        recParametro("IDParametro").value = IDParametro
    End If
    recParametro("NumeroDecimal").value = value
    recParametro.Update
    
    recParametro.Close
    Set recParametro = Nothing
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Parametro.SaveSystemParameterNumberDecimal", "Error al grabar el Parámetro del Sistema." & vbCr & vbCr & "IDParametro: " & IDParametro
End Sub

'///////////////////////////////////////////////////////////////////////
' PARAMETROS DEL USUARIO - INICIO
'///////////////////////////////////////////////////////////////////////

Private Function Usuario_Guardar_Inicializar(ByVal IDParametro As String) As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    mrecUsuarioParametro.Filter = "IDUsuario = " & pUsuario.IDUsuario & " AND IDParametro = '" & IDParametro & "'"
    If mrecUsuarioParametro.EOF Then
        mrecUsuarioParametro.AddNew
        mrecUsuarioParametro("IDUsuario").value = pUsuario.IDUsuario
        mrecUsuarioParametro("IDParametro").value = IDParametro
    End If
    Usuario_Guardar_Inicializar = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Parametro.Usuario_Guardar_Inicializar", "Error al grabar el Parámetro del Usuario." & vbCr & vbCr & "IDParametro: " & IDParametro, , , False
End Function

Private Function Usuario_Guardar_Terminar(ByVal IDParametro As String) As Boolean

RESTART:
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    mrecUsuarioParametro.Update
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_VALUE_CHANGED_SINCE_LAST_READ Then
        mrecUsuarioParametro.CancelUpdate
        mrecUsuarioParametro.Requery
        Resume RESTART
    Else
        ShowErrorMessage "Classes.Parametro.Usuario_Guardar_Terminar", "Error al grabar el Parámetro del Usuario." & vbCr & vbCr & "IDParametro: " & IDParametro, , , False
    End If
End Function

' ======  TEXTO  ======

Public Function Usuario_LeerTexto(ByVal IDParametro As String, Optional ByVal ValorPredeterminado As String) As String
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    mrecUsuarioParametro.Filter = "IDUsuario = " & pUsuario.IDUsuario & " AND IDParametro = '" & IDParametro & "'"
    If mrecUsuarioParametro.EOF Then
        Usuario_LeerTexto = ValorPredeterminado
    Else
        Usuario_LeerTexto = mrecUsuarioParametro("Texto").value
    End If
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Parametro.Usuario_LeerTexto", "Error al obtener el Parametro del Usuario." & vbCrLf & vbCrLf & "IDParametro: " & IDParametro
End Function

Public Sub Usuario_GuardarTexto(ByVal IDParametro As String, ByVal Valor As String)
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Usuario_Guardar_Inicializar(IDParametro) Then
        mrecUsuarioParametro("Texto").value = IIf(Trim(Valor) = "", Null, Valor)
        Call Usuario_Guardar_Terminar(IDParametro)
    End If
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Parametro.Usuario_GuardarTexto", "Error al grabar el Parámetro del Usuario." & vbCr & vbCr & "IDParametro: " & IDParametro, , , False
End Sub

' ======  NUMERO  ======

Public Function Usuario_LeerNumero(ByVal IDParametro As String, Optional ByVal ValorPredeterminado As Double) As Double
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    mrecUsuarioParametro.Filter = "IDUsuario = " & pUsuario.IDUsuario & " AND IDParametro = '" & IDParametro & "'"
    If mrecUsuarioParametro.EOF Then
        Usuario_LeerNumero = ValorPredeterminado
    Else
        Usuario_LeerNumero = mrecUsuarioParametro("Numero").value
    End If
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Parametro.Usuario_LeerNumero", "Error al obtener el Parametro del Usuario." & vbCrLf & vbCrLf & "IDParametro: " & IDParametro
End Function

Public Sub Usuario_GuardarNumero(ByVal IDParametro As String, ByVal Valor As Double)
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Usuario_Guardar_Inicializar(IDParametro) Then
        mrecUsuarioParametro("Numero").value = Valor
        Call Usuario_Guardar_Terminar(IDParametro)
    End If
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Parametro.Usuario_GuardarTexto", "Error al grabar el Parámetro del Usuario." & vbCr & vbCr & "IDParametro: " & IDParametro, , , False
End Sub

' ======  FECHAHORA  ======

Public Function Usuario_LeerFechaHora(ByVal IDParametro As String, Optional ByVal ValorPredeterminado As Double) As Double
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    mrecUsuarioParametro.Filter = "IDUsuario = " & pUsuario.IDUsuario & " AND IDParametro = '" & IDParametro & "'"
    If mrecUsuarioParametro.EOF Then
        Usuario_LeerFechaHora = ValorPredeterminado
    Else
        Usuario_LeerFechaHora = mrecUsuarioParametro("FechaHora").value
    End If
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Parametro.Usuario_LeerFechaHora", "Error al obtener el Parametro del Usuario." & vbCrLf & vbCrLf & "IDParametro: " & IDParametro
End Function

Public Sub Usuario_GuardarFechaHora(ByVal IDParametro As String, ByVal Valor As Date)
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Usuario_Guardar_Inicializar(IDParametro) Then
        mrecUsuarioParametro("FechaHora").value = Valor
        Call Usuario_Guardar_Terminar(IDParametro)
    End If
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Parametro.Usuario_GuardarFechaHora", "Error al grabar el Parámetro del Usuario." & vbCr & vbCr & "IDParametro: " & IDParametro, , , False
End Sub

' ======  SI / NO  ======

Public Function Usuario_LeerSiNo(ByVal IDParametro As String, Optional ByVal ValorPredeterminado As Boolean) As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    mrecUsuarioParametro.Filter = "IDParametro = '" & IDParametro & "'"
    If mrecUsuarioParametro.EOF Then
        Usuario_LeerSiNo = ValorPredeterminado
    Else
        Usuario_LeerSiNo = mrecUsuarioParametro("SiNo").value
    End If
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Parametro.Usuario_LeerSiNo", "Error al obtener el Parametro del Usuario." & vbCrLf & vbCrLf & "IDParametro: " & IDParametro
End Function

Public Sub Usuario_GuardarSiNo(ByVal IDParametro As String, ByVal Valor As Boolean)
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Usuario_Guardar_Inicializar(IDParametro) Then
        mrecUsuarioParametro("SiNo").value = Valor
        Call Usuario_Guardar_Terminar(IDParametro)
    End If
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Parametro.Usuario_GuardarSiNo", "Error al grabar el Parámetro del Usuario." & vbCr & vbCr & "IDParametro: " & IDParametro, , , False
End Sub

'///////////////////////////////////////////////////////////////////////
' PARAMETROS DEL USUARIO - FIN
'///////////////////////////////////////////////////////////////////////

Private Sub Class_Terminate()
    If Not mrecUsuarioParametro Is Nothing Then
        If mrecUsuarioParametro.State = adStateOpen Then
            If Not (mrecUsuarioParametro.BOF Or mrecUsuarioParametro.EOF) Then
                If mrecUsuarioParametro.EditMode <> adEditNone Then
                    mrecUsuarioParametro.CancelUpdate
                End If
            End If
            mrecUsuarioParametro.Close
        End If
        Set mrecUsuarioParametro = Nothing
    End If
End Sub

'///////////////////////////////////////////////////////////////////////
' PARAMETROS DEL SISTEMA - INICIO
'///////////////////////////////////////////////////////////////////////

Private Function LoadSystemParameterText(ByRef recParametro As ADODB.Recordset, ByVal ParameterID As String, ByVal DefaultValue As String) As String
    recParametro.Filter = "IDParametro = '" & ParameterID & "'"
    If recParametro.EOF Then
        LoadSystemParameterText = DefaultValue
    Else
        If IsNull(recParametro("Texto").value) Then
            LoadSystemParameterText = DefaultValue
        Else
            LoadSystemParameterText = recParametro("Texto").value
        End If
    End If
End Function

Private Function LoadSystemParameterNumberInteger(ByRef recParametro As ADODB.Recordset, ByVal ParameterID As String, ByVal DefaultValue As Long) As Long
    recParametro.Filter = "IDParametro = '" & ParameterID & "'"
    If recParametro.EOF Then
        LoadSystemParameterNumberInteger = DefaultValue
    Else
        If IsNull(recParametro("NumeroEntero").value) Then
            LoadSystemParameterNumberInteger = DefaultValue
        Else
            LoadSystemParameterNumberInteger = recParametro("NumeroEntero").value
        End If
    End If
End Function

Private Function LoadSystemParameterNumberDecimal(ByRef recParametro As ADODB.Recordset, ByVal ParameterID As String, ByVal DefaultValue As Single) As Single
    recParametro.Filter = "IDParametro = '" & ParameterID & "'"
    If recParametro.EOF Then
        LoadSystemParameterNumberDecimal = DefaultValue
    Else
        If IsNull(recParametro("NumeroDecimal").value) Then
            LoadSystemParameterNumberDecimal = DefaultValue
        Else
            LoadSystemParameterNumberDecimal = recParametro("NumeroDecimal").value
        End If
    End If
End Function

Private Function LoadSystemParameterCurrency(ByRef recParametro As ADODB.Recordset, ByVal ParameterID As String, ByVal DefaultValue As Currency) As Currency
    recParametro.Filter = "IDParametro = '" & ParameterID & "'"
    If recParametro.EOF Then
        LoadSystemParameterCurrency = DefaultValue
    Else
        If IsNull(recParametro("Moneda").value) Then
            LoadSystemParameterCurrency = DefaultValue
        Else
            LoadSystemParameterCurrency = recParametro("Moneda").value
        End If
    End If
End Function

Private Function LoadSystemParameterDateTime(ByRef recParametro As ADODB.Recordset, ByVal ParameterID As String, ByVal DefaultValue As Date) As Date
    recParametro.Filter = "IDParametro = '" & ParameterID & "'"
    If recParametro.EOF Then
        LoadSystemParameterDateTime = DefaultValue
    Else
        If IsNull(recParametro("FechaHora").value) Then
            LoadSystemParameterDateTime = DefaultValue
        Else
            LoadSystemParameterDateTime = recParametro("FechaHora").value
        End If
    End If
End Function

Private Function LoadSystemParameterBoolean(ByRef recParametro As ADODB.Recordset, ByVal ParameterID As String, ByVal DefaultValue As Boolean) As Boolean
    recParametro.Filter = "IDParametro = '" & ParameterID & "'"
    If recParametro.EOF Then
        LoadSystemParameterBoolean = DefaultValue
    Else
        If IsNull(recParametro("SiNo").value) Then
            LoadSystemParameterBoolean = DefaultValue
        Else
            LoadSystemParameterBoolean = recParametro("SiNo").value
        End If
    End If
End Function

'///////////////////////////////////////////////////////////////////////
' PARAMETROS DEL SISTEMA - FIN
'///////////////////////////////////////////////////////////////////////

Public Sub SaveListViewSettings(ByVal Name As String, ByRef ListView As MSComctlLib.ListView)
    Dim ColumnHeader As MSComctlLib.ColumnHeader
    
    On Error GoTo ErrorHandler
    
    DoEvents
    pParametro.Usuario_GuardarNumero Name & "_ListView_SortKey", ListView.SortKey
    pParametro.Usuario_GuardarNumero Name & "_ListView_SortOrder", ListView.SortOrder
    For Each ColumnHeader In ListView.ColumnHeaders
        pParametro.Usuario_GuardarNumero Name & "_ListView_Column_" & ColumnHeader.Key & "_Width", ColumnHeader.Width
        pParametro.Usuario_GuardarNumero Name & "_ListView_Column_" & ColumnHeader.Key & "_Position", ColumnHeader.Position
    Next ColumnHeader
    
ErrorHandler:
End Sub

Public Sub GetListViewSettings(ByVal Name As String, ByRef ListView As MSComctlLib.ListView)
    Dim ColumnHeader As MSComctlLib.ColumnHeader
    
    On Error GoTo ErrorHandler
    
    ListView.SortKey = pParametro.Usuario_LeerNumero(Name & "_ListView_SortKey", ListView.SortKey)
    ListView.SortOrder = pParametro.Usuario_LeerNumero(Name & "_ListView_SortOrder", ListView.SortOrder)
    'TUVE QUE DESHABILITAR LA OPCION DE CAMBIAR LA POSICION DE UNA COLUMNA
    'PORQUE EL ACTIVEX (COMMON CONTROLS) SE CUELGA AL HACERLO CON EL MOUSE
    ListView.AllowColumnReorder = False
    For Each ColumnHeader In ListView.ColumnHeaders
        ColumnHeader.Width = pParametro.Usuario_LeerNumero(Name & "_ListView_Column_" & ColumnHeader.Key & "_Width", ColumnHeader.Width)
        'ColumnHeader.Position = pParametro.Usuario_LeerNumero(Name & "_ListView_Column_" & ColumnHeader.Key & "_Position", ColumnHeader.Position)
    Next ColumnHeader
    
ErrorHandler:
End Sub

Public Sub CopyListViewSettings(ByRef ListViewSource As MSComctlLib.ListView, ByRef ListViewDestination As MSComctlLib.ListView)
    Dim ColumnIndex As Long
    
    On Error GoTo ErrorHandler
    
    ListViewDestination.SortKey = ListViewSource.SortKey
    ListViewDestination.SortOrder = ListViewSource.SortOrder
    For ColumnIndex = 1 To ListViewDestination.ColumnHeaders.Count
        If ColumnIndex > ListViewSource.ColumnHeaders.Count Then
            Exit For
        End If
        ListViewDestination.ColumnHeaders(ColumnIndex).Width = ListViewSource.ColumnHeaders(ColumnIndex).Width
        ListViewDestination.ColumnHeaders(ColumnIndex).Position = ListViewSource.ColumnHeaders(ColumnIndex).Position
    Next ColumnIndex
    
ErrorHandler:
End Sub

Public Sub SaveTrueDBGridSettings(ByVal Name As String, ByRef TrueDBGrid As TrueOleDBGrid80.TDBGrid, ByVal SortColumn As Integer, ByVal SortOrderAscending As Boolean)
    Dim Column As TrueOleDBGrid80.Column

    On Error GoTo ErrorHandler
    DoEvents
    Call pParametro.Usuario_GuardarNumero(Name & "_TrueDBGrid_SortColumn", SortColumn)
    Call pParametro.Usuario_GuardarSiNo(Name & "_TrueDBGrid_SortOrder", SortOrderAscending)
    For Each Column In TrueDBGrid.Columns
        Call pParametro.Usuario_GuardarNumero(Name & "_TrueDBGrid_Column_" & Column.DataField & "_Width", Column.Width)
    Next Column

ErrorHandler:
End Sub

Public Sub GetTrueDBGridSettings(ByVal Name As String, ByRef TrueDBGrid As TrueOleDBGrid80.TDBGrid, ByRef SortColumn As Integer, ByRef SortOrderAscending As Boolean)
    Dim Column As TrueOleDBGrid80.Column

    On Error GoTo ErrorHandler

    SortColumn = pParametro.Usuario_LeerNumero(Name & "_TrueDBGrid_SortColumn", SortColumn)
    SortOrderAscending = pParametro.Usuario_LeerSiNo(Name & "_TrueDBGrid_SortOrder", SortOrderAscending)

    TrueDBGrid.RowDividerStyle = IIf(pParametro.ListView_GridLines, dbgCustomColor, dbgNoDividers)
    TrueDBGrid.RowDividerColor = &HECE9D8

    For Each Column In TrueDBGrid.Columns
        Column.DividerStyle = IIf(pParametro.ListView_GridLines, dbgCustomColor, dbgNoDividers)
        Column.DividerColor = &HECE9D8
        Column.Width = pParametro.Usuario_LeerNumero(Name & "_TrueDBGrid_Column_" & Column.DataField & "_Width", Column.Width)
    Next Column

ErrorHandler:
End Sub

Public Sub SaveCoolBarSettings(ByVal Name As String, ByRef CoolBar As ComCtl3.CoolBar)
    Dim Band As ComCtl3.Band

    On Error GoTo ErrorHandler

    DoEvents
    For Each Band In CoolBar.Bands
        pParametro.Usuario_GuardarNumero Name & "_CoolBar_Band_" & Band.Key & "_Width", Band.Width
        If Band.Position > 1 Then
            pParametro.Usuario_GuardarSiNo Name & "_CoolBar_Band_" & Band.Key & "_NewRow", Band.NewRow
        End If
    Next Band

ErrorHandler:
End Sub

Public Sub GetCoolBarSettings(ByVal Name As String, ByRef CoolBar As ComCtl3.CoolBar)
    Dim Band As ComCtl3.Band

    On Error GoTo ErrorHandler

    For Each Band In CoolBar.Bands
        Band.Width = pParametro.Usuario_LeerNumero(Name & "_CoolBar_Band_" & Band.Key & "_Width", Band.Width)
        If Band.Position > 1 Then
            Band.NewRow = pParametro.Usuario_LeerSiNo(Name & "_CoolBar_Band_" & Band.Key & "_NewRow", Band.NewRow)
        End If
    Next Band

ErrorHandler:
End Sub

Public Function PlanillaViajeEmail_SendReportFormat_Extension() As String
    Dim ValueReturn As String
    
    Select Case PlanillaViajeEmail_SendReportFormat
        Case crEFTCrystalReport
            ValueReturn = "rpt"
        Case crEFTExactRichText
            ValueReturn = "rtf"
        Case crEFTExcel97
            ValueReturn = "xls"
        Case crEFTHTML40
            ValueReturn = "htm"
        Case crEFTPortableDocFormat
            ValueReturn = "pdf"
        Case crEFTWordForWindows
            ValueReturn = "doc"
        Case crEFTXML
            ValueReturn = "xml"
    End Select
    
    PlanillaViajeEmail_SendReportFormat_Extension = ValueReturn
End Function
