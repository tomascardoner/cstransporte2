VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPasajeHistorico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pasaje de Datos a Histórico"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PasajeAHistorico.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   7155
   Begin VB.Frame fraProcesos 
      Caption         =   "Procesos:"
      Height          =   2475
      Left            =   180
      TabIndex        =   3
      Top             =   3420
      Width           =   3975
      Begin VB.CheckBox chkBackupDatabase 
         Caption         =   "Copia de Seguridad de los Datos"
         Height          =   210
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Value           =   1  'Checked
         Width           =   3555
      End
      Begin VB.CheckBox chkEmail 
         Caption         =   "E-mails (Eliminar)"
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   2100
         Value           =   1  'Checked
         Width           =   3555
      End
      Begin VB.CheckBox chkVehiculoMantenimientoAccion 
         Caption         =   "Acciones de Mantenimiento de Vehículos"
         Height          =   210
         Left            =   180
         TabIndex        =   8
         Top             =   1740
         Value           =   1  'Checked
         Width           =   3555
      End
      Begin VB.CheckBox chkPersonaRespuesta 
         Caption         =   "Respuestas de Personas"
         Height          =   210
         Left            =   180
         TabIndex        =   7
         Top             =   1380
         Value           =   1  'Checked
         Width           =   3555
      End
      Begin VB.CheckBox chkFeriado 
         Caption         =   "Feriados"
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   1020
         Value           =   1  'Checked
         Width           =   3555
      End
      Begin VB.CheckBox chkCuentaCorriente_Viaje 
         Caption         =   "Cuentas Corriente y Viajes"
         Height          =   210
         Left            =   180
         TabIndex        =   5
         Top             =   660
         Value           =   1  'Checked
         Width           =   3555
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Procesar"
      Height          =   435
      Left            =   5700
      TabIndex        =   11
      Top             =   6120
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker dtpFechaCierre 
      Height          =   315
      Left            =   2340
      TabIndex        =   2
      Top             =   1200
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   116916225
      CurrentDate     =   37897
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   180
      TabIndex        =   10
      Top             =   6180
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label lblFechaCierre 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Cierre (Inclusive):"
      Height          =   210
      Left            =   180
      TabIndex        =   1
      Top             =   1260
      Width           =   1995
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Left            =   180
      Picture         =   "PasajeAHistorico.frx":000C
      Top             =   300
      Width           =   480
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"PasajeAHistorico.frx":044E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   6195
   End
End
Attribute VB_Name = "frmPasajeHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFechaCierreMaxima As Date
Private mDatabasePath As String
Private mDatabaseName As String

Private Sub Form_Load()
    mDatabasePath = CSM_File.GetPath(pDatabase.GetPrimaryFile_PathAndFileName())
    mDatabaseName = pParametro.Database_DatabaseHistory

    mFechaCierreMaxima = DateAdd("m", -2, Date)
    mFechaCierreMaxima = DateSerial(Year(mFechaCierreMaxima), Month(mFechaCierreMaxima), 1)
    mFechaCierreMaxima = DateAdd("d", -1, mFechaCierreMaxima)
    
    dtpFechaCierre.Value = mFechaCierreMaxima
End Sub

Private Sub cmdOK_Click()
    Dim ScriptText As String
    
    If DateDiff("d", mFechaCierreMaxima, dtpFechaCierre.Value) > 0 Then
        MsgBox "La fecha de Cierre debe ser menor o igual al " & Format(mFechaCierreMaxima, "Short Date") & ".", vbInformation, App.Title
        dtpFechaCierre.SetFocus
        Exit Sub
    End If
    
    If chkCuentaCorriente_Viaje.Value = vbUnchecked And chkFeriado.Value = vbUnchecked And chkPersonaRespuesta.Value = vbUnchecked And chkVehiculoMantenimientoAccion.Value = vbUnchecked Then
        MsgBox "Debe seleccionar alguno de los Procesos.", vbInformation, App.Title
        Exit Sub
    End If
    
    If MsgBox("¿Desea iniciar el Proceso?" & vbCr & vbCr & "Recuerde que no debe haber usuarios utilizando el Sistema y no podrá detener el proceso hasta que finalice.", vbExclamation + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
        Exit Sub
    End If
    
    cmdOK.Enabled = False
    
    'REALIZO EL BACKUP
    If chkBackupDatabase.Value = vbChecked Then
        
        ScriptText = "BACKUP DATABASE [" & pParametro.Database_Database & "] TO DISK = N'" & mDatabasePath & "History_Backup\" & CleanInvalidCharsByAllowed(pParametro.CompanyName, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") & Format(Date, "yyyymmdd") & "_Backup.bak' WITH INIT, NOUNLOAD, NAME = N'" & pParametro.Database_Database & " backup', SKIP, STATS = 10, DESCRIPTION = N'Backup antes del cierre del " & dtpFechaCierre.Value & "', NOFORMAT"
        
        lblStatus.Visible = True
        lblStatus.Caption = "Realizando copia de seguridad de la Base de Datos..."
        DoEvents
        
        If Not pDatabase.ExecuteScript(ScriptText, "Error al realizar la copia de seguridad de la Base de Datos") Then
            lblStatus.Visible = False
            lblStatus.Caption = ""
            cmdOK.Enabled = True
            Exit Sub
        End If
    End If
    
'    'CREO LA BASE DE DATOS
'    lblStatus.Visible = True
'    lblStatus.Caption = "Creando la Base de Datos..."
'    DoEvents
'    If Not CreateDatabase(mDatabaseName, mDatabasePath, mDatabasePath) Then
'        lblStatus.Visible = False
'        lblStatus.Caption = ""
'        cmdOK.Enabled = True
'        Exit Sub
'    End If
    
'    'CREO LOS DEFAULTS
'    lblStatus.Visible = True
'    lblStatus.Caption = "Creando los Defaults..."
'    DoEvents
'    If Not CreateDefaults() Then
'        lblStatus.Caption = "Eliminando la Base de Datos..."
'        DoEvents
'        Call RemoveDatabase
'        lblStatus.Visible = False
'        lblStatus.Caption = ""
'        cmdOK.Enabled = True
'        Exit Sub
'    End If
    
'    'CREO LAS RULES
'    lblStatus.Visible = True
'    lblStatus.Caption = "Creando las Rules..."
'    DoEvents
'    If Not CreateRules() Then
'        lblStatus.Caption = "Eliminando la Base de Datos..."
'        DoEvents
'        Call RemoveDatabase
'        lblStatus.Visible = False
'        lblStatus.Caption = ""
'        cmdOK.Enabled = True
'        Exit Sub
'    End If
    
'    'CREO LOS USER DEFINED DATATYPES
'    lblStatus.Visible = True
'    lblStatus.Caption = "Creando los User Defined Datatypes..."
'    DoEvents
'    If Not CreateUserDefinedDatatypes() Then
'        lblStatus.Caption = "Eliminando la Base de Datos..."
'        DoEvents
'        Call RemoveDatabase
'        lblStatus.Visible = False
'        lblStatus.Caption = ""
'        cmdOK.Enabled = True
'        Exit Sub
'    End If
    
'    'CREO LAS TABLAS
'    lblStatus.Visible = True
'    lblStatus.Caption = "Creando las Tablas..."
'    DoEvents
'    If Not CreateTables() Then
'        lblStatus.Caption = "Eliminando la Base de Datos..."
'        DoEvents
'        Call RemoveDatabase
'        lblStatus.Visible = False
'        lblStatus.Caption = ""
'        cmdOK.Enabled = True
'        Exit Sub
'    End If
    
'    'CREO LAS PRIMARY KEYS
'    lblStatus.Visible = True
'    lblStatus.Caption = "Creando las Primary Keys..."
'    DoEvents
'    If Not CreatePrimaryKeys() Then
'        lblStatus.Caption = "Eliminando la Base de Datos..."
'        DoEvents
'        Call RemoveDatabase
'        lblStatus.Visible = False
'        lblStatus.Caption = ""
'        cmdOK.Enabled = True
'        Exit Sub
'    End If
    
'    'BINDEO LOS DEFAULTS Y LAS RULES
'    lblStatus.Visible = True
'    lblStatus.Caption = "Bindeando los Defaults y las Rules..."
'    DoEvents
'    If Not BindDefaultsAndRules() Then
'        lblStatus.Caption = "Eliminando la Base de Datos..."
'        DoEvents
'        Call RemoveDatabase
'        lblStatus.Visible = False
'        lblStatus.Caption = ""
'        cmdOK.Enabled = True
'        Exit Sub
'    End If
    
    'CUENTACORRIENTE - VIAJE - VIAJEDETALLE
    If chkCuentaCorriente_Viaje.Value = vbChecked Then
        'INICIO LA TRANSACCION
        ScriptText = "SET XACT_ABORT ON" & vbCr & vbCr
        ScriptText = ScriptText & "BEGIN TRANSACTION" & vbCr & vbCr
        
        'COPIO FRANCO
        ScriptText = ScriptText & "INSERT INTO [" & mDatabaseName & "]..Franco" & vbCr
        ScriptText = ScriptText & "(Fecha, IDPersona, Importe, IDMovimientoCuentaCorriente, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)" & vbCr
        ScriptText = ScriptText & "SELECT Franco.Fecha, Franco.IDPersona, Franco.Importe, Franco.IDMovimientoCuentaCorriente, Franco.FechaHoraCreacion, Franco.IDUsuarioCreacion, Franco.FechaHoraModificacion, Franco.IDUsuarioModificacion" & vbCr
        ScriptText = ScriptText & "FROM Franco INNER JOIN CuentaCorriente ON Franco.IDMovimientoCuentaCorriente = CuentaCorriente.IDMovimiento" & vbCr
        ScriptText = ScriptText & "WHERE CuentaCorriente.FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        'ELIMINO FRANCO
        ScriptText = ScriptText & "DELETE" & vbCr
        ScriptText = ScriptText & "FROM Franco" & vbCr
        ScriptText = ScriptText & "FROM Franco INNER JOIN CuentaCorriente ON Franco.IDMovimientoCuentaCorriente = CuentaCorriente.IDMovimiento" & vbCr
        ScriptText = ScriptText & "WHERE CuentaCorriente.FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        'COPIO VIAJE
        ScriptText = ScriptText & "INSERT INTO [" & mDatabaseName & "]..Viaje" & vbCr
        ScriptText = ScriptText & "(FechaHora, IDRuta, RutaOtra, IDPersona, Kilometro, Duracion, Importe, ImporteContado, IDCuentaCorrienteCaja, Charter, IDVehiculo, IDConductor, AcreditaSueldo, IDConductor2, AcreditaSueldo2, DiaSemanaBase, Estado, AsientoOcupado, Notas, Personal, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion, FechaHoraEnProgreso, IDUsuarioEnProgreso, FechaHoraFinalizado, IDUsuarioFinalizado, FechaHoraCancelado, IDUsuarioCancelado)" & vbCr
        ScriptText = ScriptText & "SELECT FechaHora, IDRuta, RutaOtra, IDPersona, Kilometro, Duracion, Importe, ImporteContado, IDCuentaCorrienteCaja, Charter, IDVehiculo, IDConductor, AcreditaSueldo, IDConductor2, AcreditaSueldo2, DiaSemanaBase, Estado, AsientoOcupado, Notas, Personal, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion, FechaHoraEnProgreso, IDUsuarioEnProgreso, FechaHoraFinalizado, IDUsuarioFinalizado, FechaHoraCancelado, IDUsuarioCancelado" & vbCr
        ScriptText = ScriptText & "FROM Viaje" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        'COPIO VIAJEDETALLE
        ScriptText = ScriptText & "INSERT INTO [" & mDatabaseName & "]..ViajeDetalle" & vbCr
        ScriptText = ScriptText & "(FechaHora, IDRuta, Indice, OcupanteTipo, Estado, Prioridad, Orden, Asiento, Realizado, IDPersona, IDListaPrecio, IDOrigen, Sube, IDDestino, Baja, ValorDeclarado, ImporteSeguro, Importe, ImporteContado, ImporteCuentaCorriente, ImprimirSaldo, IDCuentaCorrienteCaja, ForzarDebito, IDPersonaCuentaCorriente, Facturar, FacturarNotas, FacturaNumero, IDPersonaRecibe, PagaQuienRecibe, Recibe, Descripcion, Horario, Telefono, DejarTraer, Entregada, EntregadaFechaHora, Retira, ReservaTipo, CreadoEnProgreso, ModificadoEnProgreso, ReservadoPor, CanceladoPor, CanceladoFechaHora, Notas, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion, IDUsuarioCancelacion)" & vbCr
        ScriptText = ScriptText & "SELECT FechaHora, IDRuta, Indice, OcupanteTipo, Estado, Prioridad, Orden, Asiento, Realizado, IDPersona, IDListaPrecio, IDOrigen, Sube, IDDestino, Baja, ValorDeclarado, ImporteSeguro, Importe, ImporteContado, ImporteCuentaCorriente, ImprimirSaldo, IDCuentaCorrienteCaja, ForzarDebito, IDPersonaCuentaCorriente, Facturar, FacturarNotas, FacturaNumero, IDPersonaRecibe, PagaQuienRecibe, Recibe, Descripcion, Horario, Telefono, DejarTraer, Entregada, EntregadaFechaHora, Retira, ReservaTipo, CreadoEnProgreso, ModificadoEnProgreso, ReservadoPor, CanceladoPor, CanceladoFechaHora, Notas, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion, IDUsuarioCancelacion" & vbCr
        ScriptText = ScriptText & "FROM ViajeDetalle" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        'COPIO VIAJEDETALLE_COMISION
        ScriptText = ScriptText & "INSERT INTO [" & mDatabaseName & "]..ViajeDetalle_Comision" & vbCr
        ScriptText = ScriptText & "(FechaHora, IDRuta, Indice, RendicionFechaHora, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)" & vbCr
        ScriptText = ScriptText & "SELECT FechaHora, IDRuta, Indice, RendicionFechaHora, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion" & vbCr
        ScriptText = ScriptText & "FROM ViajeDetalle_Comision" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        'COPIO VIAJEDETALLE_CONEXION
        ScriptText = ScriptText & "INSERT INTO [" & mDatabaseName & "]..ViajeDetalle_Conexion" & vbCr
        ScriptText = ScriptText & "(FechaHora, IDRuta, Indice, Conexion_FechaHora, Conexion_IDRuta, Conexion_Indice)" & vbCr
        ScriptText = ScriptText & "SELECT FechaHora, IDRuta, Indice, Conexion_FechaHora, Conexion_IDRuta, Conexion_Indice" & vbCr
        ScriptText = ScriptText & "FROM ViajeDetalle_Conexion" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        'PARA LOS MOVIMIENTOS DE CUENTACORRIENTE ANTERIORES A LA FECHA DE CIERRE
        'QUE PERTENEZCAN A VIAJES Y A DETALLES DE VIAJES POSTERIORES A LA FECHA DE CIERRE:
        ScriptText = ScriptText & "DECLARE @IDGrupoAjustes int" & vbCr
        ScriptText = ScriptText & "DECLARE @MaxID int" & vbCr
        ScriptText = ScriptText & "DECLARE @CuentaCorrienteDebito_Temp table (IDKey int NOT NULL IDENTITY PRIMARY KEY, IDMovimiento int NOT NULL, IDCuentaCorrienteCaja int NOT NULL, IDPersona int NULL, FechaHora smalldatetime NOT NULL, Importe money NOT NULL)" & vbCr & vbCr
        'OBTENGO LOS DATOS DE ESOS MOVIMIENTOS DE LA TABLA
        ScriptText = ScriptText & "INSERT INTO @CuentaCorrienteDebito_Temp" & vbCr
        ScriptText = ScriptText & "(IDMovimiento, IDCuentaCorrienteCaja, IDPersona, FechaHora, Importe)" & vbCr
        ScriptText = ScriptText & "SELECT IDMovimiento, IDCuentaCorrienteCaja, IDPersona, FechaHora, Importe" & vbCr
        ScriptText = ScriptText & "FROM CuentaCorriente" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00') AND Viaje_FechaHora >= convert(smalldatetime, '" & Format(DateAdd("d", 1, dtpFechaCierre.Value), "yyyy/mm/dd") & " 00:00:00')" & vbCr & vbCr
        ScriptText = ScriptText & "SET @IDGrupoAjustes = (SELECT IDCuentaCorrienteGrupo FROM CuentaCorrienteGrupo WHERE Nombre = 'Ajustes')" & vbCr
        ScriptText = ScriptText & "SET @MaxID = (SELECT ISNULL(MAX(IDMovimiento), 0) + 1 FROM CuentaCorriente)" & vbCr & vbCr
        'GENERO AJUSTES CON EL MISMO IMPORTE Y LA MISMA FECHA
        ScriptText = ScriptText & "INSERT INTO CuentaCorriente" & vbCr
        ScriptText = ScriptText & "(IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)" & vbCr
        ScriptText = ScriptText & "SELECT @MaxID + IDKey, @IDGrupoAjustes, IDCuentaCorrienteCaja, IDPersona, FechaHora, 'Ajuste por Cierre', Importe, NULL, NULL, 0, NULL, NULL, NULL, getdate(), 'administrator', getdate(), 'administrator'" & vbCr
        ScriptText = ScriptText & "FROM @CuentaCorrienteDebito_Temp" & vbCr & vbCr
        ScriptText = ScriptText & "SET @MaxID = (SELECT ISNULL(MAX(IDMovimiento), 0) + 1 FROM CuentaCorriente)" & vbCr & vbCr
        'GENERO AJUSTES CON EL IMPORTE CONTRARIO EN EL DIA SIGUIENTE A LA FECHA DE CIERRE
        ScriptText = ScriptText & "INSERT INTO CuentaCorriente" & vbCr
        ScriptText = ScriptText & "(IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)" & vbCr
        ScriptText = ScriptText & "SELECT @MaxID + IDKey, @IDGrupoAjustes, IDCuentaCorrienteCaja, IDPersona, convert(smalldatetime, '" & Format(DateAdd("d", 1, dtpFechaCierre.Value), "yyyy/mm/dd") & " 00:00:00'), 'Ajuste por Cierre', Importe * -1, NULL, NULL, 0, NULL, NULL, NULL, getdate(), 'administrator', getdate(), 'administrator'" & vbCr
        ScriptText = ScriptText & "FROM @CuentaCorrienteDebito_Temp" & vbCr & vbCr & vbCr
        'CAMBIO LA FECHA DEL MOVIMIENTO ORIGINAL AL DIA SIGUIENTE A LA FECHA DE CIERRE
        ScriptText = ScriptText & "UPDATE CuentaCorriente" & vbCr
        ScriptText = ScriptText & "SET FechaHoraModificacion = getdate(), FechaHora = convert(smalldatetime, '" & Format(DateAdd("d", 1, dtpFechaCierre.Value), "yyyy/mm/dd") & " 00:00:00')" & vbCr
        ScriptText = ScriptText & "WHERE IDMovimiento IN (SELECT IDMovimiento FROM @CuentaCorrienteDebito_Temp)" & vbCr & vbCr
        
        'GENERO LOS SALDOS INICIALES DE LAS CUENTAS CORRIENTE (POR CADA GRUPO, CAJA Y PERSONA)
        ScriptText = ScriptText & "DECLARE @CuentaCorriente_Saldo table (IDKey int NOT NULL IDENTITY PRIMARY KEY, IDCuentaCorrienteGrupo int NOT NULL, IDCuentaCorrienteCaja int NOT NULL, IDPersona int NULL, Saldo money NOT NULL)" & vbCr & vbCr
        
        ScriptText = ScriptText & "SET @MaxID = (SELECT ISNULL(MAX(IDMovimiento), 0) + 1 FROM CuentaCorriente)" & vbCr & vbCr
        
        ScriptText = ScriptText & "INSERT INTO @CuentaCorriente_Saldo" & vbCr
        ScriptText = ScriptText & "(IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, Saldo)" & vbCr
        ScriptText = ScriptText & "SELECT IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, SUM(Importe)" & vbCr
        ScriptText = ScriptText & "FROM CuentaCorriente" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr
        ScriptText = ScriptText & "GROUP BY IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona" & vbCr
        ScriptText = ScriptText & "HAVING SUM(Importe) <> 0" & vbCr & vbCr
        
        ScriptText = ScriptText & "INSERT INTO CuentaCorriente" & vbCr
        ScriptText = ScriptText & "(IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)" & vbCr
        ScriptText = ScriptText & "SELECT @MaxID + IDKey, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00'), 'Saldo Anterior por Cierre', Saldo, NULL, NULL, 1, NULL, NULL, NULL, getdate(), '" & pUsuario.IDUsuario & "', getdate(), '" & pUsuario.IDUsuario & "'" & vbCr
        ScriptText = ScriptText & "FROM @CuentaCorriente_Saldo" & vbCr & vbCr
        
        'COPIO CUENTACORRIENTE
        ScriptText = ScriptText & "INSERT INTO [" & mDatabaseName & "]..CuentaCorriente" & vbCr
        ScriptText = ScriptText & "(IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, Viaje_ConductorNumero, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)" & vbCr
        ScriptText = ScriptText & "SELECT IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, Viaje_ConductorNumero, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion" & vbCr
        ScriptText = ScriptText & "FROM CuentaCorriente" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00') AND SaldoAnterior = 0" & vbCr & vbCr
        
        'ELIMINO VIAJEDETALLE_COMISION
        ScriptText = ScriptText & "DELETE" & vbCr
        ScriptText = ScriptText & "FROM ViajeDetalle_Comision" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        'ELIMINO VIAJEDETALLE_CONEXION
        ScriptText = ScriptText & "DELETE" & vbCr
        ScriptText = ScriptText & "FROM ViajeDetalle_Conexion" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
                
        'ELIMINO VIAJEDETALLE
        ScriptText = ScriptText & "DELETE" & vbCr
        ScriptText = ScriptText & "FROM ViajeDetalle" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        'ELIMINO VIAJE
        ScriptText = ScriptText & "DELETE" & vbCr
        ScriptText = ScriptText & "FROM Viaje" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        'ELIMINO CUENTACORRIENTE
        ScriptText = ScriptText & "DELETE" & vbCr
        ScriptText = ScriptText & "FROM CuentaCorriente" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora < convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00') OR (FechaHora = convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00') AND SaldoAnterior = 0)" & vbCr & vbCr
        
        'CIERRO LA TRANSACCION
        ScriptText = ScriptText & "COMMIT TRANSACTION" & vbCr & vbCr
        ScriptText = ScriptText & "SET XACT_ABORT OFF"
        
        lblStatus.Visible = True
        lblStatus.Caption = "Procesando Cuentas Corriente y Viajes..."
        DoEvents
        
        If Not pDatabase.ExecuteScript(ScriptText, "Error al procesar la Cuenta Corriente y los Viajes.") Then
            If MsgBox("¿Desea continuar ejecutando los procesos restantes?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                lblStatus.Visible = False
                lblStatus.Caption = ""
                cmdOK.Enabled = True
                Exit Sub
            End If
        End If
    End If
    
    'FERIADO
    If chkFeriado.Value = vbChecked Then
        ScriptText = "SET XACT_ABORT ON" & vbCr & vbCr
        ScriptText = ScriptText & "BEGIN TRANSACTION" & vbCr & vbCr
        
        ScriptText = ScriptText & "INSERT INTO [" & mDatabaseName & "]..Feriado" & vbCr
        ScriptText = ScriptText & "(Fecha, Nombre, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)" & vbCr
        ScriptText = ScriptText & "SELECT Fecha, Nombre, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion" & vbCr
        ScriptText = ScriptText & "FROM Feriado" & vbCr
        ScriptText = ScriptText & "WHERE Fecha <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        ScriptText = ScriptText & "DELETE" & vbCr
        ScriptText = ScriptText & "FROM Feriado" & vbCr
        ScriptText = ScriptText & "WHERE Fecha <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        ScriptText = ScriptText & "COMMIT TRANSACTION" & vbCr & vbCr
        ScriptText = ScriptText & "SET XACT_ABORT OFF"
        
        lblStatus.Visible = True
        lblStatus.Caption = "Procesando Feriados..."
        DoEvents
        
        If Not pDatabase.ExecuteScript(ScriptText, "Error al Procesar los Feriados") Then
            If MsgBox("¿Desea continuar ejecutando los procesos restantes?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                lblStatus.Visible = False
                lblStatus.Caption = ""
                cmdOK.Enabled = True
                Exit Sub
            End If
        End If
    End If
    
    'PERSONARESPUESTA
    If chkPersonaRespuesta.Value = vbChecked Then
        ScriptText = "SET XACT_ABORT ON" & vbCr & vbCr
        ScriptText = ScriptText & "BEGIN TRANSACTION" & vbCr & vbCr
        
        ScriptText = ScriptText & "INSERT INTO [" & mDatabaseName & "]..PersonaRespuesta" & vbCr
        ScriptText = ScriptText & "(IDPersona, FechaHora, Respuesta, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)" & vbCr
        ScriptText = ScriptText & "SELECT IDPersona, FechaHora, Respuesta, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion" & vbCr
        ScriptText = ScriptText & "FROM PersonaRespuesta" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        ScriptText = ScriptText & "DELETE" & vbCr
        ScriptText = ScriptText & "FROM PersonaRespuesta" & vbCr
        ScriptText = ScriptText & "WHERE FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        ScriptText = ScriptText & "COMMIT TRANSACTION" & vbCr & vbCr
        ScriptText = ScriptText & "SET XACT_ABORT OFF"
        
        lblStatus.Visible = True
        lblStatus.Caption = "Procesando Respuestas de Personas..."
        DoEvents
        
        If Not pDatabase.ExecuteScript(ScriptText, "Error al Procesar las Respuestas de las Personas") Then
            If MsgBox("¿Desea continuar ejecutando los procesos restantes?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                lblStatus.Visible = False
                lblStatus.Caption = ""
                cmdOK.Enabled = True
                Exit Sub
            End If
        End If
    End If
    
    'VEHICULOMANTENIMIENTOACCION
    If chkVehiculoMantenimientoAccion.Value = vbChecked Then
        ScriptText = "SET XACT_ABORT ON" & vbCr & vbCr
        ScriptText = ScriptText & "BEGIN TRANSACTION" & vbCr & vbCr
        
        ScriptText = ScriptText & "DECLARE @AccionesACopiar table (IDVehiculoMantenimientoAccion int NOT NULL PRIMARY KEY)" & vbCr
        ScriptText = ScriptText & "DECLARE @IDVehiculo int" & vbCr
        ScriptText = ScriptText & "DECLARE @IDVehiculoMantenimientoGrupo int" & vbCr
        ScriptText = ScriptText & "DECLARE @IDVehiculoMantenimientoAccionLast int" & vbCr
        ScriptText = ScriptText & "DECLARE VehiculoMantenimientoCursor" & vbCr
        ScriptText = ScriptText & "CURSOR LOCAL FORWARD_ONLY KEYSET" & vbCr
        ScriptText = ScriptText & "FOR SELECT VehiculoMantenimientoAccion.IDVehiculo, VehiculoMantenimientoAccion.IDVehiculoMantenimientoGrupo" & vbCr
        ScriptText = ScriptText & "FROM VehiculoMantenimiento INNER JOIN VehiculoMantenimientoAccion ON VehiculoMantenimiento.IDVehiculoMantenimientoGrupo = VehiculoMantenimientoAccion.IDVehiculoMantenimientoGrupo AND VehiculoMantenimiento.IDVehiculo = VehiculoMantenimientoAccion.IDVehiculo" & vbCr
        ScriptText = ScriptText & "WHERE (Tipo = 'KI' OR Tipo = 'DI') AND FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr
        ScriptText = ScriptText & "GROUP BY VehiculoMantenimientoAccion.IDVehiculo, VehiculoMantenimientoAccion.IDVehiculoMantenimientoGrupo" & vbCr & vbCr
        
        ScriptText = ScriptText & "INSERT INTO @AccionesACopiar" & vbCr
        ScriptText = ScriptText & "(IDVehiculoMantenimientoAccion)" & vbCr
        ScriptText = ScriptText & "SELECT VehiculoMantenimientoAccion.IDVehiculoMantenimientoAccion" & vbCr
        ScriptText = ScriptText & "FROM VehiculoMantenimiento INNER JOIN VehiculoMantenimientoAccion ON VehiculoMantenimiento.IDVehiculoMantenimientoGrupo = VehiculoMantenimientoAccion.IDVehiculoMantenimientoGrupo AND VehiculoMantenimiento.IDVehiculo = VehiculoMantenimientoAccion.IDVehiculo" & vbCr
        ScriptText = ScriptText & "WHERE (VehiculoMantenimiento.Tipo = 'FE' OR VehiculoMantenimiento.Tipo = 'NI')" & vbCr
        ScriptText = ScriptText & "AND VehiculoMantenimientoAccion.FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr & vbCr
        
        ScriptText = ScriptText & "OPEN VehiculoMantenimientoCursor" & vbCr
        ScriptText = ScriptText & "FETCH NEXT FROM VehiculoMantenimientoCursor INTO @IDVehiculo, @IDVehiculoMantenimientoGrupo" & vbCr
        ScriptText = ScriptText & "WHILE @@FETCH_STATUS = 0" & vbCr
        ScriptText = ScriptText & "BEGIN" & vbCr
        ScriptText = ScriptText & "IF (SELECT COUNT(IDVehiculoMantenimientoAccion) FROM VehiculoMantenimientoAccion WHERE IDVehiculo = @IDVehiculo AND IDVehiculoMantenimientoGrupo = @IDVehiculoMantenimientoGrupo AND FechaHora > convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')) > 0" & vbCr
        ScriptText = ScriptText & "BEGIN" & vbCr
        ScriptText = ScriptText & "INSERT INTO @AccionesACopiar" & vbCr
        ScriptText = ScriptText & "(IDVehiculoMantenimientoAccion)" & vbCr
        ScriptText = ScriptText & "SELECT IDVehiculoMantenimientoAccion" & vbCr
        ScriptText = ScriptText & "FROM VehiculoMantenimientoAccion" & vbCr
        ScriptText = ScriptText & "WHERE IDVehiculo = @IDVehiculo" & vbCr
        ScriptText = ScriptText & "AND IDVehiculoMantenimientoGrupo = @IDVehiculoMantenimientoGrupo" & vbCr
        ScriptText = ScriptText & "AND FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr
        ScriptText = ScriptText & "END" & vbCr
        ScriptText = ScriptText & "ELSE" & vbCr
        ScriptText = ScriptText & "BEGIN" & vbCr
        ScriptText = ScriptText & "SET @IDVehiculoMantenimientoAccionLast = (SELECT TOP 1 IDVehiculoMantenimientoAccion FROM VehiculoMantenimientoAccion WHERE IDVehiculo = @IDVehiculo AND IDVehiculoMantenimientoGrupo = @IDVehiculoMantenimientoGrupo AND VehiculoMantenimientoAccion.FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00') ORDER BY FechaHora DESC)" & vbCr
        ScriptText = ScriptText & "INSERT INTO @AccionesACopiar" & vbCr
        ScriptText = ScriptText & "(IDVehiculoMantenimientoAccion)" & vbCr
        ScriptText = ScriptText & "SELECT VehiculoMantenimientoAccion.IDVehiculoMantenimientoAccion" & vbCr
        ScriptText = ScriptText & "FROM VehiculoMantenimientoAccion" & vbCr
        ScriptText = ScriptText & "WHERE IDVehiculo = @IDVehiculo" & vbCr
        ScriptText = ScriptText & "AND IDVehiculoMantenimientoGrupo = @IDVehiculoMantenimientoGrupo" & vbCr
        ScriptText = ScriptText & "AND FechaHora <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')" & vbCr
        ScriptText = ScriptText & "AND IDVehiculoMantenimientoAccion <> @IDVehiculoMantenimientoAccionLast" & vbCr
        ScriptText = ScriptText & "END" & vbCr
        ScriptText = ScriptText & "FETCH NEXT FROM VehiculoMantenimientoCursor INTO @IDVehiculo, @IDVehiculoMantenimientoGrupo" & vbCr
        ScriptText = ScriptText & "END" & vbCr
        ScriptText = ScriptText & "CLOSE VehiculoMantenimientoCursor" & vbCr
        ScriptText = ScriptText & "DEALLOCATE VehiculoMantenimientoCursor" & vbCr & vbCr
        
        ScriptText = ScriptText & "INSERT INTO [" & mDatabaseName & "]..VehiculoMantenimientoAccion" & vbCr
        ScriptText = ScriptText & "(IDVehiculoMantenimientoAccion, IDVehiculo, IDVehiculoMantenimientoGrupo, IDConductor, FechaHora, Kilometraje, Litros, Importe, Notas, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)" & vbCr
        ScriptText = ScriptText & "SELECT IDVehiculoMantenimientoAccion, IDVehiculo, IDVehiculoMantenimientoGrupo, IDConductor, FechaHora, Kilometraje, Litros, Importe, Notas, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion" & vbCr
        ScriptText = ScriptText & "FROM VehiculoMantenimientoAccion" & vbCr
        ScriptText = ScriptText & "WHERE IDVehiculoMantenimientoAccion IN (SELECT IDVehiculoMantenimientoAccion FROM @AccionesACopiar)" & vbCr & vbCr
        
        ScriptText = ScriptText & "DELETE" & vbCr
        ScriptText = ScriptText & "FROM VehiculoMantenimientoAccion" & vbCr
        ScriptText = ScriptText & "WHERE IDVehiculoMantenimientoAccion IN (SELECT IDVehiculoMantenimientoAccion FROM @AccionesACopiar)" & vbCr & vbCr
        
        ScriptText = ScriptText & "COMMIT TRANSACTION" & vbCr & vbCr
        ScriptText = ScriptText & "SET XACT_ABORT OFF"
    
        lblStatus.Visible = True
        lblStatus.Caption = "Procesando Acciones de Mantenimiento de Vehículos..."
        DoEvents
        
        If Not pDatabase.ExecuteScript(ScriptText, "Error al Procesar las Acciones de Mantenimiento de Vehículos") Then
            If MsgBox("¿Desea continuar ejecutando los procesos restantes?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                lblStatus.Visible = False
                lblStatus.Caption = ""
                cmdOK.Enabled = True
                Exit Sub
            End If
        End If
    End If
    
    'EMAIL
    If chkEmail.Value = vbChecked Then
        ScriptText = "DELETE" & vbCr
        ScriptText = ScriptText & "FROM EmailMessage" & vbCr
        ScriptText = ScriptText & "WHERE DateTime <= convert(smalldatetime, '" & Format(dtpFechaCierre.Value, "yyyy/mm/dd") & " 23:59:00')"
        
        lblStatus.Visible = True
        lblStatus.Caption = "Eliminando E-mails..."
        DoEvents
        
        Call pDatabase.ExecuteScript(ScriptText, "Error al Eliminar los E-mails")
    End If
        
    'COMPRIMO LA BASE DE DATOS DE ORIGEN
    lblStatus.Visible = True
    lblStatus.Caption = "Comprimiendo la Base de Datos de Origen..."
    DoEvents
    Call ShrinkDatabase(pParametro.Database_Database, 10)
    
    'COMPRIMO LA BASE DE DATOS DE DESTINO
    lblStatus.Visible = True
    lblStatus.Caption = "Comprimiendo la Base de Datos de Destino..."
    DoEvents
    Call ShrinkDatabase(mDatabaseName, 0)
    
'    'DETTACHEO LA BASE DE DATOS
'    lblStatus.Visible = True
'    lblStatus.Caption = "Detacheando la Base de Datos..."
'    DoEvents
'    Call DettachDatabase
    
    lblStatus.Visible = False
    lblStatus.Caption = ""
    cmdOK.Enabled = True
    
    MsgBox "El Proceso ha finalizado exitosamente.", vbInformation, App.Title
    
    Unload frmPasajeHistorico
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPasajeHistorico = Nothing
End Sub

Private Function CreateDatabase(ByVal DatabaseName As String, ByVal DataFilePath As String, ByVal LogFilePath As String) As Boolean
    Dim ScriptText As String
    
    'CREO LA BASE DE DATOS
    ScriptText = "CREATE DATABASE [" & DatabaseName & "] ON (NAME = N'" & DatabaseName & "_Data', FILENAME = N'" & DataFilePath & DatabaseName & "_Data.mdf', SIZE = 1, FILEGROWTH = 10%) LOG ON (NAME = N'" & DatabaseName & "_Log', FILENAME = N'" & LogFilePath & DatabaseName & "_Log.ldf', SIZE = 1, FILEGROWTH = 10%) COLLATE Modern_Spanish_CI_AS"
    
    CreateDatabase = pDatabase.ExecuteScript(ScriptText, "Error al Crear la Base de Datos.")
End Function

Private Function CreateDefaults() As Boolean
    Dim ScriptText As String
    
    ScriptText = "CREATE DEFAULT [def_DateTimeCurrent] AS getdate()"
    
    CreateDefaults = pDatabase.ExecuteScript(ScriptText, "Error al Crear los Defaults.", , , mDatabaseName)
End Function

Private Function CreateRules() As Boolean
    Dim ScriptText As String
    
    ScriptText = "CREATE RULE [rul_DiaSemana] AS @DiaSemana BETWEEN 1 AND 7"
    If Not pDatabase.ExecuteScript(ScriptText, "Error al Crear las Rules.", , , mDatabaseName) Then
        Exit Function
    End If
    ScriptText = "CREATE RULE [rul_EntidadTipo] AS @EntidadTipo IN ('PC', 'PO', 'PA', 'VE', 'VI', 'VD', 'RU', 'CO')"
    If Not pDatabase.ExecuteScript(ScriptText, "Error al Crear las Rules.", , , mDatabaseName) Then
        Exit Function
    End If
    CreateRules = True
End Function

Private Function CreateUserDefinedDatatypes() As Boolean
    Dim ScriptText As String
    
    ScriptText = "SET XACT_ABORT ON" & vbCr & vbCr
    
    ScriptText = ScriptText & "BEGIN TRANSACTION" & vbCr & vbCr
    
    ScriptText = ScriptText & "EXEC sp_addtype N'udt_Activo', N'bit', N'not null'" & vbCr
    ScriptText = ScriptText & "EXEC sp_addtype N'udt_DiaSemana', N'tinyint', N'not null'" & vbCr
    ScriptText = ScriptText & "EXEC sp_bindrule N'[dbo].[rul_DiaSemana]', N'[udt_DiaSemana]'" & vbCr
    ScriptText = ScriptText & "EXEC sp_addtype N'udt_EntidadTipo', N'char (2)', N'not null'" & vbCr
    ScriptText = ScriptText & "EXEC sp_bindrule N'[dbo].[rul_EntidadTipo]', N'[udt_EntidadTipo]'" & vbCr
    ScriptText = ScriptText & "EXEC sp_addtype N'udt_FechaHoraCreacion', N'smalldatetime', N'not null'" & vbCr
    ScriptText = ScriptText & "EXEC sp_bindefault N'[dbo].[def_DateTimeCurrent]', N'[udt_FechaHoraCreacion]'" & vbCr
    ScriptText = ScriptText & "EXEC sp_addtype N'udt_FechaHoraModificacion', N'smalldatetime', N'not null'" & vbCr
    ScriptText = ScriptText & "EXEC sp_addtype N'udt_IDUsuario', N'char (30)', N'not null'" & vbCr
    ScriptText = ScriptText & "EXEC sp_addtype N'udt_Notas', N'varchar (8000)', N'null'" & vbCr & vbCr
    
    ScriptText = ScriptText & "COMMIT TRANSACTION" & vbCr & vbCr
    
    ScriptText = ScriptText & "SET XACT_ABORT OFF"
    
    CreateUserDefinedDatatypes = pDatabase.ExecuteScript(ScriptText, "Error al Crear los User Defined Datatypes.", , , mDatabaseName)
End Function

Private Function CreateTables() As Boolean
    Dim ScriptText As String
    
    ScriptText = "SET XACT_ABORT ON" & vbCr & vbCr
    
    ScriptText = ScriptText & "BEGIN TRANSACTION" & vbCr & vbCr
    
    If chkCuentaCorriente_Viaje.Value = vbChecked Then
        ScriptText = ScriptText & "CREATE TABLE [dbo].[CuentaCorriente] ([IDMovimiento] [int] NOT NULL, [IDCuentaCorrienteGrupo] [int] NOT NULL, [IDCuentaCorrienteCaja] [int] NOT NULL, [IDPersona] [int] NULL, [FechaHora] [smalldatetime] NOT NULL, [Descripcion] [varchar] (255) COLLATE Modern_Spanish_CI_AS NOT NULL, [Importe] [money] NOT NULL, [IDPersonaOrigen] [int] NULL, [Notas] [udt_Notas] NULL, [SaldoAnterior] [bit] NOT NULL, [Viaje_FechaHora] [smalldatetime] NULL, [Viaje_IDRuta] [char](20) COLLATE Modern_Spanish_CI_AS NULL, [Viaje_Indice] [int] NULL, [FechaHoraCreacion] [udt_FechaHoraCreacion] NOT NULL, [IDUsuarioCreacion] [udt_IDUsuario] NULL, [FechaHoraModificacion] [udt_FechaHoraModificacion] NULL, [IDUsuarioModificacion] [udt_IDUsuario] NULL) ON [PRIMARY]" & vbCr & vbCr
        ScriptText = ScriptText & "CREATE TABLE [dbo].[Viaje] ([FechaHora] [smalldatetime] NOT NULL, [IDRuta] [char] (20) COLLATE Modern_Spanish_CI_AS NOT NULL, [RutaOtra] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL, [IDPersona] [int] NULL, [Kilometro] [smallint] NULL, [Duracion] [smallint] NULL, [Importe] [smallmoney] NULL, [ImporteContado] [smallmoney] NULL, [IDCuentaCorrienteCaja] [int] NULL, [AcreditaSueldo] [bit] NOT NULL, [Charter] [bit] NOT NULL, [IDVehiculo] [int] NULL, [IDConductor] [int] NULL, [DiaSemanaBase] [udt_DiaSemana] NOT NULL, [Estado] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL, [AsientoOcupado] [smallint] NOT NULL, [Notas] [udt_Notas] NULL"
        ScriptText = ScriptText & ", [FechaHoraCreacion] [udt_FechaHoraCreacion] NOT NULL, [IDUsuarioCreacion] [udt_IDUsuario] NOT NULL, [FechaHoraModificacion] [udt_FechaHoraModificacion] NOT NULL, [IDUsuarioModificacion] [udt_IDUsuario] NOT NULL, [FechaHoraEnProgreso] [smalldatetime] NULL, [IDUsuarioEnProgreso] [udt_IDUsuario] NULL, [FechaHoraFinalizado] [smalldatetime] NULL, [IDUsuarioFinalizado] [udt_IDUsuario] NULL, [FechaHoraCancelado] [smalldatetime] NULL, [IDUsuarioCancelado] [udt_IDUsuario] NULL) ON [PRIMARY]" & vbCr & vbCr
        ScriptText = ScriptText & "CREATE TABLE [dbo].[ViajeDetalle] ([FechaHora] [smalldatetime] NOT NULL, [IDRuta] [char] (20) COLLATE Modern_Spanish_CI_AS NOT NULL, [Indice] [int] NOT NULL, [OcupanteTipo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL, [Estado] [char] (3) COLLATE Modern_Spanish_CI_AS NULL, [Prioridad] [int] NULL, [Orden] [int] NULL, [Asiento] [tinyint] NULL, [Realizado] [bit] NULL, [IDPersona] [int] NOT NULL, [IDListaPrecio] [int] NULL, [IDOrigen] [int] NOT NULL, [Sube] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL, [IDDestino] [int] NOT NULL, [Baja] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL, [Importe] [smallmoney] NOT NULL, [ImporteContado] [smallmoney] NOT NULL, [ImporteCuentaCorriente] [smallmoney] NOT NULL, [ImprimirSaldo] [bit] NOT NULL, [IDCuentaCorrienteCaja] [int] NULL, [ForzarDebito] [bit] NOT NULL, [IDPersonaCuentaCorriente] [int] NULL"
        ScriptText = ScriptText & ", [Facturar] [bit] NOT NULL, [FacturarNotas] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL, [FacturaNumero] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL, [Recibe] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL, [Descripcion] [varchar] (100) COLLATE Modern_Spanish_CI_AS NULL, [Telefono] [varchar] (30) COLLATE Modern_Spanish_CI_AS NULL, [DejarTraer] [char] (1) COLLATE Modern_Spanish_CI_AS NULL, [Entregada] [bit] NOT NULL, [EntregadaFechaHora] [smalldatetime] NULL, [ReservaTipo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL, [CreadoEnProgreso] [bit] NOT NULL, [ModificadoEnProgreso] [bit] NOT NULL, [ReservadoPor] [varchar] (100) COLLATE Modern_Spanish_CI_AS NULL, [CanceladoPor] [varchar] (100) COLLATE Modern_Spanish_CI_AS NULL, [CanceladoFechaHora] [smalldatetime] NULL, [Notas] [udt_Notas] NULL"
        ScriptText = ScriptText & ", [FechaHoraCreacion] [udt_FechaHoraCreacion] NOT NULL, [IDUsuarioCreacion] [udt_IDUsuario] NOT NULL, [FechaHoraModificacion] [smalldatetime] NOT NULL, [IDUsuarioModificacion] [udt_IDUsuario] NOT NULL, [IDUsuarioCancelacion] [udt_IDUsuario] NULL) ON [PRIMARY]" & vbCr & vbCr
    End If
    If chkFeriado.Value = vbChecked Then
        ScriptText = ScriptText & "CREATE TABLE [dbo].[Feriado] ([Fecha] [smalldatetime] NOT NULL, [Nombre] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL, [FechaHoraCreacion] [udt_FechaHoraCreacion] NOT NULL, [IDUsuarioCreacion] [udt_IDUsuario] NOT NULL, [FechaHoraModificacion] [udt_FechaHoraModificacion] NOT NULL, [IDUsuarioModificacion] [udt_IDUsuario] NOT NULL) ON [PRIMARY]" & vbCr & vbCr
    End If
    If chkPersonaRespuesta.Value = vbChecked Then
        ScriptText = ScriptText & "CREATE TABLE [dbo].[PersonaRespuesta] ([IDPersona] [int] NOT NULL, [FechaHora] [smalldatetime] NOT NULL, [Respuesta] [varchar] (500) COLLATE Modern_Spanish_CI_AS NOT NULL, [Activo] [bit] NOT NULL, [FechaHoraCreacion] [udt_FechaHoraCreacion] NOT NULL, [IDUsuarioCreacion] [udt_IDUsuario] NOT NULL, [FechaHoraModificacion] [udt_FechaHoraModificacion] NOT NULL, [IDUsuarioModificacion] [udt_IDUsuario] NOT NULL) ON [PRIMARY]" & vbCr & vbCr
    End If
    If chkVehiculoMantenimientoAccion.Value = vbChecked Then
        ScriptText = ScriptText & "CREATE TABLE [dbo].[VehiculoMantenimientoAccion] ([IDVehiculoMantenimientoAccion] [int] NOT NULL, [IDVehiculo] [int] NOT NULL, [IDVehiculoMantenimientoGrupo] [int] NOT NULL, [IDConductor] [int] NULL, [FechaHora] [smalldatetime] NOT NULL, [Kilometraje] [int] NULL, [Litros] [float] NULL, [Importe] [smallmoney] NULL, [Notas] [udt_Notas] NULL, [FechaHoraCreacion] [udt_FechaHoraCreacion] NOT NULL, [IDUsuarioCreacion] [udt_IDUsuario] NOT NULL, [FechaHoraModificacion] [udt_FechaHoraModificacion] NOT NULL, [IDUsuarioModificacion] [udt_IDUsuario] NOT NULL) ON [PRIMARY]" & vbCr & vbCr
    End If
'    If chkAlarma.Value = vbChecked Then
'        ScriptText = ScriptText & "CREATE TABLE [dbo].[Alarma] ([IDAlarma] [int] NOT NULL, [Nombre] [varchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL, [Fecha] [smalldatetime] NOT NULL, [Preaviso] [smallint] NOT NULL, [Notas] [udt_Notas] NULL, [Activo] [udt_Activo] NOT NULL, [FechaHoraCreacion] [udt_FechaHoraCreacion] NOT NULL, [IDUsuarioCreacion] [udt_IDUsuario] NOT NULL, [FechaHoraModificacion] [udt_FechaHoraModificacion] NOT NULL, [IDUsuarioModificacion] [udt_IDUsuario] NOT NULL) ON [PRIMARY]" & vbCr & vbCr
'    End If
'    If chkPersonaAlarma.Value = vbChecked Then
'        ScriptText = ScriptText & "CREATE TABLE [dbo].[PersonaAlarma] ([IDPersona] [int] NOT NULL, [IDPersonaAlarmaGrupo] [int] NOT NULL, [Fecha] [smalldatetime] NOT NULL, [Preaviso] [smallint] NOT NULL, [Notas] [udt_Notas] NULL, [Activo] [udt_Activo] NOT NULL, [FechaHoraCreacion] [udt_FechaHoraCreacion] NOT NULL, [IDUsuarioCreacion] [udt_IDUsuario] NOT NULL, [FechaHoraModificacion] [udt_FechaHoraModificacion] NOT NULL, [IDUsuarioModificacion] [udt_IDUsuario] NOT NULL) ON [PRIMARY]" & vbCr & vbCr
'    End If
    
    ScriptText = ScriptText & "COMMIT TRANSACTION" & vbCr & vbCr
    
    ScriptText = ScriptText & "SET XACT_ABORT OFF"
    
    CreateTables = pDatabase.ExecuteScript(ScriptText, "Error al Crear las Tablas.", , , mDatabaseName)
End Function

Private Function CreatePrimaryKeys() As Boolean
    Dim ScriptText As String
    
    ScriptText = "SET XACT_ABORT ON" & vbCr & vbCr
    
    ScriptText = ScriptText & "BEGIN TRANSACTION" & vbCr & vbCr
    
    If chkCuentaCorriente_Viaje.Value = vbChecked Then
        ScriptText = ScriptText & "ALTER TABLE [dbo].[CuentaCorriente] ADD CONSTRAINT [PK__CuentaCorriente] PRIMARY KEY NONCLUSTERED ([IDMovimiento]) WITH FILLFACTOR = 10 ON [PRIMARY]" & vbCr & vbCr
        ScriptText = ScriptText & "ALTER TABLE [dbo].[Viaje] ADD CONSTRAINT [PK__Viaje] PRIMARY KEY NONCLUSTERED ([FechaHora], [IDRuta]) WITH FILLFACTOR = 10 ON [PRIMARY]" & vbCr & vbCr
        ScriptText = ScriptText & "ALTER TABLE [dbo].[ViajeDetalle] ADD CONSTRAINT [PK__ViajeDetalle] PRIMARY KEY NONCLUSTERED ([FechaHora], [IDRuta], [Indice]) WITH FILLFACTOR = 10 ON [PRIMARY]" & vbCr & vbCr
    End If
    If chkFeriado.Value = vbChecked Then
        ScriptText = ScriptText & "ALTER TABLE [dbo].[Feriado] WITH NOCHECK ADD CONSTRAINT [PK__Feriado] PRIMARY KEY CLUSTERED ([Fecha]) WITH FILLFACTOR = 10 ON [PRIMARY]" & vbCr & vbCr
    End If
    If chkPersonaRespuesta.Value = vbChecked Then
        ScriptText = ScriptText & "ALTER TABLE [dbo].[PersonaRespuesta] WITH NOCHECK ADD CONSTRAINT [PK__PersonaRespuesta] PRIMARY KEY CLUSTERED ([IDPersona], [FechaHora]) ON [PRIMARY]" & vbCr & vbCr
    End If
    If chkVehiculoMantenimientoAccion.Value = vbChecked Then
        ScriptText = ScriptText & "ALTER TABLE [dbo].[VehiculoMantenimientoAccion] WITH NOCHECK ADD CONSTRAINT [PK__VehiculoMantenimientoAccion] PRIMARY KEY CLUSTERED ([IDVehiculoMantenimientoAccion]) WITH FILLFACTOR = 10 ON [PRIMARY]" & vbCr & vbCr
    End If
'    If chkAlarma.Value = vbChecked Then
'        ScriptText = ScriptText & "ALTER TABLE [dbo].[Alarma] WITH NOCHECK ADD CONSTRAINT [PK__Alarma] PRIMARY KEY  CLUSTERED ([IDAlarma]) WITH FILLFACTOR = 10 ON [PRIMARY]" & vbCr & vbCr
'    End If
'    If chkPersonaAlarma.Value = vbChecked Then
'        ScriptText = ScriptText & "ALTER TABLE [dbo].[PersonaAlarma] WITH NOCHECK ADD CONSTRAINT [PK__PersonaAlarma] PRIMARY KEY  CLUSTERED ([IDPersona], [IDPersonaAlarmaGrupo]) WITH FILLFACTOR = 10 ON [PRIMARY]" & vbCr & vbCr
'    End If
    
    ScriptText = ScriptText & "COMMIT TRANSACTION" & vbCr & vbCr
    
    ScriptText = ScriptText & "SET XACT_ABORT OFF"
    
    CreatePrimaryKeys = pDatabase.ExecuteScript(ScriptText, "Error al Crear las Primary Keys.", , , mDatabaseName)
End Function

Private Function BindDefaultsAndRules() As Boolean
    Dim ScriptText As String
    
    ScriptText = "SET XACT_ABORT ON" & vbCr & vbCr
    
    ScriptText = ScriptText & "BEGIN TRANSACTION" & vbCr & vbCr
    
    If chkCuentaCorriente_Viaje.Value = vbChecked Then
        ScriptText = ScriptText & "EXEC sp_bindefault N'[dbo].[def_DateTimeCurrent]', N'[CuentaCorriente].[FechaHoraCreacion]'" & vbCr
        ScriptText = ScriptText & "EXEC sp_bindrule N'[dbo].[rul_DiaSemana]', N'[Viaje].[DiaSemanaBase]'" & vbCr
        ScriptText = ScriptText & "EXEC sp_bindefault N'[dbo].[def_DateTimeCurrent]', N'[Viaje].[FechaHoraCreacion]'" & vbCr
        ScriptText = ScriptText & "EXEC sp_bindefault N'[dbo].[def_DateTimeCurrent]', N'[ViajeDetalle].[FechaHoraCreacion]'" & vbCr
    End If
    If chkFeriado.Value = vbChecked Then
        ScriptText = ScriptText & "EXEC sp_bindefault N'[dbo].[def_DateTimeCurrent]', N'[Feriado].[FechaHoraCreacion]'" & vbCr
    End If
    If chkPersonaRespuesta.Value = vbChecked Then
        ScriptText = ScriptText & "EXEC sp_bindefault N'[dbo].[def_DateTimeCurrent]', N'[PersonaRespuesta].[FechaHoraCreacion]'" & vbCr
    End If
    If chkVehiculoMantenimientoAccion.Value = vbChecked Then
        ScriptText = ScriptText & "EXEC sp_bindefault N'[dbo].[def_DateTimeCurrent]', N'[VehiculoMantenimientoAccion].[FechaHoraCreacion]'" & vbCr
    End If
'    If chkAlarma.Value = vbChecked Then
'        ScriptText = ScriptText & "EXEC sp_bindefault N'[dbo].[def_DateTimeCurrent]', N'[Alarma].[FechaHoraCreacion]'" & vbCr
'    End If
'    If chkPersonaAlarma.Value = vbChecked Then
'        ScriptText = ScriptText & "EXEC sp_bindefault N'[dbo].[def_DateTimeCurrent]', N'[PersonaAlarma].[FechaHoraCreacion]'" & vbCr
'        ScriptText = ScriptText & "EXEC sp_bindefault N'[dbo].[def_DateTimeCurrent]', N'[PersonaAlarma].[FechaHoraModificacion]'" & vbCr
'    End If
    
    ScriptText = ScriptText & "COMMIT TRANSACTION" & vbCr & vbCr
    
    ScriptText = ScriptText & "SET XACT_ABORT OFF"
    
    BindDefaultsAndRules = pDatabase.ExecuteScript(ScriptText, "Error al Crear las Tablas.", , , mDatabaseName)
End Function

Private Function RemoveDatabase() As Boolean
    Dim ScriptText As String
    
    ScriptText = "DROP DATABASE [" & mDatabaseName & "]"
    
    RemoveDatabase = pDatabase.ExecuteScript(ScriptText, "Error al Eliminar la Base de Datos.", , , "master")
End Function

Private Function ShrinkDatabase(ByVal DatabaseName As String, ByVal PercentFree As Long) As Boolean
    Dim ScriptText As String
    
    ScriptText = "DBCC SHRINKDATABASE (N'" & DatabaseName & "', " & PercentFree & ")"
    
    ShrinkDatabase = pDatabase.ExecuteScript(ScriptText, "Error al Comprimir los Archivos de la Base de Datos.", , , "master")
End Function

Private Function DettachDatabase() As Boolean
    Dim ScriptText As String
    Dim recResult As ADODB.Recordset
    
    ScriptText = "DECLARE @result int" & vbCr & vbCr
    
    ScriptText = ScriptText & "EXEC @result = sp_detach_db '" & mDatabaseName & "', 'true'" & vbCr & vbCr
    
    ScriptText = ScriptText & "SELECT @result AS result"
    
    If pDatabase.ExecuteScript(ScriptText, "Error al Detachar la Base de Datos.", True, recResult, "master") Then
        If recResult("result").Value = 1 Then
            ShowErrorMessage "Forms.CierreEjercicio.DettachDatabase", "Error al Detachar la Base de Datos."
        End If
        recResult.Close
        Set recResult = Nothing
    End If
End Function
