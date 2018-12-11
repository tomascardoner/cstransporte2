VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmReporteParametro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametro del Reporte"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ReporteParametro.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6150
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPersonaUltimo 
      Caption         =   "&Ultimo"
      Height          =   315
      Left            =   5400
      TabIndex        =   23
      Top             =   1140
      Width           =   555
   End
   Begin VB.TextBox txtWeekday 
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1140
      Width           =   1050
   End
   Begin VB.CommandButton cmdDateToday 
      Height          =   315
      Left            =   4440
      Picture         =   "ReporteParametro.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   1140
      Width           =   315
   End
   Begin VB.CommandButton cmdDatePrevious 
      Height          =   315
      Left            =   1140
      Picture         =   "ReporteParametro.frx":0156
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   1140
      Width           =   300
   End
   Begin VB.CommandButton cmdDateNext 
      Height          =   315
      Left            =   4140
      Picture         =   "ReporteParametro.frx":06E0
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   1140
      Width           =   300
   End
   Begin VB.ComboBox cboYear 
      Height          =   330
      Left            =   3060
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1140
      Width           =   855
   End
   Begin VB.ComboBox cboMonth 
      Height          =   330
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1140
      Width           =   1815
   End
   Begin VB.TextBox txtCurrency 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1140
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1140
      Width           =   1455
   End
   Begin VB.PictureBox picOcupanteTipo 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1140
      ScaleHeight     =   195
      ScaleWidth      =   2655
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
      Begin VB.OptionButton optOcupanteTipoPasajero 
         Caption         =   "Pasajero"
         Height          =   210
         Left            =   1260
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optOcupanteTipoComision 
         Caption         =   "Comisión"
         Height          =   210
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtLong 
      Height          =   315
      Left            =   1140
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1140
      Width           =   855
   End
   Begin VB.CheckBox chkBoolean 
      Height          =   210
      Left            =   1140
      TabIndex        =   13
      Top             =   1200
      Width           =   195
   End
   Begin VB.TextBox txtParameter 
      BackColor       =   &H8000000F&
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   660
      Width           =   4815
   End
   Begin VB.TextBox txtReporte 
      BackColor       =   &H8000000F&
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   180
      Width           =   4815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3420
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4740
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   315
      Left            =   4800
      TabIndex        =   10
      Top             =   1140
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   105840642
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Top             =   1140
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   105840641
      CurrentDate     =   36950
   End
   Begin MSDataListLib.DataCombo datcboValue 
      Height          =   330
      Left            =   1140
      TabIndex        =   11
      Top             =   1140
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboValue 
      Height          =   330
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1140
      Width           =   4815
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   210
      Left            =   180
      TabIndex        =   0
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label lblParameter 
      AutoSize        =   -1  'True
      Caption         =   "Parámetro:"
      Height          =   210
      Left            =   180
      TabIndex        =   18
      Top             =   720
      Width           =   780
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      Caption         =   "Reporte:"
      Height          =   210
      Left            =   180
      TabIndex        =   16
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmReporteParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mReporteParametro As ReporteParametro
Private mKeyDecimal As Boolean

Private Sub cmdPersonaUltimo_Click()
    If frmMDI.cboPersona.ListIndex > -1 Then
        datcboValue.BoundText = Val(frmMDI.cboPersona.ItemData(frmMDI.cboPersona.ListIndex))
    End If
    cmdOK.SetFocus
End Sub

Public Sub LoadDataAndShow(ByVal ReportName As String, ByRef ReporteParametro As ReporteParametro)
    Dim Index As Long
    
    Set mReporteParametro = ReporteParametro
    
    Load Me
    
    txtReporte.Text = ReportName
    
    txtParameter.Text = mReporteParametro.Nombre
    
    'OCULTO E INICIALIZO TODOS LOS CONTROLES
    cmdDatePrevious.Visible = False
    txtWeekday.Visible = False
    dtpDate.Visible = False
    dtpDate.Value = Date
    dtpDate_Change
    cmdDateNext.Visible = False
    cmdDateToday.Visible = False
    dtpTime.Value = "00:00:00"
    dtpTime.Visible = False
    cboMonth.ListIndex = -1
    cboMonth.Visible = False
    cboYear.ListIndex = -1
    cboYear.Visible = False
    
    datcboValue.BoundText = ""
    datcboValue.Visible = False
    cboValue.ListIndex = -1
    cboValue.Visible = False
    cmdPersonaUltimo.Visible = False
    chkBoolean.Value = vbUnchecked
    chkBoolean.Visible = False
    txtLong.Text = ""
    txtLong.Visible = False
    txtCurrency.Text = ""
    txtCurrency.Visible = False
    picOcupanteTipo.Visible = False
    
    Select Case mReporteParametro.Tipo
        Case REPORTE_PARAMETRO_TIPO_DAY_OF_WEEK
            For Index = 1 To 7
                cboValue.AddItem WeekdayName(Index)
            Next Index
            cboValue.Visible = True
            If Not IsEmpty(mReporteParametro.Valor) Then
                cboValue.ListIndex = mReporteParametro.Valor - 1
            End If
        Case REPORTE_PARAMETRO_TIPO_DATE_TIME
            If Not IsEmpty(mReporteParametro.Valor) Then
                dtpDate.Value = mReporteParametro.Valor
                dtpDate_Change
                dtpTime.Value = mReporteParametro.Valor
            End If
            cmdDatePrevious.Visible = True
            txtWeekday.Visible = True
            dtpDate.Visible = True
            cmdDateNext.Visible = True
            cmdDateToday.Visible = True
            dtpTime.Left = cmdDateToday.Left + cmdDateToday.Width + 45
            dtpTime.Visible = True
        Case REPORTE_PARAMETRO_TIPO_DATE
            If Not IsEmpty(mReporteParametro.Valor) Then
                dtpDate.Value = mReporteParametro.Valor
                dtpDate_Change
            End If
            cmdDatePrevious.Visible = True
            txtWeekday.Visible = True
            dtpDate.Visible = True
            cmdDateNext.Visible = True
            cmdDateToday.Visible = True
        Case REPORTE_PARAMETRO_TIPO_TIME
            If Not IsEmpty(mReporteParametro.Valor) Then
                dtpTime.Value = mReporteParametro.Valor
            End If
            dtpTime.Left = txtLong.Left
            dtpTime.Visible = True
        Case REPORTE_PARAMETRO_TIPO_YEAR_MONTH_FROM, REPORTE_PARAMETRO_TIPO_YEAR_MONTH_TO
            For Index = 1 To 12
                cboMonth.AddItem MonthName(Index)
            Next Index
            For Index = 2000 To 2099
                cboYear.AddItem Index
            Next Index
            If Not IsEmpty(mReporteParametro.Valor) Then
                cboMonth.ListIndex = Month(mReporteParametro.Valor) - 1
                cboYear.ListIndex = Year(mReporteParametro.Valor) - 2000
            End If
            cboMonth.Visible = True
            cboYear.Visible = True
        Case REPORTE_PARAMETRO_TIPO_PERSONA
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDPersona, Apellido + (CASE ISNULL(Nombre, '') WHEN '' THEN '' ELSE ', ' + Nombre END) AS ApellidoNombre FROM Persona WHERE Activo = 1 ORDER BY ApellidoNombre", "IDPersona", "ApellidoNombre", "Personas", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
            datcboValue.Width = txtParameter.Width - cmdPersonaUltimo.Width
            cmdPersonaUltimo.Visible = True
        Case REPORTE_PARAMETRO_TIPO_PERSONA_CLIENTE
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDPersona, Apellido + (CASE ISNULL(Nombre, '') WHEN '' THEN '' ELSE ', ' + Nombre END) AS ApellidoNombre FROM Persona WHERE Activo = 1 AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CLIENTE & "' ORDER BY ApellidoNombre", "IDPersona", "ApellidoNombre", "Clientes", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
            datcboValue.Width = txtParameter.Width - cmdPersonaUltimo.Width
            cmdPersonaUltimo.Visible = True
        Case REPORTE_PARAMETRO_TIPO_PERSONA_CONDUCTOR
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDPersona, Apellido + (CASE ISNULL(Nombre, '') WHEN '' THEN '' ELSE ', ' + Nombre END) AS ApellidoNombre FROM Persona WHERE Activo = 1 AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "' ORDER BY ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
            datcboValue.Width = txtParameter.Width - cmdPersonaUltimo.Width
            cmdPersonaUltimo.Visible = True
        Case REPORTE_PARAMETRO_TIPO_PERSONA_ADMINISTRATIVO
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDPersona, Apellido + (CASE ISNULL(Nombre, '') WHEN '' THEN '' ELSE ', ' + Nombre END) AS ApellidoNombre FROM Persona WHERE Activo = 1 AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO & "' ORDER BY ApellidoNombre", "IDPersona", "ApellidoNombre", "Administrativos", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
            datcboValue.Width = txtParameter.Width - cmdPersonaUltimo.Width
            cmdPersonaUltimo.Visible = True
        Case REPORTE_PARAMETRO_TIPO_PERSONA_ALARMA_GRUPO
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDPersonaAlarmaGrupo, Nombre FROM PersonaAlarmaGrupo WHERE Activo = 1 ORDER BY Nombre", "IDPersonaAlarmaGrupo", "Nombre", "Grupos de Alarmas de Personas", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
        Case REPORTE_PARAMETRO_TIPO_RUTA
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT RTRIM(IDRuta) AS IDRuta FROM Ruta WHERE Activo = 1" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta", "IDRuta", "IDRuta", "Rutas", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
        Case REPORTE_PARAMETRO_TIPO_VEHICULO
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDVehiculo, Nombre FROM Vehiculo WHERE Activo = 1 ORDER BY Nombre", "IDVehiculo", "Nombre", "Vehículos", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
        Case REPORTE_PARAMETRO_TIPO_VEHICULO_MANTENIMIENTO_GRUPO
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDVehiculoMantenimientoGrupo, Nombre FROM VehiculoMantenimientoGrupo WHERE Activo = 1 ORDER BY Nombre", "IDVehiculoMantenimientoGrupo", "Nombre", "Grupos de Mantenimiento de Vehículos", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
        Case REPORTE_PARAMETRO_TIPO_VIAJE_ESTADO
            cboValue.AddItem "Activo"
            cboValue.AddItem "En Progreso"
            cboValue.AddItem "Finalizado"
            cboValue.AddItem "Cancelado"
            cboValue.Visible = True
            If Not IsEmpty(mReporteParametro.Valor) Then
                cboValue.ListIndex = mReporteParametro.Valor
            End If
        Case REPORTE_PARAMETRO_TIPO_PERSONA_TIPO
            cboValue.AddItem ENTIDAD_TIPO_PERSONA_CLIENTE_NOMBRE
            cboValue.AddItem ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO_NOMBRE
            cboValue.AddItem ENTIDAD_TIPO_PERSONA_CONDUCTOR_NOMBRE
            cboValue.Visible = True
            If Not IsEmpty(mReporteParametro.Valor) Then
                Select Case mReporteParametro.Valor
                    Case ENTIDAD_TIPO_PERSONA_CLIENTE
                        cboValue.ListIndex = 0
                    Case ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO
                        cboValue.ListIndex = 1
                    Case ENTIDAD_TIPO_PERSONA_CONDUCTOR
                        cboValue.ListIndex = 2
                End Select
            End If
        Case REPORTE_PARAMETRO_TIPO_CUENTACORRIENTE_GRUPO
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDCuentaCorrienteGrupo, Nombre FROM CuentaCorrienteGrupo WHERE Activo = 1 ORDER BY Nombre", "IDCuentaCorrienteGrupo", "Nombre", "Grupos", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
        Case REPORTE_PARAMETRO_TIPO_CUENTACORRIENTE_CAJA
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
        Case REPORTE_PARAMETRO_TIPO_ALARMA
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDAlarma, Nombre FROM Alarma WHERE Activo = 1 ORDER BY Nombre", "IDAlarma", "Nombre", "Alarmas", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
        Case REPORTE_PARAMETRO_TIPO_BOOLEAN
            chkBoolean.Visible = True
            chkBoolean.Value = IIf(mReporteParametro.Valor, vbChecked, vbUnchecked)
        Case REPORTE_PARAMETRO_TIPO_NUMBER_LONG
            txtLong.Visible = True
            txtLong.Text = mReporteParametro.Valor
        Case REPORTE_PARAMETRO_TIPO_CURRENCY
            txtCurrency.Visible = True
            txtCurrency.Text = Format(mReporteParametro.Valor, "Currency")
        Case REPORTE_PARAMETRO_TIPO_VIAJE_DETALLE_OCUPANTE_TIPO
            picOcupanteTipo.Visible = True
            optOcupanteTipoComision.Value = (mReporteParametro.Valor = OCUPANTE_TIPO_COMISION)
            optOcupanteTipoPasajero.Value = (mReporteParametro.Valor = OCUPANTE_TIPO_PASAJERO)
        Case REPORTE_PARAMETRO_TIPO_LISTA_PRECIO
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDListaPrecio, Nombre FROM ListaPrecio WHERE Activo = 1 ORDER BY Nombre", "IDListaPrecio", "Nombre", "Listas de Precios", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
        Case REPORTE_PARAMETRO_TIPO_LUGAR
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDLugar, Nombre FROM Lugar WHERE Activo = 1 ORDER BY Nombre", "IDLugar", "Nombre", "Lugares", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
        Case REPORTE_PARAMETRO_TIPO_LUGAR_GRUPO
            If Not CSM_Control_DataCombo.FillFromSQL(datcboValue, "SELECT IDLugarGrupo, Nombre FROM LugarGrupo WHERE Activo = 1 ORDER BY Nombre", "IDLugarGrupo", "Nombre", "Grupos de Lugares", cscpItemOrNone, mReporteParametro.Valor) Then
                Hide
                Exit Sub
            End If
            datcboValue.Visible = True
    End Select
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show vbModal, frmMDI
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    Select Case mReporteParametro.Tipo
        Case REPORTE_PARAMETRO_TIPO_DAY_OF_WEEK
            If cboValue.ListIndex = -1 Then
                mReporteParametro.Valor = Empty
                mReporteParametro.ValorLeyenda = ""
            Else
                mReporteParametro.Valor = cboValue.ListIndex + 1
                mReporteParametro.ValorLeyenda = cboValue.Text
            End If
        Case REPORTE_PARAMETRO_TIPO_DATE_TIME
            mReporteParametro.Valor = CDate(Format(dtpDate.Value, "Short Date") & " " & Format(dtpTime.Value, "Short Time"))
            mReporteParametro.ValorLeyenda = Format(dtpDate.Value, "Short Date") & " " & Format(dtpTime.Value, "Short Time")
        Case REPORTE_PARAMETRO_TIPO_DATE
            mReporteParametro.Valor = dtpDate.Value
            mReporteParametro.ValorLeyenda = Format(mReporteParametro.Valor, "Short Date")
        Case REPORTE_PARAMETRO_TIPO_TIME
            mReporteParametro.Valor = dtpTime.Value
            mReporteParametro.ValorLeyenda = Format(mReporteParametro.Valor, "Short Time")
        Case REPORTE_PARAMETRO_TIPO_YEAR_MONTH_FROM, REPORTE_PARAMETRO_TIPO_YEAR_MONTH_TO
            If cboMonth.ListIndex = -1 Then
                MsgBox "Debe seleccionar el Mes", vbInformation, App.Title
                cboMonth.SetFocus
                Exit Sub
            End If
            If cboYear.ListIndex = -1 Then
                MsgBox "Debe seleccionar el Año", vbInformation, App.Title
                cboYear.SetFocus
                Exit Sub
            End If
            If mReporteParametro.Tipo = REPORTE_PARAMETRO_TIPO_YEAR_MONTH_FROM Then
                mReporteParametro.Valor = DateSerial(cboYear.ListIndex + 2000, cboMonth.ListIndex + 1, 1)
            Else
                mReporteParametro.Valor = DateAdd("s", -1, DateSerial(cboYear.ListIndex + 2000, cboMonth.ListIndex + 2, 1))
            End If
            mReporteParametro.ValorLeyenda = cboMonth.Text & " de " & cboYear.Text
        Case REPORTE_PARAMETRO_TIPO_PERSONA, REPORTE_PARAMETRO_TIPO_PERSONA_CLIENTE, REPORTE_PARAMETRO_TIPO_PERSONA_CONDUCTOR, REPORTE_PARAMETRO_TIPO_PERSONA_ADMINISTRATIVO, REPORTE_PARAMETRO_TIPO_PERSONA_ALARMA_GRUPO, REPORTE_PARAMETRO_TIPO_VEHICULO, REPORTE_PARAMETRO_TIPO_VEHICULO_MANTENIMIENTO_GRUPO, REPORTE_PARAMETRO_TIPO_CUENTACORRIENTE_GRUPO, REPORTE_PARAMETRO_TIPO_CUENTACORRIENTE_CAJA, REPORTE_PARAMETRO_TIPO_ALARMA, REPORTE_PARAMETRO_TIPO_LISTA_PRECIO, REPORTE_PARAMETRO_TIPO_LUGAR, REPORTE_PARAMETRO_TIPO_LUGAR_GRUPO
            If Val(datcboValue.BoundText) = 0 Then
                mReporteParametro.Valor = Empty
                mReporteParametro.ValorLeyenda = ""
            Else
                mReporteParametro.Valor = Val(datcboValue.BoundText)
                mReporteParametro.ValorLeyenda = datcboValue.Text
            End If
        Case REPORTE_PARAMETRO_TIPO_RUTA
            If datcboValue.BoundText = "" Then
                mReporteParametro.Valor = Empty
                mReporteParametro.ValorLeyenda = ""
            Else
                mReporteParametro.Valor = datcboValue.BoundText
                mReporteParametro.ValorLeyenda = datcboValue.Text
            End If
        Case REPORTE_PARAMETRO_TIPO_VIAJE_ESTADO
            mReporteParametro.Valor = Switch(cboValue.ListIndex = -1, Empty, cboValue.ListIndex = 0, VIAJE_ESTADO_ACTIVO, cboValue.ListIndex = 1, VIAJE_ESTADO_EN_PROGRESO, cboValue.ListIndex = 2, VIAJE_ESTADO_FINALIZADO, cboValue.ListIndex = 3, VIAJE_ESTADO_CANCELADO)
            mReporteParametro.ValorLeyenda = cboValue.Text
        Case REPORTE_PARAMETRO_TIPO_PERSONA_TIPO
            mReporteParametro.Valor = Switch(cboValue.ListIndex = -1, Empty, cboValue.ListIndex = 0, ENTIDAD_TIPO_PERSONA_CLIENTE, cboValue.ListIndex = 1, ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO, cboValue.ListIndex = 2, ENTIDAD_TIPO_PERSONA_CONDUCTOR)
            mReporteParametro.ValorLeyenda = cboValue.Text
        Case REPORTE_PARAMETRO_TIPO_BOOLEAN
            mReporteParametro.Valor = (chkBoolean.Value = vbChecked)
            mReporteParametro.ValorLeyenda = IIf(chkBoolean.Value = vbChecked, "Sí", "No")
        Case REPORTE_PARAMETRO_TIPO_NUMBER_LONG
            mReporteParametro.Valor = Val(txtLong.Text)
            mReporteParametro.ValorLeyenda = Val(txtLong.Text)
        Case REPORTE_PARAMETRO_TIPO_CURRENCY
            mReporteParametro.Valor = CCur(txtCurrency.Text)
            mReporteParametro.ValorLeyenda = Format(CCur(txtCurrency.Text), "Currency")
        Case REPORTE_PARAMETRO_TIPO_VIAJE_DETALLE_OCUPANTE_TIPO
            mReporteParametro.Valor = Switch(optOcupanteTipoComision.Value, OCUPANTE_TIPO_COMISION, optOcupanteTipoPasajero.Value, OCUPANTE_TIPO_PASAJERO, optOcupanteTipoComision.Value = False And optOcupanteTipoPasajero.Value = False, "")
            mReporteParametro.ValorLeyenda = Switch(optOcupanteTipoComision.Value, "Comisión", optOcupanteTipoPasajero.Value, "Pasajero", optOcupanteTipoComision.Value = False And optOcupanteTipoPasajero.Value = False, "")
    End Select
    
    Tag = "OK"
    Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mReporteParametro = Nothing
    Set frmReporteParametro = Nothing
End Sub

Private Sub txtCurrency_GotFocus()
    CSM_Control_TextBox.SelAllText txtCurrency
End Sub

Private Sub txtCurrency_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtCurrency_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48 And KeyAscii <> 45) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtCurrency.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtCurrency.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtCurrency_LostFocus()
    If Not IsNumeric(txtCurrency.Text) Then
        txtCurrency.Text = Val(txtCurrency.Text)
    End If
    txtCurrency.Text = Format(CCur(txtCurrency.Text), "Currency")
End Sub

Private Sub txtLong_GotFocus()
    CSM_Control_TextBox.SelAllText txtLong
End Sub

Private Sub txtLong_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLong_LostFocus()
    txtLong.Text = Val(txtLong.Text)
End Sub

Private Sub dtpDate_Change()
    txtWeekday.Text = WeekdayName(Weekday(dtpDate.Value))
End Sub

Private Sub cmdDatePrevious_Click()
    dtpDate.Value = DateAdd("d", -1, dtpDate.Value)
    dtpDate.SetFocus
    dtpDate_Change
End Sub

Private Sub cmdDateNext_Click()
    dtpDate.Value = DateAdd("d", 1, dtpDate.Value)
    dtpDate.SetFocus
    dtpDate_Change
End Sub

Private Sub cmdDateToday_Click()
    Dim OldValue As Date
    
    OldValue = dtpDate.Value
    dtpDate.Value = Date
    dtpDate.SetFocus
    If OldValue <> dtpDate.Value Then
        dtpDate_Change
    End If
End Sub
