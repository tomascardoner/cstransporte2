VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmViajeDetalleAsistencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asistencia al"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajeDetalleAsistencia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   9030
   Begin VB.TextBox txtFacturaNumero 
      Height          =   315
      Left            =   5280
      MaxLength       =   20
      TabIndex        =   19
      Top             =   3660
      Width           =   2115
   End
   Begin VB.TextBox txtOperacion 
      Height          =   315
      Left            =   6780
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.ComboBox cboCuotas 
      Height          =   330
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1800
      Width           =   630
   End
   Begin VB.TextBox txtRetira 
      Height          =   315
      Left            =   5280
      MaxLength       =   50
      TabIndex        =   28
      Top             =   4560
      Width           =   3615
   End
   Begin VB.CheckBox chkForzarDebito 
      Height          =   210
      Left            =   7140
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4140
      Width           =   195
   End
   Begin VB.CheckBox chkEntregada 
      Alignment       =   1  'Right Justify
      Caption         =   "Entregada:"
      Height          =   210
      Left            =   4350
      TabIndex        =   24
      Top             =   4140
      Width           =   1125
   End
   Begin VB.ComboBox cboRealizado 
      Height          =   330
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtPersonaCuentaCorriente 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3240
      Width           =   3315
   End
   Begin VB.TextBox txtOrden 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtBaja 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   5100
      Width           =   3015
   End
   Begin VB.TextBox txtDestino 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   4740
      Width           =   3015
   End
   Begin VB.TextBox txtSube 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4140
      Width           =   3015
   End
   Begin VB.TextBox txtOrigen 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3780
      Width           =   3015
   End
   Begin VB.TextBox txtRuta 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1740
      Width           =   3015
   End
   Begin VB.TextBox txtHora 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1095
   End
   Begin VB.TextBox txtFecha 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2100
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1335
   End
   Begin VB.CommandButton cmdCuentaCorrienteCaja 
      Caption         =   "..."
      Height          =   315
      Left            =   8640
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Cajas"
      Top             =   2820
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtImporteCuentaCorriente 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtAsiento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox txtDiaSemana 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1050
   End
   Begin VB.TextBox txtSaldoActual 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   7260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtImporteContado 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5280
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1455
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   53
      Top             =   780
      Width           =   8775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7680
      TabIndex        =   30
      Top             =   5100
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   29
      Top             =   5100
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboCuentaCorrienteCaja 
      Height          =   330
      Left            =   5280
      TabIndex        =   14
      Top             =   2820
      Width           =   3315
      _ExtentX        =   5847
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
   Begin MSComCtl2.DTPicker dtpEntregadaHora 
      Height          =   315
      Left            =   7500
      TabIndex        =   26
      Top             =   4080
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
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
      Format          =   92930050
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpEntregadaFecha 
      Height          =   315
      Left            =   5760
      TabIndex        =   25
      Top             =   4080
      Visible         =   0   'False
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
      Format          =   92930049
      CurrentDate     =   36950
   End
   Begin MSDataListLib.DataCombo datcboMedioPago 
      Height          =   330
      Left            =   6780
      TabIndex        =   4
      Top             =   1440
      Width           =   2115
      _ExtentX        =   3731
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
   Begin VB.Label lblFacturaNumero 
      AutoSize        =   -1  'True
      Caption         =   "Factura Nº:"
      Height          =   210
      Left            =   4380
      TabIndex        =   18
      Top             =   3720
      Width           =   825
   End
   Begin VB.Label lblOperacion 
      AutoSize        =   -1  'True
      Caption         =   "Operación:"
      Height          =   210
      Left            =   5940
      TabIndex        =   7
      Top             =   1860
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblCuotas 
      AutoSize        =   -1  'True
      Caption         =   "Cuotas:"
      Height          =   210
      Left            =   4380
      TabIndex        =   5
      Top             =   1860
      Width           =   555
   End
   Begin VB.Label lblRetira 
      AutoSize        =   -1  'True
      Caption         =   "Retira:"
      Height          =   210
      Left            =   4380
      TabIndex        =   27
      Top             =   4620
      Width           =   465
   End
   Begin VB.Label lblForzarDebito 
      AutoSize        =   -1  'True
      Caption         =   "Debitar Viaje:"
      Height          =   210
      Left            =   6060
      TabIndex        =   22
      Top             =   4140
      Width           =   960
   End
   Begin VB.Label lblPersonaCuentaCorriente 
      AutoSize        =   -1  'True
      Caption         =   "Debitar a:"
      Height          =   210
      Left            =   4380
      TabIndex        =   16
      Top             =   3300
      Width           =   690
   End
   Begin VB.Label lblOrden 
      AutoSize        =   -1  'True
      Caption         =   "Orden:"
      Height          =   210
      Left            =   120
      TabIndex        =   38
      Top             =   2340
      Width           =   495
   End
   Begin VB.Label lblRealizado 
      AutoSize        =   -1  'True
      Caption         =   "Realizado:"
      Height          =   210
      Left            =   4380
      TabIndex        =   20
      Top             =   4140
      Width           =   750
   End
   Begin VB.Label lblBaja 
      AutoSize        =   -1  'True
      Caption         =   "Baja:"
      Height          =   210
      Left            =   120
      TabIndex        =   50
      Top             =   5160
      Width           =   360
   End
   Begin VB.Label lblSube 
      AutoSize        =   -1  'True
      Caption         =   "Sube:"
      Height          =   210
      Left            =   120
      TabIndex        =   46
      Top             =   4200
      Width           =   420
   End
   Begin VB.Label lblCuentaCorrienteCaja 
      AutoSize        =   -1  'True
      Caption         =   "Caja:"
      Height          =   210
      Left            =   4380
      TabIndex        =   13
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label lblImporteCuentaCorriente 
      AutoSize        =   -1  'True
      Caption         =   "P. Cta. Cte.:"
      Height          =   210
      Left            =   4380
      TabIndex        =   9
      Top             =   2460
      Width           =   840
   End
   Begin VB.Label lblAsiento 
      AutoSize        =   -1  'True
      Caption         =   "Asiento N°"
      Height          =   210
      Left            =   120
      TabIndex        =   40
      Top             =   2820
      Width           =   765
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   4200
      Y1              =   960
      Y2              =   5460
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "&Fecha:"
      Height          =   210
      Left            =   120
      TabIndex        =   31
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "&Hora:"
      Height          =   210
      Left            =   120
      TabIndex        =   34
      Top             =   1440
      Width           =   390
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "&Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   36
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblSaldoActual 
      Alignment       =   2  'Center
      Caption         =   "Saldo Actual:"
      Height          =   210
      Left            =   7320
      TabIndex        =   11
      Top             =   2160
      Width           =   1380
   End
   Begin VB.Label lblOrigen 
      AutoSize        =   -1  'True
      Caption         =   "&Origen:"
      Height          =   210
      Left            =   120
      TabIndex        =   44
      Top             =   3840
      Width           =   525
   End
   Begin VB.Label lblDestino 
      AutoSize        =   -1  'True
      Caption         =   "&Destino:"
      Height          =   210
      Left            =   120
      TabIndex        =   48
      Top             =   4800
      Width           =   585
   End
   Begin VB.Label lblImporteContado 
      AutoSize        =   -1  'True
      Caption         =   "P. Contado:"
      Height          =   210
      Left            =   4380
      TabIndex        =   2
      Top             =   1500
      Width           =   825
   End
   Begin VB.Label lblImporte 
      AutoSize        =   -1  'True
      Caption         =   "&Importe:"
      Height          =   210
      Left            =   4380
      TabIndex        =   0
      Top             =   1080
      Width           =   570
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos de la Asistencia"
      Height          =   210
      Left            =   780
      TabIndex        =   52
      Top             =   300
      Width           =   2805
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "ViajeDetalleAsistencia.frx":054A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "&Pasajero:"
      Height          =   210
      Left            =   120
      TabIndex        =   42
      Top             =   3300
      Width           =   675
   End
End
Attribute VB_Name = "frmViajeDetalleAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mViajeDetalle As ViajeDetalle

Private mKeyDecimal As Boolean

Private mMedioPago As MedioPago

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef ViajeDetalle As ViajeDetalle)
    Dim Viaje As Viaje
    Dim Persona As Persona
    Dim Lugar As Lugar
    Dim CuentaCorrienteCaja As CuentaCorrienteCaja
    Dim ShowSaldo As Boolean
    
    Set mViajeDetalle = ViajeDetalle
    Set ViajeDetalle = Nothing
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    Caption = "Asistencia " & IIf(mViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO, "al Pasajero", "la Comisión")
    
    Set Viaje = New Viaje
    Viaje.FechaHora = mViajeDetalle.FechaHora
    Viaje.IDRuta = mViajeDetalle.IDRuta
    If Not Viaje.Load() Then
        Unload Me
        Exit Sub
    End If
    
    chkForzarDebito.Enabled = pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_INASISTENCIA_NODEBITAR, False)
    
    With mViajeDetalle
        txtDiaSemana.Text = .FechaHora_WeekdayName
        txtFecha.Text = .FechaHora_FormattedAsDate
        txtHora.Text = .FechaHora_FormattedAsTime
        txtRuta.Text = Viaje.Ruta_DisplayName
        txtOrden.Text = .Orden
        txtAsiento.Text = IIf(.Asiento = -1, "", .Asiento)
        txtPersona.Tag = .IDPersona
        
        Set Persona = New Persona
        Persona.IDPersona = .IDPersona
        If Persona.Load() Then
            txtPersona.Text = Persona.ApellidoNombre
        End If
        
        SetLastPersona .IDPersona, txtPersona.Text
        
        Set Lugar = New Lugar
        Lugar.IDLugar = .IDOrigen
        If Lugar.Load() Then
            txtOrigen.Text = Lugar.Nombre
        Else
            txtOrigen.Text = ""
        End If
        txtSube.Text = .Sube
        
        Lugar.IDLugar = .IDDestino
        If Lugar.Load() Then
            txtDestino.Text = Lugar.Nombre
        Else
            txtDestino.Text = ""
        End If
        Set Lugar = Nothing
        txtBaja.Text = .Baja
        
        txtImporte.Text = .Importe_Formatted
        txtImporteContado.Text = .ImporteContado_Formatted
        txtImporteContado_LostFocus
        
        'MEDIO DE PAGO
        If Not CSM_Control_DataCombo.FillFromSQL(datcboMedioPago, "SELECT IDMedioPago, Nombre FROM MedioPago WHERE Activo = 1 OR IDMedioPago = " & .IDMedioPago & " ORDER BY Nombre", "IDMedioPago", "Nombre", "Medios de Pago", cscpItemOrfirst, IIf(.IDMedioPago = 0, pParametro.MedioPago_Predeterminado_ID, .IDMedioPago)) Then
            Unload Me
            Exit Sub
        End If
        cboCuotas.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboCuotas, .Cuotas_Formatted, cscpItemOrfirst)
        txtOperacion.Text = .Operacion
        
        txtImporteCuentaCorriente.Text = .ImporteCuentaCorriente_Formatted
        
        If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL, False) Then
            If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 OR IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & " ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, .IDCuentaCorrienteCaja) Then
                Unload Me
                Exit Sub
            End If
        Else
            If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT CuentaCorrienteCaja.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre FROM CuentaCorrienteCaja LEFT JOIN Persona ON CuentaCorrienteCaja.IDPersona = Persona.IDPersona WHERE (CuentaCorrienteCaja.Activo = 1 OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & ") AND (CuentaCorrienteCaja.MostrarSiempre = 1 OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & pUsuario.IDCuentaCorrienteCaja & " OR Persona.IDPersona = " & Viaje.IDConductor & " OR Persona.IDPersona = " & Viaje.IDConductor2 & " OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & ") ORDER BY CuentaCorrienteCaja.Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, IIf(.IDCuentaCorrienteCaja = 0, pUsuario.IDCuentaCorrienteCaja, .IDCuentaCorrienteCaja)) Then
                Unload Me
                Exit Sub
            End If
        End If
        
        If .ImporteContado <> 0 Then
            Set CuentaCorrienteCaja = New CuentaCorrienteCaja
            CuentaCorrienteCaja.IDPersona = Viaje.IDConductor
            CuentaCorrienteCaja.NoMatchRaiseError = False
            Call CuentaCorrienteCaja.LoadByPersona
            If .IDCuentaCorrienteCaja <> pUsuario.IDCuentaCorrienteCaja And .IDCuentaCorrienteCaja <> CuentaCorrienteCaja.IDCuentaCorrienteCaja Then
                If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_IMPORTE_OTRO_USUARIO, False) Then
                    txtImporteContado.Enabled = False
                    datcboMedioPago.Enabled = False
                    cboCuotas.Enabled = False
                    txtOperacion.Enabled = False
                End If
                If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_CAJA_OTRO_USUARIO, False) Then
                    datcboCuentaCorrienteCaja.Enabled = False
                End If
            End If
            Set CuentaCorrienteCaja = Nothing
        End If
        
        Set Viaje = Nothing
        
        txtFacturaNumero.Text = .FacturaNumero
                
        If .Realizado = 0 Then
            cboRealizado.ListIndex = 1
        Else
            cboRealizado.ListIndex = .Realizado
        End If
        lblRealizado.Visible = (.OcupanteTipo = OCUPANTE_TIPO_PASAJERO)
        cboRealizado.Visible = (.OcupanteTipo = OCUPANTE_TIPO_PASAJERO)
        chkForzarDebito.Value = IIf(mViajeDetalle.ForzarDebito, vbChecked, vbUnchecked)
        
        chkEntregada.Value = vbChecked
        chkEntregada.Visible = (mViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_COMISION)
        
        dtpEntregadaFecha.Visible = (mViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_COMISION)
        dtpEntregadaHora.Visible = (mViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_COMISION)
        lblRetira.Visible = (mViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_COMISION)
        txtRetira.Visible = (mViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_COMISION)
        If .Entregada Then
            dtpEntregadaFecha.Value = .EntregadaFechaHora
            dtpEntregadaHora.Value = .EntregadaFechaHora
            txtRetira.Text = .Retira
        Else
            dtpEntregadaFecha.Value = Date
            dtpEntregadaHora.Value = Time
            txtRetira.Text = ""
        End If
        
        If .IDPersonaCuentaCorriente = 0 Then
            txtPersonaCuentaCorriente.Text = ""
        Else
            Persona.IDPersona = .IDPersonaCuentaCorriente
            If Persona.Load() Then
                txtPersonaCuentaCorriente.Text = Persona.ApellidoNombre
            End If
        End If
        Persona.ViajeActual_FechaHora = .FechaHora
        Persona.ViajeActual_IDRuta = .IDRuta
        Persona.ViajeActual_Indice = .Indice
        Persona.LoadSaldoActual
        txtSaldoActual.Tag = Persona.SaldoActual
        
        CalcularImporteCuentaCorriente
        
        'AVISA QUE DEBE
        If Persona.SaldoActual < 0 And Not Persona.PermiteViajarSinPagar Then
            Load frmPersonaSaldo
            frmPersonaSaldo.txtPersona.Tag = Persona.IDPersona
            frmPersonaSaldo.txtPersona.Text = " " & Persona.ApellidoNombre
            frmPersonaSaldo.txtSaldo.Text = Persona.SaldoActual_Formatted
            frmPersonaSaldo.FillListView
            frmPersonaSaldo.Show
            ShowSaldo = True
        End If
        Set Persona = Nothing
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
    
    If ShowSaldo Then
        frmPersonaSaldo.SetFocus
    End If
    
    'SI CORRESPONDE, AVISO QUE LE QUEDAN POCOS VIAJES PREPAGOS
    Dim PrepagosRestantes As Integer
    
    If mViajeDetalle.ListaPrecio.PrepagoEs Then
        If mViajeDetalle.ListaPrecio.PrepagoReservasCantidad > 0 Then
            PrepagosRestantes = mViajeDetalle.ObtenerPrepagosRestantes
            If PrepagosRestantes <= pParametro.Persona_Prepago_AvisoCantidadRestante Then
                If PrepagosRestantes = 1 Then
                    MsgBox "Informe al Cliente que le queda sólo 1 Viaje Prepago.", vbExclamation, App.Title
                Else
                    MsgBox "Informe al Cliente que le quedan sólo " & PrepagosRestantes & " Viajes Prepagos.", vbExclamation, App.Title
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    dtpEntregadaFecha.CalendarTitleBackColor = vbDesktop
    
    cboRealizado.AddItem "--"
    cboRealizado.AddItem "Sí"
    cboRealizado.AddItem "No"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mMedioPago = Nothing
    Set mViajeDetalle = Nothing
End Sub

Private Sub txtImporteContado_GotFocus()
    CSM_Control_TextBox.SelAllText txtImporteContado
End Sub

Private Sub txtImporteContado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtImporteContado_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtImporteContado.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtImporteContado.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtImporteContado_Change()
    CalcularImporteCuentaCorriente
End Sub

Private Sub txtImporteContado_LostFocus()
    If Not IsNumeric(txtImporteContado.Text) Then
        txtImporteContado.Text = Val(txtImporteContado.Text)
    End If
    txtImporteContado.Text = Format(CCur(txtImporteContado.Text), "Currency")

    lblCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0)
    datcboCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0)
End Sub

Private Sub datcboMedioPago_Change()
    If (mViajeDetalle.IDRuta <> pParametro.Ruta_ID_Otra) And Val(datcboMedioPago.BoundText) > 0 Then
        Set mMedioPago = New MedioPago
        mMedioPago.IDMedioPago = Val(datcboMedioPago.BoundText)
        If mMedioPago.Load() Then
            lblCuotas.Visible = mMedioPago.UtilizaOperacion
            If mMedioPago.UtilizaOperacion Then
                If mMedioPago.MedioPagoPlan.LoadCuotas Then
                    CSM_Control_ComboBox.FillFromCollection cboCuotas, mMedioPago.MedioPagoPlan.CCuotas, "Cuota", "Cuota", cscpCurrentOrFirst
                End If
            End If
            cboCuotas.Visible = mMedioPago.UtilizaOperacion
            lblOperacion.Visible = mMedioPago.UtilizaOperacion
            txtOperacion.Visible = mMedioPago.UtilizaOperacion
        
            lblCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0 And mMedioPago.IDCuentaCorrienteCaja = 0)
            datcboCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0 And mMedioPago.IDCuentaCorrienteCaja = 0)
            cmdCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0 And mMedioPago.IDCuentaCorrienteCaja = 0)
        End If
    Else
        lblCuotas.Visible = False
        cboCuotas.Visible = False
        lblOperacion.Visible = False
        txtOperacion.Visible = False
    End If
End Sub

Private Sub chkEntregada_Click()
    dtpEntregadaFecha.Visible = (chkEntregada.Value = vbChecked)
    dtpEntregadaHora.Visible = (chkEntregada.Value = vbChecked)
    lblRetira.Visible = (chkEntregada.Value = vbChecked)
    txtRetira.Visible = (chkEntregada.Value = vbChecked)
End Sub

Private Sub cboRealizado_Click()
    lblForzarDebito.Visible = (mViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO And cboRealizado.ListIndex = 2)
    chkForzarDebito.Visible = (mViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO And cboRealizado.ListIndex = 2)
    If cboRealizado.ListIndex = 2 Then
        chkForzarDebito.Value = vbChecked
    End If
End Sub

Private Sub txtFacturaNumero_GotFocus()
    CSM_Control_TextBox.SelAllText txtFacturaNumero
End Sub

Private Sub txtRetira_GotFocus()
    CSM_Control_TextBox.SelAllText txtRetira
End Sub

Private Sub txtRetira_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRetira_LostFocus()
    txtRetira.Text = UCase(txtRetira.Text)
    txtRetira.Text = CleanInvalidSpaces(txtRetira.Text)
End Sub

Private Sub cmdOK_Click()
    Dim Viaje As Viaje
    Dim AsientoAsignado As Long
    
    If Not IsNumeric(txtImporteContado.Text) Then
        MsgBox "El Importe de Contado ingresado es incorrecto.", vbInformation, App.Title
        txtImporteContado.SetFocus
        Exit Sub
    End If
    If CCur(txtImporteContado.Text) < 0 Then
        MsgBox "El Importe de Contado debe ser mayor o igual a cero.", vbInformation, App.Title
        txtImporteContado.SetFocus
        Exit Sub
    End If
    If CCur(txtImporteContado.Text) > 0 And mMedioPago.IDCuentaCorrienteCaja = 0 And Val(datcboCuentaCorrienteCaja.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Caja.", vbInformation, App.Title
        On Error Resume Next
        datcboCuentaCorrienteCaja.SetFocus
        Exit Sub
    End If
    If pParametro.ViajeDetalle_MedioPago_UtilizaOperacion_ObligaFacturaNumero Then
        If mMedioPago.UtilizaOperacion And Len(Trim(txtFacturaNumero.Text)) = 0 Then
            MsgBox "Debe especificar el Nº de Factura.", vbInformation, App.Title
            txtFacturaNumero.SetFocus
            Exit Sub
        End If
    End If
    
    If cboRealizado.ListIndex = 2 Then
        Call ViajeDetalle_ShowViajeVuelta(mViajeDetalle, "Hay %1 Reserva(s) sin Asistencia para este Pasajero en el mismo Día.")
    End If
    
    'SI PASO DE NO REALIZADO A OTRO ESTADO DE REALIZADO,
    'Y NO PERMITE RESERVAS CONDICIONALES, VERIFICO QUE HAYA LUGAR
    If mViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO And Not pParametro.Permitir_Reservas_Condicionales Then
        If mViajeDetalle.Realizado = 2 And cboRealizado.ListIndex <> 2 Then
            Set Viaje = New Viaje
            Viaje.FechaHora = mViajeDetalle.FechaHora
            Viaje.IDRuta = mViajeDetalle.IDRuta
            Viaje.IDOrigen = mViajeDetalle.IDOrigen
            Viaje.IDDestino = mViajeDetalle.IDDestino
            AsientoAsignado = Viaje.Asiento_Asignar_GetAsiento(0)
            Select Case AsientoAsignado
                Case -1
                    MsgBox "No se puede actualizar la Reserva porque no hay más lugar en este Viaje.", vbExclamation, App.Title
                    Set Viaje = Nothing
                    Exit Sub
                Case -2
                Case -3
                Case Else
            End Select
            Set Viaje = Nothing
        End If
    End If
    
    With mViajeDetalle
        .ImporteContado = CCur(txtImporteContado.Text)
        If .IDRuta <> pParametro.Ruta_ID_Otra Then
            .IDMedioPago = Val(datcboMedioPago.BoundText)
            If mMedioPago.UtilizaOperacion Then
                .Cuotas = Val(cboCuotas.Text)
                .Operacion = txtOperacion.Text
            Else
                .Cuotas = 0
                .Operacion = ""
            End If
        Else
            .IDMedioPago = 0
            .Cuotas = 0
            .Operacion = ""
        End If
        .ImporteCuentaCorriente = CCur(txtImporteCuentaCorriente.Text)
        If CCur(txtImporteContado.Text) > 0 Then
            If mMedioPago.IDCuentaCorrienteCaja = 0 Then
                .IDCuentaCorrienteCaja = Val(datcboCuentaCorrienteCaja.BoundText)
            Else
                .IDCuentaCorrienteCaja = mMedioPago.IDCuentaCorrienteCaja
            End If
        Else
            .IDCuentaCorrienteCaja = 0
        End If
        .FacturaNumero = txtFacturaNumero.Text
        If mViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
            .Realizado = cboRealizado.ListIndex
            .ForzarDebito = (chkForzarDebito.Value = vbChecked)
            .Retira = ""
        Else
            .Entregada = (chkEntregada.Value = vbChecked)
            If .Entregada Then
                .EntregadaFechaHora = CDate(Format(dtpEntregadaFecha.Value, "Short Date") & " " & Format(dtpEntregadaHora.Value, "Short Time"))
                .Retira = txtRetira.Text
            Else
                .EntregadaFechaHora = DATE_TIME_FIELD_NULL_VALUE
                .Retira = ""
            End If
        End If
        If Not .Realizar() Then
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CalcularImporteCuentaCorriente()
    Dim ImporteRestante As Currency
    
    On Error Resume Next
    
    If CCur(txtImporteContado.Text) >= CCur(txtImporte.Text) Then
        ImporteRestante = 0
    Else
        ImporteRestante = CCur(txtImporte.Text) - CCur(txtImporteContado.Text)
    End If
    
    If CCur(txtSaldoActual.Tag) > 0 And CCur(txtImporte.Text) > 0 Then
        If CCur(txtSaldoActual.Tag) >= ImporteRestante Then
            txtImporteCuentaCorriente.Text = Format(ImporteRestante, "Currency")
        Else
            txtImporteCuentaCorriente.Text = Format(CCur(txtSaldoActual.Tag), "Currency")
        End If
        txtSaldoActual.Text = Format(CCur(txtSaldoActual.Tag) - CCur(txtImporteCuentaCorriente.Text), "Currency")
    Else
        txtImporteCuentaCorriente.Text = Format(0, "Currency")
        txtSaldoActual.Text = Format(CCur(txtSaldoActual.Tag), "Currency")
    End If
End Sub

Public Sub FillComboBoxCuentaCorrienteCaja()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboCuentaCorrienteCaja.BoundText)
    Set recData = datcboCuentaCorrienteCaja.RowSource
    recData.Requery
    Set recData = Nothing
    datcboCuentaCorrienteCaja.BoundText = KeySave
End Sub

Public Sub FillComboBoxMedioPago()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboMedioPago.BoundText)
    Set recData = datcboMedioPago.RowSource
    recData.Requery
    Set recData = Nothing
    datcboMedioPago.BoundText = KeySave
End Sub

