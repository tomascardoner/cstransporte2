VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmViajeDetallePropiedadPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle del Pago"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajeDetallePropiedadPago.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboCuotas 
      Height          =   330
      Left            =   4380
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   630
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2460
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3780
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCuentaCorrienteCaja 
      Caption         =   "..."
      Height          =   315
      Left            =   4740
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Cajas"
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtOperacion 
      Height          =   315
      Left            =   1380
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1860
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1380
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1020
      Width           =   1395
   End
   Begin VB.CommandButton cmdHoy 
      Height          =   315
      Left            =   3600
      Picture         =   "ViajeDetallePropiedadPago.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   180
      Width           =   315
   End
   Begin VB.CommandButton cmdAnterior 
      Height          =   315
      Left            =   1380
      Picture         =   "ViajeDetallePropiedadPago.frx":0156
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   180
      Width           =   300
   End
   Begin VB.CommandButton cmdSiguiente 
      Height          =   315
      Left            =   3300
      Picture         =   "ViajeDetallePropiedadPago.frx":06E0
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   180
      Width           =   300
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   1680
      TabIndex        =   15
      Top             =   180
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
      Format          =   105512961
      CurrentDate     =   36950
   End
   Begin MSDataListLib.DataCombo datcboMedioPago 
      Height          =   330
      Left            =   1380
      TabIndex        =   3
      Top             =   1440
      Width           =   2145
      _ExtentX        =   3784
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
   Begin MSDataListLib.DataCombo datcboCuentaCorrienteCaja 
      Height          =   330
      Left            =   1380
      TabIndex        =   9
      Top             =   2280
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
   Begin MSComCtl2.DTPicker dtpHora 
      Height          =   315
      Left            =   1380
      TabIndex        =   19
      Top             =   600
      Width           =   915
      _ExtentX        =   1614
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
      CustomFormat    =   "HH:mm"
      Format          =   105512963
      UpDown          =   -1  'True
      CurrentDate     =   36494
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "&Hora:"
      Height          =   210
      Left            =   120
      TabIndex        =   18
      Top             =   660
      Width           =   390
   End
   Begin VB.Label lblCuotas 
      AutoSize        =   -1  'True
      Caption         =   "Cuotas:"
      Height          =   210
      Left            =   3720
      TabIndex        =   4
      Top             =   1500
      Width           =   555
   End
   Begin VB.Label lblCuentaCorrienteCaja 
      AutoSize        =   -1  'True
      Caption         =   "Caja:"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   2340
      Width           =   360
   End
   Begin VB.Label lblMedioPago 
      AutoSize        =   -1  'True
      Caption         =   "Medio de Pago:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label lblOperacion 
      AutoSize        =   -1  'True
      Caption         =   "Operación:"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblImporte 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   570
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "&Fecha:"
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmViajeDetallePropiedadPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mKeyDecimal As Boolean

Private mPago As CuentaCorriente

Private mMedioPago As MedioPago

Public Sub LoadData(ByRef ViajeDetalle As ViajeDetalle, ByRef Pago As CuentaCorriente)
    Dim CuentaCorrienteCaja As CuentaCorrienteCaja
    
    Set mPago = Pago
    
    With mPago
        If .IsNew Then
            dtpFecha.Value = Date
            dtpHora.Value = Time
        Else
            dtpFecha.Value = .FechaHora_FormattedAsDate
            dtpHora.Value = .FechaHora_FormattedAsTime
        End If
        txtImporte.Text = .Importe_Formatted
        txtImporte_LostFocus
        
        'MEDIO DE PAGO
        If Not CSM_Control_DataCombo.FillFromSQL(datcboMedioPago, "SELECT IDMedioPago, Nombre FROM MedioPago WHERE Activo = 1 OR IDMedioPago = " & .IDMedioPago & " ORDER BY Nombre", "IDMedioPago", "Nombre", "Medios de Pago", cscpItemOrfirst, IIf(.IDMedioPago = 0, pParametro.MedioPago_Predeterminado_ID, .IDMedioPago)) Then
            Unload Me
            Exit Sub
        End If
        cboCuotas.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboCuotas, .Cuotas_Formatted, cscpItemOrfirst)
        txtOperacion.Text = .Operacion
        
        If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL, False) Then
            If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 OR IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & " ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, .IDCuentaCorrienteCaja) Then
                Unload Me
                Exit Sub
            End If
        Else
            If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT CuentaCorrienteCaja.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre FROM CuentaCorrienteCaja LEFT JOIN Persona ON CuentaCorrienteCaja.IDPersona = Persona.IDPersona WHERE (CuentaCorrienteCaja.Activo = 1 OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & ") AND (CuentaCorrienteCaja.MostrarSiempre = 1 OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & pUsuario.IDCuentaCorrienteCaja & " OR Persona.IDPersona = " & ViajeDetalle.Viaje.IDConductor & " OR Persona.IDPersona = " & ViajeDetalle.Viaje.IDConductor2 & " OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & ") ORDER BY CuentaCorrienteCaja.Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, IIf(.IDCuentaCorrienteCaja = 0, pUsuario.IDCuentaCorrienteCaja, .IDCuentaCorrienteCaja)) Then
                Unload Me
                Exit Sub
            End If
        End If
        
        If Not ViajeDetalle.IsNew Then
            'ESTOY EDITANDO UN PAGO EXISTENTE
            If ViajeDetalle.Viaje.IDConductor <> 0 Then
                Set CuentaCorrienteCaja = New CuentaCorrienteCaja
                CuentaCorrienteCaja.IDPersona = ViajeDetalle.Viaje.IDConductor
                Call CuentaCorrienteCaja.LoadByPersona
                If .IDCuentaCorrienteCaja <> pUsuario.IDCuentaCorrienteCaja And .IDCuentaCorrienteCaja <> CuentaCorrienteCaja.IDCuentaCorrienteCaja Then
                    If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_IMPORTE_OTRO_USUARIO, False) Then
                        txtImporte.Enabled = False
                        datcboMedioPago.Enabled = False
                        cboCuotas.Enabled = False
                        txtOperacion.Enabled = False
                    End If
                    If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_CAJA_OTRO_USUARIO, False) Then
                        datcboCuentaCorrienteCaja.Enabled = False
                    End If
                End If
                Set CuentaCorrienteCaja = Nothing
            Else
                If .IDCuentaCorrienteCaja <> pUsuario.IDCuentaCorrienteCaja Then
                    If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_IMPORTE_OTRO_USUARIO, False) Then
                        txtImporte.Enabled = False
                        datcboMedioPago.Enabled = False
                        cboCuotas.Enabled = False
                        txtOperacion.Enabled = False
                    End If
                    If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_CAJA_OTRO_USUARIO, False) Then
                        datcboCuentaCorrienteCaja.Enabled = False
                    End If
                End If
            End If
        End If
        
        If ViajeDetalle.IsNew Then
            If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_ADD_FECHAHORARUTA, False) Then
                cmdAnterior.Visible = False
                dtpFecha.Enabled = False
                cmdSiguiente.Visible = False
                cmdHoy.Visible = False
            Else
                cmdAnterior.Visible = True
                dtpFecha.Enabled = True
                cmdSiguiente.Visible = True
                cmdHoy.Visible = True
            End If
        Else
            If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_FECHAHORARUTA, False) Then
                cmdAnterior.Visible = False
                dtpFecha.Enabled = False
                cmdSiguiente.Visible = False
                cmdHoy.Visible = False
            Else
                cmdAnterior.Visible = True
                dtpFecha.Enabled = True
                cmdSiguiente.Visible = True
                cmdHoy.Visible = True
            End If
        End If
        
    End With
End Sub

Private Sub cmdAnterior_Click()
    dtpFecha.Value = DateAdd("d", -1, dtpFecha.Value)
    dtpFecha.SetFocus
End Sub

Private Sub cmdSiguiente_Click()
    dtpFecha.Value = DateAdd("d", 1, dtpFecha.Value)
    dtpFecha.SetFocus
End Sub

Private Sub cmdHoy_Click()
    Dim OldValue As Date
    
    OldValue = dtpFecha.Value
    dtpFecha.Value = Date
    dtpFecha.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mMedioPago = Nothing
End Sub

Private Sub txtImporte_GotFocus()
    CSM_Control_TextBox.SelAllText txtImporte
End Sub

Private Sub txtImporte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtImporte.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtImporte.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtImporte_LostFocus()
    If Not IsNumeric(txtImporte.Text) Then
        txtImporte.Text = Val(txtImporte.Text)
    End If
    txtImporte.Text = Format(CCur(txtImporte.Text), "Currency")
End Sub

Private Sub datcboMedioPago_Change()
    If Val(datcboMedioPago.BoundText) > 0 Then
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
        
            lblCuentaCorrienteCaja.Visible = (mMedioPago.IDCuentaCorrienteCaja = 0)
            datcboCuentaCorrienteCaja.Visible = (mMedioPago.IDCuentaCorrienteCaja = 0)
            cmdCuentaCorrienteCaja.Visible = (mMedioPago.IDCuentaCorrienteCaja = 0)
        End If
    Else
        lblCuotas.Visible = False
        cboCuotas.Visible = False
        lblOperacion.Visible = False
        txtOperacion.Visible = False
    End If
End Sub

Private Sub txtOperacion_GotFocus()
    CSM_Control_TextBox.SelAllText txtOperacion
End Sub

Private Sub txtOperacion_LostFocus()
    txtOperacion.Text = CleanInvalidSpaces(txtOperacion.Text)
End Sub

Private Sub cmdOK_Click()
    If CCur(txtImporte.Text) <= 0 Then
        MsgBox "El Importe debe ser mayor a cero.", vbInformation, App.Title
        txtImporte.SetFocus
        Exit Sub
    End If
    If Val(datcboMedioPago.BoundText) = 0 Then
        MsgBox "Debe especificar el Medio de Pago.", vbInformation, App.Title
        datcboMedioPago.SetFocus
        Exit Sub
    End If
    If Val(datcboCuentaCorrienteCaja.BoundText) = 0 Then
        MsgBox "Debe especificar la Caja.", vbInformation, App.Title
        datcboCuentaCorrienteCaja.SetFocus
        Exit Sub
    End If
    
    mPago.FechaHora = CDate(Format(dtpFecha.Value, "Short Date") & " " & Format(dtpHora.Value, "Short Time"))
    mPago.Importe = CCur(txtImporte.Text)
    mPago.IDMedioPago = Val(datcboMedioPago.BoundText)
    mPago.Cuotas = Val(cboCuotas.Text)
    mPago.Operacion = txtOperacion.Text
    mPago.IDCuentaCorrienteCaja = Val(datcboCuentaCorrienteCaja.BoundText)
    mPago.IDCuentaCorrienteGrupo = pParametro.CuentaCorrienteGrupo_ID_ViajeCredito
    
    Me.Tag = "OK"
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = "CANCEL"
    Me.Hide
End Sub
