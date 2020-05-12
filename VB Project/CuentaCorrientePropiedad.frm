VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmCuentaCorrientePropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CuentaCorrientePropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   4800
   Begin VB.ComboBox cboCuotas 
      Height          =   330
      Left            =   4020
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   5040
      Width           =   630
   End
   Begin VB.TextBox txtOperacion 
      Height          =   315
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   23
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame fraTipo 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1080
      TabIndex        =   13
      Top             =   4260
      Width           =   2475
      Begin VB.OptionButton optIngreso 
         Caption         =   "Ingreso"
         Height          =   210
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1035
      End
      Begin VB.OptionButton optEgreso 
         Caption         =   "Egreso"
         Height          =   210
         Left            =   1200
         TabIndex        =   15
         Top             =   0
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   3660
      Picture         =   "CuentaCorrientePropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdPersonaClear 
      Height          =   315
      Left            =   3780
      Picture         =   "CuentaCorrientePropiedad.frx":0B74
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Borrar"
      Top             =   3180
      Width           =   315
   End
   Begin VB.CommandButton cmdPersona 
      Height          =   315
      Left            =   3420
      Picture         =   "CuentaCorrientePropiedad.frx":10FE
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Buscar..."
      Top             =   3180
      Width           =   315
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3180
      Width           =   2355
   End
   Begin VB.CommandButton cmdUltimo 
      Caption         =   "&Ultimo"
      Height          =   315
      Left            =   4140
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3180
      Width           =   555
   End
   Begin MSDataListLib.DataCombo datcboGrupo 
      Height          =   330
      Left            =   1080
      TabIndex        =   5
      Top             =   2340
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
   Begin VB.CommandButton cmdCaja 
      Caption         =   "..."
      Height          =   315
      Left            =   4440
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Cajas"
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdGrupo 
      Caption         =   "..."
      Height          =   315
      Left            =   4440
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Grupos"
      Top             =   2340
      Width           =   255
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   1080
      MaxLength       =   255
      TabIndex        =   11
      Top             =   3720
      Width           =   3555
   End
   Begin VB.CommandButton cmdSiguiente 
      Height          =   315
      Left            =   3000
      Picture         =   "CuentaCorrientePropiedad.frx":1688
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   1380
      Width           =   300
   End
   Begin VB.CommandButton cmdAnterior 
      Height          =   315
      Left            =   1080
      Picture         =   "CuentaCorrientePropiedad.frx":1C12
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   1380
      Width           =   300
   End
   Begin VB.CommandButton cmdHoy 
      Height          =   315
      Left            =   3300
      Picture         =   "CuentaCorrientePropiedad.frx":219C
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   1380
      Width           =   315
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   17
      Top             =   4620
      Width           =   1455
   End
   Begin VB.TextBox txtIDMovimiento 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   960
      Width           =   1035
   End
   Begin VB.TextBox txtNotas 
      Height          =   825
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   5820
      Width           =   3555
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   36
      Top             =   780
      Width           =   4515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3420
      TabIndex        =   27
      Top             =   6780
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2100
      TabIndex        =   26
      Top             =   6780
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Top             =   1380
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
      Format          =   100401153
      CurrentDate     =   36950
   End
   Begin MSDataListLib.DataCombo datcboCaja 
      Height          =   330
      Left            =   1080
      TabIndex        =   7
      Top             =   2760
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
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
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
      Format          =   100401155
      UpDown          =   -1  'True
      CurrentDate     =   36494
   End
   Begin MSDataListLib.DataCombo datcboMedioPago 
      Height          =   330
      Left            =   1080
      TabIndex        =   19
      Top             =   5040
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
   Begin VB.Label lblCuotas 
      AutoSize        =   -1  'True
      Caption         =   "Cuotas:"
      Height          =   210
      Left            =   3360
      TabIndex        =   20
      Top             =   5100
      Width           =   555
   End
   Begin VB.Label lblMedioPago 
      AutoSize        =   -1  'True
      Caption         =   "Medio Pago:"
      Height          =   210
      Left            =   120
      TabIndex        =   18
      Top             =   5100
      Width           =   870
   End
   Begin VB.Label lblOperacion 
      AutoSize        =   -1  'True
      Caption         =   "Operación:"
      Height          =   210
      Left            =   120
      TabIndex        =   22
      Top             =   5460
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "&Hora:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1860
      Width           =   390
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   4260
      Width           =   345
   End
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "&Persona:"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   645
   End
   Begin VB.Label lblCaja 
      AutoSize        =   -1  'True
      Caption         =   "Caja:"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   2820
      Width           =   360
   End
   Begin VB.Label lblGrupo 
      AutoSize        =   -1  'True
      Caption         =   "&Grupo:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblDescripcion 
      AutoSize        =   -1  'True
      Caption         =   "&Descripción:"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   3780
      Width           =   900
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "&Fecha:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblIDMovimiento 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   210
      Left            =   120
      TabIndex        =   28
      Top             =   1020
      Width           =   180
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   210
      Left            =   120
      TabIndex        =   24
      Top             =   5880
      Width           =   465
   End
   Begin VB.Label lblImporte 
      AutoSize        =   -1  'True
      Caption         =   "&Importe:"
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   570
   End
   Begin VB.Label lblLegend 
      Caption         =   "Ingrese aquí los Datos del Movimiento de Cta. Cte."
      Height          =   390
      Left            =   780
      TabIndex        =   35
      Top             =   180
      Width           =   2145
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "CuentaCorrientePropiedad.frx":22E6
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmCuentaCorrientePropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCuentaCorriente As CuentaCorriente
Private mNew As Boolean

Private mKeyDecimal As Boolean

Public IsHistory As Boolean

Private mMedioPago As MedioPago

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef CuentaCorriente As CuentaCorriente)
    Dim Persona As Persona
    Dim Notes As Boolean
    Dim All As Boolean
    
    Set mCuentaCorriente = CuentaCorriente
    mNew = (mCuentaCorriente.IDMovimiento = 0)
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    If IsHistory Then
        Notes = False
        All = False
    Else
        If mNew Then
            Notes = True
            All = True
        Else
            If mCuentaCorriente.IDCuentaCorrienteGrupo = pParametro.CuentaCorrienteGrupo_ID_ViajeDebito Or mCuentaCorriente.IDCuentaCorrienteGrupo = pParametro.CuentaCorrienteGrupo_ID_ViajeCredito Or mCuentaCorriente.SaldoAnterior Then
                Notes = False
                All = False
            Else
                If DateDiff("n", Now, mCuentaCorriente.FechaHora) < -pParametro.CuentaCorriente_MovimientoAnterior_Minutos Then
                    'MOVIMIENTO ANTERIOR
                    
                    'NOTAS
                    If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_MODIFY_NOTES_ANTERIOR, False) Then
                        'PERMITIDO NOTAS
                        Notes = True
                    Else
                        'NO PERMITIDO
                        Notes = False
                    End If
                    
                    If mCuentaCorriente.IDCuentaCorrienteGrupo = pParametro.CuentaCorrienteGrupo_ID_Transferencia Then
                        'TRANSFERENCIAS
                        All = pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_MODIFY_TRANSFER_ANTERIOR, False)
                    Else
                        'OTROS MOVIMIENTOS
                        All = pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_MODIFY_ANTERIOR, False)
                    End If
                Else
                    'MOVIMIENTO ACTUAL
                    
                    'NOTAS
                    If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_MODIFY_NOTES_ACTUAL, False) Then
                        'PERMITIDO NOTAS
                        Notes = True
                    Else
                        'NO PERMITIDO
                        Notes = False
                    End If
                    
                    If mCuentaCorriente.IDCuentaCorrienteGrupo = pParametro.CuentaCorrienteGrupo_ID_Transferencia Then
                        'TRANSFERENCIAS
                        All = pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_MODIFY_TRANSFER_ACTUAL, False)
                    Else
                        'OTROS MOVIMIENTOS
                        All = pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_MODIFY_ACTUAL, False)
                    End If
                End If
            End If
        End If
    End If
    EnableControls Notes, All
    
    With mCuentaCorriente
        txtIDMovimiento.Text = IIf(mNew, "", .IDMovimiento)
        dtpFecha.Value = IIf(mNew, Date, CDate(Format(.FechaHora, "Short Date")))
        dtpHora.Value = IIf(mNew, Time, CDate(Format(.FechaHora, "Short Time")))
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboGrupo, "SELECT IDCuentaCorrienteGrupo, Nombre FROM CuentaCorrienteGrupo WHERE Activo = 1 AND ((IDCuentaCorrienteGrupo <> " & pParametro.CuentaCorrienteGrupo_ID_ViajeDebito & " AND IDCuentaCorrienteGrupo <> " & pParametro.CuentaCorrienteGrupo_ID_ViajeCredito & ") OR IDCuentaCorrienteGrupo = " & .IDCuentaCorrienteGrupo & ") ORDER BY Nombre", "IDCuentaCorrienteGrupo", "Nombre", "Grupos", cscpItemOrNone, .IDCuentaCorrienteGrupo) Then
            Unload Me
            Exit Sub
        End If
        
        If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL, False) Then
            If Not CSM_Control_DataCombo.FillFromSQL(datcboCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 OR IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & " ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, .IDCuentaCorrienteCaja) Then
                Unload Me
                Exit Sub
            End If
        Else
            If mNew Then
                If Not CSM_Control_DataCombo.FillFromSQL(datcboCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 OR IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & " OR IDCuentaCorrienteCaja = " & pUsuario.IDCuentaCorrienteCaja & " OR IDCuentaCorrienteCaja = " & pParametro.CuentaCorrienteCaja_ID_ViajeDebito & " ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, pUsuario.IDCuentaCorrienteCaja) Then
                    Unload Me
                    Exit Sub
                End If
            Else
                If Not CSM_Control_DataCombo.FillFromSQL(datcboCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 OR IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & " OR IDCuentaCorrienteCaja = " & pUsuario.IDCuentaCorrienteCaja & " OR IDCuentaCorrienteCaja = " & pParametro.CuentaCorrienteCaja_ID_ViajeDebito & " ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, .IDCuentaCorrienteCaja) Then
                    Unload Me
                    Exit Sub
                End If
            End If
        End If
        
        txtPersona.Tag = .IDPersona
        If mNew Or .IDPersona = 0 Then
            txtPersona.Text = ""
        Else
            Set Persona = New Persona
            Persona.IDPersona = .IDPersona
            If Not Persona.Load() Then
                Set Persona = Nothing
                Unload Me
                Exit Sub
            End If
            txtPersona.Text = Persona.ApellidoNombre
            Set Persona = Nothing
            
            SetLastPersona .IDPersona, txtPersona.Text
        End If
        
        txtDescripcion.Text = .Descripcion
        
        'MEDIO DE PAGO
        If Not CSM_Control_DataCombo.FillFromSQL(datcboMedioPago, "SELECT IDMedioPago, Nombre FROM MedioPago WHERE Activo = 1 OR IDMedioPago = " & .IDMedioPago & " ORDER BY Nombre", "IDMedioPago", "Nombre", "Medios de Pago", cscpItemOrfirst, IIf(.IDMedioPago = 0, pParametro.MedioPago_Predeterminado_ID, .IDMedioPago)) Then
            Unload Me
            Exit Sub
        End If
        cboCuotas.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboCuotas, .Cuotas_Formatted, cscpItemOrfirst)
        txtOperacion.Text = .Operacion
        
        If mNew Then
            optIngreso.Value = False
            optEgreso.Value = False
            txtImporte.Text = ""
        Else
            If .Importe >= 0 Then
                optIngreso.Value = True
                txtImporte.Text = Format(.Importe, "Currency")
            Else
                optEgreso.Value = True
                txtImporte.Text = Format(.Importe * -1, "Currency")
            End If
        End If
        txtNotas.Text = .Notas
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mCuentaCorriente = Nothing
    Set mMedioPago = Nothing
    Set frmCuentaCorrientePropiedad = Nothing
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
    dtpFecha.Value = Date
    dtpFecha.SetFocus
End Sub

Private Sub cmdGrupo_Click()
    If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmCuentaCorrienteGrupo.Show
        On Error Resume Next
        Set frmCuentaCorrienteGrupo.lvwData.SelectedItem = frmCuentaCorrienteGrupo.lvwData.ListItems(KEY_STRINGER & Val(datcboGrupo.BoundText))
        frmCuentaCorrienteGrupo.lvwData.SelectedItem.EnsureVisible
        If frmCuentaCorrienteGrupo.WindowState = vbMinimized Then
            frmCuentaCorrienteGrupo.WindowState = vbNormal
        End If
        frmCuentaCorrienteGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdPersona_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(txtPersona.Tag), UCase(Left(txtPersona.Text, 1)), Me.Name, "", "")
    End If
End Sub

Private Sub cmdPersonaClear_Click()
    txtPersona.Tag = 0
    txtPersona.Text = ""
End Sub

Private Sub cmdUltimo_Click()
    If frmMDI.cboPersona.ListIndex > -1 Then
        PersonaSelected Val(frmMDI.cboPersona.ItemData(frmMDI.cboPersona.ListIndex)), ""
    End If
    cmdPersona.SetFocus
End Sub

Private Sub txtDescripcion_GotFocus()
    CSM_Control_TextBox.SelAllText txtDescripcion
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
        If mMedioPago.Load Then
            lblCuotas.Visible = mMedioPago.UtilizaOperacion
            If mMedioPago.UtilizaOperacion Then
                If mMedioPago.MedioPagoPlan.LoadCuotas Then
                    CSM_Control_ComboBox.FillFromCollection cboCuotas, mMedioPago.MedioPagoPlan.CCuotas, "Cuota", "Cuota", cscpCurrentOrFirst
                End If
            End If
            cboCuotas.Visible = mMedioPago.UtilizaOperacion
            lblOperacion.Visible = mMedioPago.UtilizaOperacion
            txtOperacion.Visible = mMedioPago.UtilizaOperacion
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

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub cmdOK_Click()
    If Val(datcboGrupo.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo.", vbInformation, App.Title
        datcboGrupo.SetFocus
        Exit Sub
    End If
    If Val(datcboCaja.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Caja.", vbInformation, App.Title
        datcboCaja.SetFocus
        Exit Sub
    End If
    If txtDescripcion.Text = "" Then
        MsgBox "Debe ingresar la Descripción.", vbInformation, App.Title
        txtDescripcion.SetFocus
        Exit Sub
    End If
    If optIngreso.Value = False And optEgreso.Value = False Then
        MsgBox "Debe seleccionar el Tipo de Movimiento.", vbInformation, App.Title
        Exit Sub
    End If
    If Not IsNumeric(txtImporte.Text) Then
        MsgBox "El Importe ingresado es incorrecto.", vbInformation, App.Title
        txtImporte.SetFocus
        Exit Sub
    End If
    If Val(datcboMedioPago.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Medio de Pago.", vbInformation, App.Title
        datcboMedioPago.SetFocus
        Exit Sub
    End If
    Select Case DateDiff("n", Now, CDate(dtpFecha.Value & " " & dtpHora.Value))
        Case Is < -pParametro.CuentaCorriente_MovimientoAnterior_Minutos
            'ANTERIOR AL TIEMPO ESPECIFICADO EN LOS PARAMETROS (10 HORAS)
            If mNew Then
                If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_ADD_ANTERIOR, False) Then
                    MsgBox "No está Autorizado a Agregar Movimientos con Fecha anterior.", vbExclamation, App.Title
                    dtpFecha.SetFocus
                    Exit Sub
                End If
            Else
                If dtpFecha.Enabled Then
                    If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_MODIFY_ANTERIOR, False) Then
                        MsgBox "No está Autorizado a Modificar Movimientos con Fecha anterior.", vbExclamation, App.Title
                        dtpFecha.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Case Is > 0
            'POSTERIOR
            MsgBox "No se permite una Fecha posterior a la actual.", vbExclamation, App.Title
            dtpFecha.SetFocus
            Exit Sub
        Case Else
            'ACTUAL
            If mNew Then
                If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_ADD_ACTUAL, False) Then
                    MsgBox "No está Autorizado a Agregar Movimientos con Fecha actual.", vbExclamation, App.Title
                    dtpFecha.SetFocus
                    Exit Sub
                End If
            Else
                If dtpFecha.Enabled Then
                    If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_MODIFY_ACTUAL, False) Then
                        MsgBox "No está Autorizado a Modificar Movimientos con Fecha actual.", vbExclamation, App.Title
                        dtpFecha.SetFocus
                        Exit Sub
                    End If
                End If
            End If
    End Select
    
    mCuentaCorriente.IDCuentaCorrienteGrupo = Val(datcboGrupo.BoundText)
    mCuentaCorriente.IDCuentaCorrienteCaja = Val(datcboCaja.BoundText)
    mCuentaCorriente.IDPersona = Val(txtPersona.Tag)
    mCuentaCorriente.FechaHora = CDate(Format(dtpFecha.Value, "Short Date") & " " & Format(dtpHora.Value, "Short Time"))
    mCuentaCorriente.Descripcion = txtDescripcion.Text
    If optIngreso.Value Then
        mCuentaCorriente.Importe = CCur(txtImporte.Text)
    Else
        mCuentaCorriente.Importe = CCur(txtImporte.Text) * -1
    End If
    mCuentaCorriente.IDMedioPago = Val(datcboMedioPago.BoundText)
    If mMedioPago.UtilizaOperacion Then
        mCuentaCorriente.Cuotas = Val(cboCuotas.Text)
        mCuentaCorriente.Operacion = txtOperacion.Text
    Else
        mCuentaCorriente.Cuotas = 0
        mCuentaCorriente.Operacion = ""
    End If
    mCuentaCorriente.Notas = txtNotas.Text
    If Not mCuentaCorriente.Update Then
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mCuentaCorriente
End Sub

Public Sub FillComboBoxCuentaCorrienteGrupo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboGrupo.BoundText)
    Set recData = datcboGrupo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboGrupo.BoundText = KeySave
End Sub

Public Sub FillComboBoxCuentaCorrienteCaja()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboCaja.BoundText)
    Set recData = datcboCaja.RowSource
    recData.Requery
    Set recData = Nothing
    datcboCaja.BoundText = KeySave
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

Public Sub PersonaSelected(ByVal IDPersona As Long, ByVal Tag As String)
    Dim Persona As Persona

    txtPersona.Tag = IDPersona
    
    Set Persona = New Persona
    Persona.IDPersona = IDPersona
    If Persona.Load() Then
        txtPersona.Text = Persona.ApellidoNombre
    End If
    Set Persona = Nothing
End Sub

Private Sub EnableControls(ByVal Notes As Boolean, ByVal All As Boolean)
    If Not mNew Then
        If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL, False) Then
            If mCuentaCorriente.IDCuentaCorrienteCaja <> pUsuario.IDCuentaCorrienteCaja Then
                All = False
            End If
        End If
        If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_MODIFY_CAJA_OFICINA, False) Then
            If mCuentaCorriente.IDCuentaCorrienteCaja = pParametro.CuentaCorrienteCaja_ID_ViajeDebito Then
                All = False
            End If
        End If
        If All Then
            Notes = All
        End If
    End If
    
    cmdAnterior.Enabled = All
    dtpFecha.Enabled = All
    cmdSiguiente.Enabled = All
    cmdHoy.Enabled = All
    dtpHora.Enabled = All
    datcboGrupo.Enabled = All
    datcboCaja.Enabled = All
    cmdPersona.Enabled = All
    cmdPersonaClear.Enabled = All
    cmdUltimo.Enabled = All
    txtDescripcion.Enabled = All
    optIngreso.Enabled = All
    optEgreso.Enabled = All
    txtImporte.Enabled = All
    
    datcboMedioPago.Enabled = All
    cboCuotas.Enabled = All
    txtOperacion.Enabled = All
    
    txtNotas.Enabled = Notes
    cmdOK.Visible = Notes
    cmdCancel.Caption = IIf(Notes, "Cancelar", "Cerrar")
End Sub
