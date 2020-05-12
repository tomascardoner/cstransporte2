VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPersonaPrepagoPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PersonaPrepagoPropiedad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   5175
   Begin VB.TextBox txtFacturaNumero 
      Height          =   315
      Left            =   1140
      MaxLength       =   20
      TabIndex        =   21
      Top             =   4620
      Width           =   2115
   End
   Begin VB.TextBox txtImporteFinal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   26
      Top             =   5700
      Width           =   1455
   End
   Begin VB.CommandButton cmdCuentaCorrienteCaja 
      Caption         =   "..."
      Height          =   315
      Left            =   4800
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Cajas"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtImporteOriginal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1140
      MaxLength       =   20
      TabIndex        =   13
      Top             =   3300
      Width           =   1455
   End
   Begin VB.TextBox txtOperacion 
      Height          =   315
      Left            =   1140
      MaxLength       =   20
      TabIndex        =   19
      Top             =   4200
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.ComboBox cboCuotas 
      Height          =   330
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3780
      Width           =   630
   End
   Begin VB.CommandButton cmdFechaInicio_Hoy 
      Height          =   315
      Left            =   3660
      Picture         =   "PersonaPrepagoPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   2280
      Width           =   315
   End
   Begin VB.CommandButton cmdFechaInicio_Anterior 
      Height          =   315
      Left            =   1440
      Picture         =   "PersonaPrepagoPropiedad.frx":0694
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   2280
      Width           =   300
   End
   Begin VB.CommandButton cmdFechaInicio_Siguiente 
      Height          =   315
      Left            =   3360
      Picture         =   "PersonaPrepagoPropiedad.frx":0C1E
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   2280
      Width           =   300
   End
   Begin VB.TextBox txtFechaFin 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   4080
      Picture         =   "PersonaPrepagoPropiedad.frx":11A8
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   960
      Width           =   3555
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   27
      Top             =   6300
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   28
      Top             =   6300
      Width           =   1215
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   4935
   End
   Begin MSDataListLib.DataCombo datcboRutaGrupo 
      Height          =   330
      Left            =   1500
      TabIndex        =   1
      Top             =   1380
      Width           =   3555
      _ExtentX        =   6271
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
   Begin MSDataListLib.DataCombo datcboListaPrecio 
      Height          =   330
      Left            =   1500
      TabIndex        =   3
      Top             =   1800
      Width           =   3555
      _ExtentX        =   6271
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
   Begin MSComCtl2.DTPicker dtpFechaInicio 
      Height          =   315
      Left            =   1740
      TabIndex        =   6
      Top             =   2280
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
      CurrentDate     =   41760
      MaxDate         =   73050
      MinDate         =   41640
   End
   Begin MSDataListLib.DataCombo datcboMedioPago 
      Height          =   330
      Left            =   1140
      TabIndex        =   15
      Top             =   3780
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
   Begin MSDataListLib.DataCombo datcboCuentaCorrienteCaja 
      Height          =   330
      Left            =   1140
      TabIndex        =   23
      Top             =   5160
      Width           =   3675
      _ExtentX        =   6482
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
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Width           =   825
   End
   Begin VB.Label lblImporte 
      AutoSize        =   -1  'True
      Caption         =   "Importe final:"
      Height          =   210
      Left            =   120
      TabIndex        =   25
      Top             =   5760
      Width           =   915
   End
   Begin VB.Label lblCuentaCorrienteCaja 
      AutoSize        =   -1  'True
      Caption         =   "Caja:"
      Height          =   210
      Left            =   120
      TabIndex        =   22
      Top             =   5220
      Width           =   360
   End
   Begin VB.Label lblImporteOriginal 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   570
   End
   Begin VB.Label lblOperacion 
      AutoSize        =   -1  'True
      Caption         =   "Operación:"
      Height          =   210
      Left            =   120
      TabIndex        =   18
      Top             =   4260
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblMedioPago 
      AutoSize        =   -1  'True
      Caption         =   "Medio Pago:"
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   870
   End
   Begin VB.Label lblCuotas 
      AutoSize        =   -1  'True
      Caption         =   "Cuotas:"
      Height          =   210
      Left            =   3420
      TabIndex        =   16
      Top             =   3840
      Width           =   555
   End
   Begin VB.Label lblListaPrecio 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Precios:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1860
      Width           =   1200
   End
   Begin VB.Label lblFechaInicio 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de inicio:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   2340
      Width           =   1125
   End
   Begin VB.Label lblFechaFin 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de fin:"
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   2820
      Width           =   945
   End
   Begin VB.Label lblFechaFin_Inclusive 
      AutoSize        =   -1  'True
      Caption         =   "inclusive"
      Height          =   210
      Left            =   3180
      TabIndex        =   11
      Top             =   2820
      Width           =   630
   End
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "Pasajero:"
      Height          =   210
      Left            =   120
      TabIndex        =   32
      Top             =   1020
      Width           =   675
   End
   Begin VB.Label lblRutaGrupo 
      AutoSize        =   -1  'True
      Caption         =   "Grupos de Rutas:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Image imgIcon2 
      Height          =   480
      Left            =   480
      Picture         =   "PersonaPrepagoPropiedad.frx":17D2
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "PersonaPrepagoPropiedad.frx":209C
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Datos del Prepago de la Persona"
      Height          =   210
      Left            =   1140
      TabIndex        =   30
      Top             =   240
      Width           =   2355
   End
End
Attribute VB_Name = "frmPersonaPrepagoPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPersonaPrepago As PersonaPrepago
Private mMedioPago As MedioPago

Private mListaPrecio_PrepagoVencimiento As String

Private mKeyDecimal As Boolean

Private Sub cboCuotas_Click()
    Call CalcularImporteFinal
End Sub

Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mPersonaPrepago
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef PersonaPrepago As PersonaPrepago)
    Set mPersonaPrepago = PersonaPrepago
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    txtPersona.Text = mPersonaPrepago.Persona.ApellidoNombre
    
    With mPersonaPrepago
        If Not CSM_Control_DataCombo.FillFromSQL(datcboRutaGrupo, "SELECT IDRutaGrupo, Nombre FROM RutaGrupo WHERE Activo = 1 OR IDRutaGrupo = " & .IDRutaGrupo & " ORDER BY Nombre", "IDRutaGrupo", "Nombre", "Grupos de Rutas") Then
            Unload Me
            Exit Sub
        End If
        datcboRutaGrupo.BoundText = .IDRutaGrupo
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboListaPrecio, "SELECT IDListaPrecio, Nombre FROM ListaPrecio WHERE PrepagoEs = 1 AND (Activo = 1 OR IDListaPrecio = " & .IDListaPrecio & IIf(pCPermiso.ListaPrecioWhere <> "", " AND " & Replace(pCPermiso.ListaPrecioWhere, "%TABLENAME%", "ListaPrecio"), "") & ") ORDER BY Nombre", "IDListaPrecio", "Nombre", "Listas de Precios") Then
            Unload Me
            Exit Sub
        End If
        datcboListaPrecio.BoundText = .IDListaPrecio
        
        If Val(datcboListaPrecio.BoundText) = 0 Then
            dtpFechaInicio.Value = Date
        Else
            dtpFechaInicio.Value = .FechaInicio
            txtFechaFin.Text = .FechaFin_Formatted
        End If
        
        txtImporteOriginal.Text = .ImporteOriginal_Formatted
        
        'MEDIO DE PAGO
        If Not CSM_Control_DataCombo.FillFromSQL(datcboMedioPago, "SELECT IDMedioPago, Nombre FROM MedioPago WHERE Activo = 1 OR IDMedioPago = " & .IDMedioPago & " ORDER BY Nombre", "IDMedioPago", "Nombre", "Medios de Pago", cscpItemOrfirst, IIf(.IDMedioPago = 0, pParametro.MedioPago_Predeterminado_ID, .IDMedioPago)) Then
            Unload Me
            Exit Sub
        End If
        cboCuotas.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboCuotas, .Cuotas_Formatted, cscpItemOrfirst)
        txtOperacion.Text = .Operacion
        txtFacturaNumero.Text = .FacturaNumero
        
        'CAJA DE CUENTA CORRIENTE
        If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL, False) Then
            If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 OR IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & " ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, .IDCuentaCorrienteCaja) Then
                Unload Me
                Exit Sub
            End If
        Else
            If .IsNew Then
                If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 OR IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & " OR IDCuentaCorrienteCaja = " & pUsuario.IDCuentaCorrienteCaja & " OR IDCuentaCorrienteCaja = " & pParametro.CuentaCorrienteCaja_ID_ViajeDebito & " ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, pUsuario.IDCuentaCorrienteCaja) Then
                    Unload Me
                    Exit Sub
                End If
            Else
                If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 OR IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & " OR IDCuentaCorrienteCaja = " & pUsuario.IDCuentaCorrienteCaja & " OR IDCuentaCorrienteCaja = " & pParametro.CuentaCorrienteCaja_ID_ViajeDebito & " ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, .IDCuentaCorrienteCaja) Then
                    Unload Me
                    Exit Sub
                End If
            End If
        End If
        
        txtImporteFinal.Text = .Importe_Formatted
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub datcboListaPrecio_Change()
    Dim ListaPrecio As ListaPrecio
    
    If Val(datcboListaPrecio.BoundText) > 0 Then
        Set ListaPrecio = New ListaPrecio
        ListaPrecio.IDListaPrecio = Val(datcboListaPrecio.BoundText)
        If ListaPrecio.Load() Then
            mListaPrecio_PrepagoVencimiento = ListaPrecio.PrepagoVencimiento
            Call dtpFechaInicio_Change
        End If
        Set ListaPrecio = Nothing
    End If
End Sub

Private Sub cmdFechaInicio_Anterior_Click()
    dtpFechaInicio.Value = DateAdd("d", -1, dtpFechaInicio.Value)
    dtpFechaInicio.SetFocus
    Call dtpFechaInicio_Change
End Sub

Private Sub dtpFechaInicio_Change()
    Select Case mListaPrecio_PrepagoVencimiento
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA1
            txtFechaFin.Text = DateAdd("d", 6, dtpFechaInicio.Value)
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA2
            txtFechaFin.Text = DateAdd("d", 13, dtpFechaInicio.Value)
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA15
            txtFechaFin.Text = DateAdd("d", 14, dtpFechaInicio.Value)
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA30
            txtFechaFin.Text = DateAdd("d", 29, dtpFechaInicio.Value)
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA45
            txtFechaFin.Text = DateAdd("d", 44, dtpFechaInicio.Value)
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES1
            txtFechaFin.Text = DateAdd("d", -1, DateAdd("m", 1, dtpFechaInicio.Value))
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES2
            txtFechaFin.Text = DateAdd("d", -1, DateAdd("m", 2, dtpFechaInicio.Value))
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES3
            txtFechaFin.Text = DateAdd("d", -1, DateAdd("m", 3, dtpFechaInicio.Value))
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES6
            txtFechaFin.Text = DateAdd("d", -1, DateAdd("m", 6, dtpFechaInicio.Value))
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_ANIO1
            txtFechaFin.Text = DateAdd("d", -1, DateAdd("yyyy", 1, dtpFechaInicio.Value))
        Case Else
            txtFechaFin.Text = ""
    End Select
End Sub

Private Sub cmdFechaInicio_Siguiente_Click()
    dtpFechaInicio.Value = DateAdd("d", 1, dtpFechaInicio.Value)
    dtpFechaInicio.SetFocus
    Call dtpFechaInicio_Change
End Sub

Private Sub cmdFechaInicio_Hoy_Click()
    Dim OldValue As Date
    
    OldValue = dtpFechaInicio.Value
    dtpFechaInicio.Value = Date
    dtpFechaInicio.SetFocus
    Call dtpFechaInicio_Change
End Sub

Private Sub txtImporteOriginal_GotFocus()
    CSM_Control_TextBox.SelAllText txtImporteOriginal
End Sub

Private Sub txtImporteOriginal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtImporteOriginal_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtImporteOriginal.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtImporteOriginal.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtImporteOriginal_Change()
    Call CalcularImporteFinal
End Sub

Private Sub txtImporteOriginal_LostFocus()
    If Not IsNumeric(txtImporteOriginal.Text) Then
        txtImporteOriginal.Text = Val(txtImporteOriginal.Text)
    End If
    txtImporteOriginal.Text = Format(CCur(txtImporteOriginal.Text), "Currency")
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

Private Sub txtFacturaNumero_GotFocus()
    CSM_Control_TextBox.SelAllText txtFacturaNumero
End Sub

Private Sub cmdOK_Click()
    Dim ListaPrecio As ListaPrecio
    Dim MovimientoCredito As CuentaCorriente
    Dim MovimientoDebito As CuentaCorriente
    
    If datcboRutaGrupo.BoundText = "" Then
        MsgBox "Debe seleccionar la Ruta.", vbInformation, App.Title
        datcboRutaGrupo.SetFocus
        Exit Sub
    End If
    If Val(datcboListaPrecio.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Lista de Precios.", vbInformation, App.Title
        datcboListaPrecio.SetFocus
        Exit Sub
    End If
    If mPersonaPrepago.IsNew Then
        If DateDiff("d", Date, dtpFechaInicio.Value) < 0 Then
            If MsgBox("La fecha especificada es anterior a la fecha actual." & vbCr & vbCr & "¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                dtpFechaInicio.SetFocus
                Exit Sub
            End If
        End If
        If DateDiff("d", Date, dtpFechaInicio.Value) > 30 Then
            If MsgBox("La fecha especificada es posterior a 30 días." & vbCr & vbCr & "¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                dtpFechaInicio.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    If Not IsNumeric(txtImporteOriginal.Text) Then
        MsgBox "El Importe ingresado es incorrecto.", vbInformation, App.Title
        txtImporteOriginal.SetFocus
        Exit Sub
    End If
    If CCur(txtImporteOriginal.Text) <= 0 Then
        MsgBox "El Importe debe ser mayor a cero.", vbInformation, App.Title
        txtImporteOriginal.SetFocus
        Exit Sub
    End If
    If Val(datcboMedioPago.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Medio de Pago.", vbInformation, App.Title
        datcboMedioPago.SetFocus
        Exit Sub
    End If
    If pParametro.ViajeDetalle_MedioPago_UtilizaOperacion_ObligaFacturaNumero Then
        If mMedioPago.UtilizaOperacion And Len(Trim(txtFacturaNumero.Text)) = 0 Then
            MsgBox "Debe especificar el Nº de Factura.", vbInformation, App.Title
            txtFacturaNumero.SetFocus
            Exit Sub
        End If
    End If
    If Val(datcboCuentaCorrienteCaja.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Caja.", vbInformation, App.Title
        datcboCuentaCorrienteCaja.SetFocus
        Exit Sub
    End If
    
    '///////////////////////////////////////////////////
    'GENERO LOS DOS MOVIMIENTOS DE CUENTA CORRIENTE
    Set ListaPrecio = New ListaPrecio
    ListaPrecio.IDListaPrecio = Val(datcboListaPrecio.BoundText)
    If Not ListaPrecio.Load Then
        Set ListaPrecio = Nothing
        Exit Sub
    End If
    
    'CREDITO
    Set MovimientoCredito = New CuentaCorriente
    With MovimientoCredito
        .RefreshListSkip = True
        If mPersonaPrepago.IDMovimiento_Credito <> 0 Then
            .NoMatchRaiseError = False
            .IDMovimiento = mPersonaPrepago.IDMovimiento_Credito
            If Not .Load() Then
                Set ListaPrecio = Nothing
                Set MovimientoCredito = Nothing
                Exit Sub
            End If
        End If
        .IDCuentaCorrienteGrupo = ListaPrecio.IDCuentaCorrienteGrupo_Credito
        .IDCuentaCorrienteCaja = Val(datcboCuentaCorrienteCaja.BoundText)
        .IDPersona = mPersonaPrepago.IDPersona
        .IDPersonaOrigen = 0
        If .IsNew Or .NoMatch Then
            .FechaHora = Now
        End If
        .Descripcion = "Prepago - Rutas: " & datcboRutaGrupo.Text & " - Fecha de Inicio: " & dtpFechaInicio.Value
        .Importe = CCur(txtImporteFinal.Text)
        .IDMedioPago = Val(datcboMedioPago.BoundText)
        If mMedioPago.UtilizaOperacion Then
            .Cuotas = cboCuotas.Text
            .Operacion = txtOperacion.Text
        Else
            .Cuotas = 1
            .Operacion = ""
        End If
        If Not .Update Then
            Set ListaPrecio = Nothing
            Set MovimientoCredito = Nothing
            Exit Sub
        End If
    End With
    
    'CREDITO
    Set MovimientoDebito = New CuentaCorriente
    With MovimientoDebito
        If mPersonaPrepago.IDMovimiento_Debito <> 0 Then
            .NoMatchRaiseError = False
            .IDMovimiento = mPersonaPrepago.IDMovimiento_Debito
            If Not .Load() Then
                Set ListaPrecio = Nothing
                Set MovimientoDebito = Nothing
                Exit Sub
            End If
        End If
        .IDCuentaCorrienteGrupo = ListaPrecio.IDCuentaCorrienteGrupo_Debito
        .IDCuentaCorrienteCaja = pParametro.CuentaCorrienteCaja_ID_ViajeDebito
        .IDPersona = mPersonaPrepago.IDPersona
        .IDPersonaOrigen = 0
        If .IsNew Or .NoMatch Then
            .FechaHora = Now
        End If
        .Descripcion = "Prepago - Rutas: " & datcboRutaGrupo.Text & " - Fecha de Inicio: " & dtpFechaInicio.Value
        .Importe = CCur(txtImporteFinal.Text) * -1
        .IDMedioPago = pParametro.MedioPago_Predeterminado_ID
        .Cuotas = 1
        .Operacion = ""
        If Not .Update Then
            Set ListaPrecio = Nothing
            Set MovimientoDebito = Nothing
            Exit Sub
        End If
    End With
        
    'PREPAGO
    With mPersonaPrepago
        .IDRutaGrupo = Val(datcboRutaGrupo.BoundText)
        .IDListaPrecio = Val(datcboListaPrecio.BoundText)
        If Val(datcboListaPrecio.BoundText) > 0 Then
            .IDListaPrecio = Val(datcboListaPrecio.BoundText)
            .FechaInicio = CDate(dtpFechaInicio.Value)
            .FechaFin = CDate(txtFechaFin.Text & " 23:59:00")
        End If
        .ImporteOriginal = CCur(txtImporteOriginal.Text)
        .Importe = CCur(txtImporteFinal.Text)
        .IDMedioPago = Val(datcboMedioPago.BoundText)
        If mMedioPago.UtilizaOperacion Then
            .Cuotas = Val(cboCuotas.Text)
            .Operacion = txtOperacion.Text
        Else
            .Cuotas = 0
            .Operacion = ""
        End If
        .FacturaNumero = txtFacturaNumero.Text
        .IDCuentaCorrienteCaja = Val(datcboCuentaCorrienteCaja.BoundText)
        .IDMovimiento_Credito = MovimientoCredito.IDMovimiento
        .IDMovimiento_Debito = MovimientoDebito.IDMovimiento
        
        If Not .Update() Then
            MovimientoCredito.Delete
            MovimientoDebito.Delete
            Set MovimientoCredito = Nothing
            Set MovimientoDebito = Nothing
            Exit Sub
        End If
    
        Set MovimientoCredito = Nothing
        Set MovimientoDebito = Nothing
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mPersonaPrepago = Nothing
    Set mMedioPago = Nothing
    Set frmPersonaPrepagoPropiedad = Nothing
End Sub

Public Sub FillComboBoxRuta()
    Dim KeySave As String
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboRutaGrupo.BoundText)
    Set recData = datcboRutaGrupo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboRutaGrupo.BoundText = KeySave
End Sub

Public Sub FillComboBoxListaPrecio()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboListaPrecio.BoundText)
    Set recData = datcboListaPrecio.RowSource
    recData.Requery
    Set recData = Nothing
    datcboListaPrecio.BoundText = KeySave
End Sub

Private Sub CalcularImporteFinal()
    If IsNumeric(txtImporteOriginal.Text) And Val(datcboMedioPago.BoundText) <> 0 Then
        If mMedioPago.UtilizaOperacion Then
            If cboCuotas.ListIndex > -1 Then
                txtImporteFinal.Text = Format(CCur(txtImporteOriginal.Text) * mMedioPago.MedioPagoPlan.CCuotas(KEY_STRINGER & cboCuotas.Text).CoeficientePrepago, "Currency")
            Else
                txtImporteFinal.Text = Format(CCur(txtImporteOriginal.Text), "Currency")
            End If
        Else
            txtImporteFinal.Text = Format(CCur(txtImporteOriginal.Text), "Currency")
        End If
    Else
        txtImporteFinal.Text = Format(0, "Currency")
    End If
End Sub
