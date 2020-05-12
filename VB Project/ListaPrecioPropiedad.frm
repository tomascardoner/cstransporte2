VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmListaPrecioPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ListaPrecioPropiedad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   6765
   Begin VB.Frame fraPrepago 
      Height          =   2355
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   6495
      Begin VB.CommandButton cmdCuentaCorrienteGrupo_Debito 
         Caption         =   "..."
         Height          =   315
         Left            =   6000
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Grupos de Cuenta Corriente"
         Top             =   1860
         Width           =   255
      End
      Begin VB.CommandButton cmdCuentaCorrienteGrupo_Credito 
         Caption         =   "..."
         Height          =   315
         Left            =   6000
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Grupos de Cuenta Corriente"
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkPrepago 
         Caption         =   "Prepago"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtPrepagoReservasCantidad 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   11
         Tag             =   "NUMERIC|"
         Top             =   900
         Width           =   735
      End
      Begin VB.ComboBox cboPrepagoVencimiento 
         Height          =   330
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   540
         Width           =   2250
      End
      Begin MSDataListLib.DataCombo datcboCuentaCorrienteGrupo_Credito 
         Height          =   330
         Left            =   2640
         TabIndex        =   14
         Top             =   1440
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
      Begin MSDataListLib.DataCombo datcboCuentaCorrienteGrupo_Debito 
         Height          =   330
         Left            =   2640
         TabIndex        =   17
         Top             =   1860
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
      Begin VB.Label lblCuentaCorrienteGrupo_Debito 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Cta.Cte. para Débito:"
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   2205
      End
      Begin VB.Label lblCuentaCorrienteGrupo_Credito 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Cta.Cte. para Crédito:"
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   1500
         Width           =   2265
      End
      Begin VB.Label lblPrepagoReservasCantidad_Ilimitada 
         AutoSize        =   -1  'True
         Caption         =   "(0 = ilimitadas)"
         Height          =   210
         Left            =   2880
         TabIndex        =   12
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label lblPrepagoReservasCantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad de Reservas:"
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1650
      End
      Begin VB.Label lblPrepagoVencimiento 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento:"
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   645
      Left            =   1140
      MaxLength       =   8000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1980
      Width           =   5475
   End
   Begin VB.TextBox txtLeyenda 
      Height          =   315
      Left            =   1140
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1500
      Width           =   5475
   End
   Begin VB.TextBox txtNotas 
      Height          =   705
      Left            =   1140
      MaxLength       =   8000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   5340
      Width           =   5475
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo"
      Height          =   210
      Left            =   120
      TabIndex        =   21
      Top             =   6180
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   1140
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1020
      Width           =   5475
   End
   Begin VB.Frame fraLine 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   120
      TabIndex        =   25
      Top             =   780
      Width           =   6495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   23
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   22
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lblDescripcion 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label lblLeyenda 
      AutoSize        =   -1  'True
      Caption         =   "Leyenda:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   210
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   465
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos de la Lista de Precios"
      Height          =   210
      Left            =   780
      TabIndex        =   24
      Top             =   300
      Width           =   3195
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "ListaPrecioPropiedad.frx":054A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmListaPrecioPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mListaPrecio As ListaPrecio

Private mKeyDecimal As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef ListaPrecio As ListaPrecio)
    Set mListaPrecio = ListaPrecio

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mListaPrecio
        txtNombre.Text = .Nombre
        txtLeyenda.Text = .Leyenda
        txtDescripcion.Text = .Descripcion
        chkPrepago.Value = IIf(.PrepagoEs, vbChecked, vbUnchecked)
        chkPrepago_Click
        cboPrepagoVencimiento.ListIndex = .PrepagoVencimiento_ListIndex
        txtPrepagoReservasCantidad.Text = .PrepagoReservasCantidad
        
        Call CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteGrupo_Credito, "SELECT IDCuentaCorrienteGrupo, Nombre FROM CuentaCorrienteGrupo WHERE Activo = 1 AND ((IDCuentaCorrienteGrupo <> " & pParametro.CuentaCorrienteGrupo_ID_ViajeDebito & " AND IDCuentaCorrienteGrupo <> " & pParametro.CuentaCorrienteGrupo_ID_ViajeCredito & ") OR IDCuentaCorrienteGrupo = " & .IDCuentaCorrienteGrupo_Credito & ") ORDER BY Nombre", "IDCuentaCorrienteGrupo", "Nombre", "Grupos de cuenta Corriente", cscpItemOrNone, .IDCuentaCorrienteGrupo_Credito)
        Call CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteGrupo_Debito, "SELECT IDCuentaCorrienteGrupo, Nombre FROM CuentaCorrienteGrupo WHERE Activo = 1 AND ((IDCuentaCorrienteGrupo <> " & pParametro.CuentaCorrienteGrupo_ID_ViajeDebito & " AND IDCuentaCorrienteGrupo <> " & pParametro.CuentaCorrienteGrupo_ID_ViajeCredito & ") OR IDCuentaCorrienteGrupo = " & .IDCuentaCorrienteGrupo_Debito & ") ORDER BY Nombre", "IDCuentaCorrienteGrupo", "Nombre", "Grupos de cuenta Corriente", cscpItemOrNone, .IDCuentaCorrienteGrupo_Debito)
        
        txtNotas.Text = .Notas
        chkActivo.Value = IIf(.Activo, vbChecked, vbUnchecked)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub Form_Load()
    cboPrepagoVencimiento.AddItem LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA1_NOMBRE
    cboPrepagoVencimiento.AddItem LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA2_NOMBRE
    cboPrepagoVencimiento.AddItem LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA15_NOMBRE
    cboPrepagoVencimiento.AddItem LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA30_NOMBRE
    cboPrepagoVencimiento.AddItem LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA45_NOMBRE
    cboPrepagoVencimiento.AddItem LISTAPRECIO_PREPAGO_VENCIMIENTO_MES1_NOMBRE
    cboPrepagoVencimiento.AddItem LISTAPRECIO_PREPAGO_VENCIMIENTO_MES2_NOMBRE
    cboPrepagoVencimiento.AddItem LISTAPRECIO_PREPAGO_VENCIMIENTO_MES3_NOMBRE
    cboPrepagoVencimiento.AddItem LISTAPRECIO_PREPAGO_VENCIMIENTO_MES6_NOMBRE
    cboPrepagoVencimiento.AddItem LISTAPRECIO_PREPAGO_VENCIMIENTO_ANIO1_NOMBRE
    
    Call CSM_Control_TextBox.PrepareAll(Me)
    
    chkPrepago_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    mKeyDecimal = CSM_Control_TextBox.CheckKeyDown(ActiveControl, KeyCode)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(ActiveControl, KeyAscii, mKeyDecimal)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mListaPrecio = Nothing
End Sub

Private Sub txtNombre_GotFocus()
    CSM_Control_TextBox.SelAllText txtNombre
End Sub

Private Sub txtNombre_Change()
    Caption = "Propiedades" & IIf(Trim(txtNombre.Text) = "", "", " de " & txtNombre.Text)
End Sub

Private Sub txtLeyenda_GotFocus()
    CSM_Control_TextBox.SelAllText txtLeyenda
End Sub

Private Sub txtDescripcion_GotFocus()
    CSM_Control_TextBox.SelAllText txtDescripcion
    cmdOK.Default = False
End Sub

Private Sub txtDescripcion_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub chkPrepago_Click()
    lblPrepagoVencimiento.Visible = (chkPrepago.Value = vbChecked)
    cboPrepagoVencimiento.Visible = (chkPrepago.Value = vbChecked)
    lblPrepagoReservasCantidad.Visible = (chkPrepago.Value = vbChecked)
    txtPrepagoReservasCantidad.Visible = (chkPrepago.Value = vbChecked)
    lblPrepagoReservasCantidad_Ilimitada.Visible = (chkPrepago.Value = vbChecked)
    
    lblCuentaCorrienteGrupo_Credito.Visible = (chkPrepago.Value = vbChecked)
    datcboCuentaCorrienteGrupo_Credito.Visible = (chkPrepago.Value = vbChecked)
    cmdCuentaCorrienteGrupo_Credito.Visible = (chkPrepago.Value = vbChecked)
    
    lblCuentaCorrienteGrupo_Debito.Visible = (chkPrepago.Value = vbChecked)
    datcboCuentaCorrienteGrupo_Debito.Visible = (chkPrepago.Value = vbChecked)
    cmdCuentaCorrienteGrupo_Debito.Visible = (chkPrepago.Value = vbChecked)
End Sub

Private Sub txtPrepagoReservasCantidad_GotFocus()
    CSM_Control_TextBox.SelAllText txtPrepagoReservasCantidad
End Sub

Private Sub txtPrepagoReservasCantidad_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtPrepagoReservasCantidad)
End Sub

Private Sub cmdCuentaCorrienteGrupo_Credito_Click()
    If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmCuentaCorrienteGrupo.Show
        On Error Resume Next
        Set frmCuentaCorrienteGrupo.lvwData.SelectedItem = frmCuentaCorrienteGrupo.lvwData.ListItems(KEY_STRINGER & Val(datcboCuentaCorrienteGrupo_Credito.BoundText))
        frmCuentaCorrienteGrupo.lvwData.SelectedItem.EnsureVisible
        If frmCuentaCorrienteGrupo.WindowState = vbMinimized Then
            frmCuentaCorrienteGrupo.WindowState = vbNormal
        End If
        frmCuentaCorrienteGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdCuentaCorrienteGrupo_Debito_Click()
    If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmCuentaCorrienteGrupo.Show
        On Error Resume Next
        Set frmCuentaCorrienteGrupo.lvwData.SelectedItem = frmCuentaCorrienteGrupo.lvwData.ListItems(KEY_STRINGER & Val(datcboCuentaCorrienteGrupo_Debito.BoundText))
        frmCuentaCorrienteGrupo.lvwData.SelectedItem.EnsureVisible
        If frmCuentaCorrienteGrupo.WindowState = vbMinimized Then
            frmCuentaCorrienteGrupo.WindowState = vbNormal
        End If
        frmCuentaCorrienteGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub cmdOK_Click()
    If Trim(txtNombre.Text) = "" Then
        MsgBox "Debe ingresar el Nombre de la Lista de Precios.", vbInformation, App.Title
        txtNombre.SetFocus
        txtNombre_GotFocus
        Exit Sub
    End If
    
    If chkPrepago.Value = vbChecked Then
        If cboPrepagoVencimiento.ListIndex = -1 Then
            MsgBox "Debe especificar los días de vencimiento del Prepago.", vbInformation, App.Title
            cboPrepagoVencimiento.SetFocus
            Exit Sub
        End If
        If txtPrepagoReservasCantidad.Text = "" Then
            MsgBox "Debe especificar la Cantidad de Reservas del Prepago (0 = Sin límite).", vbInformation, App.Title
            txtPrepagoReservasCantidad.SetFocus
            Exit Sub
        End If
        If Val(datcboCuentaCorrienteGrupo_Credito.BoundText) = 0 Then
            MsgBox "Debe especificar el Grupo de Cuenta Corriente para los Créditos.", vbInformation, App.Title
            datcboCuentaCorrienteGrupo_Credito.SetFocus
            Exit Sub
        End If
        If Val(datcboCuentaCorrienteGrupo_Debito.BoundText) = 0 Then
            MsgBox "Debe especificar el Grupo de Cuenta Corriente para los Débitos.", vbInformation, App.Title
            datcboCuentaCorrienteGrupo_Debito.SetFocus
            Exit Sub
        End If
    End If
    
    With mListaPrecio
        .Nombre = txtNombre.Text
        .Leyenda = txtLeyenda.Text
        .Descripcion = txtDescripcion.Text
        .PrepagoEs = (chkPrepago.Value = vbChecked)
        If chkPrepago.Value = vbChecked Then
            .PrepagoVencimiento_ListIndex = cboPrepagoVencimiento.ListIndex
            .PrepagoReservasCantidad = Val(txtPrepagoReservasCantidad.Text)
            .IDCuentaCorrienteGrupo_Credito = Val(datcboCuentaCorrienteGrupo_Credito.BoundText)
            .IDCuentaCorrienteGrupo_Debito = Val(datcboCuentaCorrienteGrupo_Debito.BoundText)
        Else
            .PrepagoVencimiento = ""
            .PrepagoReservasCantidad = 0
            .IDCuentaCorrienteGrupo_Credito = 0
            .IDCuentaCorrienteGrupo_Debito = 0
        End If
        .Notas = txtNotas.Text
        .Activo = (chkActivo.Value = vbChecked)
        If Not .Update() Then
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Set frmListaPrecioPropiedad = Nothing
End Sub

Public Sub FillComboBoxCuentaCorrienteGrupo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    'CRÉDITO
    KeySave = Val(datcboCuentaCorrienteGrupo_Credito.BoundText)
    Set recData = datcboCuentaCorrienteGrupo_Credito.RowSource
    recData.Requery
    Set recData = Nothing
    datcboCuentaCorrienteGrupo_Credito.BoundText = KeySave

    'DÉBITO
    KeySave = Val(datcboCuentaCorrienteGrupo_Debito.BoundText)
    Set recData = datcboCuentaCorrienteGrupo_Debito.RowSource
    recData.Requery
    Set recData = Nothing
    datcboCuentaCorrienteGrupo_Debito.BoundText = KeySave
End Sub
