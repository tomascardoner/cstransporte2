VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMedioPagoPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MedioPagoPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdMedioPagoPlan 
      Caption         =   "..."
      Height          =   315
      Left            =   4260
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Planes de Medios de Pago"
      Top             =   2460
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdCaja 
      Caption         =   "..."
      Height          =   315
      Left            =   4260
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Cajas"
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUtilizaOperacion 
      Alignment       =   1  'Right Justify
      Caption         =   "Utiliza Operación:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtAbreviatura 
      Height          =   315
      Left            =   1140
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1020
      Width           =   615
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   1140
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1500
      Width           =   3375
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
      TabIndex        =   15
      Top             =   780
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3300
      TabIndex        =   13
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboCaja 
      Height          =   330
      Left            =   1140
      TabIndex        =   9
      Top             =   2940
      Width           =   3075
      _ExtentX        =   5424
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
   Begin MSDataListLib.DataCombo datcboMedioPagoPlan 
      Height          =   330
      Left            =   1140
      TabIndex        =   6
      Top             =   2460
      Width           =   3075
      _ExtentX        =   5424
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
   Begin VB.Label lblMedioPagoPlan 
      AutoSize        =   -1  'True
      Caption         =   "Plan:"
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label lblCaja 
      AutoSize        =   -1  'True
      Caption         =   "Caja:"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label lblAbreviatura 
      AutoSize        =   -1  'True
      Caption         =   "Abreviatura:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   600
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos del Medio de Pago"
      Height          =   210
      Left            =   780
      TabIndex        =   14
      Top             =   300
      Width           =   2955
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMedioPagoPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMedioPago As MedioPago

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef MedioPago As MedioPago)
    Set mMedioPago = MedioPago
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mMedioPago
        txtAbreviatura.Text = .Abreviatura
        txtNombre.Text = .Nombre
        chkUtilizaOperacion.Value = IIf(.UtilizaOperacion, vbChecked, vbUnchecked)
        chkUtilizaOperacion_Click
        Call CSM_Control_DataCombo.FillFromSQL(datcboMedioPagoPlan, "SELECT IDMedioPagoPlan, Nombre FROM MedioPagoPlan WHERE Activo = 1 OR IDMedioPagoPlan = " & .IDMedioPagoPlan & " ORDER BY Nombre", "IDMedioPagoPlan", "Nombre", "Planes de Medios de Pago", cscpItemOrNone, .IDMedioPagoPlan)
        Call CSM_Control_DataCombo.FillFromSQL(datcboCaja, "(SELECT 1 AS Orden, 0 AS IDCuentaCorrienteCaja, '" & CSM_Constant.ITEM_NONE_CHARS20 & "' AS Nombre) UNION (SELECT 2 AS Orden, IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 OR IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & ") ORDER BY Orden, Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrfirst, .IDCuentaCorrienteCaja)

        chkActivo.Value = IIf(.Activo, vbChecked, vbUnchecked)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub chkUtilizaOperacion_Click()
    lblMedioPagoPlan.Visible = (chkUtilizaOperacion.Value = vbChecked)
    datcboMedioPagoPlan.Visible = (chkUtilizaOperacion.Value = vbChecked)
    cmdMedioPagoPlan.Visible = (chkUtilizaOperacion.Value = vbChecked)
End Sub

Private Sub txtAbreviatura_GotFocus()
    CSM_Control_TextBox.SelAllText txtAbreviatura
End Sub

Private Sub txtNombre_GotFocus()
    CSM_Control_TextBox.SelAllText txtNombre
End Sub

Private Sub txtNombre_Change()
    Caption = "Propiedades" & IIf(Trim(txtNombre.Text) = "", "", " de " & txtNombre.Text)
End Sub

Private Sub cmdOK_Click()
    If Trim(txtAbreviatura.Text) = "" Then
        MsgBox "Debe ingresar la Abreviatura del Medio de Pago.", vbInformation, App.Title
        txtAbreviatura.SetFocus
        txtAbreviatura_GotFocus
        Exit Sub
    End If
    If Trim(txtNombre.Text) = "" Then
        MsgBox "Debe ingresar el Nombre del Medio de Pago.", vbInformation, App.Title
        txtNombre.SetFocus
        txtNombre_GotFocus
        Exit Sub
    End If
    If (chkUtilizaOperacion.Value = vbChecked) And Val(datcboMedioPagoPlan.BoundText) = 0 Then
        MsgBox "Debe especificar el Plan.", vbInformation, App.Title
        datcboMedioPagoPlan.SetFocus
        Exit Sub
    End If
    
    With mMedioPago
        .Abreviatura = Trim(txtAbreviatura.Text)
        .Nombre = txtNombre.Text
        .UtilizaOperacion = (chkUtilizaOperacion.Value = vbChecked)
        If .UtilizaOperacion Then
            .IDMedioPagoPlan = Val(datcboMedioPagoPlan.BoundText)
        End If
        .IDCuentaCorrienteCaja = Val(datcboCaja.BoundText)
        .Activo = (chkActivo.Value = vbChecked)
        If Not .Update Then
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mMedioPago = Nothing
    Set frmMedioPagoPropiedad = Nothing
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
