VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCuentaCorriente 
   Caption         =   "Cuenta Corriente"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CuentaCorriente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   11100
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   60
      TabIndex        =   17
      Top             =   5220
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PIN"
            ImageIndex      =   1
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   1410
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   2487
      BandCount       =   7
      FixedOrder      =   -1  'True
      _CBWidth        =   11100
      _CBHeight       =   1410
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinWidth1       =   5670
      MinHeight1      =   570
      Width1          =   5670
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterPersona"
      MinWidth2       =   4710
      MinHeight2      =   360
      Width2          =   4710
      Key2            =   "FilterPersona"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Child3          =   "picFilterFecha"
      MinWidth3       =   6705
      MinHeight3      =   360
      Width3          =   6705
      Key3            =   "FilterFecha"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Child4          =   "picFilterGrupo"
      MinWidth4       =   3225
      MinHeight4      =   360
      Width4          =   3225
      Key4            =   "FilterGrupo"
      NewRow4         =   0   'False
      AllowVertical4  =   0   'False
      Child5          =   "picFilterCaja"
      MinWidth5       =   3225
      MinHeight5      =   360
      Width5          =   3225
      Key5            =   "FilterCaja"
      NewRow5         =   0   'False
      AllowVertical5  =   0   'False
      Child6          =   "picFilterTipo"
      MinWidth6       =   1605
      MinHeight6      =   330
      Width6          =   1605
      Key6            =   "picFilterTipo"
      NewRow6         =   0   'False
      AllowVertical6  =   0   'False
      Child7          =   "picMedioPago"
      MinWidth7       =   3795
      MinHeight7      =   360
      Width7          =   3795
      Key7            =   "MedioPago"
      NewRow7         =   0   'False
      AllowVertical7  =   0   'False
      Begin VB.PictureBox picMedioPago 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   5445
         ScaleHeight     =   360
         ScaleWidth      =   5565
         TabIndex        =   32
         Top             =   1020
         Width           =   5565
         Begin VB.ComboBox cboMedioPago 
            Height          =   330
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   0
            Width           =   2610
         End
         Begin VB.Label lblMedioPago 
            AutoSize        =   -1  'True
            Caption         =   "Medio de Pago:"
            Height          =   210
            Left            =   0
            TabIndex        =   34
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.PictureBox picFilterTipo 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3615
         ScaleHeight     =   330
         ScaleWidth      =   1605
         TabIndex        =   29
         Top             =   1035
         Width           =   1605
         Begin VB.ComboBox cboFilterTipo 
            Height          =   330
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   0
            Width           =   1170
         End
         Begin VB.Label lblFilterTipo 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   210
            Left            =   0
            TabIndex        =   31
            Top             =   60
            Width           =   345
         End
      End
      Begin VB.PictureBox picFilterCaja 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   3225
         TabIndex        =   25
         Top             =   1020
         Width           =   3225
         Begin VB.ComboBox cboCaja 
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   0
            Width           =   2610
         End
         Begin VB.Label lblCaja 
            AutoSize        =   -1  'True
            Caption         =   "Caja:"
            Height          =   210
            Left            =   0
            TabIndex        =   27
            Top             =   60
            Width           =   360
         End
      End
      Begin VB.PictureBox picFilterGrupo 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   7785
         ScaleHeight     =   360
         ScaleWidth      =   3225
         TabIndex        =   22
         Top             =   630
         Width           =   3225
         Begin VB.ComboBox cboGrupo 
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   0
            Width           =   2610
         End
         Begin VB.Label lblGrupo 
            AutoSize        =   -1  'True
            Caption         =   "Grupo:"
            Height          =   210
            Left            =   0
            TabIndex        =   23
            Top             =   60
            Width           =   495
         End
      End
      Begin VB.PictureBox picFilterFecha 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   7395
         TabIndex        =   5
         Top             =   630
         Width           =   7395
         Begin VB.CommandButton cmdHoyHasta 
            Height          =   315
            Left            =   6360
            Picture         =   "CuentaCorriente.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdSiguienteHasta 
            Height          =   315
            Left            =   6060
            Picture         =   "CuentaCorriente.frx":06D4
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdAnteriorHasta 
            Height          =   315
            Left            =   4320
            Picture         =   "CuentaCorriente.frx":0C5E
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdHoyDesde 
            Height          =   315
            Left            =   3720
            Picture         =   "CuentaCorriente.frx":11E8
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdSiguienteDesde 
            Height          =   315
            Left            =   3420
            Picture         =   "CuentaCorriente.frx":1332
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdAnteriorDesde 
            Height          =   315
            Left            =   1680
            Picture         =   "CuentaCorriente.frx":18BC
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.ComboBox cboFecha 
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   0
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   315
            Left            =   1980
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   107151361
            CurrentDate     =   36950
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   315
            Left            =   4620
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   107151361
            CurrentDate     =   36950
         End
         Begin VB.Label lblFechaAnd 
            AutoSize        =   -1  'True
            Caption         =   "y"
            Height          =   210
            Left            =   4140
            TabIndex        =   10
            Top             =   60
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   210
            Left            =   0
            TabIndex        =   9
            Top             =   60
            Width           =   495
         End
      End
      Begin VB.PictureBox picFilterPersona 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   6300
         ScaleHeight     =   360
         ScaleWidth      =   4710
         TabIndex        =   4
         Top             =   135
         Width           =   4710
         Begin VB.CommandButton cmdPersonaClear 
            Height          =   315
            Left            =   3780
            Picture         =   "CuentaCorriente.frx":1E46
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Borrar"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton cmdPersona 
            Height          =   315
            Left            =   3420
            Picture         =   "CuentaCorriente.frx":23D0
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Buscar..."
            Top             =   0
            Width           =   315
         End
         Begin VB.TextBox txtPersona 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   2715
         End
         Begin VB.CommandButton cmdUltimo 
            Caption         =   "Ultimo"
            Height          =   315
            Left            =   4140
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   555
         End
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            Caption         =   "Persona:"
            Height          =   210
            Left            =   0
            TabIndex        =   21
            Top             =   60
            Width           =   645
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   1005
         ButtonWidth     =   2170
         ButtonHeight    =   1005
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Nuevo"
               Key             =   "NEW"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Propiedades"
               Key             =   "PROPERTIES"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Eliminar"
               Key             =   "DELETE"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Seleccionar"
               Key             =   "SELECT"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "PRINT"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_CLIENTE"
                     Text            =   "Composición de Saldo Cliente"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_LISTADO"
                     Text            =   "Listado"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_LISTADO_COMPLETO"
                     Text            =   "Listado Completo"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5295
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   661
            MinWidth        =   661
            Key             =   "PIN"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   185
            MinWidth        =   176
            Key             =   "TEXT"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18150
            Key             =   "INFO"
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ListView lvwData 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1500
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "IDMovimiento"
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "FechaHora"
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Grupo"
         Text            =   "Grupo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Caja"
         Text            =   "Caja"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Cliente"
         Text            =   "Cliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Descripcion"
         Text            =   "Descripcion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Realizado"
         Text            =   "Realizado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "PersonaOrigen"
         Text            =   "Pasajero"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Key             =   "Importe"
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "MedioPago"
         Text            =   "Medio de Pago"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Key             =   "Acumulado"
         Text            =   "Acumulado"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmCuentaCorriente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean
Private mEntidadTipo As String
Private mCIDGrupos As Collection
Private mCIDCajas As Collection

Private mrecData As ADODB.Recordset

Public DatabaseName As String
Public IsHistory As Boolean

Public FormWaitingForSelect As String

Public Sub LoadDataAndShow()
    Load Me
    
    If Not FillListView(0) Then
        Unload Me
        Exit Sub
    End If

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub ForceRefresh()
    FillListView 0
End Sub

Public Function FillListView(ByVal IDMovimiento As Long) As Boolean
    Dim Persona As Persona
    Dim CuentaCorrienteCaja As CuentaCorrienteCaja
    Dim ListItem As MSComctlLib.ListItem
    Dim KeySave As String
    
    Dim SQL_Select_SaldoAnterior As String
    Dim SQL_From_SaldoAnterior As String
    Dim SQL_Where_SaldoAnterior As String
    
    Dim SQL_Select As String
    Dim SQL_From As String
    Dim SQL_Where As String
    
    Dim SaldoAcumulado As Currency
    
    If mLoading Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If IDMovimiento = 0 Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & IDMovimiento
    End If
    
    'PERSONAL
    If pPersonal Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "(Viaje.Personal IS NULL OR Viaje.Personal = 0)"
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorriente.SaldoAnterior = 0"
    End If
    
    If Val(txtPersona.Tag) > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorriente.IDPersona = " & Val(txtPersona.Tag)
    End If
    
    If cboGrupo.ListIndex = 0 Then
        'SELECCIONADOS TODOS LOS GRUPOS, FILTRO POR CAMPO OCULTAR DE LOS GRUPOS SI NO TIENE PERMISOS
        If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_GRUPO_HIDDEN_SHOW, False) Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorrienteGrupo.Ocultar = 0"
        End If
    Else
        'SELECCIONADO UN GRUPO, FILTRO POR ESE
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorriente.IDCuentaCorrienteGrupo = " & cboGrupo.ItemData(cboGrupo.ListIndex)
    End If
    
    'CAJA
    If cboCaja.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorriente.IDCuentaCorrienteCaja = " & cboCaja.ItemData(cboCaja.ListIndex) & " AND (CuentaCorriente.IDCuentaCorrienteCaja <> " & pParametro.CuentaCorrienteCaja_ID_ViajeDebito & " OR CuentaCorriente.Importe > 0)"
    End If
    
    Select Case cboFilterTipo.ListIndex
        Case 1
            SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorriente.Importe >= 0"
        Case 2
            SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorriente.Importe < 0"
    End Select
    
    'MEDIO PAGO
    If cboMedioPago.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorriente.IDMedioPago = " & cboMedioPago.ItemData(cboMedioPago.ListIndex)
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    'SALDO INICIAL
    If cboFecha.ListIndex > 0 Then
        If cboFecha.ListIndex <> 3 Then
            SQL_Select_SaldoAnterior = "(SELECT NULL AS IDMovimiento, NULL AS FechaHora, NULL AS FacturaNumero, NULL AS IDCuentaCorrienteGrupo, NULL AS CuentaCorrienteGrupo, NULL AS IDCuentaCorrienteCaja, NULL AS CuentaCorrienteCaja, NULL AS Persona, 'SALDO ANTERIOR:' AS Descripcion, NULL AS Realizado, ISNULL(convert(money, sum(Importe)), 0) AS Importe, NULL AS MedioPago, NULL AS PersonaOrigen" & vbCr
            If pPersonal Then
                SQL_From_SaldoAnterior = "FROM ([" & DatabaseName & "]..CuentaCorriente INNER JOIN CuentaCorrienteGrupo ON CuentaCorriente.IDCuentaCorrienteGrupo = CuentaCorrienteGrupo.IDCuentaCorrienteGrupo) LEFT JOIN [" & DatabaseName & "]..Viaje ON CuentaCorriente.Viaje_FechaHora = Viaje.FechaHora AND CuentaCorriente.Viaje_IDRuta = Viaje.IDRuta" & vbCr
            Else
                SQL_From_SaldoAnterior = "FROM [" & DatabaseName & "]..CuentaCorriente INNER JOIN CuentaCorrienteGrupo ON CuentaCorriente.IDCuentaCorrienteGrupo = CuentaCorrienteGrupo.IDCuentaCorrienteGrupo" & vbCr
            End If
            SQL_Where_SaldoAnterior = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "convert(char(10), CuentaCorriente.FechaHora, 111) < '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & "')" & vbCr
        End If
    End If
    
    'DATE FILTER
    Select Case cboFecha.ListIndex
        Case 0  'ALL
        Case 1  'EQUAL
            SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorriente.FechaHora BETWEEN '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & " 00:00:00' AND '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & " 23:59:00'"
        Case 2  'GREATER OR EQUAL
            SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorriente.FechaHora >= '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & " 00:00:00'"
        Case 3  'MINOR OR EQUAL
            SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorriente.FechaHora <= '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & " 23:59:00'"
        Case 4  'BETWEEN
            SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "CuentaCorriente.FechaHora BETWEEN '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & " 00:00:00' AND '" & Format(dtpFechaHasta.Value, "yyyy/mm/dd") & " 23:59:00'"
    End Select
    
    If SQL_Where <> "" Then
        SQL_Where = SQL_Where & vbCr
    End If
    
    SQL_Select = "SELECT CuentaCorriente.IDMovimiento, CuentaCorriente.FechaHora, ViajeDetalle.FacturaNumero, CuentaCorriente.IDCuentaCorrienteGrupo, CuentaCorrienteGrupo.Nombre AS CuentaCorrienteGrupo, CuentaCorriente.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre AS CuentaCorrienteCaja, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, CuentaCorriente.Descripcion, ViajeDetalle.Realizado, CuentaCorriente.Importe, MedioPago.Nombre AS MedioPago, Pasajero.Apellido + (CASE ISNULL(Pasajero.Nombre, '') WHEN '' THEN '' ELSE ', ' + Pasajero.Nombre END) AS PersonaOrigen" & vbCr
    If pPersonal Then
        SQL_From = "FROM (((((([" & DatabaseName & "]..CuentaCorriente INNER JOIN CuentaCorrienteGrupo ON CuentaCorriente.IDCuentaCorrienteGrupo = CuentaCorrienteGrupo.IDCuentaCorrienteGrupo) INNER JOIN CuentaCorrienteCaja ON CuentaCorriente.IDCuentaCorrienteCaja = CuentaCorrienteCaja.IDCuentaCorrienteCaja) LEFT JOIN MedioPago ON CuentaCorriente.IDMedioPago = MedioPago.IDMedioPago) LEFT JOIN Persona ON CuentaCorriente.IDPersona = Persona.IDPersona) LEFT JOIN Persona AS Pasajero ON CuentaCorriente.IDPersonaOrigen = Pasajero.IDPersona) LEFT JOIN [" & DatabaseName & "]..ViajeDetalle ON CuentaCorriente.Viaje_FechaHora = ViajeDetalle.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalle.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalle.Indice) LEFT JOIN [" & DatabaseName & "]..Viaje AS Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta" & vbCr
    Else
        SQL_From = "FROM ((((([" & DatabaseName & "]..CuentaCorriente INNER JOIN CuentaCorrienteGrupo ON CuentaCorriente.IDCuentaCorrienteGrupo = CuentaCorrienteGrupo.IDCuentaCorrienteGrupo) INNER JOIN CuentaCorrienteCaja ON CuentaCorriente.IDCuentaCorrienteCaja = CuentaCorrienteCaja.IDCuentaCorrienteCaja) LEFT JOIN MedioPago ON CuentaCorriente.IDMedioPago = MedioPago.IDMedioPago) LEFT JOIN Persona ON CuentaCorriente.IDPersona = Persona.IDPersona) LEFT JOIN Persona AS Pasajero ON CuentaCorriente.IDPersonaOrigen = Pasajero.IDPersona) LEFT JOIN [" & DatabaseName & "]..ViajeDetalle ON CuentaCorriente.Viaje_FechaHora = ViajeDetalle.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalle.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalle.Indice" & vbCr
    End If
    
    lvwData.ListItems.Clear
    Set mCIDGrupos = New Collection
    Set mCIDCajas = New Collection
    
    Set mrecData = New ADODB.Recordset
    If SQL_Select_SaldoAnterior <> "" Then
        mrecData.Source = "(" & SQL_Select_SaldoAnterior & SQL_From_SaldoAnterior & SQL_Where_SaldoAnterior & ") UNION" & vbCr & SQL_Select & SQL_From & SQL_Where & "ORDER BY FechaHora, IDMovimiento"
    Else
        mrecData.Source = SQL_Select & SQL_From & SQL_Where & "ORDER BY CuentaCorriente.FechaHora, CuentaCorriente.IDMovimiento"
    End If
    mrecData.MaxRecords = pParametro.Recordset_MaxRecords
    mrecData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With mrecData
        If Not .EOF Then
            Do While Not .EOF
                mCIDGrupos.Add .Fields("IDCuentaCorrienteGrupo").Value
                mCIDCajas.Add .Fields("IDCuentaCorrienteCaja").Value
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & Val(.Fields("IDMovimiento").Value & ""), .Fields("IDMovimiento").Value & "")
                ListItem.SubItems(1) = Format(.Fields("FechaHora").Value, "Short Date") & " " & Format(.Fields("FechaHora").Value, "Short Time")
                ListItem.SubItems(2) = .Fields("CuentaCorrienteGrupo").Value & ""
                ListItem.SubItems(3) = .Fields("CuentaCorrienteCaja").Value & ""
                ListItem.SubItems(4) = .Fields("Persona").Value & ""
                ListItem.SubItems(5) = .Fields("Descripcion").Value
                ListItem.SubItems(6) = IIf(IsNull(.Fields("Realizado").Value), "", IIf(.Fields("Realizado").Value, "Sí", "No"))
                ListItem.SubItems(7) = .Fields("PersonaOrigen").Value & ""
                ListItem.SubItems(8) = Format(.Fields("Importe").Value, "Currency")
                SaldoAcumulado = SaldoAcumulado + .Fields("Importe").Value
                ListItem.SubItems(9) = .Fields("MedioPago").Value & ""
                ListItem.SubItems(10) = Format(SaldoAcumulado, "Currency")
                .MoveNext
            Loop
            
            If SQL_Select_SaldoAnterior <> "" Then
                stbMain.Panels("TEXT").Text = .RecordCount - 1 & " items" & IIf(.RecordCount >= .MaxRecords, " (Limitados)", "") & " "
            Else
                stbMain.Panels("TEXT").Text = .RecordCount & " items" & IIf(.RecordCount >= .MaxRecords, " (Limitados)", "") & " "
            End If
        Else
            stbMain.Panels("TEXT").Text = "No hay items. "
        End If
    End With
    
    stbMain.Panels("INFO").Text = ""
    
    If Val(txtPersona.Tag) > 0 Then
        Set Persona = New Persona
        Persona.IDPersona = Val(txtPersona.Tag)
        If Persona.LoadSaldoActual() Then
            stbMain.Panels("INFO").Text = "Saldo Actual: " & Persona.SaldoActual_Formatted
        End If
        Set Persona = Nothing
    End If
    If cboCaja.ListIndex > 0 Then
        Set CuentaCorrienteCaja = New CuentaCorrienteCaja
        CuentaCorrienteCaja.IDCuentaCorrienteCaja = cboCaja.ItemData(cboCaja.ListIndex)
        If CuentaCorrienteCaja.LoadSaldoActual() Then
            stbMain.Panels("INFO").Text = "Saldo Actual: " & CuentaCorrienteCaja.SaldoActual_Formatted
        End If
        Set CuentaCorrienteCaja = Nothing
    End If
    
    On Error Resume Next
    Set lvwData.SelectedItem = lvwData.ListItems(KeySave)
    lvwData.SelectedItem.EnsureVisible

    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    FillListView = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.CuentaCorriente.FillListView", "Error al obtener la lista de Movimientos."
End Function

Private Sub cboCaja_Click()
    If cboCaja.ListIndex > 0 Then
        txtPersona.Tag = 0
        txtPersona.Text = ""
    End If
    FillListView 0
End Sub

Private Sub cboFecha_Click()
    cmdAnteriorDesde.Visible = (cboFecha.ListIndex > 0)
    dtpFechaDesde.Visible = (cboFecha.ListIndex > 0)
    cmdSiguienteDesde.Visible = (cboFecha.ListIndex > 0)
    cmdHoyDesde.Visible = (cboFecha.ListIndex > 0)
    
    lblFechaAnd.Visible = (cboFecha.ListIndex = 4)
    
    cmdAnteriorHasta.Visible = (cboFecha.ListIndex = 4)
    dtpFechaHasta.Visible = (cboFecha.ListIndex = 4)
    cmdSiguienteHasta.Visible = (cboFecha.ListIndex = 4)
    cmdHoyHasta.Visible = (cboFecha.ListIndex = 4)
    
    FillListView 0
End Sub

Private Sub cboFilterTipo_Click()
    FillListView 0
End Sub

Private Sub cboMedioPago_Click()
    FillListView 0
End Sub

Private Sub cboGrupo_Click()
    FillListView 0
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
End Sub

Private Sub cmdPersona_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(txtPersona.Tag), UCase(Left(txtPersona.Text, 1)), Me.Name, "", "")
    End If
End Sub

Private Sub cmdPersonaClear_Click()
    If Val(txtPersona.Tag) <> 0 Then
        txtPersona.Tag = 0
        txtPersona.Text = ""
        FillListView 0
    End If
End Sub

Private Sub cmdUltimo_Click()
    If frmMDI.cboPersona.ListIndex > -1 Then
        PersonaSelected Val(frmMDI.cboPersona.ItemData(frmMDI.cboPersona.ListIndex)), "PP"
    End If
    cmdPersona.SetFocus
End Sub

Private Sub dtpFechaDesde_Change()
    FillListView 0
End Sub

Private Sub cmdAnteriorDesde_Click()
    dtpFechaDesde.Value = DateAdd("d", -1, dtpFechaDesde.Value)
    dtpFechaDesde.SetFocus
    dtpFechaDesde_Change
End Sub

Private Sub cmdSiguienteDesde_Click()
    dtpFechaDesde.Value = DateAdd("d", 1, dtpFechaDesde.Value)
    dtpFechaDesde.SetFocus
    dtpFechaDesde_Change
End Sub

Private Sub cmdHoyDesde_Click()
    Dim OldValue As Date
    
    OldValue = dtpFechaDesde.Value
    dtpFechaDesde.Value = Date
    dtpFechaDesde.SetFocus
    If OldValue <> dtpFechaDesde.Value Then
        dtpFechaDesde_Change
    End If
End Sub

Private Sub dtpFechaHasta_Change()
    FillListView 0
End Sub

Private Sub cmdAnteriorHasta_Click()
    dtpFechaHasta.Value = DateAdd("d", -1, dtpFechaHasta.Value)
    dtpFechaHasta.SetFocus
    dtpFechaHasta_Change
End Sub

Private Sub cmdSiguienteHasta_Click()
    dtpFechaHasta.Value = DateAdd("d", 1, dtpFechaHasta.Value)
    dtpFechaHasta.SetFocus
    dtpFechaHasta_Change
End Sub

Private Sub cmdHoyHasta_Click()
    Dim OldValue As Date
    
    OldValue = dtpFechaHasta.Value
    dtpFechaHasta.Value = Date
    dtpFechaHasta.SetFocus
    If OldValue <> dtpFechaHasta.Value Then
        dtpFechaHasta_Change
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyN
                tlbMain_ButtonClick tlbMain.Buttons.Item("NEW")
            Case vbKeyP
                tlbMain_ButtonClick tlbMain.Buttons.Item("PROPERTIES")
            Case vbKeyE
                tlbMain_ButtonClick tlbMain.Buttons.Item("DELETE")
            Case vbKeyS
                tlbMain_ButtonClick tlbMain.Buttons.Item("SELECT")
        End Select
    End If
End Sub

Private Sub Form_Load()
    Dim Persona As Persona
    
    mLoading = True
        
    lvwData.GridLines = pParametro.ListView_GridLines
    
    If pParametro.MedioPago_Predeterminado_ID = 0 Then
        cbrMain.Bands("MedioPago").Visible = False
    End If
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL TOOLBAR
    Set tlbMain.ImageList = frmMDI.ilsFormToolbar
    Set tlbMain.HotImageList = frmMDI.ilsFormToolbarHot
    tlbMain.Buttons("NEW").Image = "NEW"
    tlbMain.Buttons("PROPERTIES").Image = "PROPERTIES"
    tlbMain.Buttons("DELETE").Image = "DELETE"
    tlbMain.Buttons("SELECT").Image = "SELECT"
    tlbMain.Buttons("PRINT").Image = "PRINT"
    cbrMain.Bands("Toolbar").MinWidth = CSM_Control_Toolbar.GetTotalWidth(tlbMain)
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwData.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL PIN
    Set tlbPin.ImageList = frmMDI.ilsFormPin
    '//////////////////////////////////////////////////////////
    
    cboFecha.AddItem "<Todas>"
    cboFecha.AddItem "="
    cboFecha.AddItem ">="
    cboFecha.AddItem "<="
    cboFecha.AddItem "Entre"
    cboFecha.ListIndex = 1
    
    dtpFechaDesde.Value = Date
    dtpFechaHasta.Value = Date
    
    FillComboBoxCuentaCorrienteGrupo
    FillComboBoxCuentaCorrienteCaja
    FillComboBoxMedioPago
    
    cboFilterTipo.AddItem ITEM_ALL_MALE
    cboFilterTipo.AddItem "Ingresos"
    cboFilterTipo.AddItem "Egresos"
    cboFilterTipo.ListIndex = 0
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "CuentaCorriente", cbrMain
    pParametro.GetListViewSettings "CuentaCorriente", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").Value = pParametro.Usuario_LeerNumero("CuentaCorriente_Pin", tlbPin.Buttons("PIN").Value)
    If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    mLoading = False
    
    If frmMDI.cboPersona.ListIndex > -1 Then
        Set Persona = New Persona
        Persona.IDPersona = Val(frmMDI.cboPersona.ItemData(frmMDI.cboPersona.ListIndex))
        If Persona.Load() Then
            Select Case Persona.EntidadTipo
                Case ENTIDAD_TIPO_PERSONA_CLIENTE
                    txtPersona.Tag = Val(frmMDI.cboPersona.ItemData(frmMDI.cboPersona.ListIndex))
                    txtPersona.Text = frmMDI.cboPersona.Text
                Case ENTIDAD_TIPO_PERSONA_CONDUCTOR
                    If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CONDUCTOR_SELECT, False) Then
                        txtPersona.Tag = Val(frmMDI.cboPersona.ItemData(frmMDI.cboPersona.ListIndex))
                        txtPersona.Text = frmMDI.cboPersona.Text
                    End If
                Case ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO
                    If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_ADMINISTRATIVO_SELECT, False) Then
                        txtPersona.Tag = Val(frmMDI.cboPersona.ItemData(frmMDI.cboPersona.ListIndex))
                        txtPersona.Text = frmMDI.cboPersona.Text
                    End If
            End Select
        End If
        
        Set Persona = Nothing
        
        FillListView 0
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mCIDGrupos = Nothing
    Set mCIDCajas = Nothing
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "CuentaCorriente", cbrMain
    pParametro.SaveListViewSettings "CuentaCorriente", lvwData
    pParametro.Usuario_GuardarNumero "CuentaCorriente_Pin", tlbPin.Buttons("PIN").Value
    If Not mrecData Is Nothing Then
        If mrecData.State = adStateOpen Then
            If Not (mrecData.BOF Or mrecData.EOF) Then
                If mrecData.EditMode <> adEditNone Then
                    mrecData.CancelUpdate
                End If
            End If
            mrecData.Close
        End If
        Set mrecData = Nothing
    End If
End Sub

Public Sub FillComboBoxCuentaCorrienteGrupo()
    Dim KeySave As Long
    
    If cboGrupo.ListCount > 0 Then
        KeySave = cboGrupo.ItemData(cboGrupo.ListIndex)
    End If
    Call CSM_Control_ComboBox.FillFromSQL(cboGrupo, "(SELECT 0 AS IDCuentaCorrienteGrupo, '<Todos>' AS Nombre, 1 AS Orden) UNION (SELECT IDCuentaCorrienteGrupo, Nombre, 2 AS Orden FROM CuentaCorrienteGrupo WHERE Activo = 1" & IIf(pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_GRUPO_HIDDEN_SHOW, False), "", " AND Ocultar = 0") & IIf(pCPermiso.CuentaCorrienteGrupoWhere <> "", " AND " & Replace(pCPermiso.CuentaCorrienteGrupoWhere, "%TABLENAME%", "ListaPrecio"), "") & ") ORDER BY Orden, Nombre", "IDCuentaCorrienteGrupo", "Nombre", "Grupos de Cuenta Corriente", cscpItemOrfirst, KeySave)
End Sub

Public Sub FillComboBoxCuentaCorrienteCaja()
    Dim KeySave As Long
    
    If cboCaja.ListCount > 0 Then
        KeySave = cboCaja.ItemData(cboCaja.ListIndex)
    End If
    Call CSM_Control_ComboBox.FillFromSQL(cboCaja, "(SELECT 0 AS IDCuentaCorrienteCaja, '<Todos>' AS Nombre, 1 AS Orden) UNION (SELECT IDCuentaCorrienteCaja, Nombre, 2 AS Orden FROM CuentaCorrienteCaja WHERE Activo = 1" & IIf(pCPermiso.CuentaCorrienteCajaWhere = "", "", " AND " & pCPermiso.CuentaCorrienteCajaWhere) & ") ORDER BY Orden, Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas de Cuenta Corriente", cscpItemOrfirst, KeySave)
End Sub

Public Sub FillComboBoxMedioPago()
    Dim KeySave As Long
    
    If cboMedioPago.ListCount > 0 Then
        KeySave = cboMedioPago.ItemData(cboMedioPago.ListIndex)
    End If
    Call CSM_Control_ComboBox.FillFromSQL(cboMedioPago, "(SELECT 0 AS IDMedioPago, '<Todos>' AS Nombre, 1 AS Orden) UNION (SELECT IDMedioPago, Nombre, 2 AS Orden FROM MedioPago WHERE Activo = 1) ORDER BY Orden, Nombre", "IDMedioPago", "Nombre", "Medios de Pago", cscpItemOrfirst, KeySave)
End Sub

Private Sub lvwData_DblClick()
    If GetFormIndex(FormWaitingForSelect) > 0 Then
        tlbMain_ButtonClick tlbMain.Buttons.Item("SELECT")
    Else
        tlbMain_ButtonClick tlbMain.Buttons.Item("PROPERTIES")
    End If
End Sub

Private Sub lvwData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        lvwData_DblClick
    End If
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim FormIndex As Long
    Dim CuentaCorriente As CuentaCorriente
    Dim DebitoCredito As Boolean
    Dim IDMovimiento As Long
    Dim recData As ADODB.Recordset
    Dim PermisoHabilitado As Boolean
    
    Select Case Button.Key
        Case "NEW"
            If IsHistory Then
                MsgBox "No se pueden crear nuevos items históricos.", vbInformation, App.Title
                lvwData.SetFocus
                Exit Sub
            End If
            
            PermisoHabilitado = (pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_ADD_ANTERIOR, False) Or pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_ADD_ACTUAL, False))
            If PermisoHabilitado Then
                Screen.MousePointer = vbHourglass
                
                Set CuentaCorriente = New CuentaCorriente
                frmCuentaCorrientePropiedad.IsHistory = IsHistory
                frmCuentaCorrientePropiedad.LoadDataAndShow Me, CuentaCorriente
                Set CuentaCorriente = Nothing
                
                Screen.MousePointer = vbDefault
            Else
                MsgBox "No está autorizado a realizar esta acción.", vbExclamation, App.Title
            End If
        Case "PROPERTIES"
            If lvwData.SelectedItem Is Nothing Then
                MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                lvwData.SetFocus
                Exit Sub
            End If
            If lvwData.SelectedItem.Text = "" Then
                MsgBox "No se puede modificar el Saldo Anterior.", vbInformation, App.Title
                lvwData.SetFocus
                Exit Sub
            End If

            Screen.MousePointer = vbHourglass
            
            Set CuentaCorriente = New CuentaCorriente
            CuentaCorriente.IDMovimiento = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
            CuentaCorriente.IsHistory = IsHistory
            If Not CuentaCorriente.Load() Then
                Set CuentaCorriente = Nothing
                lvwData.SetFocus
                Exit Sub
            End If
            
            If CuentaCorriente.SaldoAnterior Then
                Set CuentaCorriente = Nothing
                MsgBox "No se pueden modificar los items pertenecientes a Saldos Anteriores.", vbExclamation, App.Title
                lvwData.SetFocus
                Exit Sub
            End If
            
            frmCuentaCorrientePropiedad.IsHistory = IsHistory
            frmCuentaCorrientePropiedad.LoadDataAndShow Me, CuentaCorriente
            
            Set CuentaCorriente = Nothing
            
            Screen.MousePointer = vbDefault
        Case "DELETE"
            If IsHistory Then
                MsgBox "No se pueden eliminar items históricos.", vbInformation, App.Title
                lvwData.SetFocus
                Exit Sub
            End If
            If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If lvwData.SelectedItem.Text = "" Then
                    MsgBox "No se puede eliminar el Saldo Anterior.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If mCIDGrupos(lvwData.SelectedItem.Index) = pParametro.CuentaCorrienteGrupo_ID_ViajeDebito Then
                    DebitoCredito = True
                    If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_DELETE_DEBITO_CREDITO, False) Then
                        MsgBox "No se pueden eliminar los Movimientos pertenecientes al grupo '" & lvwData.SelectedItem.SubItems(2) & "'." & vbCr & "Estos Movimientos se eliminan automáticamente al eliminar o cancelar los Viajes.", vbExclamation, App.Title
                        lvwData.SetFocus
                        Exit Sub
                    End If
                End If
                If mCIDGrupos(lvwData.SelectedItem.Index) = pParametro.CuentaCorrienteGrupo_ID_ViajeCredito Then
                    DebitoCredito = True
                    If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_DELETE_DEBITO_CREDITO, False) Then
                        MsgBox "No se pueden eliminar los items pertenecientes al grupo '" & lvwData.SelectedItem.SubItems(2) & "'." & vbCr & "Estos Movimientos se eliminan automáticamente al eliminar los Viajes o los pagos en los mismos.", vbExclamation, App.Title
                        lvwData.SetFocus
                        Exit Sub
                    End If
                End If
                
                If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL, False) Then
                    If mCIDCajas(lvwData.SelectedItem.Index) <> pUsuario.IDCuentaCorrienteCaja And mCIDCajas(lvwData.SelectedItem.Index) <> pParametro.CuentaCorrienteCaja_ID_ViajeDebito Then
                        MsgBox "No se pueden eliminar los Movimientos de esta Caja.", vbExclamation, App.Title
                        lvwData.SetFocus
                        Exit Sub
                    End If
                End If
                
                IDMovimiento = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                Set CuentaCorriente = New CuentaCorriente
                CuentaCorriente.IDMovimiento = IDMovimiento
                If Not CuentaCorriente.Load() Then
                    Set CuentaCorriente = Nothing
                    Exit Sub
                End If
                
                If CuentaCorriente.SaldoAnterior Then
                    Set CuentaCorriente = Nothing
                    MsgBox "No se pueden eliminar los items pertenecientes a Saldos Anteriores.", vbExclamation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                If MsgBox("¿Desea eliminar el Movimiento seleccionado?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    If DebitoCredito Then
                        If pTrapErrors Then
                            On Error GoTo ErrorHandler
                        End If
                        
                        Set recData = New ADODB.Recordset
                        recData.Source = "SELECT FechaHora, IDRuta, Orden FROM ViajeDetalle WHERE IDMovimientoDebito = " & IDMovimiento & " OR IDMovimientoCredito = " & IDMovimiento
                        recData.Open , pDatabase.Connection, adOpenStatic, adLockReadOnly, adCmdText
                        If Not recData.EOF Then
                            MsgBox "No se puede Eliminar este Movimiento porque está Relacionado a una Reserva o Comisión." & vbCr & vbCr & "Fecha/Hora: " & Format(recData("FechaHora").Value, "Short Date") & " " & Format(recData("FechaHora").Value, "Short Time") & vbCr & "Ruta: " & recData("IDRuta").Value, vbExclamation, App.Title
                            recData.Close
                            Set recData = Nothing
                            
                            On Error GoTo 0
                            Exit Sub
                        End If
                        recData.Close
                        Set recData = Nothing
                        
                        On Error GoTo 0
                    End If
                    
                    CuentaCorriente.Delete
                    Set CuentaCorriente = Nothing
                End If
            End If
        Case "SELECT"
            FormIndex = GetFormIndex(FormWaitingForSelect)
            If FormIndex >= 0 Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If lvwData.SelectedItem.Text = "" Then
                    MsgBox "No se puede Seleccionar el Saldo Anterior.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                Forms(FormIndex).CuentaCorrienteSelected Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                Forms(FormIndex).SetFocus
                If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
                    Unload Me
                End If
                Screen.MousePointer = vbDefault
            End If
            FormWaitingForSelect = ""
        Case "PRINT"
            Call tlbMain_ButtonMenuClick(tlbMain.Buttons("PRINT").ButtonMenus("PRINT_CLIENTE"))
    End Select
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.CuentaCorriente.Delete", "Error al verificar si el Movimiento pertenece a una Reserva." & vbCr & vbCr & "IDMovimiento: " & IDMovimiento
End Sub

Private Sub tlbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim Reporte As Reporte
    Dim ReporteSubTitle As String
    
    If Not mrecData Is Nothing Then
        Select Case ButtonMenu.Parent.Key
            Case "PRINT"
                Select Case cboFecha.ListIndex
                    Case 0
                    Case 1
                        ReporteSubTitle = "Del día " & dtpFechaDesde.Value
                    Case 2
                        ReporteSubTitle = "Desde el día " & dtpFechaDesde.Value
                    Case 3
                        ReporteSubTitle = "Hasta el día " & dtpFechaDesde.Value
                    Case 4
                        ReporteSubTitle = dtpFechaDesde.Value & " al " & dtpFechaHasta.Value
                End Select
                If Val(txtPersona.Tag) > 0 Then
                    ReporteSubTitle = ReporteSubTitle & IIf(ReporteSubTitle = "", "", " - Cliente: ") & txtPersona.Text
                End If
                ReporteSubTitle = ReporteSubTitle & IIf(ReporteSubTitle = "", "", vbCr) & "Usuario: " & pUsuario.Nombre
                Select Case ButtonMenu.Key
                    Case "PRINT_CLIENTE"
                        If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "CuentaCorriente_Cliente") Then
                            Set Reporte = New Reporte
                            Reporte.IDReporte = "CuentaCorriente_Cliente"
                            If Reporte.Load() Then
                                Reporte.Titulo = Reporte.Titulo & IIf(ReporteSubTitle = "", "", vbCr) & ReporteSubTitle
                                Set Reporte.Recordset = mrecData
                                If Reporte.OpenReport() Then
                                    Reporte.PrintReport pParametro.Report_Preview
                                End If
                            End If
                            Set Reporte = Nothing
                        End If
                    Case "PRINT_LISTADO"
                        If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "CuentaCorriente_Listado") Then
                            Set Reporte = New Reporte
                            Reporte.IDReporte = "CuentaCorriente_Listado"
                            If Reporte.Load() Then
                                Reporte.Titulo = Reporte.Titulo & IIf(ReporteSubTitle = "", "", vbCr) & ReporteSubTitle
                                Set Reporte.Recordset = mrecData
                                If Reporte.OpenReport() Then
                                    Reporte.PrintReport pParametro.Report_Preview
                                End If
                                Set Reporte = Nothing
                            End If
                        End If
                Case "PRINT_LISTADO_COMPLETO"
                    If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "CuentaCorriente_Listado_Completo") Then
                        Set Reporte = New Reporte
                        Reporte.IDReporte = "CuentaCorriente_Listado_Completo"
                        If Reporte.Load() Then
                            Reporte.Titulo = Reporte.Titulo & IIf(ReporteSubTitle = "", "", vbCr) & ReporteSubTitle
                            Set Reporte.Recordset = mrecData
                            If Reporte.OpenReport() Then
                                Reporte.PrintReport pParametro.Report_Preview
                            End If
                            Set Reporte = Nothing
                        End If
                    End If
            End Select
        End Select
    End If
End Sub

Private Sub tlbPin_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Value = tbrUnpressed Then
        Button.Image = 1
    Else
        Button.Image = 2
    End If
End Sub

Private Sub ResizeControls(ByVal CoolBarHeight As Single)
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    lvwData.Top = CoolBarHeight + CONTROL_SPACE
    lvwData.Left = CONTROL_SPACE
    lvwData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwData.Height = ScaleHeight - lvwData.Top - CONTROL_SPACE - stbMain.Height
    
    tlbPin.Top = ScaleHeight - 330
    tlbPin.Left = 15
End Sub

Public Sub PersonaSelected(ByVal IDPersona As Long, ByVal Tag As String)
    Dim Persona As Persona
    
    Set Persona = New Persona
    Persona.IDPersona = IDPersona
    If Not Persona.Load() Then
        Set Persona = Nothing
        Exit Sub
    End If
    Select Case Persona.EntidadTipo
        Case ENTIDAD_TIPO_PERSONA_CLIENTE
        Case ENTIDAD_TIPO_PERSONA_CONDUCTOR
            If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CONDUCTOR_SELECT, False) Then
                MsgBox "No puede seleccionar Personas de tipo Conductor.", vbExclamation, App.Title
                Set Persona = Nothing
                On Error Resume Next
                lvwData.SetFocus
                On Error GoTo 0
                Exit Sub
            End If
        Case ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO
            If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_ADMINISTRATIVO_SELECT, False) Then
                MsgBox "No puede seleccionar Personas de tipo Administrativo.", vbExclamation, App.Title
                Set Persona = Nothing
                On Error Resume Next
                lvwData.SetFocus
                On Error GoTo 0
                Exit Sub
            End If
    End Select
    
    mEntidadTipo = Persona.EntidadTipo
    txtPersona.Tag = IDPersona
    txtPersona.Text = Persona.ApellidoNombre
    Set Persona = Nothing
    
    On Error Resume Next
    lvwData.SetFocus
    On Error GoTo 0
    
    cboCaja.ListIndex = 0
    
    FillListView 0
End Sub
