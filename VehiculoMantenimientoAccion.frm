VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVehiculoMantenimientoAccion 
   Caption         =   "Acciones de Mantenimiento de Vehículos"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   Icon            =   "VehiculoMantenimientoAccion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   10050
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   60
      TabIndex        =   18
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
      Height          =   1380
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   2434
      BandCount       =   4
      FixedOrder      =   -1  'True
      _CBWidth        =   10050
      _CBHeight       =   1380
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinWidth1       =   4410
      MinHeight1      =   540
      Width1          =   4410
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterFecha"
      MinWidth2       =   6705
      MinHeight2      =   360
      Width2          =   6705
      Key2            =   "FilterFecha"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Child3          =   "picFilterVehiculo"
      MinWidth3       =   4200
      MinHeight3      =   360
      Width3          =   4200
      Key3            =   "FilterVehiculo"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Child4          =   "picFilterGrupo"
      MinWidth4       =   3945
      MinHeight4      =   330
      Width4          =   3945
      Key4            =   "FilterGrupo"
      NewRow4         =   0   'False
      AllowVertical4  =   0   'False
      Begin VB.PictureBox picFilterGrupo 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6015
         ScaleHeight     =   330
         ScaleWidth      =   3945
         TabIndex        =   21
         Top             =   1005
         Width           =   3945
         Begin VB.ComboBox cboFilterGrupo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   0
            Width           =   3330
         End
         Begin VB.Label lblFilterGrupo 
            AutoSize        =   -1  'True
            Caption         =   "Grupo:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   23
            Top             =   60
            Width           =   495
         End
      End
      Begin VB.PictureBox picFilterVehiculo 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   5625
         TabIndex        =   10
         Top             =   990
         Width           =   5625
         Begin VB.ComboBox cboFilterVehiculo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   780
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   0
            Width           =   3330
         End
         Begin VB.Label lblVehiculo 
            AutoSize        =   -1  'True
            Caption         =   "Vehículo:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   11
            Top             =   60
            Width           =   675
         End
      End
      Begin VB.PictureBox picFilterFecha 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   9795
         TabIndex        =   4
         Top             =   600
         Width           =   9795
         Begin VB.TextBox txtDiaSemana 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.CommandButton cmdHoyHasta 
            Height          =   315
            Left            =   6360
            Picture         =   "VehiculoMantenimientoAccion.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdSiguienteHasta 
            Height          =   315
            Left            =   6060
            Picture         =   "VehiculoMantenimientoAccion.frx":0A14
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdAnteriorHasta 
            Height          =   315
            Left            =   4320
            Picture         =   "VehiculoMantenimientoAccion.frx":0F9E
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdHoyDesde 
            Height          =   315
            Left            =   3720
            Picture         =   "VehiculoMantenimientoAccion.frx":1528
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdSiguienteDesde 
            Height          =   315
            Left            =   3420
            Picture         =   "VehiculoMantenimientoAccion.frx":1672
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdAnteriorDesde 
            Height          =   315
            Left            =   1680
            Picture         =   "VehiculoMantenimientoAccion.frx":1BFC
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.ComboBox cboFecha 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   0
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   315
            Left            =   1980
            TabIndex        =   6
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
            Format          =   151257089
            CurrentDate     =   36950
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   315
            Left            =   4620
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
            Format          =   151257089
            CurrentDate     =   36950
         End
         Begin VB.Label lblFechaAnd 
            AutoSize        =   -1  'True
            Caption         =   "y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4140
            TabIndex        =   9
            Top             =   60
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   8
            Top             =   60
            Width           =   495
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   540
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   953
         ButtonWidth     =   1931
         ButtonHeight    =   953
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Nuevo"
               Key             =   "NEW"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Propiedades"
               Key             =   "PROPERTIES"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Eliminar"
               Key             =   "DELETE"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Seleccionar"
               Key             =   "SELECT"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6030
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   661
            MinWidth        =   661
            Key             =   "PIN"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16510
            Key             =   "TEXT"
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
      Left            =   60
      TabIndex        =   0
      Top             =   1440
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Vehiculo"
         Text            =   "Vehículo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Grupo"
         Text            =   "Grupo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Fecha"
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Kilometraje"
         Text            =   "Kilometraje"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Litros"
         Text            =   "Litros"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Importe"
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmVehiculoMantenimientoAccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mLoading As Boolean

Public FormWaitingForSelect As String
Public SelectTag As String

Public Sub FillListView(ByVal IDVehiculoMantenimientoAccion As Long)
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    Dim ImporteTotal As Currency
    Dim LitrosTotal As Double
    
    If mLoading Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    If IDVehiculoMantenimientoAccion = 0 Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & IDVehiculoMantenimientoAccion
    End If
    
    SQL_Where = ""
    
    If cboFecha.ListIndex > 0 Then
        If cboFecha.ListIndex < 7 Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "convert(char(10), VehiculoMantenimientoAccion.FechaHora, 111) " & cboFecha.Text & " '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & "'"
        Else
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "convert(char(10), VehiculoMantenimientoAccion.FechaHora, 111) BETWEEN '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & "' AND '" & Format(dtpFechaHasta.Value, "yyyy/mm/dd") & "'"
        End If
    End If
    
    If cboFilterVehiculo.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "VehiculoMantenimientoAccion.IDVehiculo = " & cboFilterVehiculo.ItemData(cboFilterVehiculo.ListIndex)
    End If
    
    If cboFilterGrupo.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "VehiculoMantenimientoAccion.IDVehiculoMantenimientoGrupo = " & cboFilterGrupo.ItemData(cboFilterGrupo.ListIndex)
    End If
        
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Select Case lvwData.SortKey
        Case 0  'VEHICULO
            SQL_OrderBy = " ORDER BY Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoGrupo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoAccion.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 1  'GRUPO
            SQL_OrderBy = " ORDER BY VehiculoMantenimientoGrupo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoAccion.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'FECHA
            SQL_OrderBy = " ORDER BY VehiculoMantenimientoAccion.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoGrupo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 3  'KILOMETRAJE
            SQL_OrderBy = " ORDER BY VehiculoMantenimientoAccion.Kilometraje" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoGrupo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoAccion.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 4  'LITROS
            SQL_OrderBy = " ORDER BY VehiculoMantenimientoAccion.Litros" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoGrupo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoAccion.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 5  'IMPORTE
            SQL_OrderBy = " ORDER BY VehiculoMantenimientoAccion.Importe" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoGrupo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoAccion.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    lvwData.ListItems.Clear
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT VehiculoMantenimientoAccion.IDVehiculoMantenimientoAccion, Vehiculo.Nombre AS Vehiculo, VehiculoMantenimientoGrupo.Nombre AS Grupo, VehiculoMantenimientoAccion.FechaHora, VehiculoMantenimientoAccion.Kilometraje, VehiculoMantenimientoAccion.Litros, VehiculoMantenimientoAccion.Importe FROM (VehiculoMantenimientoAccion INNER JOIN Vehiculo ON VehiculoMantenimientoAccion.IDVehiculo = Vehiculo.IDVehiculo) INNER JOIN VehiculoMantenimientoGrupo ON VehiculoMantenimientoAccion.IDVehiculoMantenimientoGrupo = VehiculoMantenimientoGrupo.IDVehiculoMantenimientoGrupo" & SQL_Where & SQL_OrderBy
    recData.MaxRecords = pParametro.Recordset_MaxRecords
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDVehiculoMantenimientoAccion").Value, .Fields("Vehiculo").Value)
                ListItem.SubItems(1) = .Fields("Grupo").Value
                ListItem.SubItems(2) = Format(.Fields("FechaHora").Value, "Short Date") & " " & Format(.Fields("FechaHora").Value, "Short Time")
                ListItem.SubItems(3) = IIf(IsNull(.Fields("Kilometraje").Value), "", .Fields("Kilometraje").Value)
                ListItem.SubItems(4) = IIf(IsNull(.Fields("Litros").Value), "", .Fields("Litros").Value)
                ListItem.SubItems(5) = Format(.Fields("Importe").Value, "Currency")
                If Not IsNull(.Fields("Importe").Value) Then
                    ImporteTotal = ImporteTotal + .Fields("Importe").Value
                End If
                If Not IsNull(.Fields("Litros").Value) Then
                    LitrosTotal = LitrosTotal + .Fields("Litros").Value
                End If
                .MoveNext
            Loop
            
            stbMain.Panels("TEXT").Text = .RecordCount & " items" & IIf(.RecordCount = .MaxRecords, " (Limitados)", "") & " - Importe Total: " & Format(ImporteTotal, "Currency") & " - Total Litros: " & Format(LitrosTotal, "#,##0.0")
        Else
            stbMain.Panels("TEXT").Text = "No hay items."
        End If
        .Close
    End With
    Set recData = Nothing
    
    On Error Resume Next
    lvwData.SelectedItem.Selected = False
    Set lvwData.SelectedItem = lvwData.ListItems(KeySave)

    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.VehiculoMantenimientoAccion.FillListView", "Error al leer la lista de Acciones de Mantenimiento de Vehículos."
End Sub

Public Sub FillComboBoxVehiculo()
    Dim recVehiculo As ADODB.Recordset
    Dim KeySave As Long
    
    If cboFilterVehiculo.ListIndex > -1 Then
        KeySave = cboFilterVehiculo.ItemData(cboFilterVehiculo.ListIndex)
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recVehiculo = New ADODB.Recordset
    recVehiculo.Source = "SELECT IDVehiculo, Nombre FROM Vehiculo WHERE Activo = 1 ORDER BY Nombre"
    recVehiculo.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    cboFilterVehiculo.Clear
    cboFilterVehiculo.AddItem ITEM_ALL_MALE
    cboFilterVehiculo.ItemData(cboFilterVehiculo.NewIndex) = 0
    Do While Not recVehiculo.EOF
        cboFilterVehiculo.AddItem recVehiculo("Nombre").Value
        cboFilterVehiculo.ItemData(cboFilterVehiculo.NewIndex) = recVehiculo("IDVehiculo").Value
        recVehiculo.MoveNext
    Loop
    recVehiculo.Close
    Set recVehiculo = Nothing

    cboFilterVehiculo.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(cboFilterVehiculo, KeySave, cscpItemOrfirst)
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.VehiculoMantenimientoAccion.FillComboBoxVehiculo", "Error al leer la lista de Vehículos."
End Sub

Public Sub FillComboBoxVehiculoMantenimientoGrupo()
    Dim recVehiculoMantenimientoGrupo As ADODB.Recordset
    Dim KeySave As Long
    
    If cboFilterGrupo.ListIndex > -1 Then
        KeySave = cboFilterGrupo.ItemData(cboFilterGrupo.ListIndex)
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recVehiculoMantenimientoGrupo = New ADODB.Recordset
    recVehiculoMantenimientoGrupo.Source = "SELECT IDVehiculoMantenimientoGrupo, Nombre FROM VehiculoMantenimientoGrupo WHERE Activo = 1 ORDER BY Nombre"
    recVehiculoMantenimientoGrupo.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    cboFilterGrupo.Clear
    cboFilterGrupo.AddItem ITEM_ALL_MALE
    cboFilterGrupo.ItemData(cboFilterGrupo.NewIndex) = 0
    Do While Not recVehiculoMantenimientoGrupo.EOF
        cboFilterGrupo.AddItem recVehiculoMantenimientoGrupo("Nombre").Value
        cboFilterGrupo.ItemData(cboFilterGrupo.NewIndex) = recVehiculoMantenimientoGrupo("IDVehiculoMantenimientoGrupo").Value
        recVehiculoMantenimientoGrupo.MoveNext
    Loop
    recVehiculoMantenimientoGrupo.Close
    Set recVehiculoMantenimientoGrupo = Nothing

    cboFilterGrupo.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(cboFilterGrupo, KeySave, cscpItemOrfirst)
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.VehiculoMantenimientoAccion.FillComboBoxVehiculoMantenimientoGrupo", "Error al leer la lista de Grupos de Mantenimiento de Vehículos."
End Sub

Private Sub cboFecha_Click()
    txtDiaSemana.Visible = (cboFecha.ListIndex > 0 And cboFecha.ListIndex < 7)
    cmdAnteriorDesde.Visible = (cboFecha.ListIndex > 0)
    dtpFechaDesde.Visible = (cboFecha.ListIndex > 0)
    cmdSiguienteDesde.Visible = (cboFecha.ListIndex > 0)
    cmdHoyDesde.Visible = (cboFecha.ListIndex > 0)
    
    lblFechaAnd.Visible = (cboFecha.ListIndex = 7)
    
    cmdAnteriorHasta.Visible = (cboFecha.ListIndex = 7)
    dtpFechaHasta.Visible = (cboFecha.ListIndex = 7)
    cmdSiguienteHasta.Visible = (cboFecha.ListIndex = 7)
    cmdHoyHasta.Visible = (cboFecha.ListIndex = 7)
    
    cmdAnteriorDesde.Left = 1680
    dtpFechaDesde.Left = 1980
    cmdSiguienteDesde.Left = 3420
    cmdHoyDesde.Left = 3720
    
    If cboFecha.ListIndex > 0 And cboFecha.ListIndex < 7 Then
        cmdAnteriorDesde.Left = cmdAnteriorDesde.Left + txtDiaSemana.Width
        dtpFechaDesde.Left = dtpFechaDesde.Left + txtDiaSemana.Width
        cmdSiguienteDesde.Left = cmdSiguienteDesde.Left + txtDiaSemana.Width
        cmdHoyDesde.Left = cmdHoyDesde.Left + txtDiaSemana.Width
    End If
    
    FillListView 0
End Sub

Private Sub cboFilterGrupo_Click()
    FillListView 0
End Sub

Private Sub cboFilterVehiculo_Click()
    FillListView 0
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
End Sub

Private Sub dtpFechaDesde_Change()
    txtDiaSemana.Text = WeekdayName(Weekday(dtpFechaDesde.Value))
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
            Case vbKeyD
                tlbMain_ButtonClick tlbMain.Buttons.Item("DETAIL")
            Case vbKeyS
                tlbMain_ButtonClick tlbMain.Buttons.Item("SELECT")
        End Select
    End If
End Sub

Private Sub Form_Load()
    lvwData.GridLines = pParametro.ListView_GridLines
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL TOOLBAR
    Set tlbMain.ImageList = frmMDI.ilsFormToolbar
    Set tlbMain.HotImageList = frmMDI.ilsFormToolbarHot
    tlbMain.Buttons("NEW").Image = "NEW"
    tlbMain.Buttons("PROPERTIES").Image = "PROPERTIES"
    tlbMain.Buttons("DELETE").Image = "DELETE"
    tlbMain.Buttons("SELECT").Image = "SELECT"
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwData.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL PIN
    Set tlbPin.ImageList = frmMDI.ilsFormPin
    '//////////////////////////////////////////////////////////
    
    mLoading = True
    
    cboFecha.AddItem "<Todas>"
    cboFecha.AddItem "="
    cboFecha.AddItem ">"
    cboFecha.AddItem ">="
    cboFecha.AddItem "<"
    cboFecha.AddItem "<="
    cboFecha.AddItem "<>"
    cboFecha.AddItem "Entre"
    cboFecha.ListIndex = 1
    
    dtpFechaDesde.Value = Date
    txtDiaSemana.Text = WeekdayName(Weekday(dtpFechaDesde.Value))
    dtpFechaHasta.Value = Date
    
    FillComboBoxVehiculo
    FillComboBoxVehiculoMantenimientoGrupo
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "VehiculoMantenimientoAccion", cbrMain
    pParametro.GetListViewSettings "VehiculoMantenimientoAccion", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").Value = pParametro.Usuario_LeerNumero("VehiculoMantenimientoAccion_Pin", tlbPin.Buttons("PIN").Value)
    If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    mLoading = False

    FillListView 0
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "VehiculoMantenimientoAccion", cbrMain
    pParametro.SaveListViewSettings "VehiculoMantenimientoAccion", lvwData
    pParametro.Usuario_GuardarNumero "VehiculoMantenimientoAccion_Pin", tlbPin.Buttons("PIN").Value
    Set frmVehiculoMantenimientoAccion = Nothing
End Sub

Private Sub lvwData_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    '//////////////////////////////////////////////////////////////////
    'Debido a que no se puede ordenar por Disponibles, ignoro el click en la columna
    If ColumnHeader.Key = "Disponibles" Then
        Exit Sub
    End If
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = 0
    If ColumnHeader.Index - 1 = lvwData.SortKey Then
        lvwData.SortOrder = IIf(lvwData.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwData.SortKey = ColumnHeader.Index - 1
        lvwData.SortOrder = lvwAscending
    End If
    ColumnHeader.Icon = lvwData.SortOrder + 1
    FillListView 0
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
    Dim VehiculoMantenimientoAccion As VehiculoMantAccion
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO_ACCION_ADD) Then
                Set VehiculoMantenimientoAccion = New VehiculoMantAccion
                
                Screen.MousePointer = vbHourglass
                frmVehiculoMantenimientoAccionPropiedad.LoadDataAndShow Me, VehiculoMantenimientoAccion
                Screen.MousePointer = vbDefault
                
                Set VehiculoMantenimientoAccion = Nothing
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO_ACCION_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Set VehiculoMantenimientoAccion = New VehiculoMantAccion
                VehiculoMantenimientoAccion.IDVehiculoMantenimientoAccion = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                If VehiculoMantenimientoAccion.Load() Then
                    Screen.MousePointer = vbHourglass
                    frmVehiculoMantenimientoAccionPropiedad.LoadDataAndShow Me, VehiculoMantenimientoAccion
                    Screen.MousePointer = vbDefault
                End If
                Set VehiculoMantenimientoAccion = Nothing
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO_ACCION_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                If MsgBox("¿Desea eliminar la Acción de Mantenimiento de Vehículos seleccionada?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    Set VehiculoMantenimientoAccion = New VehiculoMantAccion
                    VehiculoMantenimientoAccion.IDVehiculoMantenimientoAccion = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                    If VehiculoMantenimientoAccion.Load() Then
                        VehiculoMantenimientoAccion.Delete
                    End If
                    Set VehiculoMantenimientoAccion = Nothing
                    
                    lvwData.SetFocus
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
                
                Screen.MousePointer = vbHourglass
                Forms(FormIndex).VehiculoMantenimientoAccionSelected Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1)), SelectTag
                Forms(FormIndex).SetFocus
                If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
                    Unload Me
                End If
                Screen.MousePointer = vbDefault
            End If
            FormWaitingForSelect = ""
    End Select
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
