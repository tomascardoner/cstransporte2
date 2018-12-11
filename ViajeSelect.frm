VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmViajeSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione el Viaje"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajeSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6345
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3300
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4020
      TabIndex        =   6
      Top             =   3300
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   1575
      Left            =   180
      TabIndex        =   8
      Top             =   1500
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   2778
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "DiaSemana"
         Text            =   "Día"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Fecha"
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Hora"
         Text            =   "Hora"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "IDRuta"
         Text            =   "Ruta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Estado"
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Asientos"
         Text            =   "Asientos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Notas"
         Text            =   "Notas"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblFinData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lblInicioData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   660
      Width           =   2655
   End
   Begin VB.Label lblVehiculoData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblFin 
      Caption         =   "Selección Fin:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblInicio 
      Caption         =   "Selección Inicio:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblVehiculo 
      Caption         =   "Vehículo:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmViajeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mButtonClicked As String

Public Function LoadDataAndShow(ByVal VehiculoNombre As String, ByVal FechaHoraInicio As String, ByVal FechaHoraFin As String, ByRef CViajes As Collection) As Viaje
    Dim Viaje As Viaje
    Dim ListItem As ListItem
    
    If CViajes.Count = 1 Then
        Set Viaje = CViajes(1)
    Else
        lblVehiculoData.Caption = VehiculoNombre
        lblInicioData.Caption = FechaHoraInicio
        lblFinData.Caption = FechaHoraFin
    
        For Each Viaje In CViajes
            With Viaje
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .FechaHora & KEY_DELIMITER & .IDRuta, WeekdayName(Weekday(.FechaHora)))
                ListItem.SubItems(1) = .FechaHora_FormattedAsDate
                ListItem.SubItems(2) = .FechaHora_FormattedAsTime
                ListItem.SubItems(3) = .Ruta_DisplayName
                ListItem.SubItems(4) = .Estado_ToString
                ListItem.SubItems(5) = .AsientoLibre
                ListItem.SubItems(6) = .Notas
            End With
        Next Viaje
        
        Me.Show vbModal, frmMDI
        If mButtonClicked = "OK" Then
            Set Viaje = CViajes(lvwData.SelectedItem.Index)
        Else
            Set Viaje = Nothing
        End If
        Unload Me
    End If
    
    Set LoadDataAndShow = Viaje
End Function

Private Sub cmdAceptar_Click()
    If lvwData.SelectedItem Is Nothing Then
        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
        lvwData.SetFocus
        Exit Sub
    End If
    
    mButtonClicked = "OK"
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    mButtonClicked = "CANCEL"
    Me.Hide
End Sub

Private Sub Form_Load()
    lvwData.GridLines = pParametro.ListView_GridLines
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwData.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    '//////////////////////////////////////////////////////////
    
    pParametro.GetListViewSettings Mid(Me.Name, 4), lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pParametro.SaveListViewSettings Mid(Me.Name, 4), lvwData
    Set frmViajeSelect = Nothing
End Sub

Private Sub lvwData_DblClick()
    Call cmdAceptar_Click
End Sub
