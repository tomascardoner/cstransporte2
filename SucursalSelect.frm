VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSucursalSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccione las Sucursales"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6420
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5100
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "IDSucursal"
         Text            =   "ID"
         Object.Width           =   1535
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Nombre"
         Text            =   "Nombre"
         Object.Width           =   3625
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Email"
         Text            =   "E-mail"
         Object.Width           =   7435
      EndProperty
   End
End
Attribute VB_Name = "frmSucursalSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SucursalNombres As String
Public SucursalEmails As String

Public Sub FillListView(ByVal IDSucursal As String)
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT Sucursal.IDSucursal, Sucursal.Nombre, Sucursal.Email FROM Sucursal WHERE Sucursal.Activo = 1" & IIf(pParametro.IDSucursal = "", "", " AND Sucursal.IDSucursal <> '" & ReplaceQuote(pParametro.IDSucursal) & "'")
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & RTrim(.Fields("IDSucursal").Value), RTrim(.Fields("IDSucursal").Value))
                ListItem.SubItems(1) = .Fields("Nombre").Value
                ListItem.SubItems(2) = .Fields("Email").Value
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set recData = Nothing
    
    On Error Resume Next
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.Ruta.FillListView", "Error al obtener la Lista de Sucursales."
End Sub

Private Sub cmdOK_Click()
    Dim ListItem As MSComctlLib.ListItem
    
    For Each ListItem In lvwData.ListItems
        If ListItem.Checked Then
            SucursalNombres = SucursalNombres & IIf(SucursalNombres = "", "", "; ") & "Sucursal " & ListItem.SubItems(1)
            SucursalEmails = SucursalEmails & IIf(SucursalEmails = "", "", "; ") & ListItem.SubItems(2)
        End If
    Next ListItem
    
    If SucursalNombres = "" Then
        MsgBox "Debe seleecionar al menos una sucursal.", vbInformation, App.Title
        lvwData.SetFocus
        Exit Sub
    End If
    
    Tag = "OK"
    Hide
End Sub

Private Sub cmdCancel_Click()
    Tag = "CANCEL"
    Hide
End Sub

Private Sub Form_Load()
    FillListView 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSucursalSelect = Nothing
End Sub
