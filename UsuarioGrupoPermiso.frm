VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmUsuarioGrupoPermiso 
   Caption         =   "Permisos"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UsuarioGrupoPermiso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   7080
   Begin VB.CommandButton cmdCheckNone 
      Caption         =   "Ninguno"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   180
      Width           =   1035
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "Todos"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   180
      Width           =   1035
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   4215
      Left            =   420
      TabIndex        =   1
      Top             =   1260
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Categoria"
         Text            =   "Categoría"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Acción"
         Text            =   "Acción"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   4875
      Left            =   240
      TabIndex        =   4
      Top             =   780
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8599
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "GENERAL"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rutas"
            Key             =   "RUTAS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Listas de Precios"
            Key             =   "LISTA_PRECIOS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cajas de Cuenta Corriente"
            Key             =   "CAJAS_CUENTACORRIENTE"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Grupos de Cuenta Corriente"
            Key             =   "GRUPOS_CUENTACORRIENTE"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Reportes"
            Key             =   "REPORTES"
            ImageVarType    =   2
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
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "UsuarioGrupoPermiso.frx":014A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione aquí los Permisos del Grupo de Usuarios"
      Height          =   210
      Left            =   780
      TabIndex        =   0
      Top             =   300
      Width           =   3765
   End
End
Attribute VB_Name = "frmUsuarioGrupoPermiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIDUsuarioGrupo As Long
Private mCPermiso As CPermiso

Public Sub LoadDataAndShow(ByVal IDUsuarioGrupo As Long)
    Dim UsuarioGrupo As UsuarioGrupo
    
    mIDUsuarioGrupo = IDUsuarioGrupo
    
    Load Me
    
    Set UsuarioGrupo = New UsuarioGrupo
    UsuarioGrupo.IDUsuarioGrupo = mIDUsuarioGrupo
    If Not UsuarioGrupo.Load() Then
        Unload Me
        Exit Sub
    End If
    Caption = "Permisos del Grupo de Usuarios: " & UsuarioGrupo.Nombre
    Set UsuarioGrupo = Nothing
    
    Set mCPermiso = New CPermiso
    mCPermiso.IDUsuarioGrupo = mIDUsuarioGrupo
    mCPermiso.OpenRecordset
    mCPermiso.Load
    
    If Not FillListView(mIDUsuarioGrupo, "") Then
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
    FillListView mIDUsuarioGrupo, ""
End Sub

Public Function FillListView(ByVal IDUsuarioGrupo As Long, ByVal IDPermiso As String) As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim KeySave As String
    Dim Permiso As Permiso
    
    Dim PermisoPrefixLen_Rutas As Integer
    Dim PermisoPrefixLen_ListaPrecios As Integer
    Dim PermisoPrefixLen_CuentaCorrienteCaja As Integer
    Dim PermisoPrefixLen_CuentaCorrienteGrupo As Integer
    Dim PermisoPrefixLen_Reportes As Integer
    
    Dim PermisoPrefix As String
    Dim PermisoPrefixLen As Integer
    
    If IDUsuarioGrupo <> mIDUsuarioGrupo Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If IDPermiso = "" Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & IDPermiso
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    lvwData.ListItems.Clear
    
    If tabMain.SelectedItem.Key = "GENERAL" Then
        '===========================
        'SOLO PARA LA OPCION GENERAL
        PermisoPrefixLen_Rutas = Len(PERMISO_RUTA_RUTA)
        PermisoPrefixLen_ListaPrecios = Len(PERMISO_LISTA_PRECIO_LISTA_PRECIO)
        PermisoPrefixLen_CuentaCorrienteCaja = Len(PERMISO_CUENTA_CORRIENTE_CAJA_CAJA)
        PermisoPrefixLen_CuentaCorrienteGrupo = Len(PERMISO_CUENTA_CORRIENTE_GRUPO_GRUPO)
        PermisoPrefixLen_Reportes = Len(PERMISO_REPORTE_REPORTE)
        
        For Each Permiso In mCPermiso
            With Permiso
                If Left(.IDPermiso, PermisoPrefixLen_Rutas) <> PERMISO_RUTA_RUTA And _
                    Left(.IDPermiso, PermisoPrefixLen_ListaPrecios) <> PERMISO_LISTA_PRECIO_LISTA_PRECIO And _
                    Left(.IDPermiso, PermisoPrefixLen_CuentaCorrienteCaja) <> PERMISO_CUENTA_CORRIENTE_CAJA_CAJA And _
                    Left(.IDPermiso, PermisoPrefixLen_CuentaCorrienteGrupo) <> PERMISO_CUENTA_CORRIENTE_GRUPO_GRUPO And _
                    Left(.IDPermiso, PermisoPrefixLen_Reportes) <> PERMISO_REPORTE_REPORTE Then
                    
                    Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .IDPermiso, .Categoria)
                    ListItem.SubItems(1) = .Descripcion
                    ListItem.Checked = .Permitido
                End If
            End With
        Next Permiso
    Else
        '========================
        'TODAS LAS DEMAS OPCIONES
        Select Case tabMain.SelectedItem.Key
            Case "RUTAS"
                PermisoPrefix = PERMISO_RUTA_RUTA
            Case "LISTA_PRECIOS"
                PermisoPrefix = PERMISO_LISTA_PRECIO_LISTA_PRECIO
            Case "CAJAS_CUENTACORRIENTE"
                PermisoPrefix = PERMISO_CUENTA_CORRIENTE_CAJA_CAJA
            Case "GRUPOS_CUENTACORRIENTE"
                PermisoPrefix = PERMISO_CUENTA_CORRIENTE_GRUPO_GRUPO
            Case "REPORTES"
                PermisoPrefix = PERMISO_REPORTE_REPORTE
        End Select
        PermisoPrefixLen = Len(PermisoPrefix)

        For Each Permiso In mCPermiso
            With Permiso
                If Left(.IDPermiso, PermisoPrefixLen) = PermisoPrefix Then
                    Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .IDPermiso, .Categoria)
                    ListItem.SubItems(1) = .Descripcion
                    ListItem.Checked = .Permitido
                End If
            End With
        Next Permiso
    End If
    
    Set Permiso = Nothing
    
    On Error Resume Next
    Set lvwData.SelectedItem = lvwData.ListItems(KeySave)
    
    Screen.MousePointer = vbDefault
    FillListView = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.UsuarioGrupoPermiso", "Error al obtener la Lista de Permisos."
End Function

Private Sub cmdCheckAll_Click()
    If MsgBox("ATENCION: Se habilitarán TODOS los permisos a este grupo de usuarios." & vbCr & vbCr & "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        CheckItems True
    End If
End Sub

Private Sub cmdCheckNone_Click()
    If MsgBox("ATENCION: Se quitarán TODOS los permisos a este grupo de usuarios." & vbCr & vbCr & "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        CheckItems False
    End If
End Sub

Private Sub Form_Load()
    lvwData.GridLines = pParametro.ListView_GridLines
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetListViewSettings "UsuarioGrupoPermiso", lvwData
End Sub

Private Sub Form_Resize()
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    tabMain.Left = CONTROL_SPACE
    tabMain.Width = ScaleWidth - tabMain.Left - CONTROL_SPACE
    tabMain.Height = ScaleHeight - tabMain.Top - CONTROL_SPACE
    
    lvwData.Left = tabMain.Left + (CONTROL_SPACE * 2)
    lvwData.Width = tabMain.Width - lvwData.Left - CONTROL_SPACE
    lvwData.Height = ScaleHeight - lvwData.Top - (CONTROL_SPACE * 3)
    
    cmdCheckNone.Left = ScaleWidth - CONTROL_SPACE - cmdCheckNone.Width
    cmdCheckAll.Left = cmdCheckNone.Left - CONTROL_SPACE - cmdCheckAll.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mCPermiso = Nothing
    Visible = False
    WindowState = vbNormal
    pParametro.SaveListViewSettings "UsuarioGrupoPermiso", lvwData
    Set frmUsuarioGrupoPermiso = Nothing
End Sub

Private Sub lvwData_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim UsuarioGrupoPermiso As UsuarioGrupoPermiso
    Dim IDPermiso As String
    
    If pCPermiso.GotPermission(PERMISO_USUARIO_GRUPO_PERMISSION_MODIFY) Then
        If Item Is Nothing Then
            MsgBox "No hay ningún Permiso seleccionado.", vbInformation, App.Title
            lvwData.SetFocus
            Exit Sub
        End If
        
        IDPermiso = CSM_String.GetSubString(Mid(Item.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)
        
        Set UsuarioGrupoPermiso = New UsuarioGrupoPermiso
        UsuarioGrupoPermiso.NoMatchRaiseError = False
        UsuarioGrupoPermiso.RefreshList = False
        UsuarioGrupoPermiso.IDUsuarioGrupo = mIDUsuarioGrupo
        UsuarioGrupoPermiso.IDPermiso = IDPermiso
        If Not UsuarioGrupoPermiso.Load() Then
            Item.Checked = Not Item.Checked
            Set UsuarioGrupoPermiso = Nothing
            Exit Sub
        End If
        
        If Item.Checked Then
            If UsuarioGrupoPermiso.NoMatch Then
                If Not UsuarioGrupoPermiso.AddNew() Then
                    Item.Checked = False
                    Set UsuarioGrupoPermiso = Nothing
                    Exit Sub
                End If
            End If
        Else
            If Not UsuarioGrupoPermiso.NoMatch Then
                If Not UsuarioGrupoPermiso.Delete() Then
                    Item.Checked = True
                    Set UsuarioGrupoPermiso = Nothing
                    Exit Sub
                End If
            End If
        End If
        Set UsuarioGrupoPermiso = Nothing
        
        If mIDUsuarioGrupo = pUsuario.IDUsuarioGrupo Then
            pCPermiso(IDPermiso).Permitido = Item.Checked
        End If
        mCPermiso(IDPermiso).Permitido = Item.Checked
    Else
        Item.Checked = Not Item.Checked
    End If
End Sub

Private Sub CheckItems(ByVal Value As Boolean)
    Dim Item As MSComctlLib.ListItem
    
    For Each Item In lvwData.ListItems
        If Item.Checked <> Value Then
            Item.Checked = Value
            lvwData_ItemCheck Item
        End If
    Next Item
End Sub

Private Sub tabMain_Click()
    FillListView mIDUsuarioGrupo, ""
End Sub
