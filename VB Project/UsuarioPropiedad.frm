VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmUsuarioPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UsuarioPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   4815
   Begin VB.CommandButton cmdPersonaClear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      Picture         =   "UsuarioPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Borrar"
      Top             =   3540
      Width           =   315
   End
   Begin VB.CommandButton cmdPersona 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4020
      Picture         =   "UsuarioPropiedad.frx":0AD4
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Buscar..."
      Top             =   3540
      Width           =   315
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3540
      Width           =   2955
   End
   Begin VB.CommandButton cmdViewPassword 
      Caption         =   "?"
      Height          =   315
      Left            =   3720
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2100
      Width           =   255
   End
   Begin VB.CommandButton cmdGrupo 
      Caption         =   "..."
      Height          =   315
      Left            =   4440
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Grupos"
      Top             =   2820
      Width           =   255
   End
   Begin VB.TextBox txtConfirma 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   2460
      Width           =   2595
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2100
      Width           =   2595
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   1080
      MaxLength       =   255
      TabIndex        =   5
      Top             =   1740
      Width           =   3615
   End
   Begin VB.TextBox txtNotas 
      Height          =   945
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   3900
      Width           =   3615
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo"
      Height          =   210
      Left            =   120
      TabIndex        =   22
      Top             =   4980
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1380
      Width           =   3615
   End
   Begin VB.TextBox txtLoginName 
      Height          =   315
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1020
      Width           =   2595
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   26
      Top             =   780
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3480
      TabIndex        =   24
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   23
      Top             =   5400
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboGrupo 
      Height          =   330
      Left            =   1080
      TabIndex        =   12
      Top             =   2820
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
   Begin MSDataListLib.DataCombo datcboEmpresa 
      Height          =   330
      Left            =   1080
      TabIndex        =   15
      Top             =   3180
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
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "Persona:"
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   645
   End
   Begin VB.Label lblEmpresa 
      AutoSize        =   -1  'True
      Caption         =   "Empresa:"
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   675
   End
   Begin VB.Label lblGrupo 
      AutoSize        =   -1  'True
      Caption         =   "&Grupo:"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblConfirma 
      AutoSize        =   -1  'True
      Caption         =   "Confirma:"
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   690
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "&Contraseña:"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label lblDescripcion 
      AutoSize        =   -1  'True
      Caption         =   "&Descripción:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   210
      Left            =   120
      TabIndex        =   20
      Top             =   3900
      Width           =   465
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos del Usuario"
      Height          =   210
      Left            =   780
      TabIndex        =   25
      Top             =   300
      Width           =   2460
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "UsuarioPropiedad.frx":105E
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblLoginName 
      AutoSize        =   -1  'True
      Caption         =   "&Login Name:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   885
   End
End
Attribute VB_Name = "frmUsuarioPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mUsuario As Usuario
Private mNew As Boolean

Private Const DUMMY_PASSWORD = "- ¿Creías que era tan fácil? -"

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Usuario As Usuario)
    Dim Persona As Persona
    
    Set mUsuario = Usuario
    mNew = (mUsuario.IDUsuario = 0)

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mUsuario
        txtLoginName.Text = .LoginName
        txtNombre.Text = .Nombre
        txtDescripcion.Text = .Descripcion
        txtPassword.Text = IIf(mNew, "", DUMMY_PASSWORD)
        txtConfirma.Text = IIf(mNew, "", DUMMY_PASSWORD)
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboGrupo, "SELECT IDUsuarioGrupo, Nombre FROM UsuarioGrupo WHERE IDUsuarioGrupo <> " & USUARIO_GRUPO_ID_ADMINISTRADORES & " AND (Activo = 1 OR IDUsuarioGrupo = " & .IDUsuarioGrupo & ") ORDER BY Nombre", "IDUsuarioGrupo", "Nombre", "Grupos de Usuarios", cscpItemOrNone, .IDUsuarioGrupo) Then
            Unload Me
            Exit Sub
        End If
        If Not CSM_Control_DataCombo.FillFromSQL(datcboEmpresa, "SELECT IDEmpresa, Nombre FROM Empresa WHERE Activo = 1 OR IDEmpresa = " & .IDEmpresa & " ORDER BY Nombre", "IDEmpresa", "Nombre", "Empresas", cscpItemOrNone, .IDEmpresa) Then
            Unload Me
            Exit Sub
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
        End If
        
        txtNotas.Text = .Notas
        chkActivo.Value = IIf(.Activo, vbChecked, vbUnchecked)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    
    cmdViewPassword.Visible = (pUsuario.IDUsuario = USUARIO_ID_ADMINISTRATOR)
    
    SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mUsuario = Nothing
    Set frmUsuarioPropiedad = Nothing
End Sub

Private Sub txtLoginName_GotFocus()
    CSM_Control_TextBox.SelAllText txtLoginName
End Sub

Private Sub txtNombre_Change()
    Caption = "Propiedades" & IIf(Trim(txtNombre.Text) = "", "", " de " & txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
    CSM_Control_TextBox.SelAllText txtNombre
End Sub

Private Sub txtDescripcion_GotFocus()
    CSM_Control_TextBox.SelAllText txtDescripcion
End Sub

Private Sub txtPassword_GotFocus()
    CSM_Control_TextBox.SelAllText txtPassword
End Sub

Private Sub cmdViewPassword_Click()
    If txtPassword.Text <> DUMMY_PASSWORD Then
        MsgBox "La contraseña del Usuario es: " & txtPassword.Text, vbInformation, App.Title
    Else
        MsgBox "La contraseña del Usuario es: " & mUsuario.Password, vbInformation, App.Title
    End If
End Sub

Private Sub txtConfirma_GotFocus()
    CSM_Control_TextBox.SelAllText txtConfirma
End Sub

Private Sub cmdGrupo_Click()
    If pCPermiso.GotPermission(PERMISO_USUARIO_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmUsuarioGrupo.Show
        On Error Resume Next
        Set frmUsuarioGrupo.lvwData.SelectedItem = frmUsuarioGrupo.lvwData.ListItems(KEY_STRINGER & datcboGrupo.BoundText)
        frmUsuarioGrupo.lvwData.SelectedItem.EnsureVisible
        If frmUsuarioGrupo.WindowState = vbMinimized Then
            frmUsuarioGrupo.WindowState = vbNormal
        End If
        frmUsuarioGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdPersona_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(txtPersona.Tag), UCase(Left(txtPersona.Text, 1)), Me.Name, ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO, "")
    End If
End Sub

Private Sub cmdPersonaClear_Click()
    txtPersona.Tag = 0
    txtPersona.Text = ""
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub cmdOK_Click()
    If Trim(txtLoginName.Text) = "" Then
        MsgBox "Debe ingresar el Login Name del Usuario.", vbInformation, App.Title
        txtLoginName.SetFocus
        txtLoginName_GotFocus
        Exit Sub
    End If
    If Trim(txtNombre.Text) = "" Then
        MsgBox "Debe ingresar el Nombre del Usuario.", vbInformation, App.Title
        txtNombre.SetFocus
        txtNombre_GotFocus
        Exit Sub
    End If
    If Trim(txtPassword.Text) = "" Then
        MsgBox "Debe ingresar la Contraseña.", vbInformation, App.Title
        txtPassword.SetFocus
        txtPassword_GotFocus
        Exit Sub
    End If
    If Len(txtPassword.Text) < 4 Then
        MsgBox "La Contraseña debe contener al menos 4 letras o números.", vbInformation, App.Title
        txtPassword.SetFocus
        txtPassword_GotFocus
        Exit Sub
    End If
    If Trim(txtConfirma.Text) = "" Then
        MsgBox "Debe confirmar la Contraseña.", vbInformation, App.Title
        txtConfirma.SetFocus
        txtConfirma_GotFocus
        Exit Sub
    End If
    If txtPassword.Text <> txtConfirma.Text Then
        MsgBox "La Contraseña no coincide con la confirmación.", vbInformation, App.Title
        txtConfirma.SetFocus
        txtConfirma_GotFocus
        Exit Sub
    End If
    If Val(datcboGrupo.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo.", vbInformation, App.Title
        datcboGrupo.SetFocus
        Exit Sub
    End If
    If Val(datcboEmpresa.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Empresa.", vbInformation, App.Title
        datcboEmpresa.SetFocus
        Exit Sub
    End If
    
    With mUsuario
        .LoginName = txtLoginName.Text
        .Nombre = txtNombre.Text
        .Descripcion = txtDescripcion.Text
        If txtPassword.Text <> DUMMY_PASSWORD Then
            .Password = txtPassword.Text
        End If
        .IDUsuarioGrupo = Val(datcboGrupo.BoundText)
        .IDEmpresa = Val(datcboEmpresa.BoundText)
        .IDPersona = Val(txtPersona.Tag)
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
End Sub

Public Sub FillComboBoxUsuarioGrupo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboGrupo.BoundText)
    Set recData = datcboGrupo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboGrupo.BoundText = KeySave
End Sub

Public Sub FillComboBoxEmpresa()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboEmpresa.BoundText)
    Set recData = datcboEmpresa.RowSource
    recData.Requery
    Set recData = Nothing
    datcboEmpresa.BoundText = KeySave
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

