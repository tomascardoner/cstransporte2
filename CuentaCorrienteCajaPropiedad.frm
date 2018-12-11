VERSION 5.00
Begin VB.Form frmCuentaCorrienteCajaPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   5100
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
   Icon            =   "CuentaCorrienteCajaPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   4680
   Begin VB.CheckBox chkOcultarSaldo 
      Alignment       =   1  'Right Justify
      Caption         =   "Ocultar Saldo:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1635
   End
   Begin VB.CheckBox chkMostrarSiempre 
      Alignment       =   1  'Right Justify
      Caption         =   "Mostrar Siempre:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1635
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000F&
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
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1500
      Width           =   2955
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
      Left            =   3840
      Picture         =   "CuentaCorrienteCajaPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Buscar..."
      Top             =   1500
      Width           =   315
   End
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
      Left            =   4200
      Picture         =   "CuentaCorrienteCajaPropiedad.frx":0AD4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Borrar"
      Top             =   1500
      Width           =   315
   End
   Begin VB.TextBox txtNotas 
      Height          =   945
      Left            =   900
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2820
      Width           =   3615
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   900
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1020
      Width           =   3615
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
      TabIndex        =   14
      Top             =   780
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3300
      TabIndex        =   12
      Top             =   4620
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   11
      Top             =   4620
      Width           =   1215
   End
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "&Persona:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   2880
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
      Caption         =   "Ingrese aquí los Datos de la Caja"
      Height          =   210
      Left            =   780
      TabIndex        =   13
      Top             =   300
      Width           =   2355
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmCuentaCorrienteCajaPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCuentaCorrienteCaja As CuentaCorrienteCaja
Private mNew As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef CuentaCorrienteCaja As CuentaCorrienteCaja)
    Dim Persona As Persona
    
    Set mCuentaCorrienteCaja = CuentaCorrienteCaja
    mNew = (mCuentaCorrienteCaja.IDCuentaCorrienteCaja = 0)

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mCuentaCorrienteCaja
        txtNombre.Text = .Nombre
        
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
        chkMostrarSiempre.Value = IIf(.MostrarSiempre, vbChecked, vbUnchecked)
        chkOcultarSaldo.Value = IIf(.OcultarSaldo, vbChecked, vbUnchecked)
        
        txtNotas.Text = .Notas
        chkActivo.Value = IIf(.Activo, vbChecked, vbUnchecked)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtNombre.Text) = "" Then
        MsgBox "Debe ingresar el Nombre de la Caja.", vbInformation, App.Title
        txtNombre.SetFocus
        txtNombre_GotFocus
        Exit Sub
    End If
    
    With mCuentaCorrienteCaja
        .Nombre = txtNombre.Text
        .IDPersona = Val(txtPersona.Tag)
        .MostrarSiempre = (chkMostrarSiempre.Value = vbChecked)
        .OcultarSaldo = (chkOcultarSaldo.Value = vbChecked)
        .Notas = txtNotas.Text
        .Activo = (chkActivo.Value = vbChecked)
        If Not .Update Then
            Exit Sub
        End If
    End With
    
    Unload Me
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

Private Sub Form_Unload(Cancel As Integer)
    Set mCuentaCorrienteCaja = Nothing
    Set frmCuentaCorrienteCajaPropiedad = Nothing
End Sub

Private Sub txtNombre_Change()
    Caption = "Propiedades" & IIf(Trim(txtNombre.Text) = "", "", " de " & txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
    CSM_Control_TextBox.SelAllText txtNombre
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
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
