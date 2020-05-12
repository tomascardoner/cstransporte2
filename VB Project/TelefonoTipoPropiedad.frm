VERSION 5.00
Begin VB.Form frmTelefonoTipoPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "TelefonoTipoPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   4680
   Begin VB.TextBox txtDiscadoSufijo 
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
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1860
      Width           =   795
   End
   Begin VB.TextBox txtDiscadoPrefijo 
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
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1440
      Width           =   795
   End
   Begin VB.TextBox txtNotas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   900
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2280
      Width           =   3615
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   3420
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox txtNombre 
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
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1020
      Width           =   3615
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   12
      Top             =   780
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3300
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1980
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblDiscadoSufijo 
      AutoSize        =   -1  'True
      Caption         =   "Sufijo:"
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
      TabIndex        =   4
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label lblDiscadoPrefijo 
      AutoSize        =   -1  'True
      Caption         =   "Prefijo:"
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
      Top             =   1500
      Width           =   495
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
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
      TabIndex        =   6
      Top             =   2340
      Width           =   465
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre:"
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
      TabIndex        =   0
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos del Tipo de Teléfono"
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
      Left            =   780
      TabIndex        =   11
      Top             =   300
      Width           =   3105
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmTelefonoTipoPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTelefonoTipo As TelefonoTipo
Private mNew As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef TelefonoTipo As TelefonoTipo)
    Set mTelefonoTipo = TelefonoTipo
    mNew = (mTelefonoTipo.IDTelefonoTipo = 0)

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mTelefonoTipo
        txtNombre.Text = .Nombre
        txtDiscadoPrefijo.Text = .DiscadoPrefijo
        txtDiscadoSufijo.Text = .DiscadoSufijo
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
        MsgBox "Debe ingresar el Nombre del Tipo de Teléfono.", vbInformation, App.Title
        txtNombre.SetFocus
        txtNombre_GotFocus
        Exit Sub
    End If
    If txtDiscadoPrefijo.Text <> "" Then
        If Not IsNumeric(txtDiscadoPrefijo.Text) Then
            MsgBox "El Prefijo debe ser un valor numérico.", vbInformation, App.Title
            txtDiscadoPrefijo.SetFocus
            txtDiscadoPrefijo_GotFocus
            Exit Sub
        End If
    End If
    If txtDiscadoSufijo.Text <> "" Then
        If Not IsNumeric(txtDiscadoSufijo.Text) Then
            MsgBox "El Sufijo debe ser un valor numérico.", vbInformation, App.Title
            txtDiscadoSufijo.SetFocus
            txtDiscadoSufijo_GotFocus
            Exit Sub
        End If
    End If
    
    With mTelefonoTipo
        .Nombre = txtNombre.Text
        .DiscadoPrefijo = txtDiscadoPrefijo.Text
        .DiscadoSufijo = txtDiscadoSufijo.Text
        .Notas = txtNotas.Text
        .Activo = (chkActivo.Value = vbChecked)
        If mNew Then
            If Not .AddNew Then
                Exit Sub
            End If
        Else
            If Not .Update Then
                Exit Sub
            End If
        End If
    End With
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mTelefonoTipo = Nothing
    Set frmTelefonoTipoPropiedad = Nothing
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

Private Sub txtDiscadoPrefijo_GotFocus()
    CSM_Control_TextBox.SelAllText txtDiscadoPrefijo
End Sub

Private Sub txtDiscadoPrefijo_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDiscadoPrefijo_LostFocus()
    txtDiscadoPrefijo.Text = CleanNotNumericChars(txtDiscadoPrefijo.Text)
End Sub

Private Sub txtDiscadoSufijo_GotFocus()
    CSM_Control_TextBox.SelAllText txtDiscadoSufijo
End Sub

Private Sub txtDiscadoSufijo_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDiscadoSufijo_LostFocus()
    txtDiscadoSufijo.Text = CleanNotNumericChars(txtDiscadoSufijo.Text)
End Sub
