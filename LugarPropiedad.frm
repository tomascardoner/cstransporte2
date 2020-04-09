VERSION 5.00
Begin VB.Form frmLugarPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LugarPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   4680
   Begin VB.TextBox txtUbicacionLongitud 
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
      Left            =   900
      MaxLength       =   11
      TabIndex        =   5
      Tag             =   "DECIMAL|EMPTY|ZERO|NEGATIVE|999.999999"
      Top             =   1980
      Width           =   1695
   End
   Begin VB.TextBox txtUbicacionLatitud 
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
      Left            =   900
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "DECIMAL|EMPTY|ZERO|NEGATIVE|99.999999"
      Top             =   1500
      Width           =   1695
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
      Top             =   2460
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
      Top             =   3600
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
      Top             =   4140
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
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
      Top             =   4140
      Width           =   1215
   End
   Begin VB.Label lblLongitud 
      AutoSize        =   -1  'True
      Caption         =   "Longitud:"
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
      Top             =   2040
      Width           =   660
   End
   Begin VB.Label lblUbicacionLatitud 
      AutoSize        =   -1  'True
      Caption         =   "Latitud:"
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
      Width           =   525
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
      Top             =   2520
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
      Caption         =   "Ingrese aquí los Datos del Lugar"
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
      Width           =   2325
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "LugarPropiedad.frx":054A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmLugarPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLugar As Lugar
Private mNew As Boolean
Private mKeyDecimal As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Lugar As Lugar)
    Set mLugar = Lugar
    mNew = (mLugar.IDLugar = 0)

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mLugar
        txtNombre.Text = .Nombre
        txtUbicacionLatitud.Text = IIf(.UbicacionLatitud = LOCATION_LATITUDE_NULL_VALUE, "", Format(.UbicacionLatitud, "##.######"))
        txtUbicacionLongitud.Text = IIf(.UbicacionLongitud = LOCATION_LONGITUDE_NULL_VALUE, "", Format(.UbicacionLongitud, "###.######"))
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
        MsgBox "Debe ingresar el Nombre del Lugar.", vbInformation, App.Title
        txtNombre.SetFocus
        txtNombre_GotFocus
        Exit Sub
    End If
    If txtUbicacionLatitud.Text <> "" Then
        If Not IsNumeric(txtUbicacionLatitud.Text) Then
            MsgBox "La Latitud debe ser un valor numérico.", vbInformation, App.Title
            txtUbicacionLatitud.SetFocus
            Exit Sub
        End If
    End If
    If txtUbicacionLongitud.Text <> "" Then
        If Not IsNumeric(txtUbicacionLongitud.Text) Then
            MsgBox "La Longitud debe ser un valor numérico.", vbInformation, App.Title
            txtUbicacionLongitud.SetFocus
            Exit Sub
        End If
    End If
    
    With mLugar
        .Nombre = txtNombre.Text
        .UbicacionLatitud = txtUbicacionLatitud.Text
        .UbicacionLongitud = txtUbicacionLongitud.Text
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
    Set mLugar = Nothing
    Set frmLugarPropiedad = Nothing
End Sub

Private Sub txtNombre_GotFocus()
    CSM_Control_TextBox.SelAllText txtNombre
End Sub

Private Sub txtNombre_Change()
    Caption = "Propiedades" & IIf(Trim(txtNombre.Text) = "", "", " de " & txtNombre.Text)
End Sub

Private Sub txtUbicacionLatitud_GotFocus()
    CSM_Control_TextBox.SelAllText txtUbicacionLatitud
End Sub

Private Sub txtUbicacionLatitud_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtUbicacionLatitud_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtUbicacionLatitud, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtUbicacionLatitud_LostFocus()
    txtUbicacionLatitud.Text = Replace(txtUbicacionLatitud.Text, ".", pRegionalSettings.NumberDecimalSymbol)
    Call CSM_Control_TextBox.FormatValue_ByTag(txtUbicacionLatitud)
End Sub

Private Sub txtUbicacionLongitud_GotFocus()
    CSM_Control_TextBox.SelAllText txtUbicacionLongitud
End Sub

Private Sub txtUbicacionLongitud_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtUbicacionLongitud_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtUbicacionLongitud, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtUbicacionLongitud_LostFocus()
    txtUbicacionLongitud.Text = Replace(txtUbicacionLongitud.Text, ".", pRegionalSettings.NumberDecimalSymbol)
    Call CSM_Control_TextBox.FormatValue_ByTag(txtUbicacionLongitud)
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
End Sub
