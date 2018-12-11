VERSION 5.00
Begin VB.Form frmVehiculoPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10170
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VehiculoPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   10170
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   9060
      Picture         =   "VehiculoPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdTelefonoDial 
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
      Left            =   8160
      Picture         =   "VehiculoPropiedad.frx":0B74
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Llamar"
      Top             =   1680
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtTelefonoNumero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7020
      MaxLength       =   16
      TabIndex        =   20
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtTelefonoArea 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6420
      MaxLength       =   5
      TabIndex        =   19
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtKilometrajeEstimado 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   960
      MaxLength       =   7
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtAnio 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   960
      MaxLength       =   4
      TabIndex        =   11
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtPasajero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6420
      MaxLength       =   3
      TabIndex        =   17
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtAsiento 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6420
      MaxLength       =   3
      TabIndex        =   15
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtDominio 
      Height          =   315
      Left            =   960
      MaxLength       =   7
      TabIndex        =   9
      Top             =   2400
      Width           =   1035
   End
   Begin VB.TextBox txtColor 
      Height          =   315
      Left            =   960
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox txtModelo 
      Height          =   315
      Left            =   960
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox txtNotas 
      Height          =   945
      Left            =   6420
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox txtMarca 
      Height          =   315
      Left            =   960
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo"
      Height          =   210
      Left            =   4800
      TabIndex        =   24
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   960
      MaxLength       =   50
      TabIndex        =   1
      Top             =   960
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
      TabIndex        =   28
      Top             =   780
      Width           =   9915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8820
      TabIndex        =   26
      Top             =   3660
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7500
      TabIndex        =   25
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Label lblKilometrajeEstimado 
      AutoSize        =   -1  'True
      Caption         =   "Kms. Est.:"
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   3180
      Width           =   720
   End
   Begin VB.Label lblAnio 
      AutoSize        =   -1  'True
      Caption         =   "Año:"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   2820
      Width           =   345
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   4680
      Y1              =   960
      Y2              =   4020
   End
   Begin VB.Label lblTelefono 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono:"
      Height          =   210
      Left            =   4800
      TabIndex        =   18
      Top             =   1740
      Width           =   675
   End
   Begin VB.Label lblPasajero 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad Pasajeros:"
      Height          =   210
      Left            =   4800
      TabIndex        =   16
      Top             =   1380
      Width           =   1440
   End
   Begin VB.Label lblAsiento 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad Asientos:"
      Height          =   210
      Left            =   4800
      TabIndex        =   14
      Top             =   1020
      Width           =   1365
   End
   Begin VB.Label lblDominio 
      AutoSize        =   -1  'True
      Caption         =   "Dominio:"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   2460
      Width           =   600
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      Caption         =   "&Color:"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   2100
      Width           =   420
   End
   Begin VB.Label lblModelo 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1740
      Width           =   555
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   210
      Left            =   4800
      TabIndex        =   22
      Top             =   2220
      Width           =   465
   End
   Begin VB.Label lblMarca 
      AutoSize        =   -1  'True
      Caption         =   "&Marca:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1380
      Width           =   495
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   600
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos del Vehículo"
      Height          =   210
      Left            =   780
      TabIndex        =   27
      Top             =   300
      Width           =   2535
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "VehiculoPropiedad.frx":1186
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmVehiculoPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mVehiculo As Vehiculo
Private mNew As Boolean

Private mLoading As Boolean


Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mVehiculo
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Vehiculo As Vehiculo)
    Set mVehiculo = Vehiculo
    mNew = (mVehiculo.IDVehiculo = 0)
    
    mLoading = True

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mVehiculo
        If .IDVehiculo = 0 Then
            txtNombre.Text = ""
            txtMarca.Text = ""
            txtModelo.Text = ""
            txtColor.Text = ""
            txtDominio.Text = ""
            txtAnio.Text = ""
            txtKilometrajeEstimado.Text = ""
            txtAsiento.Text = ""
            txtPasajero.Text = ""
            txtTelefonoArea.Text = ""
            txtTelefonoNumero.Text = ""
            txtNotas.Text = ""
            chkActivo.Value = vbChecked
        Else
            txtNombre.Text = .Nombre
            txtMarca.Text = .Marca
            txtModelo.Text = .Modelo
            txtColor.Text = .Color
            txtDominio.Text = .Dominio
            txtAnio.Text = IIf(.Anio = 0, "", .Anio)
            txtKilometrajeEstimado.Text = IIf(.KilometrajeEstimado = 0, "", .KilometrajeEstimado)
            txtAsiento.Text = .Asiento
            txtPasajero.Text = .Pasajero
            txtTelefonoArea.Text = .TelefonoArea
            txtTelefonoNumero.Text = .TelefonoNumero
            txtNotas = .Notas
            chkActivo = IIf(.Activo, vbChecked, vbUnchecked)
        End If
    End With
    
    mLoading = False
    
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
        MsgBox "Debe ingresar el Nombre del Vehículo.", vbInformation, App.Title
        txtNombre.SetFocus
        Exit Sub
    End If
    If Trim(txtAsiento.Text) = "" Then
        MsgBox "Debe ingresar la Cantidad de Asientos del Vehículo.", vbInformation, App.Title
        txtAsiento.SetFocus
        Exit Sub
    End If
    If Trim(txtPasajero.Text) = "" Then
        MsgBox "Debe ingresar la Cantidad de Pasajeros del Vehículo.", vbInformation, App.Title
        txtPasajero.SetFocus
        Exit Sub
    End If
    If Val(txtPasajero.Text) < Val(txtAsiento.Text) Then
        MsgBox "La Cantidad de Pasajeros del Vehículo debe ser mayor o igual a la Cantidad de Asientos del Vehículo.", vbInformation, App.Title
        txtPasajero.SetFocus
        txtPasajero_GotFocus
        Exit Sub
    End If
    If txtKilometrajeEstimado.Text <> "" Then
        If Not IsNumeric(txtKilometrajeEstimado.Text) Then
            MsgBox "El Kilometraje Estimado debe ser un valor numérico.", vbInformation, App.Title
            txtKilometrajeEstimado.SetFocus
            Exit Sub
        End If
    End If
    
    With mVehiculo
        .Nombre = txtNombre.Text
        .Marca = txtMarca.Text
        .Modelo = txtModelo.Text
        .Color = txtColor.Text
        .Dominio = txtDominio.Text
        .Anio = Int(Val(txtAnio.Text))
        .KilometrajeEstimado = Val(txtKilometrajeEstimado.Text)
        .Asiento = Int(Val(txtAsiento.Text))
        .Pasajero = Int(Val(txtPasajero.Text))
        .TelefonoArea = txtTelefonoArea.Text
        .TelefonoNumero = txtTelefonoNumero.Text
        .Notas = txtNotas.Text
        .Activo = (chkActivo.Value = vbChecked)
        If Not .Update() Then
            Exit Sub
        End If
    End With
    
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mVehiculo = Nothing
    Set frmVehiculoPropiedad = Nothing
End Sub

Private Sub txtAnio_GotFocus()
    CSM_Control_TextBox.SelAllText txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtAnio_LostFocus()
    txtAnio.Text = Val(txtAnio.Text)
    If txtAnio.Text = 0 Then
        txtAnio.Text = ""
    End If
End Sub

Private Sub txtAsiento_GotFocus()
    CSM_Control_TextBox.SelAllText txtAsiento
End Sub

Private Sub txtAsiento_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtAsiento_LostFocus()
    txtAsiento.Text = Abs(Val(txtAsiento.Text))
    If txtAsiento.Text = 0 Then
        txtAsiento.Text = ""
    End If
End Sub

Private Sub txtPasajero_GotFocus()
    CSM_Control_TextBox.SelAllText txtPasajero
End Sub

Private Sub txtPasajero_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPasajero_LostFocus()
    txtPasajero.Text = Abs(Val(txtPasajero.Text))
    If txtPasajero.Text = 0 Then
        txtPasajero.Text = ""
    End If
End Sub

Private Sub txtColor_GotFocus()
    CSM_Control_TextBox.SelAllText txtColor
End Sub

Private Sub txtDominio_GotFocus()
    CSM_Control_TextBox.SelAllText txtDominio
End Sub

Private Sub txtKilometrajeEstimado_GotFocus()
    CSM_Control_TextBox.SelAllText txtKilometrajeEstimado
End Sub

Private Sub txtKilometrajeEstimado_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtKilometrajeEstimado_LostFocus()
    txtKilometrajeEstimado.Text = Val(txtKilometrajeEstimado.Text)
End Sub

Private Sub txtMarca_GotFocus()
    CSM_Control_TextBox.SelAllText txtMarca
End Sub

Private Sub txtModelo_GotFocus()
    CSM_Control_TextBox.SelAllText txtModelo
End Sub

Private Sub txtNombre_Change()
    Caption = "Propiedades" & IIf(Trim(txtNombre.Text) = "", "", " de " & Trim(txtNombre.Text))
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

Private Sub txtTelefonoArea_GotFocus()
    CSM_Control_TextBox.SelAllText txtTelefonoArea
End Sub

Private Sub txtTelefonoArea_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTelefonoArea_LostFocus()
    txtTelefonoArea.Text = CleanNotNumericChars(txtTelefonoArea.Text)
End Sub

Private Sub txtTelefonoNumero_GotFocus()
    CSM_Control_TextBox.SelAllText txtTelefonoNumero
End Sub

Private Sub txtTelefonoNumero_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTelefonoNumero_Change()
    If pTelephony.TelephonyType <> "NONE" Then
        cmdTelefonoDial.Visible = (Trim(txtTelefonoNumero.Text) <> "" And pTelephony.Initialized)
    Else
        cmdTelefonoDial.Visible = False
    End If
    If pTelephony.Initialized And Not mLoading Then
        If Trim(txtTelefonoNumero.Text) = "" Then
            If Trim(txtTelefonoArea.Text) = pTelephony.LocationCityCode Then
                txtTelefonoArea.Text = ""
            End If
        Else
            If Trim(txtTelefonoArea.Text) = "" Then
                txtTelefonoArea.Text = pTelephony.LocationCityCode
            End If
        End If
    End If
End Sub

Private Sub txtTelefonoNumero_LostFocus()
    txtTelefonoNumero.Text = CleanNotNumericChars(txtTelefonoNumero.Text)
End Sub

Private Sub cmdTelefonoDial_Click()
    Dim TelefonoTipo As TelefonoTipo
    
    If pTelephony.TelephonyType <> "NONE" And pTelephony.Initialized Then
        Set TelefonoTipo = New TelefonoTipo
        TelefonoTipo.IDTelefonoTipo = pParametro.Vehiculo_TelefonoTipo_ID
        If TelefonoTipo.Load() Then
            Call pTelephony.DialNumber(txtTelefonoArea.Text, TelefonoTipo.DiscadoPrefijo & txtTelefonoNumero.Text & TelefonoTipo.DiscadoSufijo)
        End If
        Set TelefonoTipo = Nothing
    End If
End Sub
