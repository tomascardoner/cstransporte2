VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPersonaBuscar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar persona"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6465
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PersonaBuscar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNombre 
      Height          =   360
      Left            =   2040
      MaxLength       =   100
      TabIndex        =   7
      Top             =   1830
      Width           =   4155
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   435
      Left            =   5100
      TabIndex        =   9
      Top             =   2700
      Width           =   1155
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   435
      Left            =   3720
      TabIndex        =   8
      Top             =   2700
      Width           =   1155
   End
   Begin VB.TextBox txtApellido 
      Height          =   360
      Left            =   2040
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1320
      Width           =   4155
   End
   Begin VB.OptionButton optApellidoNombre 
      Caption         =   "por apellido y nombre:"
      Height          =   300
      Left            =   180
      TabIndex        =   3
      Top             =   900
      Width           =   2355
   End
   Begin VB.TextBox txtDocumentoNumero 
      Height          =   360
      Left            =   3120
      MaxLength       =   15
      TabIndex        =   0
      Top             =   210
      Width           =   1635
   End
   Begin VB.OptionButton optDocumento 
      Caption         =   "por documento:"
      Height          =   300
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Value           =   -1  'True
      Width           =   1755
   End
   Begin MSDataListLib.DataCombo datcboDocumentoTipo 
      Height          =   330
      Left            =   2040
      TabIndex        =   2
      Top             =   210
      Width           =   1035
      _ExtentX        =   1826
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
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   240
      Left            =   1080
      TabIndex        =   6
      Top             =   1890
      Width           =   735
   End
   Begin VB.Label lblApellido 
      AutoSize        =   -1  'True
      Caption         =   "Apellido:"
      Height          =   240
      Left            =   1080
      TabIndex        =   4
      Top             =   1380
      Width           =   750
   End
End
Attribute VB_Name = "frmPersonaBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim skipFocusChange As Boolean

Private Sub Form_Load()
    Call CSM_Control_DataCombo.FillFromSQL(datcboDocumentoTipo, "SELECT IDDocumentoTipo, Nombre, 2 AS Orden FROM DocumentoTipo WHERE Activo = 1 ORDER BY Nombre", "IDDocumentoTipo", "Nombre", "Tipos de Documento", cscpItemOrFirst, 4)
End Sub

Private Sub optDocumento_Click()
    txtDocumentoNumero.SetFocus
End Sub

Private Sub txtDocumentoNumero_GotFocus()
    Call CSM_Control_TextBox.SelAllText(txtDocumentoNumero)
End Sub

Private Sub optApellidoNombre_Click()
    If Not skipFocusChange Then
        txtApellido.SetFocus
    End If
End Sub

Private Sub txtApellido_GotFocus()
    Call CSM_Control_TextBox.SelAllText(txtApellido)
End Sub

Private Sub txtApellido_Change()
    skipFocusChange = True
    optApellidoNombre.value = True
    skipFocusChange = False
End Sub

Private Sub txtNombre_GotFocus()
    Call CSM_Control_TextBox.SelAllText(txtNombre)
End Sub

Private Sub txtNombre_Change()
    skipFocusChange = True
    optApellidoNombre.value = True
    skipFocusChange = False
End Sub

Private Sub cmdBuscar_Click()
    If optDocumento.value = False And optApellidoNombre.value = False Then
        MsgBox "Debe seleccionar una de las opciones de búsqueda.", vbInformation, App.Title
    End If
    If optDocumento.value Then
        If Val(datcboDocumentoTipo.BoundText) = 0 Then
            MsgBox "Debe seleccionar el tipo de documento.", vbInformation, App.Title
            datcboDocumentoTipo.SetFocus
            Exit Sub
        End If
        If Len(Trim(txtDocumentoNumero.Text)) < 6 Then
            MsgBox "El número de documento debe tener al menos 6 caracteres.", vbInformation, App.Title
            txtDocumentoNumero.SetFocus
            Exit Sub
        End If
    ElseIf optApellidoNombre.value Then
        If Len(Trim(txtApellido.Text)) < 4 And Len(Trim(txtNombre.Text)) = 0 Then
            MsgBox "El apellido debe tener al menos 4 caracteres.", vbInformation, App.Title
            txtApellido.SetFocus
            Exit Sub
        End If
    End If

    Tag = "OK"
    Me.Hide
End Sub

Private Sub cmdCerrar_Click()
    Tag = "CANCEL"
    Me.Hide
End Sub
