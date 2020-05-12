VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAlarmaPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "AlarmaPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   3660
      Picture         =   "AlarmaPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   60
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
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1020
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
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   3540
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox txtPreaviso 
      Alignment       =   1  'Right Justify
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
      Left            =   1020
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1860
      Width           =   1035
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
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2460
      Width           =   3615
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   12
      Top             =   780
      Width           =   4515
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
      Left            =   3420
      TabIndex        =   10
      Top             =   4140
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
      Left            =   2100
      TabIndex        =   9
      Top             =   4140
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   1020
      TabIndex        =   3
      Top             =   1440
      Width           =   1635
      _ExtentX        =   2884
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
      Format          =   16777217
      CurrentDate     =   36950
   End
   Begin VB.Label lblPreavisoUnidad 
      AutoSize        =   -1  'True
      Caption         =   "días antes."
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
      Left            =   2220
      TabIndex        =   13
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label lblPreaviso 
      AutoSize        =   -1  'True
      Caption         =   "Aviso:"
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
      Width           =   465
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
      Left            =   120
      TabIndex        =   2
      Top             =   1500
      Width           =   495
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
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
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Datos de la Alarma General"
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
      Width           =   1980
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "AlarmaPropiedad.frx":0B74
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAlarmaPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAlarma As Alarma
Private mNew As Boolean


Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mAlarma
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Alarma As Alarma)
    Set mAlarma = Alarma
    mNew = (mAlarma.IDAlarma = 0)
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mAlarma
        txtNombre.Text = .Nombre
        If mNew Then
            dtpFecha.Value = Date
        Else
            dtpFecha.Value = .Fecha
        End If
        txtPreaviso.Text = .Preaviso
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
        MsgBox "Debe ingresar el Nombre.", vbInformation, App.Title
        txtNombre.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtPreaviso.Text) Then
        MsgBox "El Aviso debe ser un valor numérico.", vbInformation, App.Title
        txtPreaviso.SetFocus
        Exit Sub
    End If
    If CLng(txtPreaviso.Text) < 0 Then
        MsgBox "El Aviso debe ser mayor o igual a cero.", vbInformation, App.Title
        txtPreaviso.SetFocus
        Exit Sub
    End If
    
    With mAlarma
        .Nombre = txtNombre.Text
        .Fecha = dtpFecha.Value
        .Preaviso = CLng(txtPreaviso.Text)
        .Notas = txtNotas.Text
        .Activo = (chkActivo.Value = vbChecked)
        If mNew Then
            If Not .AddNew() Then
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
    Set mAlarma = Nothing
    Set frmAlarmaPropiedad = Nothing
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

Private Sub txtPreaviso_GotFocus()
    CSM_Control_TextBox.SelAllText txtPreaviso
End Sub

Private Sub txtPreaviso_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPreaviso_LostFocus()
    txtPreaviso.Text = Val(txtPreaviso.Text)
End Sub

