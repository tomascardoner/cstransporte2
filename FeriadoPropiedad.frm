VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFeriadoPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   2640
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
   Icon            =   "FeriadoPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2640
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   3600
      Picture         =   "FeriadoPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   900
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1500
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
      TabIndex        =   7
      Top             =   780
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3300
      TabIndex        =   5
      Top             =   2100
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   4
      Top             =   2100
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   1020
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
      MaxDate         =   73050
      MinDate         =   36526
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "&Fecha:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   600
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos del Feriado"
      Height          =   210
      Left            =   780
      TabIndex        =   6
      Top             =   300
      Width           =   2445
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmFeriadoPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFeriado As Feriado
Private mNew As Boolean


Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mFeriado
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Feriado As Feriado)
    Set mFeriado = Feriado
    mNew = (mFeriado.Fecha = DATE_TIME_FIELD_NULL_VALUE)

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mFeriado
        If mNew Then
            dtpFecha.Value = Date
        Else
            dtpFecha.Value = .Fecha
        End If
        dtpFecha_Change
        txtNombre.Text = .Nombre
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
    With mFeriado
        .Fecha = dtpFecha.Value
        .Nombre = txtNombre.Text
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

Private Sub dtpFecha_Change()
    Caption = "Propiedades del Feriado: " & dtpFecha.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mFeriado = Nothing
    Set frmFeriadoPropiedad = Nothing
End Sub
