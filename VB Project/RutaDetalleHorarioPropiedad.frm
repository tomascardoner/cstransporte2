VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRutaDetalleHorarioPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RutaDetalleHorarioPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   5325
   Begin VB.ComboBox cboDiaSemana 
      Height          =   330
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   1635
   End
   Begin VB.TextBox txtLugar 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1380
      Width           =   3630
   End
   Begin VB.TextBox txtRuta 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3420
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   3420
      Width           =   1215
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker dtpHoraInicio 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   2340
      Width           =   975
      _ExtentX        =   1720
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
      CustomFormat    =   "HH:mm"
      Format          =   108593155
      UpDown          =   -1  'True
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpHoraFin 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
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
      CustomFormat    =   "HH:mm"
      Format          =   108593155
      UpDown          =   -1  'True
      CurrentDate     =   36494
   End
   Begin VB.Label lblDiaSemana 
      AutoSize        =   -1  'True
      Caption         =   "Día de la semana:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1980
      Width           =   1275
   End
   Begin VB.Label lblHoraInicio 
      AutoSize        =   -1  'True
      Caption         =   "Hora de inicio:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1020
   End
   Begin VB.Label lblHoraFin 
      AutoSize        =   -1  'True
      Caption         =   "Hora de fin:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   2820
      Width           =   840
   End
   Begin VB.Label lblLugar 
      AutoSize        =   -1  'True
      Caption         =   "Lugar:"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image imgIcon2 
      Height          =   480
      Left            =   180
      Picture         =   "RutaDetalleHorarioPropiedad.frx":054A
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Datos del Horario del Detalle de Ruta"
      Height          =   210
      Left            =   840
      TabIndex        =   13
      Top             =   240
      Width           =   2625
   End
End
Attribute VB_Name = "frmRutaDetalleHorarioPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRutaDetalleHorario As RutaDetalleHorario

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef RutaDetalleHorario As RutaDetalleHorario, ByVal LugarNombre As String)
    Set mRutaDetalleHorario = RutaDetalleHorario
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
        
    With RutaDetalleHorario
        txtRuta.Text = .IDRuta
        txtLugar.Text = LugarNombre
        
        cboDiaSemana.ListIndex = .DiaSemanaNumero
        dtpHoraInicio.value = .HoraInicio
        dtpHoraFin.value = .HoraFin
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
    If cboDiaSemana.ListIndex = -1 Then
        MsgBox "Debe seleccionar el día de la semana.", vbInformation, App.Title
        cboDiaSemana.SetFocus
        Exit Sub
    End If
    If dtpHoraFin.value < dtpHoraInicio.value Then
        MsgBox "La Hora de fin debe ser mayor a la Hora de inicio.", vbInformation, App.Title
        dtpHoraFin.SetFocus
        Exit Sub
    End If
    
    With mRutaDetalleHorario
        .DiaSemanaNumero = cboDiaSemana.ListIndex
        .DiaSemana = cboDiaSemana.Text
        .HoraInicio = dtpHoraInicio.value
        .HoraFin = dtpHoraFin.value
        If Not .Update() Then
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Index As Byte
    
    cboDiaSemana.AddItem CSM_Constant.ITEM_ALL_MALE
    For Index = 1 To 7
        cboDiaSemana.AddItem WeekdayName(Index)
    Next Index
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mRutaDetalleHorario = Nothing
    Set frmRutaDetalleHorarioPropiedad = Nothing
End Sub
