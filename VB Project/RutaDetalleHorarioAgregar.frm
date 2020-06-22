VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRutaDetalleHorarioAgregar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RutaDetalleHorarioAgregar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   5040
   Begin VB.Frame fraDiaSemana 
      Caption         =   "Días de la semana:"
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   1860
      Width           =   4755
      Begin VB.CheckBox chkDiaSemana 
         Caption         =   "DiaSemana"
         Height          =   210
         Index           =   7
         Left            =   2460
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkDiaSemana 
         Caption         =   "DiaSemana"
         Height          =   210
         Index           =   6
         Left            =   600
         TabIndex        =   20
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkDiaSemana 
         Caption         =   "DiaSemana"
         Height          =   210
         Index           =   5
         Left            =   600
         TabIndex        =   19
         Top             =   1860
         Width           =   1215
      End
      Begin VB.CheckBox chkDiaSemana 
         Caption         =   "DiaSemana"
         Height          =   210
         Index           =   4
         Left            =   600
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox chkDiaSemana 
         Caption         =   "DiaSemana"
         Height          =   210
         Index           =   3
         Left            =   600
         TabIndex        =   17
         Top             =   1260
         Width           =   1215
      End
      Begin VB.CheckBox chkDiaSemana 
         Caption         =   "DiaSemana"
         Height          =   210
         Index           =   2
         Left            =   600
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkDiaSemana 
         Caption         =   "DiaSemana"
         Height          =   210
         Index           =   1
         Left            =   2460
         TabIndex        =   15
         Top             =   1260
         Width           =   1215
      End
      Begin VB.OptionButton optDiaSemanaAlgunos 
         Caption         =   "Algunos:"
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1155
      End
      Begin VB.OptionButton optDiaSemanaTodos 
         Caption         =   "Todos"
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.TextBox txtLugar 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1380
      Width           =   3630
   End
   Begin VB.TextBox txtRuta 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2340
      TabIndex        =   2
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3660
      TabIndex        =   3
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4755
   End
   Begin MSComCtl2.DTPicker dtpHoraInicio 
      Height          =   315
      Left            =   1260
      TabIndex        =   8
      Top             =   4620
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
      Left            =   1260
      TabIndex        =   9
      Top             =   5040
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
   Begin VB.Label lblHoraInicio 
      AutoSize        =   -1  'True
      Caption         =   "Hora de inicio:"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   1020
   End
   Begin VB.Label lblHoraFin 
      AutoSize        =   -1  'True
      Caption         =   "Hora de fin:"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   5100
      Width           =   840
   End
   Begin VB.Label lblLugar 
      AutoSize        =   -1  'True
      Caption         =   "Lugar:"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image imgIcon2 
      Height          =   480
      Left            =   180
      Picture         =   "RutaDetalleHorarioAgregar.frx":054A
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Datos de los Horarios del Detalle de Ruta"
      Height          =   210
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   2940
   End
End
Attribute VB_Name = "frmRutaDetalleHorarioAgregar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIDRuta As String
Private mIDLugar As Long

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByVal IDRuta As String, ByVal IDLugar As Integer, ByVal LugarNombre As String)
    mIDRuta = IDRuta
    mIDLugar = IDLugar
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
        
    txtRuta.Text = IDRuta
    txtLugar.Text = LugarNombre
    
    optDiaSemanaTodos.value = True
    
    dtpHoraInicio.value = 0
    dtpHoraFin.value = 0
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub Form_Load()
    Dim Index As Byte
    
    For Index = 1 To 7
        chkDiaSemana(Index).Caption = WeekdayName(Index)
    Next Index
End Sub

Private Sub optDiaSemanaTodos_Click()
    MakeCheckBoxesVisible
End Sub

Private Sub optDiaSemanaAlgunos_Click()
    MakeCheckBoxesVisible
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim Index As Long
    Dim RutaDetalleHorario As RutaDetalleHorario
    
    If optDiaSemanaTodos.value = False And optDiaSemanaAlgunos.value = False Then
        MsgBox "Debe especificar alguna de las opciones de Días de Semana.", vbInformation, App.Title
        Exit Sub
    End If
    If optDiaSemanaAlgunos.value Then
        For Index = 1 To 7
            If chkDiaSemana(Index).value = vbChecked Then
                Exit For
            End If
            If Index = 7 Then
                MsgBox "Debe seleccionar alguno de los Días de Semana.", vbInformation, App.Title
                Exit Sub
            End If
        Next Index
    End If
    If dtpHoraFin.value < dtpHoraInicio.value Then
        MsgBox "La Hora de fin debe ser mayor a la Hora de inicio.", vbInformation, App.Title
        dtpHoraFin.SetFocus
        Exit Sub
    End If
    
    If optDiaSemanaTodos.value Then
        Set RutaDetalleHorario = New RutaDetalleHorario
        With RutaDetalleHorario
            .IDRuta = mIDRuta
            .IDLugar = mIDLugar
            .DiaSemanaNumero = 0
            .DiaSemana = CSM_Constant.ITEM_ALL_MALE
            .HoraInicio = dtpHoraInicio.value
            .HoraFin = dtpHoraFin.value
            If Not .Update() Then
                Exit Sub
            End If
        End With
        Set RutaDetalleHorario = Nothing
    Else
        For Index = 1 To 7
            If chkDiaSemana(Index).value = vbChecked Then
                Set RutaDetalleHorario = New RutaDetalleHorario
                With RutaDetalleHorario
                    .RefreshListSkip = True
                    .IDRuta = mIDRuta
                    .IDLugar = mIDLugar
                    .DiaSemanaNumero = Index
                    .DiaSemana = WeekdayName(Index)
                    .HoraInicio = dtpHoraInicio.value
                    .HoraFin = dtpHoraFin.value
                    If Not .Update() Then
                        Exit Sub
                    End If
                End With
                Set RutaDetalleHorario = Nothing
            End If
        Next Index
        RefreshList_RefreshRutaDetalleHorario mIDRuta, mIDLugar, 0
    End If
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRutaDetalleHorarioPropiedad = Nothing
End Sub

Private Sub MakeCheckBoxesVisible()
    Dim Index As Byte
    
    For Index = 1 To 7
        chkDiaSemana(Index).Visible = optDiaSemanaAlgunos.value
    Next Index
End Sub
