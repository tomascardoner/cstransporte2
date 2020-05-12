VERSION 5.00
Begin VB.Form frmViajePropiedadAuditoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auditoría del Viaje"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajePropiedadAuditoria.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCanceladoUsuario 
      BackColor       =   &H8000000F&
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4380
      Width           =   3330
   End
   Begin VB.TextBox txtCanceladoFecha 
      BackColor       =   &H8000000F&
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3960
      Width           =   2250
   End
   Begin VB.TextBox txtCanceladoFechaDiaSemana 
      BackColor       =   &H8000000F&
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1050
   End
   Begin VB.TextBox txtFinalizadoUsuario 
      BackColor       =   &H8000000F&
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3420
      Width           =   3330
   End
   Begin VB.TextBox txtFinalizadoFecha 
      BackColor       =   &H8000000F&
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3000
      Width           =   2250
   End
   Begin VB.TextBox txtFinalizadoFechaDiaSemana 
      BackColor       =   &H8000000F&
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1050
   End
   Begin VB.TextBox txtEnProgresoUsuario 
      BackColor       =   &H8000000F&
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2460
      Width           =   3330
   End
   Begin VB.TextBox txtEnProgresoFecha 
      BackColor       =   &H8000000F&
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2250
   End
   Begin VB.TextBox txtEnProgresoFechaDiaSemana 
      BackColor       =   &H8000000F&
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1050
   End
   Begin VB.TextBox txtModificacionUsuario 
      BackColor       =   &H8000000F&
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1500
      Width           =   3330
   End
   Begin VB.TextBox txtModificacionFecha 
      BackColor       =   &H8000000F&
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2250
   End
   Begin VB.TextBox txtModificacionFechaDiaSemana 
      BackColor       =   &H8000000F&
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1050
   End
   Begin VB.TextBox txtCreacionUsuario 
      BackColor       =   &H8000000F&
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   3330
   End
   Begin VB.TextBox txtCreacionFecha 
      BackColor       =   &H8000000F&
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2250
   End
   Begin VB.TextBox txtCreacionFechaDiaSemana 
      BackColor       =   &H8000000F&
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3540
      TabIndex        =   15
      Top             =   4980
      Width           =   1215
   End
   Begin VB.Label lblCanceladoFecha 
      AutoSize        =   -1  'True
      Caption         =   "Cancelado el:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label lblCanceladoUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Cancelado por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   1080
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4740
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4740
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lblFinalizadoFecha 
      AutoSize        =   -1  'True
      Caption         =   "Finalizado el:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   3060
      Width           =   915
   End
   Begin VB.Label lblFinalizadoUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Finalizado por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   1020
   End
   Begin VB.Label lblEnProgresoFecha 
      AutoSize        =   -1  'True
      Caption         =   "En Progreso el:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   2100
      Width           =   1080
   End
   Begin VB.Label lblEnProgresoUsuario 
      AutoSize        =   -1  'True
      Caption         =   "En Progreso por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   1185
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4740
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4755
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblModificacionFecha 
      AutoSize        =   -1  'True
      Caption         =   "Modificado el:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   19
      Top             =   1140
      Width           =   990
   End
   Begin VB.Label lblModificacionUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Modificado por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   1110
   End
   Begin VB.Label lblCreacionFecha 
      AutoSize        =   -1  'True
      Caption         =   "Creado el:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   17
      Top             =   180
      Width           =   735
   End
   Begin VB.Label lblCreacionUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Creado por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmViajePropiedadAuditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mViaje As Viaje

Public Sub LoadDataAndShow(ByRef Viaje As Viaje)
    Dim Usuario As Usuario

    Set mViaje = Viaje
    
    Load Me
    
    With mViaje
        Set Usuario = New Usuario
        Me.Caption = "Auditoría del Viaje - ID " & .IDViaje_Formatted
    
        'CREACION
        txtCreacionFechaDiaSemana.Text = WeekdayName(Weekday(.FechaHoraCreacion))
        txtCreacionFecha.Text = Format(.FechaHoraCreacion, "Short Date") & " " & Format(.FechaHoraCreacion, "Short Time")
        Usuario.IDUsuario = .IDUsuarioCreacion
        If Usuario.Load() Then
            txtCreacionUsuario.Text = Usuario.Nombre
        End If
        
        'MODIFICACION
        txtModificacionFechaDiaSemana.Text = WeekdayName(Weekday(.FechaHoraModificacion))
        txtModificacionFecha.Text = Format(.FechaHoraModificacion, "Short Date") & " " & Format(.FechaHoraModificacion, "Short Time")
        Usuario.IDUsuario = .IDUsuarioModificacion
        If Usuario.Load() Then
            txtModificacionUsuario.Text = Usuario.Nombre
        End If
        
        'EN PROGRESO
        If .IDUsuarioEnProgreso <> "" Then
            txtEnProgresoFechaDiaSemana.Text = WeekdayName(Weekday(.FechaHoraEnProgreso))
            txtEnProgresoFecha.Text = Format(.FechaHoraEnProgreso, "Short Date") & " " & Format(.FechaHoraEnProgreso, "Short Time")
            Usuario.IDUsuario = .IDUsuarioEnProgreso
            If Usuario.Load() Then
                txtEnProgresoUsuario.Text = Usuario.Nombre
            End If
        End If

        'FINALIZADO
        If .IDUsuarioFinalizado <> "" Then
            txtFinalizadoFechaDiaSemana.Text = WeekdayName(Weekday(.FechaHoraFinalizado))
            txtFinalizadoFecha.Text = Format(.FechaHoraFinalizado, "Short Date") & " " & Format(.FechaHoraFinalizado, "Short Time")
            Usuario.IDUsuario = .IDUsuarioFinalizado
            If Usuario.Load() Then
                txtFinalizadoUsuario.Text = Usuario.Nombre
            End If
        End If

        'CANCELADO
        If .IDUsuarioCancelado <> "" Then
            txtCanceladoFechaDiaSemana.Text = WeekdayName(Weekday(.FechaHoraCancelado))
            txtCanceladoFecha.Text = Format(.FechaHoraCancelado, "Short Date") & " " & Format(.FechaHoraCancelado, "Short Time")
            Usuario.IDUsuario = .IDUsuarioCancelado
            If Usuario.Load() Then
                txtCanceladoUsuario.Text = Usuario.Nombre
            End If
        End If

        Set Usuario = Nothing
    End With
        
    Me.Show vbModal, frmMDI
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mViaje = Nothing
    Set frmViajePropiedadAuditoria = Nothing
End Sub
