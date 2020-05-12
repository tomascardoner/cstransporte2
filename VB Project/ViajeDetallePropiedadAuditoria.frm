VERSION 5.00
Begin VB.Form frmViajeDetallePropiedadAuditoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auditoría del Detalle de Viaje"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajeDetallePropiedadAuditoria.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4980
      Width           =   1215
   End
   Begin VB.Frame fraCancelado 
      Caption         =   "Cancelación:"
      Height          =   2055
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4935
      Begin VB.CheckBox chkCanceladoForzarDebito 
         Height          =   210
         Left            =   1380
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1620
         Width           =   195
      End
      Begin VB.TextBox txtCanceladoPor 
         Height          =   315
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1140
         Width           =   3375
      End
      Begin VB.TextBox txtCanceladoHora 
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   720
         Width           =   1170
      End
      Begin VB.TextBox txtCanceladoFecha 
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   300
         Width           =   1410
      End
      Begin VB.TextBox txtCanceladoFechaDiaSemana 
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   300
         Width           =   1050
      End
      Begin VB.Label lblCanceladoForzarDebito 
         AutoSize        =   -1  'True
         Caption         =   "Debitar Viaje:"
         Height          =   210
         Left            =   180
         TabIndex        =   2
         Top             =   1620
         Width           =   960
      End
      Begin VB.Label lblCanceladoPor 
         AutoSize        =   -1  'True
         Caption         =   "Por:"
         Height          =   210
         Left            =   180
         TabIndex        =   0
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label lblCanceladoHora 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   210
         Left            =   180
         TabIndex        =   22
         Top             =   780
         Width           =   390
      End
      Begin VB.Label lblCanceladoFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   210
         Left            =   180
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraAudit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   4935
      Begin VB.TextBox txtCancelacionUsuario 
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1980
         Width           =   3330
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   300
         Width           =   1050
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
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   300
         Width           =   2250
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   720
         Width           =   3330
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1050
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
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1140
         Width           =   2250
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1560
         Width           =   3330
      End
      Begin VB.Label lblCancelacionUsuario 
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
         Left            =   180
         TabIndex        =   24
         Top             =   2040
         Width           =   1080
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
         Left            =   180
         TabIndex        =   16
         Top             =   780
         Width           =   855
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
         Left            =   180
         TabIndex        =   15
         Top             =   360
         Width           =   735
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
         Left            =   180
         TabIndex        =   14
         Top             =   1620
         Width           =   1110
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
         Left            =   180
         TabIndex        =   13
         Top             =   1200
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmViajeDetallePropiedadAuditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mViajeDetalle As ViajeDetalle

Public Sub LoadDataAndShow(ByRef ViajeDetalle As ViajeDetalle)
    Dim Usuario As Usuario

    Set mViajeDetalle = ViajeDetalle
    
    Load Me
    
    With mViajeDetalle
        Me.Caption = "Auditoría del Detalle de Viaje - ID " & .IDViajeDetalle_Formatted
    
        If .Estado = VIAJE_DETALLE_ESTADO_CANCELADO Then
            txtCanceladoFechaDiaSemana.Text = WeekdayName(Weekday(.CanceladoFechaHora))
            txtCanceladoFecha.Text = .CanceladoFechaHora_FormattedAsDate
            txtCanceladoHora.Text = .CanceladoFechaHora_FormattedAsTime
            txtCanceladoPor.Text = .CanceladoPor
        End If
        
        chkCanceladoForzarDebito.Value = IIf(.ForzarDebito, vbChecked, vbUnchecked)
        txtCreacionFechaDiaSemana.Text = WeekdayName(Weekday(.FechaHoraCreacion))
        txtCreacionFecha.Text = .FechaHoraCreacion_Formatted
    
        Set Usuario = New Usuario
        Usuario.IDUsuario = .IDUsuarioCreacion
        If Usuario.Load() Then
            txtCreacionUsuario.Text = Usuario.Nombre
        End If
        txtModificacionFechaDiaSemana.Text = WeekdayName(Weekday(.FechaHoraModificacion))
        txtModificacionFecha.Text = .FechaHoraModificacion_Formatted
    
        Usuario.IDUsuario = .IDUsuarioModificacion
        If Usuario.Load() Then
            txtModificacionUsuario.Text = Usuario.Nombre
        End If
        
        If .IDUsuarioCancelacion <> "" Then
            Usuario.IDUsuario = .IDUsuarioCancelacion
            If Usuario.Load() Then
                txtCancelacionUsuario.Text = Usuario.Nombre
            End If
        End If
        
        Set Usuario = Nothing
    End With
    
    If mViajeDetalle.Estado <> VIAJE_DETALLE_ESTADO_CANCELADO Or frmViajeDetallePropiedad.cmdOK.Visible = False Then
        txtCanceladoPor.Enabled = False
        chkCanceladoForzarDebito.Enabled = False
        cmdOK.Visible = False
        cmdCancel.Caption = "Cerrar"
    End If
    
    Me.Show vbModal, frmMDI
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtCanceladoPor.Enabled Then
        If Trim(txtCanceladoPor.Text) = "" Then
            MsgBox "Debe especificar quién Canceló la Reserva.", vbInformation, App.Title
            txtCanceladoPor.SetFocus
            Exit Sub
        End If
    
        mViajeDetalle.CanceladoPor = txtCanceladoPor.Text
        mViajeDetalle.ForzarDebito = (chkCanceladoForzarDebito.Value = vbChecked)
    End If
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmViajeDetallePropiedadAuditoria = Nothing
End Sub

Private Sub txtCanceladoPor_GotFocus()
    CSM_Control_TextBox.SelAllText txtCanceladoPor
End Sub
