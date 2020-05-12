VERSION 5.00
Begin VB.Form frmViajeDetalleConfirmacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "¿Confirma que desea guardar los cambios de esta Reserva?"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5145
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5145
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCombinacion 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2715
   End
   Begin VB.TextBox txtRecibe 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2220
      Width           =   4035
   End
   Begin VB.TextBox txtPasajeroEnvia 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1800
      Width           =   4035
   End
   Begin VB.TextBox txtRuta 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1380
      Width           =   4035
   End
   Begin VB.TextBox txtFecha 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   1215
   End
   Begin VB.TextBox txtHora 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtDia 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   540
      Width           =   1215
   End
   Begin VB.TextBox txtTipo 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "No"
      Height          =   435
      Left            =   2640
      TabIndex        =   16
      Top             =   3180
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Sí"
      Default         =   -1  'True
      Height          =   435
      Left            =   1380
      TabIndex        =   15
      Top             =   3180
      Width           =   1155
   End
   Begin VB.Label lblCombinacion 
      AutoSize        =   -1  'True
      Caption         =   "En combinación con el viaje:"
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   2700
      Width           =   2025
   End
   Begin VB.Label lblRecibe 
      AutoSize        =   -1  'True
      Caption         =   "Recibe:"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label lblPasajeroEnvia 
      AutoSize        =   -1  'True
      Caption         =   "Pasajero:"
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   1860
      Width           =   675
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "Hora:"
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   390
   End
   Begin VB.Label lblDia 
      AutoSize        =   -1  'True
      Caption         =   "Día:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   270
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   345
   End
End
Attribute VB_Name = "frmViajeDetalleConfirmacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Me.Tag = "OK"
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    Me.Tag = "CANCEL"
    Me.Hide
End Sub

