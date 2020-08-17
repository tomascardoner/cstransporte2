VERSION 5.00
Begin VB.Form frmViajeDetallePropiedadInformacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajeDetallePropiedadInformacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtIDPersona 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2100
      Width           =   1215
   End
   Begin VB.TextBox txtReservaCodigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1215
   End
   Begin VB.TextBox txtIndice 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1140
      Width           =   1215
   End
   Begin VB.TextBox txtIDViajeDetalle 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   660
      Width           =   1215
   End
   Begin VB.TextBox txtIDViaje 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label lblIDPersona 
      AutoSize        =   -1  'True
      Caption         =   "ID persona:"
      Height          =   225
      Left            =   180
      TabIndex        =   8
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label lblReservaCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Reserva:"
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblIndice 
      AutoSize        =   -1  'True
      Caption         =   "Indice:"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label lblIDViajeDetalle 
      AutoSize        =   -1  'True
      Caption         =   "ID viaje detalle:"
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label lblIDViaje 
      AutoSize        =   -1  'True
      Caption         =   "ID viaje:"
      Height          =   225
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   645
   End
End
Attribute VB_Name = "frmViajeDetallePropiedadInformacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

