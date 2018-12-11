VERSION 5.00
Begin VB.Form frmAuditoriaGenerico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auditoría"
   ClientHeight    =   3105
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
   Icon            =   "AuditoriaGenerico.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   2580
      Width           =   1215
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
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   5
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1380
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
         Top             =   1380
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
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1800
         Width           =   3330
      End
      Begin VB.Line Line1 
         X1              =   180
         X2              =   4680
         Y1              =   1200
         Y2              =   1200
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
         TabIndex        =   4
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
         TabIndex        =   1
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
         TabIndex        =   9
         Top             =   1860
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
         TabIndex        =   6
         Top             =   1440
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmAuditoriaGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadDataAndShow(ByVal Objeto As Object)
    Dim Usuario As Usuario

    Load Me
    
    With Objeto
        If .IDUsuarioCreacion <> "" Then
            Set Usuario = New Usuario
            Usuario.NoMatchRaiseError = False
            
            txtCreacionFechaDiaSemana.Text = WeekdayName(Weekday(.FechaHoraCreacion))
            txtCreacionFecha.Text = Format(.FechaHoraCreacion, "Short Date") & " " & Format(.FechaHoraCreacion, "Short Time")
            Usuario.IDUsuario = .IDUsuarioCreacion
            If Usuario.Load() Then
                If Usuario.NoMatch Then
                    txtCreacionUsuario.Text = .IDUsuarioCreacion
                Else
                    txtCreacionUsuario.Text = Usuario.Nombre
                End If
            End If
            
            txtModificacionFechaDiaSemana.Text = WeekdayName(Weekday(.FechaHoraModificacion))
            txtModificacionFecha.Text = Format(.FechaHoraModificacion, "Short Date") & " " & Format(.FechaHoraModificacion, "Short Time")
            Usuario.IDUsuario = .IDUsuarioModificacion
            If Usuario.Load() Then
                If Usuario.NoMatch Then
                    txtModificacionUsuario.Text = .IDUsuarioModificacion
                Else
                    txtModificacionUsuario.Text = Usuario.Nombre
                End If
            End If
            
            Set Usuario = Nothing
        End If
    End With
    
    Set Objeto = Nothing
    
    Me.Show vbModal, frmMDI
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAuditoriaGenerico = Nothing
End Sub
