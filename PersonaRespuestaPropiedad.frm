VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPersonaRespuestaPropiedad 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "PersonaRespuestaPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   4905
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   3780
      Picture         =   "PersonaRespuestaPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1020
      Width           =   3615
   End
   Begin VB.TextBox txtRespuesta 
      Height          =   315
      Left            =   1140
      MaxLength       =   500
      TabIndex        =   8
      Top             =   2460
      Width           =   3615
   End
   Begin VB.CommandButton cmdHoy 
      Height          =   315
      Left            =   3360
      Picture         =   "PersonaRespuestaPropiedad.frx":0B74
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   1500
      Width           =   315
   End
   Begin VB.CommandButton cmdAnterior 
      Height          =   315
      Left            =   1140
      Picture         =   "PersonaRespuestaPropiedad.frx":0CBE
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   1500
      Width           =   300
   End
   Begin VB.CommandButton cmdSiguiente 
      Height          =   315
      Left            =   3060
      Picture         =   "PersonaRespuestaPropiedad.frx":1248
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   1500
      Width           =   300
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
      Left            =   180
      TabIndex        =   9
      Top             =   2940
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   15
      Top             =   780
      Width           =   4635
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
      Left            =   3540
      TabIndex        =   11
      Top             =   3420
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
      Left            =   2220
      TabIndex        =   10
      Top             =   3420
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   1500
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
   Begin MSComCtl2.DTPicker dtpHora 
      Height          =   315
      Left            =   1140
      TabIndex        =   6
      Top             =   1980
      Width           =   1395
      _ExtentX        =   2461
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
      Format          =   16777218
      CurrentDate     =   36494
   End
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "&Pasajero:"
      Height          =   210
      Left            =   180
      TabIndex        =   12
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label lblRespuesta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Respuesta:"
      Height          =   210
      Left            =   180
      TabIndex        =   7
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "&Fecha:"
      Height          =   210
      Left            =   180
      TabIndex        =   0
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "&Hora:"
      Height          =   210
      Left            =   180
      TabIndex        =   5
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   " Datos de la Respuesta de la Persona"
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
      TabIndex        =   14
      Top             =   300
      Width           =   2715
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "PersonaRespuestaPropiedad.frx":17D2
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPersonaRespuestaPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPersonaRespuesta As PersonaRespuesta
Private mNew As Boolean

Private mLoading As Boolean


Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mPersonaRespuesta
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef PersonaRespuesta As PersonaRespuesta)
    Dim Persona As Persona
    
    Set mPersonaRespuesta = PersonaRespuesta
    mNew = (mPersonaRespuesta.FechaHora = DATE_TIME_FIELD_NULL_VALUE)
    
    mLoading = True

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    Set Persona = New Persona
    Persona.IDPersona = mPersonaRespuesta.IDPersona
    If Not Persona.Load() Then
        Set Persona = Nothing
        Unload Me
        Exit Sub
    End If
    txtPersona.Text = Persona.ApellidoNombre
    Set Persona = Nothing
    
    With mPersonaRespuesta
        If mNew Then
            dtpFecha.Value = Date
            dtpHora.Value = Time
            txtRespuesta.Text = ""
            chkActivo.Value = vbChecked
        Else
            dtpFecha.Value = Format(.FechaHora, "Short Date")
            dtpHora.Value = Format(.FechaHora, "Short Time")
            txtRespuesta.Text = .Respuesta
            chkActivo = IIf(.Activo, vbChecked, vbUnchecked)
        End If
    End With
    
    mLoading = False
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub cmdAnterior_Click()
    dtpFecha.Value = DateAdd("d", -1, dtpFecha.Value)
    dtpFecha.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHoy_Click()
    dtpFecha.Value = Date
    dtpFecha.SetFocus
End Sub

Private Sub cmdOK_Click()
    If Trim(txtRespuesta.Text) = "" Then
        MsgBox "Debe ingresar la Respuesta de la Persona.", vbInformation, App.Title
        txtRespuesta.SetFocus
        Exit Sub
    End If
    
    With mPersonaRespuesta
        .FechaHora = CDate(dtpFecha.Value & " " & dtpHora.Value)
        .Respuesta = txtRespuesta.Text
        .Activo = (chkActivo.Value = vbChecked)
        If mNew Then
            If Not .AddNew() Then
                Exit Sub
            End If
        Else
            If Not .Update() Then
                Exit Sub
            End If
        End If
    End With
    
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdSiguiente_Click()
    dtpFecha.Value = DateAdd("d", 1, dtpFecha.Value)
    dtpFecha.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mPersonaRespuesta = Nothing
    Set frmPersonaRespuestaPropiedad = Nothing
End Sub

Private Sub txtRespuesta_GotFocus()
    CSM_Control_TextBox.SelAllText txtRespuesta
End Sub

Private Sub txtRespuesta_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRespuesta_LostFocus()
    txtRespuesta.Text = UCase(txtRespuesta.Text)
    txtRespuesta.Text = CleanInvalidSpaces(txtRespuesta.Text)
End Sub

