VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonaSaldo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cliente con Saldo Negativo"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PersonaSaldo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSaldoActual 
      Caption         =   "..."
      Height          =   315
      Left            =   3360
      TabIndex        =   11
      ToolTipText     =   "Ver Movimientos..."
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtRespuesta 
      Height          =   315
      Left            =   1740
      MaxLength       =   500
      TabIndex        =   1
      Top             =   2100
      Width           =   4695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3780
      TabIndex        =   4
      Top             =   5160
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   5160
      Width           =   1275
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000F&
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
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1020
      Width           =   4695
   End
   Begin VB.TextBox txtSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
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
      Left            =   1740
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   1995
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   3519
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FechaHora"
         Text            =   "Fecha/Hora"
         Object.Width           =   2417
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Respuesta"
         Text            =   "Respuesta"
         Object.Width           =   8114
      EndProperty
   End
   Begin VB.Label lblRespuestaAnterior 
      AutoSize        =   -1  'True
      Caption         =   "Respuestas Anteriores:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   2700
      Width           =   1725
   End
   Begin VB.Label lblRespuesta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Respuesta:"
      Height          =   210
      Left            =   780
      TabIndex        =   0
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Saldo:"
      Height          =   210
      Left            =   1140
      TabIndex        =   8
      Top             =   1620
      Width           =   450
   End
   Begin VB.Label lblPersona 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   210
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label lblDescription 
      Caption         =   "Este Cliente tiene saldo negativo y no está autorizado a viajar sin pagar al Contado. Por favor, infórmelo al Cliente."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   240
      Width           =   5715
   End
   Begin VB.Image imgExclamation 
      Height          =   480
      Left            =   120
      Picture         =   "PersonaSaldo.frx":058A
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmPersonaSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Public Sub FillListView()
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    
    If mLoading Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT PersonaRespuesta.FechaHora, PersonaRespuesta.Respuesta FROM PersonaRespuesta WHERE IDPersona = " & Val(txtPersona.Tag) & " AND Activo = 1 ORDER BY FechaHora DESC"
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        Do While Not .EOF
            Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("FechaHora").Value, Format(.Fields("FechaHora").Value, "Short Date") & " " & Format(.Fields("FechaHora").Value, "Short Time"))
            ListItem.SubItems(1) = .Fields("Respuesta").Value
            .MoveNext
        Loop
        .Close
    End With
    Set recData = Nothing
    
    On Error Resume Next
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.PersonaSaldo.FillListView", "Error al obtener la Lista de Respuestas de la Persona." & vbCr & vbCr & "IDPersona: " & Val(txtPersona.Tag)
End Sub

Private Sub cmdOK_Click()
    Dim PersonaRespuesta As PersonaRespuesta
    
    If Trim(txtRespuesta.Text) = "" Then
        MsgBox "Debe ingresar la Respuesta del Cliente.", vbInformation, App.Title
        txtRespuesta.SetFocus
        Exit Sub
    End If
    
    Set PersonaRespuesta = New PersonaRespuesta
    With PersonaRespuesta
        .IDPersona = Val(txtPersona.Tag)
        .FechaHora = Now
        .Respuesta = txtRespuesta.Text
        If .AddNew() Then
            Set PersonaRespuesta = Nothing
            Unload Me
            Exit Sub
        End If
    End With
    Set PersonaRespuesta = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSaldoActual_Click()
    Dim Persona As Persona
    
    If Val(txtPersona.Tag) > 0 Then
        Set Persona = New Persona
        Persona.IDPersona = Val(txtPersona.Tag)
        If Persona.Load() Then
            Select Case Persona.EntidadTipo
                Case ENTIDAD_TIPO_PERSONA_CLIENTE
                Case ENTIDAD_TIPO_PERSONA_CONDUCTOR
                    If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CONDUCTOR_SELECT, False) Then
                        MsgBox "No puede ver los Movimientos de Personas de tipo Conductor.", vbExclamation, App.Title
                        Set Persona = Nothing
                        Exit Sub
                    End If
                Case ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO
                    If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_ADMINISTRATIVO_SELECT, False) Then
                        MsgBox "No puede ver los Movimientos de Personas de tipo Administrativo.", vbExclamation, App.Title
                        Set Persona = Nothing
                        Exit Sub
                    End If
            End Select
        End If
        Set Persona = Nothing
        
        Screen.MousePointer = vbHourglass
        Load frmCuentaCorriente
        frmCuentaCorriente.txtPersona.Tag = Val(txtPersona.Tag)
        frmCuentaCorriente.txtPersona.Text = txtPersona.Text
        frmCuentaCorriente.cboFecha.ListIndex = 2
        frmCuentaCorriente.dtpFechaDesde.Value = DateAdd("d", -30, Date)
        frmCuentaCorriente.LoadDataAndShow
        On Error Resume Next
        If frmCuentaCorriente.WindowState = vbMinimized Then
            frmCuentaCorriente.WindowState = vbNormal
        End If
        frmCuentaCorriente.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    mLoading = True
    
    lvwData.GridLines = pParametro.ListView_GridLines
    
    Top = (frmMDI.ScaleHeight - Height) / 2
    Left = (frmMDI.ScaleWidth - Width) / 2
    
    mLoading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPersonaSaldo = Nothing
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
