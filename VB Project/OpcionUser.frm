VERSION 5.00
Begin VB.Form frmOpcionUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones del Usuario"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OpcionUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   4530
   Begin VB.CheckBox chkMessenger_Enabled 
      Caption         =   "Habilitar Messenger"
      Height          =   210
      Left            =   180
      TabIndex        =   12
      Top             =   4560
      Width           =   1995
   End
   Begin VB.Frame fraViajeEstadoVencido 
      Caption         =   "Aviso de Viajes con Estado vencido:"
      Height          =   1035
      Left            =   120
      TabIndex        =   8
      Top             =   3300
      Width           =   4275
      Begin VB.TextBox txtViaje_EstadoVencido_CheckIntervalSeconds 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   11
         Top             =   600
         Width           =   570
      End
      Begin VB.CheckBox chkViaje_EstadoVencido_Check 
         Caption         =   "Avisar"
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lblViaje_EstadoVencido_CheckIntervalSeconds 
         AutoSize        =   -1  'True
         Caption         =   "Intervalo de Verificación:                 segundos."
         Height          =   210
         Left            =   480
         TabIndex        =   10
         Top             =   660
         Width           =   3330
      End
   End
   Begin VB.TextBox txtPersona_Apellido_Busqueda_Delay_Milliseconds 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   14
      Top             =   5100
      Width           =   570
   End
   Begin VB.Frame fraWarning 
      Caption         =   "Avisos al Inicio de Sesión:"
      Height          =   1515
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4275
      Begin VB.CheckBox chkAlarma_Aviso 
         Caption         =   "Alarmas Generales"
         Height          =   210
         Left            =   180
         TabIndex        =   7
         Top             =   1140
         Width           =   3915
      End
      Begin VB.CheckBox chkPersonaAlarma_Aviso 
         Caption         =   "Alarmas de Personas"
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   720
         Width           =   3915
      End
      Begin VB.CheckBox chkVehiculoMantenimiento_Aviso 
         Caption         =   "Mantenimiento de Vehículos"
         Height          =   210
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   3915
      End
   End
   Begin VB.Frame fraListViewDesign 
      Caption         =   "Diseño de Listas"
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4275
      Begin VB.CheckBox chkListView_GridLines 
         Caption         =   "Mostar Líneas Divisorias"
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   3915
      End
      Begin VB.CheckBox chkViajeDetalle_SeparateRowsByType 
         Caption         =   "Detalle de Viaje: Separador de Filas por Tipo"
         Height          =   210
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Width           =   3915
      End
      Begin VB.CheckBox chkViajeDetalle_SeparateRowsByStatus 
         Caption         =   "Detalle de Viaje: Separador de Filas por Estado"
         Height          =   210
         Left            =   180
         TabIndex        =   3
         Top             =   1140
         Width           =   3915
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1860
      TabIndex        =   16
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3180
      TabIndex        =   17
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Label lblPersona_Apellido_Busqueda_Delay_Milliseconds_Unit 
      AutoSize        =   -1  'True
      Caption         =   "milisegundos."
      Height          =   210
      Left            =   3000
      TabIndex        =   15
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblPersona_Apellido_Busqueda_Delay_Milliseconds 
      Caption         =   "Intervalo de tiempo para la Búsqueda de Personas:"
      Height          =   450
      Left            =   180
      TabIndex        =   13
      Top             =   5040
      Width           =   1935
   End
End
Attribute VB_Name = "frmOpcionUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkViaje_EstadoVencido_Check_Click()
    lblViaje_EstadoVencido_CheckIntervalSeconds.Visible = (chkViaje_EstadoVencido_Check.Value = vbChecked)
    txtViaje_EstadoVencido_CheckIntervalSeconds.Visible = (chkViaje_EstadoVencido_Check.Value = vbChecked)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not IsNumeric(txtViaje_EstadoVencido_CheckIntervalSeconds.Text) Then
        MsgBox "El Intervalo de Verificación para el Aviso de Viajes con Estado vencido debe ser un valor numérico.", vbInformation, App.Title
        txtViaje_EstadoVencido_CheckIntervalSeconds.SetFocus
        txtViaje_EstadoVencido_CheckIntervalSeconds_GotFocus
        Exit Sub
    End If
    If Not IsNumeric(txtPersona_Apellido_Busqueda_Delay_Milliseconds.Text) Then
        MsgBox "El Intervalo de tiempo para la Búsqueda de Personas debe ser un valor numérico.", vbInformation, App.Title
        txtPersona_Apellido_Busqueda_Delay_Milliseconds.SetFocus
        txtPersona_Apellido_Busqueda_Delay_Milliseconds_GotFocus
        Exit Sub
    End If
    If CLng(txtPersona_Apellido_Busqueda_Delay_Milliseconds.Text) <= 0 Then
        MsgBox "El Intervalo de tiempo para la Búsqueda de Personas debe mayor a cero.", vbInformation, App.Title
        txtPersona_Apellido_Busqueda_Delay_Milliseconds.SetFocus
        txtPersona_Apellido_Busqueda_Delay_Milliseconds_GotFocus
        Exit Sub
    End If
    
    pParametro.ListView_GridLines = (chkListView_GridLines.Value = vbChecked)
    pParametro.ViajeDetalle_SeparateRowsByType = (chkViajeDetalle_SeparateRowsByType.Value = vbChecked)
    pParametro.ViajeDetalle_SeparateRowsByStatus = (chkViajeDetalle_SeparateRowsByStatus.Value = vbChecked)
    pParametro.Usuario_GuardarSiNo "ListView_GridLines", pParametro.ListView_GridLines
    pParametro.Usuario_GuardarSiNo "ViajeDetalle_SeparateRowsByType", pParametro.ViajeDetalle_SeparateRowsByType
    pParametro.Usuario_GuardarSiNo "ViajeDetalle_SeparateRowsByStatus", pParametro.ViajeDetalle_SeparateRowsByStatus
    
    pParametro.VehiculoMantenimiento_Aviso = (chkVehiculoMantenimiento_Aviso.Value = vbChecked)
    pParametro.PersonaAlarma_Aviso = (chkPersonaAlarma_Aviso.Value = vbChecked)
    pParametro.Alarma_Aviso = (chkAlarma_Aviso.Value = vbChecked)
    pParametro.Usuario_GuardarSiNo "VehiculoMantenimiento_Aviso", pParametro.VehiculoMantenimiento_Aviso
    pParametro.Usuario_GuardarSiNo "PersonaAlarma_Aviso", pParametro.PersonaAlarma_Aviso
    pParametro.Usuario_GuardarSiNo "Alarma_Aviso", pParametro.Alarma_Aviso
    
    pParametro.Viaje_EstadoVencido_Check = (chkViaje_EstadoVencido_Check.Value = vbChecked)
    pParametro.Viaje_EstadoVencido_CheckIntervalSeconds = CLng(txtViaje_EstadoVencido_CheckIntervalSeconds.Text)
    pParametro.Usuario_GuardarSiNo "Viaje_EstadoVencido_Check", pParametro.Viaje_EstadoVencido_Check
    pParametro.Usuario_GuardarNumero "Viaje_EstadoVencido_CheckIntervalSeconds", pParametro.Viaje_EstadoVencido_CheckIntervalSeconds
    
    ' MESSENGER
    pParametro.Usuario_Messenger_Enabled = (chkMessenger_Enabled.Value = vbChecked)
    pParametro.Usuario_GuardarSiNo "Messenger_Enabled", pParametro.Usuario_Messenger_Enabled
    pMessengerEnabled = (pParametro.Messenger_Enabled And pParametro.Usuario_Messenger_Enabled And pCPermiso.GotPermission(PERMISO_MESSENGER, False))
    frmMDI.tlbMain.Buttons("MESSENGER").Visible = pMessengerEnabled
    frmMDI.tlbMain.Buttons("SEPARATOR_7").Visible = pMessengerEnabled
    
    pParametro.Persona_Apellido_Busqueda_Delay_Milliseconds = CLng(txtPersona_Apellido_Busqueda_Delay_Milliseconds.Text)
    pParametro.Usuario_GuardarNumero "Persona_Apellido_Busqueda_Delay_Milliseconds", pParametro.Persona_Apellido_Busqueda_Delay_Milliseconds
    
    Unload Me
End Sub

Private Sub Form_Load()
    chkListView_GridLines.Value = IIf(pParametro.ListView_GridLines, vbChecked, vbUnchecked)
    chkViajeDetalle_SeparateRowsByType.Value = IIf(pParametro.ViajeDetalle_SeparateRowsByType, vbChecked, vbUnchecked)
    chkViajeDetalle_SeparateRowsByStatus.Value = IIf(pParametro.ViajeDetalle_SeparateRowsByStatus, vbChecked, vbUnchecked)
    
    chkVehiculoMantenimiento_Aviso.Value = IIf(pParametro.VehiculoMantenimiento_Aviso, vbChecked, vbUnchecked)
    chkPersonaAlarma_Aviso.Value = IIf(pParametro.PersonaAlarma_Aviso, vbChecked, vbUnchecked)
    chkAlarma_Aviso.Value = IIf(pParametro.Alarma_Aviso, vbChecked, vbUnchecked)
    
    chkViaje_EstadoVencido_Check.Value = IIf(pParametro.Viaje_EstadoVencido_Check, vbChecked, vbUnchecked)
    chkViaje_EstadoVencido_Check_Click
    txtViaje_EstadoVencido_CheckIntervalSeconds.Text = pParametro.Viaje_EstadoVencido_CheckIntervalSeconds
    
    chkMessenger_Enabled.Value = IIf(pParametro.Usuario_Messenger_Enabled, vbChecked, vbUnchecked)
    
    txtPersona_Apellido_Busqueda_Delay_Milliseconds.Text = pParametro.Persona_Apellido_Busqueda_Delay_Milliseconds
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOpcionUser = Nothing
End Sub

Private Sub txtViaje_EstadoVencido_CheckIntervalSeconds_GotFocus()
    CSM_Control_TextBox.SelAllText txtViaje_EstadoVencido_CheckIntervalSeconds
End Sub

Private Sub txtViaje_EstadoVencido_CheckIntervalSeconds_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtViaje_EstadoVencido_CheckIntervalSeconds_LostFocus()
    txtViaje_EstadoVencido_CheckIntervalSeconds.Text = Val(txtViaje_EstadoVencido_CheckIntervalSeconds.Text)
    If txtViaje_EstadoVencido_CheckIntervalSeconds.Text = 0 Then
        txtViaje_EstadoVencido_CheckIntervalSeconds.Text = ""
    End If
End Sub

Private Sub txtPersona_Apellido_Busqueda_Delay_Milliseconds_GotFocus()
    CSM_Control_TextBox.SelAllText txtPersona_Apellido_Busqueda_Delay_Milliseconds
End Sub

Private Sub txtPersona_Apellido_Busqueda_Delay_Milliseconds_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPersona_Apellido_Busqueda_Delay_Milliseconds_LostFocus()
    txtPersona_Apellido_Busqueda_Delay_Milliseconds.Text = Val(txtPersona_Apellido_Busqueda_Delay_Milliseconds.Text)
    If txtPersona_Apellido_Busqueda_Delay_Milliseconds.Text = 0 Then
        txtPersona_Apellido_Busqueda_Delay_Milliseconds.Text = ""
    End If
End Sub

