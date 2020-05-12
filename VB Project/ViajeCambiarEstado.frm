VERSION 5.00
Begin VB.Form frmViajeCambiarEstado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de Estado del Viaje"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajeCambiarEstado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picEstadoNuevo 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2640
      ScaleHeight     =   3495
      ScaleWidth      =   1095
      TabIndex        =   20
      Top             =   1860
      Width           =   1095
      Begin VB.OptionButton optEstadoNuevoEnProgreso 
         Caption         =   "En Progreso"
         Height          =   795
         Left            =   0
         Picture         =   "ViajeCambiarEstado.frx":062A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   900
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optEstadoNuevoFinalizado 
         Caption         =   "Finalizado"
         Height          =   795
         Left            =   0
         Picture         =   "ViajeCambiarEstado.frx":0EF4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optEstadoNuevoCancelado 
         Caption         =   "Cancelado"
         Height          =   795
         Left            =   0
         Picture         =   "ViajeCambiarEstado.frx":17BE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2700
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optEstadoNuevoActivo 
         Caption         =   "Activo"
         Height          =   795
         Left            =   0
         Picture         =   "ViajeCambiarEstado.frx":2088
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox picEstadoActual 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   360
      ScaleHeight     =   3495
      ScaleWidth      =   1095
      TabIndex        =   19
      Top             =   1860
      Width           =   1095
      Begin VB.OptionButton optEstadoActualActivo 
         Caption         =   "Activo"
         Height          =   795
         Left            =   0
         Picture         =   "ViajeCambiarEstado.frx":296A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optEstadoActualEnProgreso 
         Caption         =   "En Progreso"
         Height          =   795
         Left            =   0
         Picture         =   "ViajeCambiarEstado.frx":324C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optEstadoActualFinalizado 
         Caption         =   "Finalizado"
         Height          =   795
         Left            =   0
         Picture         =   "ViajeCambiarEstado.frx":3B16
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optEstadoActualCancelado 
         Caption         =   "Cancelado"
         Height          =   795
         Left            =   0
         Picture         =   "ViajeCambiarEstado.frx":43E0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.TextBox txtRuta 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   780
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   960
      Width           =   3090
   End
   Begin VB.TextBox txtHora 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   780
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   540
      Width           =   1170
   End
   Begin VB.TextBox txtFecha 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   1410
   End
   Begin VB.TextBox txtFechaDiaSemana 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   780
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1500
      TabIndex        =   4
      Top             =   5820
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2820
      TabIndex        =   5
      Top             =   5820
      Width           =   1215
   End
   Begin VB.Line linEstadoNuevoCancelado 
      BorderWidth     =   2
      X1              =   1560
      X2              =   2520
      Y1              =   3840
      Y2              =   4920
   End
   Begin VB.Line linEstadoNuevoFinalizado 
      BorderWidth     =   2
      X1              =   1560
      X2              =   2520
      Y1              =   3660
      Y2              =   4020
   End
   Begin VB.Line linEstadoNuevoEnProgreso 
      BorderWidth     =   2
      X1              =   1560
      X2              =   2520
      Y1              =   3480
      Y2              =   3180
   End
   Begin VB.Line linEstadoNuevoActivo 
      BorderWidth     =   2
      X1              =   1560
      X2              =   2520
      Y1              =   3300
      Y2              =   2280
   End
   Begin VB.Image imgArrowRight 
      Height          =   480
      Left            =   1800
      Picture         =   "ViajeCambiarEstado.frx":4CAA
      Top             =   1320
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   180
      X2              =   3960
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Label lblEstadoNuevo 
      AutoSize        =   -1  'True
      Caption         =   "Estado Nuevo"
      Height          =   210
      Left            =   2700
      TabIndex        =   14
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label lblEstadoActual 
      AutoSize        =   -1  'True
      Caption         =   "Estado Actual"
      Height          =   210
      Left            =   360
      TabIndex        =   13
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   1020
      Width           =   375
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "Hora:"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   390
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   180
      Width           =   495
   End
End
Attribute VB_Name = "frmViajeCambiarEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mViaje As Viaje

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByVal FechaHora As Date, ByVal IDRuta As String)
    Dim ValueActivo As Boolean
    Dim ValueEnProgreso As Boolean
    Dim ValueFinalizado As Boolean
    Dim ValueCancelado As Boolean
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    txtFechaDiaSemana.Text = WeekdayName(Weekday(FechaHora))
    txtFecha.Text = Format(FechaHora, "Short Date")
    txtHora.Text = Format(FechaHora, "Short Time")
    txtRuta.Text = IDRuta
    
    Set mViaje = New Viaje
    mViaje.FechaHora = FechaHora
    mViaje.IDRuta = IDRuta
    If Not mViaje.Load() Then
        Unload Me
        Exit Sub
    End If
    
    'ANTERIOR
    ValueActivo = (mViaje.Estado = VIAJE_ESTADO_ACTIVO)
    optEstadoActualActivo.Visible = ValueActivo
    optEstadoActualActivo.Value = ValueActivo
    
    ValueEnProgreso = (mViaje.Estado = VIAJE_ESTADO_EN_PROGRESO)
    optEstadoActualEnProgreso.Visible = ValueEnProgreso
    optEstadoActualEnProgreso.Value = ValueEnProgreso
    
    ValueFinalizado = (mViaje.Estado = VIAJE_ESTADO_FINALIZADO)
    optEstadoActualFinalizado.Visible = ValueFinalizado
    optEstadoActualFinalizado.Value = ValueFinalizado
    
    ValueCancelado = (mViaje.Estado = VIAJE_ESTADO_CANCELADO)
    optEstadoActualCancelado.Visible = ValueCancelado
    optEstadoActualCancelado.Value = ValueCancelado
    
    'NUEVOS
    ValueActivo = ((Not ValueActivo) And pCPermiso.GotPermission(PERMISO_VIAJE_CHANGE_STATUS_SPECIAL, False))
    optEstadoNuevoActivo.Visible = ValueActivo
    linEstadoNuevoActivo.Visible = ValueActivo
    
    ValueEnProgreso = ((Not ValueEnProgreso) And (mViaje.Estado = VIAJE_ESTADO_ACTIVO Or pCPermiso.GotPermission(PERMISO_VIAJE_CHANGE_STATUS_SPECIAL, False)))
    optEstadoNuevoEnProgreso.Visible = ValueEnProgreso
    linEstadoNuevoEnProgreso.Visible = ValueEnProgreso
    
    ValueFinalizado = ((Not ValueFinalizado) And (mViaje.Estado = VIAJE_ESTADO_EN_PROGRESO Or pCPermiso.GotPermission(PERMISO_VIAJE_CHANGE_STATUS_SPECIAL, False)))
    optEstadoNuevoFinalizado.Visible = ValueFinalizado
    linEstadoNuevoFinalizado.Visible = ValueFinalizado
    
    ValueCancelado = ((Not ValueCancelado) And ((mViaje.Estado = VIAJE_ESTADO_ACTIVO And pCPermiso.GotPermission(PERMISO_VIAJE_CHANGE_STATUS_CANCEL, False)) Or pCPermiso.GotPermission(PERMISO_VIAJE_CHANGE_STATUS_SPECIAL, False)))
    optEstadoNuevoCancelado.Visible = ValueCancelado
    linEstadoNuevoCancelado.Visible = ValueCancelado
    
    If Not (ValueActivo Or ValueEnProgreso Or ValueFinalizado Or ValueCancelado) Then
        MsgBox "No se puede realizar el Cambio de Estado.", vbInformation, App.Title
        Unload Me
        Exit Sub
    End If
    
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
    Dim GenerarReservasFijas As Boolean
    
    If (optEstadoNuevoEnProgreso.Value Or optEstadoNuevoFinalizado.Value) And mViaje.IDVehiculo = 0 Then
        MsgBox "Debe especificar el Vehículo del Viaje.", vbInformation, App.Title
        Exit Sub
    End If
    If (optEstadoNuevoEnProgreso.Value Or optEstadoNuevoFinalizado.Value) And mViaje.IDConductor = 0 Then
        MsgBox "Debe especificar el Conductor del Viaje.", vbInformation, App.Title
        Exit Sub
    End If
    If optEstadoNuevoCancelado.Value Then
        GenerarReservasFijas = False
    End If
    If optEstadoActualCancelado.Value Then
        GenerarReservasFijas = True
    End If
        
    mViaje.Estado = Switch(optEstadoNuevoActivo.Value, VIAJE_ESTADO_ACTIVO, optEstadoNuevoEnProgreso.Value, VIAJE_ESTADO_EN_PROGRESO, optEstadoNuevoFinalizado.Value, VIAJE_ESTADO_FINALIZADO, optEstadoNuevoCancelado.Value, VIAJE_ESTADO_CANCELADO)
    If Not mViaje.Update Then
        Exit Sub
    End If
    
    If GenerarReservasFijas Then
        mViaje.DiaSemanaBase = Weekday(mViaje.FechaHora)
        Call mViaje.GenerarReservasFijas
    End If
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mViaje = Nothing
    Set frmViajeCambiarEstado = Nothing
End Sub
