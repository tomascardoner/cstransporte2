VERSION 5.00
Begin VB.Form frmViajeDetalleCancelar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelación de la Reserva"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajeDetalleCancelar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1380
      Width           =   3735
   End
   Begin VB.CheckBox chkCanceladoForzarDebito 
      Alignment       =   1  'Right Justify
      Caption         =   "Debitar Viaje"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1395
   End
   Begin VB.TextBox txtCanceladoPor 
      Height          =   315
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1980
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3180
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3180
      Width           =   1215
   End
   Begin VB.TextBox txtFechaDiaSemana 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1050
   End
   Begin VB.TextBox txtFecha 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1410
   End
   Begin VB.TextBox txtHora 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   540
      Width           =   1170
   End
   Begin VB.TextBox txtRuta 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   3090
   End
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "Pasajero:"
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   675
   End
   Begin VB.Label lblCanceladoPor 
      AutoSize        =   -1  'True
      Caption         =   "Cancelado por:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   180
      Width           =   495
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "Hora:"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   390
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   1020
      Width           =   375
   End
End
Attribute VB_Name = "frmViajeDetalleCancelar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mViajeDetalle As ViajeDetalle

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef ViajeDetalle As ViajeDetalle)
    Dim Persona As Persona
    
    Set mViajeDetalle = ViajeDetalle
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mViajeDetalle
        txtFechaDiaSemana.Text = .FechaHora_WeekdayName
        txtFecha.Text = .FechaHora_FormattedAsDate
        txtHora.Text = .FechaHora_FormattedAsTime
        txtRuta.Text = .IDRuta
        
        Set Persona = New Persona
        Persona.IDPersona = .IDPersona
        If Persona.Load() Then
            txtPersona.Text = Persona.ApellidoNombre
            txtCanceladoPor.Text = Persona.ApellidoNombre
        End If
        Set Persona = Nothing
    End With
    
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
    If Trim(txtCanceladoPor.Text) = "" Then
        MsgBox "Debe especificar quién Canceló la Reserva.", vbInformation, App.Title
        txtCanceladoPor.SetFocus
        Exit Sub
    End If
    
    With mViajeDetalle
        .Estado = VIAJE_DETALLE_ESTADO_CANCELADO
        .CanceladoPor = txtCanceladoPor.Text
        .ForzarDebito = (chkCanceladoForzarDebito.Value = vbChecked)
        If Not .CambiarEstado(pParametro.Viaje_Permite_RutaConexion) Then
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mViajeDetalle = Nothing
    Set frmViajeDetalleCancelar = Nothing
End Sub

Private Sub txtCanceladoPor_GotFocus()
    CSM_Control_TextBox.SelAllText txtCanceladoPor
End Sub
