VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmViajeActualizarSueldo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Sueldos de los Viajes"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajeActualizarSueldo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5850
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4500
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3180
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFechaDesde 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1260
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   47448065
      CurrentDate     =   37897
   End
   Begin MSComCtl2.DTPicker dtpHoraDesde 
      Height          =   315
      Left            =   3300
      TabIndex        =   3
      Top             =   1260
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
      Format          =   47448066
      CurrentDate     =   36494
   End
   Begin VB.Label lblFechaDesde 
      AutoSize        =   -1  'True
      Caption         =   "Actualizar desde el:"
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Label lblProcessDescription 
      Caption         =   $"ViajeActualizarSueldo.frx":000C
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5595
   End
End
Attribute VB_Name = "frmViajeActualizarSueldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    dtpFechaDesde.Value = DateAdd("d", 1, Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmViajeActualizarSueldo = Nothing
End Sub

Private Sub cmdOK_Click()
    Dim recData As ADODB.Recordset
    Dim Viaje As Viaje
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If DateDiff("d", dtpFechaDesde.Value, Now) > 60 Then
        MsgBox "La Fecha debe ser hoy o posterior.", vbInformation, App.Title
        dtpFechaDesde.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Se actualizarán los sueldos de los Viajes a partir del día " & dtpFechaDesde.Value & "." & vbCr & vbCr & "¿Desea realizar la operación?", vbExclamation + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set recData = New ADODB.Recordset
    Set recData.ActiveConnection = pDatabase.Connection
    recData.CursorType = adOpenForwardOnly
    recData.LockType = adLockReadOnly
    recData.Source = "SELECT FechaHora, IDRuta FROM Viaje WHERE FechaHora >= CONVERT(smalldatetime, '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & " " & Format(dtpHoraDesde.Value, "hh:nn:ss") & "') AND Estado = '" & VIAJE_ESTADO_FINALIZADO & "' ORDER BY FechaHora, IDRuta"
    recData.Open , , , , adCmdText
    
    Set Viaje = New Viaje
    Viaje.RefreshListSkip = True
    Do While Not recData.EOF
        With Viaje
            Viaje.FechaHora = recData("FechaHora").Value
            Viaje.IDRuta = RTrim(recData("IDRuta").Value)
            If Viaje.Load() Then
                Call Viaje.CuentaCorriente_Asignar_Sueldo
            End If
            recData.MoveNext
        End With
    Loop
    Set Viaje = Nothing
    RefreshList_Module.RefreshList_RefreshCuentaCorriente 0
    
    recData.Close
    Set recData = Nothing
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Se ha realizado la actualización de los Sueldos.", vbInformation, App.Title
    
    Unload frmViajeActualizarSueldo
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.ViajeActualizarSueldo.OK", "Error al actualizar los sueldos de los viajes."
End Sub

Private Sub cmdCancel_Click()
    Unload frmViajeActualizarSueldo
End Sub
