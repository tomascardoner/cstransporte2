VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReservaActualizarPrecio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Precios de las Reservas"
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
   Icon            =   "ReservaActualizarPrecio.frx":0000
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
      Format          =   16842753
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
      Format          =   16842754
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
      Caption         =   $"ReservaActualizarPrecio.frx":000C
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5595
   End
End
Attribute VB_Name = "frmReservaActualizarPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload frmReservaActualizarPrecio
End Sub

Private Sub cmdOK_Click()
    Dim Query As String
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If DateDiff("n", Now, CDate(dtpFechaDesde.Value & " " & Format(dtpHoraDesde.Value, "hh:nn"))) < 0 Then
        MsgBox "La Fecha y Hora debe ser posterior a este momento.", vbInformation, App.Title
        dtpFechaDesde.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Se actualizarán los precios de las Reservas a partir del día " & dtpFechaDesde.Value & " a las " & Format(dtpHoraDesde.Value, "hh:nn") & "." & vbCr & vbCr & "¿Desea realizar la operación?", vbExclamation + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Query = "UPDATE ViajeDetalle" & vbCr
    Query = Query & "SET ViajeDetalle.Importe = ListaPrecioDetalle.Importe" & vbCr
    Query = Query & "FROM ((ListaPrecioDetalle INNER JOIN RutaDetalle AS RutaDetalleOrigen ON ListaPrecioDetalle.IDRuta = RutaDetalleOrigen.IDRuta AND ListaPrecioDetalle.IDLugarGrupoOrigen = RutaDetalleOrigen.IDLugarGrupo)" & vbCr
    Query = Query & "INNER JOIN RutaDetalle AS RutaDetalleDestino ON ListaPrecioDetalle.IDRuta = RutaDetalleDestino.IDRuta AND ListaPrecioDetalle.IDLugarGrupoDestino = RutaDetalleDestino.IDLugarGrupo)" & vbCr
    Query = Query & "INNER JOIN ViajeDetalle ON ListaPrecioDetalle.IDRuta = ViajeDetalle.IDRuta AND ListaPrecioDetalle.OcupanteTipo = ViajeDetalle.OcupanteTipo AND ListaPrecioDetalle.IDListaPrecio = ViajeDetalle.IDListaPrecio AND RutaDetalleOrigen.IDLugar = ViajeDetalle.IDOrigen AND RutaDetalleDestino.IDLugar = ViajeDetalle.IDDestino" & vbCr
    Query = Query & "WHERE ViajeDetalle.FechaHora >= '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & " " & Format(dtpHoraDesde.Value, "hh:nn:ss") & "' AND ViajeDetalle.OcupanteTipo = 'PA' AND ViajeDetalle.Importe <> ListaPrecioDetalle.Importe"
    
    Call pDatabase.Connection.Execute(Query)
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Se ha realizado la actualización de las Reservas.", vbInformation, App.Title
    
    Unload frmReservaActualizarPrecio
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.ReservaActualizarPrecio.OK", "Error al actualizar los precios de las reservas."
End Sub

Private Sub Form_Load()
    dtpFechaDesde.Value = DateAdd("d", 1, Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReservaActualizarPrecio = Nothing
End Sub
