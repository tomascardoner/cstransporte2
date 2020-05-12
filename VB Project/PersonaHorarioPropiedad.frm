VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPersonaHorarioPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PersonaHorarioPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   5010
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   3960
      Picture         =   "PersonaHorarioPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdLugarOrigen 
      Caption         =   "..."
      Height          =   315
      Left            =   4620
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   3660
      Width           =   255
   End
   Begin VB.CommandButton cmdLugarDestino 
      Caption         =   "..."
      Height          =   315
      Left            =   4620
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtBaja 
      Height          =   315
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   23
      Top             =   4920
      Width           =   3615
   End
   Begin VB.TextBox txtSube 
      Height          =   315
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   18
      Top             =   4020
      Width           =   3615
   End
   Begin VB.CommandButton cmdHoyHasta 
      Height          =   315
      Left            =   3120
      Picture         =   "PersonaHorarioPropiedad.frx":0B74
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdHoyDesde 
      Height          =   315
      Left            =   3120
      Picture         =   "PersonaHorarioPropiedad.frx":0CBE
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   2700
      Width           =   315
   End
   Begin VB.TextBox txtDiaSemana 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1110
   End
   Begin VB.TextBox txtHora 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1110
   End
   Begin VB.TextBox txtRuta 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2340
      TabIndex        =   24
      Top             =   5460
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3660
      TabIndex        =   25
      Top             =   5460
      Width           =   1215
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   26
      Top             =   720
      Width           =   4815
   End
   Begin MSComCtl2.DTPicker dtpFechaDesde 
      Height          =   315
      Left            =   1260
      TabIndex        =   9
      Top             =   2700
      Width           =   1875
      _ExtentX        =   3307
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
      CheckBox        =   -1  'True
      Format          =   16777217
      CurrentDate     =   36950
   End
   Begin MSComCtl2.DTPicker dtpFechaHasta 
      Height          =   315
      Left            =   1260
      TabIndex        =   12
      Top             =   3120
      Width           =   1875
      _ExtentX        =   3307
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
      CheckBox        =   -1  'True
      Format          =   16777217
      CurrentDate     =   36950
   End
   Begin MSDataListLib.DataCombo datcboOrigen 
      Height          =   330
      Left            =   1260
      TabIndex        =   15
      Top             =   3660
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo datcboDestino 
      Height          =   330
      Left            =   1260
      TabIndex        =   20
      Top             =   4560
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblOrigen 
      AutoSize        =   -1  'True
      Caption         =   "&Origen:"
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label lblDestino 
      AutoSize        =   -1  'True
      Caption         =   "&Destino:"
      Height          =   210
      Left            =   120
      TabIndex        =   19
      Top             =   4620
      Width           =   585
   End
   Begin VB.Label lblBaja 
      AutoSize        =   -1  'True
      Caption         =   "Baja:"
      Height          =   210
      Left            =   120
      TabIndex        =   22
      Top             =   4980
      Width           =   360
   End
   Begin VB.Label lblSube 
      AutoSize        =   -1  'True
      Caption         =   "Sube:"
      Height          =   210
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   420
   End
   Begin VB.Label lblFechaHasta 
      AutoSize        =   -1  'True
      Caption         =   "Fecha &Fin:"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   3180
      Width           =   750
   End
   Begin VB.Label lblFechaDesde 
      AutoSize        =   -1  'True
      Caption         =   "Fecha &Inicio:"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   2220
      Width           =   375
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "Hora:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1860
      Width           =   390
   End
   Begin VB.Label lblDiaSemana 
      AutoSize        =   -1  'True
      Caption         =   "Día Semana:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "Pasajero:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   675
   End
   Begin VB.Image imgIcon2 
      Height          =   480
      Left            =   480
      Picture         =   "PersonaHorarioPropiedad.frx":0E08
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "PersonaHorarioPropiedad.frx":16D2
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Datos del Horario de la Persona"
      Height          =   210
      Left            =   1140
      TabIndex        =   27
      Top             =   240
      Width           =   2280
   End
End
Attribute VB_Name = "frmPersonaHorarioPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCPersonaHorario As Collection

Private mRutaDetalleIndiceMinimo As Long
Private mRutaDetallaIndiceMaximo As Long

Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mCPersonaHorario(1)
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByVal CPersonaHorario As Collection)
    Dim Index As Long
    Dim PersonaHorario As PersonaHorario
    Dim DiaSemana As Integer
    Dim Hora As Date
    Dim Ruta As Ruta
    
    Set mCPersonaHorario = CPersonaHorario
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    txtPersona.Text = mCPersonaHorario(1).Persona.ApellidoNombre
    
    dtpFechaDesde.Value = Date
    dtpFechaDesde.Value = Null
    dtpFechaHasta.Value = Date
    dtpFechaHasta.Value = Null
    
    'ORIGEN - DESTINO
    Set Ruta = New Ruta
    Ruta.IDRuta = mCPersonaHorario(1).IDRuta
    If Not Ruta.GetStatistics(0, mRutaDetalleIndiceMinimo, mRutaDetallaIndiceMaximo) Then
        Set Ruta = Nothing
        Exit Sub
    End If
    Set Ruta = Nothing
    Call CSM_Control_DataCombo.FillFromSQL(datcboOrigen, "(SELECT 0 AS IDLugar, '----------' AS Nombre, -1 AS Indice) UNION (SELECT RutaDetalle.IDLugar, Lugar.Nombre, RutaDetalle.Indice FROM RutaDetalle INNER JOIN Lugar ON RutaDetalle.IDLugar = Lugar.IDLugar WHERE RutaDetalle.IDRuta = '" & ReplaceQuote(mCPersonaHorario(1).IDRuta) & "' AND RutaDetalle.Indice < " & mRutaDetallaIndiceMaximo & " AND RutaDetalle.IDLugar <> " & pParametro.Lugar_ID_Otro & " AND Lugar.Activo = 1) ORDER BY Indice", "IDLugar", "Nombre", "Orígenes", cscpFirst)
    datcboOrigen_Change
    
    For Index = 1 To mCPersonaHorario.Count
        Set PersonaHorario = mCPersonaHorario(Index)
        
        With PersonaHorario
            If Index = 1 Then
                DiaSemana = .DiaSemana
                Hora = .Hora
            Else
                If DiaSemana <> .DiaSemana Then
                    DiaSemana = -1
                End If
                If Hora <> .Hora Then
                    Hora = DATE_TIME_FIELD_NULL_VALUE
                End If
            End If
        End With
    Next Index
    
    'DIA SEMANA
    If DiaSemana = -1 Then
        txtDiaSemana.Text = "Múltiples"
    Else
        txtDiaSemana.Text = WeekdayName(DiaSemana)
    End If
    
    'HORA
    If Hora = DATE_TIME_FIELD_NULL_VALUE Then
        txtHora.Text = "Múltiples"
    Else
        txtHora.Text = Format(Hora, "Short Time")
    End If
    
    txtRuta.Text = mCPersonaHorario(1).IDRuta
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHoyDesde_Click()
    dtpFechaDesde.Value = Date
End Sub

Private Sub cmdHoyHasta_Click()
    dtpFechaHasta.Value = Date
End Sub

Private Sub cmdLugarDestino_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR) Then
        Screen.MousePointer = vbHourglass
        frmLugar.Show
        On Error Resume Next
        Set frmLugar.lvwData.SelectedItem = frmLugar.lvwData.ListItems(KEY_STRINGER & Val(datcboDestino.BoundText))
        frmLugar.lvwData.SelectedItem.EnsureVisible
        If frmLugar.WindowState = vbMinimized Then
            frmLugar.WindowState = vbNormal
        End If
        frmLugar.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdLugarOrigen_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR) Then
        Screen.MousePointer = vbHourglass
        frmLugar.Show
        On Error Resume Next
        Set frmLugar.lvwData.SelectedItem = frmLugar.lvwData.ListItems(KEY_STRINGER & Val(datcboOrigen.BoundText))
        frmLugar.lvwData.SelectedItem.EnsureVisible
        If frmLugar.WindowState = vbMinimized Then
            frmLugar.WindowState = vbNormal
        End If
        frmLugar.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdOK_Click()
    Dim PersonaHorario As PersonaHorario
    
    If (Not IsNull(dtpFechaDesde.Value)) And Not (IsNull(dtpFechaHasta.Value)) Then
        If DateDiff("d", dtpFechaDesde.Value, dtpFechaHasta.Value) < 0 Then
            MsgBox "La Fecha de Inicio debe ser menor o igual a la Fecha de Fin.", vbInformation, App.Title
            dtpFechaHasta.SetFocus
            Exit Sub
        End If
    End If
    
    For Each PersonaHorario In mCPersonaHorario
        With PersonaHorario
            .RefreshList = False
            .FechaDesde = IIf(IsNull(dtpFechaDesde.Value), DATE_TIME_FIELD_NULL_VALUE, dtpFechaDesde.Value)
            .FechaHasta = IIf(IsNull(dtpFechaHasta.Value), DATE_TIME_FIELD_NULL_VALUE, dtpFechaHasta.Value)
            .IDOrigen = Val(datcboOrigen.BoundText)
            .Sube = txtSube.Text
            .IDDestino = Val(datcboDestino.BoundText)
            .Baja = txtBaja.Text
            
            If Not .AddNew() Then
                RefreshList_RefreshPersonaHorario .IDPersona, 1, Date, ""
                Exit Sub
            End If
        End With
    Next PersonaHorario
    
    MsgBox "Se ha generado la Reserva Fija." & vbCr & "De todos modos, recuerde verificar que haya lugar en los Viajes correspondientes.", vbExclamation, App.Title
        
    RefreshList_RefreshPersonaHorario mCPersonaHorario(1).IDPersona, 1, Time, ""
    RefreshList_RefreshViajeDetalle Date, "", 0, True
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mCPersonaHorario = Nothing
    Set frmPersonaHorarioPropiedad = Nothing
End Sub

Private Sub txtBaja_GotFocus()
    CSM_Control_TextBox.SelAllText txtBaja
End Sub

Private Sub txtSube_GotFocus()
    CSM_Control_TextBox.SelAllText txtSube
End Sub

Private Sub datcboOrigen_Change()
    Dim RutaDetalle As RutaDetalle
    Dim RutaDetalleIndice As Long
    
    Dim IDDestinoSave As Long
    
    If datcboOrigen.Text = "" Then
        Exit Sub
    End If
        
    'Busco el Detalle de la Ruta para filtrar el ComboBox de Destino a partir del Origen
    If Val(datcboOrigen.BoundText) = 0 Then
        RutaDetalleIndice = -1
    Else
        Set RutaDetalle = New RutaDetalle
        RutaDetalle.IDRuta = mCPersonaHorario(1).IDRuta
        RutaDetalle.IDLugar = Val(datcboOrigen.BoundText)
        If RutaDetalle.Load() Then
            RutaDetalleIndice = RutaDetalle.Indice
        End If
        Set RutaDetalle = Nothing
    End If

    IDDestinoSave = Val(datcboDestino.BoundText)
    Call CSM_Control_DataCombo.FillFromSQL(datcboDestino, "(SELECT 0 AS IDLugar, '----------' AS Nombre, 99999999 AS Indice) UNION (SELECT RutaDetalle.IDLugar, Lugar.Nombre, RutaDetalle.Indice FROM RutaDetalle INNER JOIN Lugar ON RutaDetalle.IDLugar = Lugar.IDLugar WHERE RutaDetalle.IDRuta = '" & ReplaceQuote(mCPersonaHorario(1).IDRuta) & "' AND RutaDetalle.Indice > " & RutaDetalleIndice & " AND RutaDetalle.IDLugar <> " & pParametro.Lugar_ID_Otro & " AND Lugar.Activo = 1) ORDER BY Indice", "IDLugar", "Nombre", "Destinos", cscpItemOrLast, IDDestinoSave)
End Sub

Public Sub FillComboBoxLugar()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboOrigen.BoundText)
    Set recData = datcboOrigen.RowSource
    recData.Requery
    Set recData = Nothing
    datcboOrigen.BoundText = KeySave

    KeySave = Val(datcboDestino.BoundText)
    Set recData = datcboDestino.RowSource
    recData.Requery
    Set recData = Nothing
    datcboDestino.BoundText = KeySave
End Sub
