VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonaDatosTransferencia 
   Caption         =   "Transferencia de Datos de Pasajeros"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   Icon            =   "PersonaDatosTransferencia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   10410
   Begin VB.CommandButton cmdBoth 
      Height          =   615
      Left            =   4920
      Picture         =   "PersonaDatosTransferencia.frx":062A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1620
      Width           =   615
   End
   Begin VB.Timer tmrListViewSettingsUpdate 
      Interval        =   5000
      Left            =   4980
      Top             =   4560
   End
   Begin MSComctlLib.ImageList ilsData 
      Left            =   4920
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PersonaDatosTransferencia.frx":0EF4
            Key             =   "PASAJERO_CONFIRMADO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PersonaDatosTransferencia.frx":148E
            Key             =   "PASAJERO_CONDICIONAL"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PersonaDatosTransferencia.frx":1A28
            Key             =   "PASAJERO_CANCELADO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PersonaDatosTransferencia.frx":1FC2
            Key             =   "COMISION"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRightToLeft 
      Height          =   615
      Left            =   4920
      Picture         =   "PersonaDatosTransferencia.frx":255C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3060
      Width           =   615
   End
   Begin VB.CommandButton cmdLeftToRight 
      Height          =   615
      Left            =   4920
      Picture         =   "PersonaDatosTransferencia.frx":2E26
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2340
      Width           =   615
   End
   Begin VB.Frame fraViajeRight 
      Height          =   1515
      Left            =   5640
      TabIndex        =   11
      Top             =   60
      Width           =   4695
      Begin VB.CommandButton cmdSelectRight 
         Caption         =   "Seleccionar"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   4455
      End
      Begin VB.TextBox txtApellidoRight 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   660
         Width           =   3570
      End
      Begin VB.TextBox txtNombreRight 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3570
      End
      Begin VB.Label lblApellidoRight 
         AutoSize        =   -1  'True
         Caption         =   "Apellido:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   600
      End
      Begin VB.Label lblNombreRight 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1140
         Width           =   600
      End
   End
   Begin VB.Frame fraViajeLeft 
      Height          =   1515
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   4695
      Begin VB.TextBox txtApellidoLeft 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   660
         Width           =   3570
      End
      Begin VB.CommandButton cmdSelectLeft 
         Caption         =   "Seleccionar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   4455
      End
      Begin VB.TextBox txtNombreLeft 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3570
      End
      Begin VB.Label lblNombreLeft 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   600
      End
      Begin VB.Label lblApellidoLeft 
         AutoSize        =   -1  'True
         Caption         =   "Apellido:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   600
      End
   End
   Begin MSComctlLib.ListView lvwDataLeft 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Orden"
         Text            =   "Orden"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Persona"
         Text            =   "Persona"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Estado"
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Asiento"
         Text            =   "Asiento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Origen"
         Text            =   "Origen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Destino"
         Text            =   "Destino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Key             =   "Tipo"
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDataRight 
      Height          =   3375
      Left            =   5640
      TabIndex        =   4
      Top             =   1680
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5953
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Orden"
         Text            =   "Orden"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Persona"
         Text            =   "Persona"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Estado"
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Asiento"
         Text            =   "Asiento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Origen"
         Text            =   "Origen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Destino"
         Text            =   "Destino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Key             =   "Tipo"
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmPersonaDatosTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFechaHoraLeft As Date
Private mIDRutaLeft As String

Private mFechaHoraRight As Date
Private mIDRutaRight As String

Public FormWaitingForSelect As String

Public Function FillListViewLeft(ByVal FechaHora As Date, ByVal IDRuta As String, ByVal Indice As Long) As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim SQL_Statement As String
    Dim ViajeDetalle As ViajeDetalle
    Dim EstadoKey As String
    Dim UltimoEstado As String
    
    If FechaHora <> mFechaHoraLeft Or IDRuta <> mIDRutaLeft Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Indice = 0 Then
        If Not lvwDataLeft.SelectedItem Is Nothing Then
            KeySave = lvwDataLeft.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & Indice
    End If
    
    SQL_Statement = "SELECT ViajeDetalle.Indice, ViajeDetalle.Estado, ViajeDetalle.Asiento, ViajeDetalle.Orden, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, Lugar_Origen.Nombre AS Origen, ViajeDetalle.Sube, Lugar_Destino.Nombre AS Destino, ViajeDetalle.Baja, ViajeDetalle.ReservaTipo "
    SQL_Statement = SQL_Statement & "FROM ((ViajeDetalle INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona) INNER JOIN Lugar AS Lugar_Origen ON ViajeDetalle.IDOrigen = Lugar_Origen.IDLugar) INNER JOIN Lugar AS Lugar_Destino ON ViajeDetalle.IDDestino = Lugar_Destino.IDLugar "
    SQL_Statement = SQL_Statement & "WHERE convert(char(10), ViajeDetalle.FechaHora, 111) + ' ' + convert(char(8), ViajeDetalle.FechaHora, 108) = '" & Format(FechaHora, "yyyy/mm/dd hh:nn:ss") & "' AND ViajeDetalle.IDRuta = '" & ReplaceQuote(IDRuta) & "' AND ViajeDetalle.OcupanteTipo = '" & OCUPANTE_TIPO_PASAJERO & "' "
    SQL_Statement = SQL_Statement & "ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Estado, ViajeDetalle.Orden"
    
    lvwDataLeft.ListItems.Clear
    Set recData = New ADODB.Recordset
    recData.Open SQL_Statement, pDBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set ViajeDetalle = New ViajeDetalle
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Select Case .Fields("Estado").Value
                    Case VIAJE_DETALLE_ESTADO_CONFIRMADO
                        EstadoKey = "CONFIRMADO"
                    Case VIAJE_DETALLE_ESTADO_CONDICIONAL
                        EstadoKey = "CONDICIONAL"
                    Case VIAJE_DETALLE_ESTADO_CANCELADO
                        EstadoKey = "CANCELADO"
                End Select
                
                '//////////////////////////////////////////////////
                'ROWS SEPARATOR BY STATUS
                If pParametro.ViajeDetalle_SeparateRowsByStatus Then
                    If UltimoEstado <> .Fields("Estado").Value Then
                        If UltimoEstado <> "" Then
                            Set ListItem = lvwDataLeft.ListItems.Add(, KEY_STRINGER & (.Fields("Indice").Value * -1), "")
                        End If
                        UltimoEstado = .Fields("Estado").Value
                    End If
                End If
                
                'Pasajero
                Set ListItem = lvwDataLeft.ListItems.Add(, KEY_STRINGER & .Fields("Indice").Value, Val(.Fields("Orden").Value & ""), , "PASAJERO_" & EstadoKey)
                ListItem.SubItems(1) = .Fields("Persona").Value
                ViajeDetalle.Estado = .Fields("Estado").Value & ""
                ListItem.SubItems(2) = ViajeDetalle.Estado_ToString
                ListItem.SubItems(3) = .Fields("Asiento").Value & ""
                ListItem.SubItems(4) = IIf(IsNull(.Fields("Sube").Value), .Fields("Origen").Value, .Fields("Sube").Value)
                ListItem.SubItems(5) = IIf(IsNull(.Fields("Baja").Value), .Fields("Destino").Value, .Fields("Baja").Value)
                ViajeDetalle.ReservaTipo = .Fields("ReservaTipo").Value
                ListItem.SubItems(6) = ViajeDetalle.ReservaTipo_ToString
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set recData = Nothing
    
    Set ViajeDetalle = Nothing
    
    On Error Resume Next
    Set lvwDataLeft.SelectedItem = lvwDataLeft.ListItems(KeySave)
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwDataLeft.SetFocus
    End If
    
    FillListViewLeft = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.ViajeDetalleTransferencia.FillListViewLeft", "Error al obtener el Detalle del Viaje."
End Function

Public Function FillListViewRight(ByVal FechaHora As Date, ByVal IDRuta As String, ByVal Indice As Long) As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim SQL_Statement As String
    Dim ViajeDetalle As ViajeDetalle
    Dim EstadoKey As String
    Dim UltimoEstado As String
    
    If FechaHora <> mFechaHoraRight Or IDRuta <> mIDRutaRight Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Indice = 0 Then
        If Not lvwDataRight.SelectedItem Is Nothing Then
            KeySave = lvwDataRight.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & Indice
    End If
    
    SQL_Statement = "SELECT ViajeDetalle.Indice, ViajeDetalle.Estado, ViajeDetalle.Asiento, ViajeDetalle.Orden, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, Lugar_Origen.Nombre AS Origen, ViajeDetalle.Sube, Lugar_Destino.Nombre AS Destino, ViajeDetalle.Baja, ViajeDetalle.ReservaTipo "
    SQL_Statement = SQL_Statement & "FROM ((ViajeDetalle INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona) INNER JOIN Lugar AS Lugar_Origen ON ViajeDetalle.IDOrigen = Lugar_Origen.IDLugar) INNER JOIN Lugar AS Lugar_Destino ON ViajeDetalle.IDDestino = Lugar_Destino.IDLugar "
    SQL_Statement = SQL_Statement & "WHERE convert(char(10), ViajeDetalle.FechaHora, 111) + ' ' + convert(char(8), ViajeDetalle.FechaHora, 108) = '" & Format(FechaHora, "yyyy/mm/dd hh:nn:ss") & "' AND ViajeDetalle.IDRuta = '" & ReplaceQuote(IDRuta) & "' AND ViajeDetalle.OcupanteTipo = '" & OCUPANTE_TIPO_PASAJERO & "' "
    SQL_Statement = SQL_Statement & "ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Estado, ViajeDetalle.Orden"
    
    lvwDataRight.ListItems.Clear
    Set recData = New ADODB.Recordset
    recData.Open SQL_Statement, pDBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set ViajeDetalle = New ViajeDetalle
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Select Case .Fields("Estado").Value
                    Case VIAJE_DETALLE_ESTADO_CONFIRMADO
                        EstadoKey = "CONFIRMADO"
                    Case VIAJE_DETALLE_ESTADO_CONDICIONAL
                        EstadoKey = "CONDICIONAL"
                    Case VIAJE_DETALLE_ESTADO_CANCELADO
                        EstadoKey = "CANCELADO"
                End Select
                
                '//////////////////////////////////////////////////
                'ROWS SEPARATOR BY STATUS
                If pParametro.ViajeDetalle_SeparateRowsByStatus Then
                    If UltimoEstado <> .Fields("Estado").Value Then
                        If UltimoEstado <> "" Then
                            Set ListItem = lvwDataRight.ListItems.Add(, KEY_STRINGER & (.Fields("Indice").Value * -1), "")
                        End If
                        UltimoEstado = .Fields("Estado").Value
                    End If
                End If
                
                'Pasajero
                Set ListItem = lvwDataRight.ListItems.Add(, KEY_STRINGER & .Fields("Indice").Value, Val(.Fields("Orden").Value & ""), , "PASAJERO_" & EstadoKey)
                ListItem.SubItems(1) = .Fields("Persona").Value
                ViajeDetalle.Estado = .Fields("Estado").Value & ""
                ListItem.SubItems(2) = ViajeDetalle.Estado_ToString
                ListItem.SubItems(3) = .Fields("Asiento").Value & ""
                ListItem.SubItems(4) = IIf(IsNull(.Fields("Sube").Value), .Fields("Origen").Value, .Fields("Sube").Value)
                ListItem.SubItems(5) = IIf(IsNull(.Fields("Baja").Value), .Fields("Destino").Value, .Fields("Baja").Value)
                ViajeDetalle.ReservaTipo = .Fields("ReservaTipo").Value
                ListItem.SubItems(6) = ViajeDetalle.ReservaTipo_ToString
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set recData = Nothing
    
    Set ViajeDetalle = Nothing
    
    On Error Resume Next
    Set lvwDataRight.SelectedItem = lvwDataRight.ListItems(KeySave)
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwDataRight.SetFocus
    End If
    
    FillListViewRight = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.ViajeDetalleTransferencia.FillListViewRight", "Error al obtener el Detalle del Viaje."
End Function

Private Sub cmdBoth_Click()
    Dim ViajeDetalle As ViajeDetalle
    Dim IndiceLeft As Long
    Dim IndiceRight As Long
    
    If mIDRutaLeft = "" Then
        MsgBox "Debe seleccionar el Viaje de la Izquierda.", vbInformation, App.Title
        cmdSelectLeft.SetFocus
        Exit Sub
    End If
    If mIDRutaRight = "" Then
        MsgBox "Debe seleccionar el Viaje de la Derecha.", vbInformation, App.Title
        cmdSelectRight.SetFocus
        Exit Sub
    End If
    If mIDRutaLeft <> mIDRutaRight Then
        MsgBox "La Ruta de ambos Viajes debe ser la misma.", vbExclamation, App.Title
        cmdSelectRight.SetFocus
        Exit Sub
    End If
    If mFechaHoraLeft = mFechaHoraRight Then
        MsgBox "El Viaje seleccionado a la Izquierda es el mismo que el seleccionado a la Derecha.", vbInformation, App.Title
        cmdSelectRight.SetFocus
        Exit Sub
    End If
    
    If lvwDataLeft.SelectedItem Is Nothing Then
        MsgBox "No hay ninguna Reserva seleccionada en el Viaje de la Izquierda.", vbInformation, App.Title
        lvwDataLeft.SetFocus
        Exit Sub
    End If
    If Val(Mid(lvwDataLeft.SelectedItem.Key, 2)) < 0 Then
        MsgBox "No hay ninguna Reserva seleccionada en el Viaje de la Izquierda.", vbInformation, App.Title
        lvwDataLeft.SetFocus
        Exit Sub
    End If
    If lvwDataRight.SelectedItem Is Nothing Then
        MsgBox "No hay ninguna Reserva seleccionada en el Viaje de la Derecha.", vbInformation, App.Title
        lvwDataRight.SetFocus
        Exit Sub
    End If
    If Val(Mid(lvwDataRight.SelectedItem.Key, 2)) < 0 Then
        MsgBox "No hay ninguna Reserva seleccionada en el Viaje de la Derecha.", vbInformation, App.Title
        lvwDataRight.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Desea intercambiar los Pasajeros seleccionados?", vbExclamation + vbYesNo, App.Title) = vbYes Then
        IndiceLeft = Val(Mid(lvwDataLeft.SelectedItem.Key, Len(KEY_STRINGER) + 1))
        IndiceRight = Val(Mid(lvwDataRight.SelectedItem.Key, Len(KEY_STRINGER) + 1))
    
        Set ViajeDetalle = New ViajeDetalle
        ViajeDetalle.RefreshList = False
        ViajeDetalle.FechaHora = mFechaHoraLeft
        ViajeDetalle.IDRuta = mIDRutaLeft
        ViajeDetalle.Indice = IndiceLeft
        If ViajeDetalle.Load() Then
            ViajeDetalle.FechaHora = mFechaHoraRight
            ViajeDetalle.IDRuta = mIDRutaRight
            ViajeDetalle.Update
            IndiceLeft = ViajeDetalle.Indice
        End If
        Set ViajeDetalle = Nothing
    
        Set ViajeDetalle = New ViajeDetalle
        ViajeDetalle.RefreshList = False
        ViajeDetalle.FechaHora = mFechaHoraRight
        ViajeDetalle.IDRuta = mIDRutaRight
        ViajeDetalle.Indice = IndiceRight
        If ViajeDetalle.Load() Then
            ViajeDetalle.FechaHora = mFechaHoraLeft
            ViajeDetalle.IDRuta = mIDRutaLeft
            ViajeDetalle.Update
            IndiceRight = ViajeDetalle.Indice
        End If
        Set ViajeDetalle = Nothing
        
        RefreshList_RefreshViajeDetalle mFechaHoraRight, mIDRutaRight, IndiceLeft
        RefreshList_RefreshViajeDetalle mFechaHoraLeft, mIDRutaLeft, IndiceRight
    End If
End Sub

Private Sub cmdLeftToRight_Click()
    Dim ViajeDetalle As ViajeDetalle
    
    If mIDRutaLeft = "" Then
        MsgBox "Debe seleccionar el Viaje de Origen.", vbInformation, App.Title
        cmdSelectLeft.SetFocus
        Exit Sub
    End If
    If mIDRutaRight = "" Then
        MsgBox "Debe seleccionar el Viaje de Destino.", vbInformation, App.Title
        cmdSelectRight.SetFocus
        Exit Sub
    End If
    If mIDRutaLeft <> mIDRutaRight Then
        MsgBox "La Ruta del Viaje de Origen debe ser la misma que la del Viaje de Destino.", vbExclamation, App.Title
        cmdSelectRight.SetFocus
        Exit Sub
    End If
    If mFechaHoraLeft = mFechaHoraRight Then
        MsgBox "El Viaje seleccionado a la Izquierda es el mismo que el seleccionado a la Derecha.", vbInformation, App.Title
        cmdSelectRight.SetFocus
        Exit Sub
    End If
    
    If lvwDataLeft.SelectedItem Is Nothing Then
        MsgBox "No hay ninguna Reserva seleccionada para Transferir.", vbInformation, App.Title
        lvwDataLeft.SetFocus
        Exit Sub
    End If
    If Val(Mid(lvwDataLeft.SelectedItem.Key, 2)) < 0 Then
        MsgBox "No hay ninguna Reserva seleccionada para Transferir.", vbInformation, App.Title
        lvwDataLeft.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Desea Transferir la Reserva seleccionada en el Viaje de la Izquierda al Viaje de la Derecha?", vbExclamation + vbYesNo, App.Title) = vbYes Then
        Set ViajeDetalle = New ViajeDetalle
        ViajeDetalle.RefreshList = False
        ViajeDetalle.FechaHora = mFechaHoraLeft
        ViajeDetalle.IDRuta = mIDRutaLeft
        ViajeDetalle.Indice = Val(Mid(lvwDataLeft.SelectedItem.Key, Len(KEY_STRINGER) + 1))
        If ViajeDetalle.Load() Then
            ViajeDetalle.FechaHora = mFechaHoraRight
            ViajeDetalle.IDRuta = mIDRutaRight
            ViajeDetalle.Update
            RefreshList_RefreshViajeDetalle mFechaHoraLeft, mIDRutaLeft, 0
            RefreshList_RefreshViajeDetalle mFechaHoraRight, mIDRutaRight, ViajeDetalle.Indice
            lvwDataRight.SetFocus
        End If
        Set ViajeDetalle = Nothing
    End If
End Sub

Private Sub cmdRightToLeft_Click()
    Dim ViajeDetalle As ViajeDetalle

    If mIDRutaRight = "" Then
        MsgBox "Debe seleccionar el Viaje de Origen.", vbInformation, App.Title
        cmdSelectRight.SetFocus
        Exit Sub
    End If
    If mIDRutaLeft = "" Then
        MsgBox "Debe seleccionar el Viaje de Destino.", vbInformation, App.Title
        cmdSelectLeft.SetFocus
        Exit Sub
    End If
    If mIDRutaLeft <> mIDRutaRight Then
        MsgBox "La Ruta del Viaje de Origen debe ser la misma que la del Viaje de Destino.", vbExclamation, App.Title
        cmdSelectLeft.SetFocus
        Exit Sub
    End If
    If mFechaHoraLeft = mFechaHoraRight Then
        MsgBox "El Viaje seleccionado a la Izquierda es el mismo que el seleccionado a la Derecha.", vbInformation, App.Title
        cmdSelectRight.SetFocus
        Exit Sub
    End If
    
    If lvwDataRight.SelectedItem Is Nothing Then
        MsgBox "No hay ninguna Reserva seleccionada para Transferir.", vbInformation, App.Title
        lvwDataRight.SetFocus
        Exit Sub
    End If
    If Val(Mid(lvwDataRight.SelectedItem.Key, 2)) < 0 Then
        MsgBox "No hay ninguna Reserva seleccionada para Transferir.", vbInformation, App.Title
        lvwDataRight.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Desea Transferir la Reserva seleccionada en el Viaje de la Derecha al Viaje de la Izquierda?", vbExclamation + vbYesNo, App.Title) = vbYes Then
        Set ViajeDetalle = New ViajeDetalle
        ViajeDetalle.RefreshList = False
        ViajeDetalle.FechaHora = mFechaHoraRight
        ViajeDetalle.IDRuta = mIDRutaRight
        ViajeDetalle.Indice = Val(Mid(lvwDataRight.SelectedItem.Key, Len(KEY_STRINGER) + 1))
        If ViajeDetalle.Load() Then
            ViajeDetalle.FechaHora = mFechaHoraLeft
            ViajeDetalle.IDRuta = mIDRutaLeft
            ViajeDetalle.Update
            RefreshList_RefreshViajeDetalle mFechaHoraLeft, mIDRutaLeft, ViajeDetalle.Indice
            RefreshList_RefreshViajeDetalle mFechaHoraRight, mIDRutaRight, 0
            lvwDataLeft.SetFocus
        End If
        Set ViajeDetalle = Nothing
    End If
End Sub

Private Sub cmdSelectLeft_Click()
    Screen.MousePointer = vbHourglass
    frmViaje.Show
    If frmViaje.WindowState = vbMinimized Then
        frmViaje.WindowState = vbNormal
    End If
    frmViaje.FormWaitingForSelect = Me.Name
    frmViaje.SelectTag = "LEFT"
    frmViaje.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelectRight_Click()
    Screen.MousePointer = vbHourglass
    frmViaje.Show
    If frmViaje.WindowState = vbMinimized Then
        frmViaje.WindowState = vbNormal
    End If
    frmViaje.FormWaitingForSelect = Me.Name
    frmViaje.SelectTag = "RIGHT"
    frmViaje.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    lvwDataLeft.GridLines = pParametro.ListView_GridLines
    lvwDataRight.GridLines = pParametro.ListView_GridLines

    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwDataLeft.SmallIcons = ilsData
    Set lvwDataRight.SmallIcons = ilsData
    '//////////////////////////////////////////////////////////
    
    ResizeAndPositionForm Me
    GetListViewSettings "ViajeDetalleTransferencia_Left", lvwDataLeft
    CopyListViewSetting lvwDataLeft, lvwDataRight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    SaveListViewSettings "ViajeDetalleTransferencia", lvwDataLeft
    Set frmViajeDetalleTransferencia = Nothing
End Sub

Private Sub Form_Resize()
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
        
    fraViajeLeft.Top = CONTROL_SPACE
    fraViajeLeft.Left = CONTROL_SPACE
    
    lvwDataLeft.Top = fraViajeLeft.Top + fraViajeLeft.Height + CONTROL_SPACE
    lvwDataLeft.Left = fraViajeLeft.Left
    lvwDataLeft.Height = ScaleHeight - lvwDataLeft.Top - CONTROL_SPACE
    lvwDataLeft.Width = (ScaleWidth - cmdLeftToRight.Width - (CONTROL_SPACE * 4)) / 2
    
    cmdBoth.Top = lvwDataLeft.Top + ((lvwDataLeft.Height - cmdBoth.Height - CONTROL_SPACE - cmdLeftToRight.Height - CONTROL_SPACE - cmdRightToLeft.Height) / 2)
    cmdBoth.Left = lvwDataLeft.Left + lvwDataLeft.Width + CONTROL_SPACE
    
    cmdLeftToRight.Top = cmdBoth.Top + cmdBoth.Height + CONTROL_SPACE
    cmdLeftToRight.Left = cmdBoth.Left

    cmdRightToLeft.Top = cmdLeftToRight.Top + cmdLeftToRight.Height + CONTROL_SPACE
    cmdRightToLeft.Left = cmdBoth.Left

    fraViajeRight.Top = CONTROL_SPACE
    fraViajeRight.Left = ScaleWidth - lvwDataLeft.Width - CONTROL_SPACE
    
    lvwDataRight.Top = lvwDataLeft.Top
    lvwDataRight.Left = fraViajeRight.Left
    lvwDataRight.Height = lvwDataLeft.Height
    lvwDataRight.Width = lvwDataLeft.Width
End Sub

Public Sub ViajeSelected(ByRef Viaje As Viaje, ByVal SelectTag As String)
    If SelectTag = "LEFT" Then
        mFechaHoraLeft = Viaje.FechaHora
        mIDRutaLeft = Viaje.IDRuta
        
        txtFechaDiaSemanaLeft.Text = Viaje.FechaHora_WeekdayName
        txtFechaLeft.Text = Viaje.FechaHora_FormattedAsDate
        txtHoraLeft.Text = Viaje.FechaHora_FormattedAsTime
        txtRutaLeft.Text = Viaje.Ruta_DisplayName
        
        FillListViewLeft mFechaHoraLeft, mIDRutaLeft, 0
    Else
        mFechaHoraRight = Viaje.FechaHora
        mIDRutaRight = Viaje.IDRuta
        
        txtFechaDiaSemanaRight.Text = Viaje.FechaHora_WeekdayName
        txtFechaRight.Text = Viaje.FechaHora_FormattedAsDate
        txtHoraRight.Text = Viaje.FechaHora_FormattedAsTime
        txtRutaRight.Text = Viaje.Ruta_DisplayName
        
        FillListViewRight mFechaHoraRight, mIDRutaRight, 0
    End If
    
    Set Viaje = Nothing
End Sub

Private Sub tmrListViewSettingsUpdate_Timer()
    CopyListViewSetting lvwDataLeft, lvwDataRight
End Sub
