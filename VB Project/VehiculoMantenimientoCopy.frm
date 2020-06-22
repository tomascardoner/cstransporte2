VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVehiculoMantenimientoCopy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Mantenimiento de Vehículos"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   Icon            =   "VehiculoMantenimientoCopy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
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
      Left            =   2100
      TabIndex        =   6
      Top             =   2220
      Width           =   1215
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
      Left            =   3420
      TabIndex        =   7
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton cmdDestination 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Vehículos"
      Top             =   1560
      Width           =   255
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   8
      Top             =   780
      Width           =   4515
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Vehículos"
      Top             =   1020
      Width           =   255
   End
   Begin MSDataListLib.DataCombo datcboSource 
      Height          =   330
      Left            =   1020
      TabIndex        =   1
      Top             =   1020
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
   Begin MSDataListLib.DataCombo datcboDestination 
      Height          =   330
      Left            =   1020
      TabIndex        =   4
      Top             =   1560
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
   Begin VB.Label lblDestination 
      AutoSize        =   -1  'True
      Caption         =   "&Destino:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1620
      Width           =   585
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "VehiculoMantenimientoCopy.frx":054A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Copie aquí los Datos del Mantenimiento del Vehículo"
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
      TabIndex        =   9
      Top             =   300
      Width           =   3720
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "&Origen:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   525
   End
End
Attribute VB_Name = "frmVehiculoMantenimientoCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload frmVehiculoMantenimientoCopy
End Sub

Private Sub cmdDestination_Click()
    If pCPermiso.GotPermission(PERMISO_VEHICULO) Then
        Screen.MousePointer = vbHourglass
        frmVehiculo.Show
        On Error Resume Next
        Set frmVehiculo.lvwData.SelectedItem = frmVehiculo.lvwData.ListItems(KEY_STRINGER & datcboDestination.BoundText)
        frmVehiculo.lvwData.SelectedItem.EnsureVisible
        If frmVehiculo.WindowState = vbMinimized Then
            frmVehiculo.WindowState = vbNormal
        End If
        frmVehiculo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdOK_Click()
    Dim cmdData As ADODB.command
    
    If datcboSource.BoundText = "" Then
        MsgBox "Debe seleccionar el Vehículo de Origen.", vbInformation, App.Title
        datcboSource.SetFocus
        Exit Sub
    End If
    If datcboDestination.BoundText = "" Then
        MsgBox "Debe seleccionar el Vehículo de Origen.", vbInformation, App.Title
        datcboDestination.SetFocus
        Exit Sub
    End If
    If datcboSource.BoundText = datcboDestination.BoundText Then
        MsgBox "El Vehículo de Origen y el de Destino, no pueden ser el mismo.", vbInformation, App.Title
        datcboDestination.SetFocus
        Exit Sub
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_VehiculoMantenimiento_Copy"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDVehiculoOrigen", adInteger, adParamInput, , Val(datcboSource.BoundText))
    cmdData.Parameters.Append cmdData.CreateParameter("IDVehiculoDestino", adInteger, adParamInput, , Val(datcboDestination.BoundText))
    cmdData.Parameters.Append cmdData.CreateParameter("IDUsuario", adChar, adParamInput, 30, pUsuario.IDUsuario)
    cmdData.Execute
    Set cmdData = Nothing
    
    RefreshList_RefreshVehiculoMantenimiento 0, 0
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Se han Copiado los Mantenimiento del Vehículo.", vbInformation, App.Title
    
    Unload frmVehiculoMantenimientoCopy
    
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.VehiculoMantenimientoCopy.OK", "Error al Copiar los Mantenimientos del Vehículo." & vbCr & vbCr & "IDVehiculoOrigen: " & Val(datcboSource.BoundText) & vbCr & "IDVehiculoDestino: " & Val(datcboDestination.BoundText)
    Set cmdData = Nothing
End Sub

Private Sub cmdSource_Click()
    If pCPermiso.GotPermission(PERMISO_VEHICULO) Then
        Screen.MousePointer = vbHourglass
        frmVehiculo.Show
        On Error Resume Next
        Set frmVehiculo.lvwData.SelectedItem = frmVehiculo.lvwData.ListItems(KEY_STRINGER & datcboSource.BoundText)
        frmVehiculo.lvwData.SelectedItem.EnsureVisible
        If frmVehiculo.WindowState = vbMinimized Then
            frmVehiculo.WindowState = vbNormal
        End If
        frmVehiculo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    If Not CSM_Control_DataCombo.FillFromSQL(datcboSource, "SELECT IDVehiculo, Nombre FROM Vehiculo WHERE Activo = 1 ORDER BY Nombre", "IDVehiculo", "Nombre", "Vehículos", cscpFirst) Then
        Unload Me
        Exit Sub
    End If
    If Not CSM_Control_DataCombo.FillFromSQL(datcboDestination, "SELECT IDVehiculo, Nombre FROM Vehiculo WHERE Activo = 1 ORDER BY Nombre", "IDVehiculo", "Nombre", "Vehículos", cscpFirst) Then
        Unload Me
        Exit Sub
    End If

    Top = (frmMDI.ScaleHeight - Height) / 2
    Left = (frmMDI.ScaleWidth - Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVehiculoMantenimientoCopy = Nothing
End Sub
