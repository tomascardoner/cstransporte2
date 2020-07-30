VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RutaDetalleHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDRutaDetalleHorario As Long
Private mIDRuta As String
Private mIDLugar As Long
Private mDiaSemanaNumero As Long
Private mDiaSemana As String
Private mHoraInicio As Date
Private mHoraFin As Date

'STATE VARIABLES
Private mIsNew As Boolean
Private mIsCopy As Boolean
Private mIsDirty As Boolean
Private mNoMatch As Boolean

'BEHAVIOR VARIABLES
Private mNoMatchRaiseError As Boolean
Private mRefreshListSkip As Boolean

'INTERNAL VARIABLES

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDRutaDetalleHorario() As Long
    IDRutaDetalleHorario = mIDRutaDetalleHorario
End Property

Public Property Let IDRutaDetalleHorario(ByVal value As Long)
    If value <> mIDRutaDetalleHorario Then
        mIsDirty = True
    End If
    mIDRutaDetalleHorario = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDRuta() As String
    IDRuta = mIDRuta
End Property

Public Property Let IDRuta(ByVal value As String)
    If value <> mIDRuta Then
        mIsDirty = True
    End If
    mIDRuta = value
End Property

Public Property Get Ruta() As Ruta
    Static oRuta As Ruta
    
    If mIDRuta = 0 Then
        Set Ruta = Nothing
    Else
        If oRuta Is Nothing Then
            Set oRuta = New Ruta
        End If
        If oRuta.IDRuta <> mIDRuta Then
            oRuta.IDRuta = mIDRuta
            Call oRuta.Load
        End If
        Set Ruta = oRuta
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDLugar() As Long
    IDLugar = mIDLugar
End Property

Public Property Let IDLugar(ByVal value As Long)
    If value <> mIDLugar Then
        mIsDirty = True
    End If
    mIDLugar = value
End Property

Public Property Get Lugar() As Lugar
    Static oLugar As Lugar
    
    If mIDLugar = 0 Then
        Set Lugar = Nothing
    Else
        If oLugar Is Nothing Then
            Set oLugar = New Lugar
        End If
        If oLugar.IDLugar <> mIDLugar Then
            oLugar.IDLugar = mIDLugar
            Call oLugar.Load
        End If
        Set Lugar = oLugar
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DiaSemanaNumero() As Long
    DiaSemanaNumero = mDiaSemanaNumero
End Property

Public Property Let DiaSemanaNumero(ByVal value As Long)
    If value <> mDiaSemanaNumero Then
        mIsDirty = True
    End If
    mDiaSemanaNumero = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DiaSemana() As String
    DiaSemana = mDiaSemana
End Property

Public Property Let DiaSemana(ByVal value As String)
    If value <> mDiaSemana Then
        mIsDirty = True
    End If
    mDiaSemana = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get HoraInicio() As Date
    HoraInicio = mHoraInicio
End Property

Public Property Let HoraInicio(ByVal value As Date)
    If value <> mHoraInicio Then
        mIsDirty = True
    End If
    mHoraInicio = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get HoraFin() As Date
    HoraFin = mHoraFin
End Property

Public Property Let HoraFin(ByVal value As Date)
    If value <> mHoraFin Then
        mIsDirty = True
    End If
    mHoraFin = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IsNew() As Boolean
    IsNew = mIsNew
End Property

Public Property Get IsCopy() As Boolean
    IsCopy = mIsCopy
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = mIsDirty
End Property

Public Sub MakeDirty()
    mIsDirty = True
End Sub

Public Property Get NoMatch() As Boolean
    NoMatch = mNoMatch
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get NoMatchRaiseError() As Boolean
    NoMatchRaiseError = mNoMatchRaiseError
End Property

Public Property Let NoMatchRaiseError(ByVal value As Boolean)
    mNoMatchRaiseError = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RefreshListSkip() As Boolean
    RefreshListSkip = mRefreshListSkip
End Property

Public Property Let RefreshListSkip(ByVal value As Boolean)
    mRefreshListSkip = value
End Property

Public Sub RefreshList()
    RefreshList_Module.RefreshList_RefreshRutaDetalleHorario IDRuta, IDLugar, IDRutaDetalleHorario
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function Load() As Boolean
    Dim cmdSP As ADODB.command
    Dim recData As ADODB.Recordset

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Call InitializeValues

    Set cmdSP = New ADODB.command
    With cmdSP
        Set .ActiveConnection = pDatabase.Connection
        .CommandText = "usp_RutaDetalleHorario_Get"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDRutaDetalleHorario", adInteger, adParamInput, , mIDRutaDetalleHorario)
    End With

    Set recData = New ADODB.Recordset
    recData.Open cmdSP, , adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

    Screen.MousePointer = vbHourglass

    Load = GetDataFromRecordset(recData)

    recData.Close
    Set recData = Nothing
    Set cmdSP = Nothing
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.RutaDetalleHorario.Load", "Error al obtener los datos del Horario del Detalle de Ruta."
End Function

Private Function GetDataFromRecordset(ByRef recData As ADODB.Recordset) As Boolean

    If (Not mNoMatchRaiseError) And recData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False

        mIDRutaDetalleHorario = recData("IDRutaDetalleHorario").value
        mIDRuta = Trim(recData("IDRuta").value)
        mIDLugar = recData("IDLugar").value
        mDiaSemanaNumero = recData("DiaSemanaNumero").value
        mDiaSemana = recData("DiaSemana").value
        mHoraInicio = recData("HoraInicio").value
        mHoraFin = recData("HoraFin").value
    End If

    Screen.MousePointer = vbDefault
    GetDataFromRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.RutaDetalleHorario.GetDataFromRecordset", "Error al obtener los datos del Horario del Detalle de Ruta."
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDRutaDetalleHorario = 0

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function Update() As Boolean
    Dim cmdSP As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set cmdSP = New ADODB.command
    With cmdSP
        Set .ActiveConnection = pDatabase.Connection
        If mIsNew Then
            .CommandText = "usp_RutaDetalleHorario_Add"
            .Parameters.Append .CreateParameter("IDRutaDetalleHorario", adInteger, adParamOutput)
        Else
            .CommandText = "usp_RutaDetalleHorario_Update"
            .Parameters.Append .CreateParameter("IDRutaDetalleHorario", adInteger, adParamInput, , mIDRutaDetalleHorario)
        End If
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDRuta", adChar, adParamInput, 20, mIDRuta)
        .Parameters.Append .CreateParameter("IDLugar", adInteger, adParamInput, , mIDLugar)
        .Parameters.Append .CreateParameter("DiaSemanaNumero", adInteger, adParamInput, , mDiaSemanaNumero)
        .Parameters.Append .CreateParameter("DiaSemana", adVarChar, adParamInput, 50, mDiaSemana)
        .Parameters.Append .CreateParameter("HoraInicio", adDate, adParamInput, , mHoraInicio)
        .Parameters.Append .CreateParameter("HoraFin", adDate, adParamInput, , mHoraFin)
        If mIsNew Then
            mIDRutaDetalleHorario = .Parameters("IDRutaDetalleHorario").value
        End If

        .Execute
    End With

    mIsCopy = False
    mIsDirty = False

    If Not mRefreshListSkip Then
        Call RefreshList
    End If

    Screen.MousePointer = vbDefault
    Update = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.RutaLugarGrupo.Update", "Error al actualizar el Horario del Detalle de Ruta."
End Function

Public Function Delete() As Boolean
    Dim cmdSP As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set cmdSP = New ADODB.command
    With cmdSP
        Set .ActiveConnection = pDatabase.Connection
        .CommandText = "usp_RutaDetalleHorario_Delete"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDRutaDetalleHorario", adInteger, adParamInput, , mIDRutaDetalleHorario)

        .Execute
    End With

    If Not mRefreshListSkip Then
        Call RefreshList
    End If

    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.RutaLugarGrupo.Delete", "Error al eliminar el Horario del Detalle de Ruta."
End Function

Public Function VerificarDisponiblilidad(ByVal Fecha As Date, ByVal Hora As Date) As Boolean
    
End Function

Private Sub InitializeValues()
    mIDRuta = ""
    mIDLugar = 0
    mDiaSemanaNumero = -1
    mDiaSemana = ""
    mHoraInicio = DATE_TIME_FIELD_NULL_VALUE
    mHoraFin = DATE_TIME_FIELD_NULL_VALUE

    mIsNew = True
    mIsCopy = False
    mIsDirty = False
End Sub

Private Sub Class_Initialize()
    Call InitializeValues
    mNoMatchRaiseError = True
    mRefreshListSkip = False
End Sub

Private Sub Class_Terminate()
    'TERMINATE CODE
End Sub
