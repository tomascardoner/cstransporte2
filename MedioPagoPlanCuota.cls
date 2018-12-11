VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MedioPagoPlanCuota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDMedioPagoPlan As Byte
Private mCuota As Byte
Private mCoeficiente As Currency
Private mCoeficientePrepago As Currency

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

Public Property Get IDMedioPagoPlan() As Byte
    IDMedioPagoPlan = mIDMedioPagoPlan
End Property

Public Property Let IDMedioPagoPlan(ByVal Value As Byte)
    If Value <> mIDMedioPagoPlan Then
        mIsDirty = True
    End If
    mIDMedioPagoPlan = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Cuota() As Byte
    Cuota = mCuota
End Property

Public Property Let Cuota(ByVal Value As Byte)
    If Value <> mCuota Then
        mIsDirty = True
    End If
    mCuota = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Coeficiente() As Currency
    Coeficiente = mCoeficiente
End Property

Public Property Let Coeficiente(ByVal Value As Currency)
    If Value <> mCoeficiente Then
        mIsDirty = True
    End If
    mCoeficiente = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get CoeficientePrepago() As Currency
    CoeficientePrepago = mCoeficientePrepago
End Property

Public Property Let CoeficientePrepago(ByVal Value As Currency)
    If Value <> mCoeficientePrepago Then
        mIsDirty = True
    End If
    mCoeficientePrepago = Value
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

Public Property Let NoMatchRaiseError(ByVal Value As Boolean)
    mNoMatchRaiseError = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RefreshListSkip() As Boolean
    RefreshListSkip = mRefreshListSkip
End Property

Public Property Let RefreshListSkip(ByVal Value As Boolean)
    mRefreshListSkip = Value
End Property

Public Sub RefreshList()
    'RefreshList_Module.MedioPagoPlanCuota IDMedioPagoPlan, Cuota
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
        .CommandText = "usp_MedioPagoPlanCuota_Get"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDMedioPagoPlan", adTinyInt, adParamInput, , mIDMedioPagoPlan)
        .Parameters.Append .CreateParameter("Cuota", adTinyInt, adParamInput, , mCuota)
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
    ShowErrorMessage "Classes.MedioPagoPlanCuota.Load", "Error al obtener los datos de la Cuota del Plan de Medios de Pago." & vbCr & vbCr & "IDMedioPagoPlan: " & IDMedioPagoPlan & vbCr & "Cuota: " & Cuota
End Function

Private Function GetDataFromRecordset(ByRef recData As ADODB.Recordset) As Boolean

    If (Not mNoMatchRaiseError) And recData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False

        mIDMedioPagoPlan = recData("IDMedioPagoPlan").Value
        mCuota = recData("Cuota").Value
        mCoeficiente = recData("Coeficiente").Value
        mCoeficientePrepago = recData("CoeficientePrepago").Value
    End If

    Screen.MousePointer = vbDefault
    GetDataFromRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.MedioPagoPlanCuota.GetDataFromRecordset", "Error al obtener los datos de la Cuota del Plan de Medios de Pago." & vbCr & vbCr & "IDMedioPagoPlan: " & IDMedioPagoPlan & vbCr & "Cuota: " & Cuota
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDMedioPagoPlan = 0
        mCuota = 0

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
            .CommandText = "usp_MedioPagoPlanCuota_Add"
        Else
            .CommandText = "usp_MedioPagoPlanCuota_Update"
        End If
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDMedioPagoPlan", adTinyInt, adParamInput, , mIDMedioPagoPlan)
        .Parameters.Append .CreateParameter("Cuota", adTinyInt, adParamInput, , mCuota)
        .Parameters.Append .CreateParameter("Coeficiente", adDecimal, adParamInput, , mCoeficiente)
        .Parameters("Coeficiente").Precision = 5
        .Parameters("Coeficiente").NumericScale = 4
        .Parameters.Append .CreateParameter("CoeficientePrepago", adDecimal, adParamInput, , mCoeficientePrepago)
        .Parameters("CoeficientePrepago").Precision = 5
        .Parameters("CoeficientePrepago").NumericScale = 4

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
    ShowErrorMessage "Classes.MedioPagoPlanCuota.Update", "Error al actualizar la Cuota del Plan de Medios de Pago." & vbCr & vbCr & "IDMedioPagoPlan: " & IDMedioPagoPlan & vbCr & "Cuota: " & Cuota
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
        .CommandText = "usp_MedioPagoPlanCuota_Delete"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDMedioPagoPlan", adTinyInt, adParamInput, , mIDMedioPagoPlan)
        .Parameters.Append .CreateParameter("Cuota", adTinyInt, adParamInput, , mCuota)

        .Execute
    End With

    If Not mRefreshListSkip Then
        Call RefreshList
    End If

    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.MedioPagoPlanCuota.Delete", "Error al eliminar la Cuota del Plan de Medios de Pago." & vbCr & vbCr & "IDMedioPagoPlan: " & IDMedioPagoPlan & vbCr & "Cuota: " & Cuota
End Function

Private Sub InitializeValues()
    mCoeficiente = 1
    mCoeficientePrepago = 1

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
