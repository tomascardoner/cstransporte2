VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmailAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mMessageID As Long
Private mAttachmentID As Long
Private mFileName As String
Private mFileExtension As String
Private mFileData As String
Private mFileSourcePath As String

'STATE VARIABLES
Private mIsNew As Boolean
Private mIsDirty As Boolean
Private mNoMatch As Boolean

'BEHAVIOR VARIABLES
Private mNoMatchRaiseError As Boolean
Private mRefreshList As Boolean

'INTERNAL VARIABLES
Private mrecData As ADODB.Recordset

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get MessageID() As Long
    MessageID = mMessageID
End Property

Public Property Let MessageID(ByVal Value As Long)
    If Value <> mMessageID Then
        mIsDirty = True
    End If
    mMessageID = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get AttachmentID() As Long
    AttachmentID = mAttachmentID
End Property

Public Property Let AttachmentID(ByVal Value As Long)
    If Value <> mAttachmentID Then
        mIsDirty = True
    End If
    mAttachmentID = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let FileName(ByVal Value As String)
    If Value <> mFileName Then
        mIsDirty = True
    End If
    mFileName = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FileExtension() As String
    FileExtension = mFileExtension
End Property

Public Property Let FileExtension(ByVal Value As String)
    If Value <> mFileExtension Then
        mIsDirty = True
    End If
    mFileExtension = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FileData() As String
    FileData = mFileData
End Property

Public Property Let FileData(ByVal Value As String)
    If Value <> mFileData Then
        mIsDirty = True
    End If
    mFileData = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FileSourcePath() As String
    FileSourcePath = mFileSourcePath
End Property

Public Property Let FileSourcePath(ByVal Value As String)
    If Value <> mFileSourcePath Then
        mIsDirty = True
    End If
    mFileSourcePath = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IsNew() As Boolean
    IsNew = mIsNew
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = mIsDirty
End Property

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

Public Property Get RefreshList() As Boolean
    RefreshList = mRefreshList
End Property

Public Property Let RefreshList(ByVal Value As Boolean)
    mRefreshList = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function OpenRecordset() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    Screen.MousePointer = vbHourglass
      
    Set mrecData = New ADODB.Recordset
    Set mrecData.ActiveConnection = pDatabase.Connection
    mrecData.CursorType = adOpenKeyset
    mrecData.LockType = adLockOptimistic
    mrecData.Source = "SELECT * FROM EmailAttachment WHERE MessageID = " & MessageID & " AND AttachmentID = " & AttachmentID
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").Value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.EmailAttachment.OpenRecordset", "Error al Abrir la Tabla EmailAttachment."
    If Not mrecData Is Nothing Then
        Set mrecData = Nothing
    End If
End Function

Public Function Load() As Boolean
    If mrecData Is Nothing Then
        If Not OpenRecordset() Then
            Exit Function
        End If
    Else
        If mrecData.State = adStateOpen Then
            If Not OpenRecordset() Then
                Exit Function
            End If
        End If
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Call InitializeValues

    If (Not NoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
        
        mFileName = mrecData("FileName").Value
        mFileExtension = mrecData("FileExtension").Value & ""
        mFileSourcePath = pSpecialFolders.Temp & mrecData("FileName").Value & "." & mFileExtension
        Call pDatabase.ReadBLOB(mrecData, "FileData", mFileSourcePath)
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.EmailAttachment.Load", "Error al obtener los datos de el Adjunto." & vbCr & vbCr & "MessageID: " & MessageID & vbCr & "AttachmentID: " & AttachmentID
End Function

Public Function Add() As Boolean
    Dim recMaxID As ADODB.Recordset

    If Not mIsDirty Then
        Add = True
        Exit Function
    End If

    If mrecData Is Nothing Then
        If Not OpenRecordset() Then
            Exit Function
        End If
    Else
        If mrecData.State = adStateOpen Then
            If Not OpenRecordset() Then
                Exit Function
            End If
        End If
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    'GET NEW ID
    Set recMaxID = New ADODB.Recordset
    Set recMaxID.ActiveConnection = pDatabase.Connection
    recMaxID.CursorType = adOpenForwardOnly
    recMaxID.LockType = adLockReadOnly
    recMaxID.Source = "SELECT MAX(AttachmentID) AS MaxID FROM EmailAttachment WHERE MessageID = " & mMessageID
    recMaxID.Open , , , , adCmdText
    mAttachmentID = Val(recMaxID("MaxID").Value & "") + 1
    recMaxID.Close
    Set recMaxID = Nothing
    
    mrecData.AddNew
    mrecData("MessageID").Value = mMessageID
    mrecData("AttachmentID").Value = mAttachmentID
    mrecData("FileName").Value = mFileName
    mrecData("FileExtension").Value = mFileExtension
    Call pDatabase.WriteBLOB(mFileSourcePath, mrecData, "FileData")
    mrecData.Update
    
    If RefreshList Then
        'RefreshList_RefreshEmailAttachment MessageID, AttachmentID
    End If
    
    Screen.MousePointer = vbDefault
    Add = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un EmailAttachment con el mismo Nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.EmailAttachment.Add", "Error al Agregar el EmailAttachment." & vbCr & vbCr & "MessageID: " & MessageID & vbCr & "AttachmentID: " & AttachmentID
        End If
    Else
        ShowErrorMessage "Classes.EmailAttachment.Add", "Error al Agregar el EmailAttachment." & vbCr & vbCr & "MessageID: " & MessageID & vbCr & "AttachmentID: " & AttachmentID
    End If
    If Not recMaxID Is Nothing Then
        If recMaxID.State = adStateOpen Then
            recMaxID.Close
        End If
        Set recMaxID = Nothing
    End If
End Function

Public Function Update() As Boolean
    If Not mIsDirty Then
        Update = True
        Exit Function
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    mrecData("FileName").Value = mFileName
    mrecData("FileExtension").Value = mFileExtension
    Call pDatabase.WriteBLOB(mFileSourcePath, mrecData, "FileData")
    mrecData.Update
    
    If RefreshList Then
        'RefreshList_RefreshEmailAttachment MessageID, AttachmentID
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un EmailAttachment con el mismo Nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.EmailAttachment.Update", "Error al Actualizar el EmailAttachment." & vbCr & vbCr & "MessageID: " & MessageID & vbCr & "AttachmentID: " & AttachmentID
        End If
    Else
        ShowErrorMessage "Classes.EmailAttachment.Update", "Error al Actualizar el EmailAttachment." & vbCr & vbCr & "MessageID: " & MessageID & vbCr & "AttachmentID: " & AttachmentID
    End If
End Function

Public Function Delete() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    mrecData.Delete
    
    If RefreshList Then
        'RefreshList_RefreshEmailAttachment 0, 0
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede Eliminar el EmailAttachment debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.EmailAttachment.Delete", "Error al Eliminar el EmailAttachment." & vbCr & vbCr & "MessageID: " & MessageID & vbCr & "AttachmentID: " & AttachmentID
        End If
    Else
        ShowErrorMessage "Classes.EmailAttachment.Delete", "Error al Eliminar el EmailAttachment." & vbCr & vbCr & "MessageID: " & MessageID & vbCr & "AttachmentID: " & AttachmentID
    End If
End Function

Private Sub InitializeValues()
    mFileName = ""
    mFileExtension = ""
    mFileData = ""
    mFileSourcePath = ""

    mIsNew = True
    mIsDirty = False
End Sub

Private Sub Class_Initialize()
    InitializeValues
    mNoMatchRaiseError = True
    mRefreshList = True
End Sub

Private Sub Class_Terminate()
    If Not mrecData Is Nothing Then
        If mrecData.State = adStateOpen Then
            mrecData.Close
        End If
        Set mrecData = Nothing
    End If
    If (Not mIsNew) And mFileSourcePath <> "" Then
        On Error Resume Next
        FileSystem.Kill mFileSourcePath
    End If
End Sub
