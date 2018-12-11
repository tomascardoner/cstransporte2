VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmailMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mMessageID As Long
Private mDateTime As Date
Private mSenderDisplayName As String
Private mSenderAddress As String
Private mRecipientToDisplayName As String
Private mRecipientToAddress As String
Private mRecipientCCDisplayName As String
Private mRecipientCCAddress As String
Private mRecipientBCCDisplayName As String
Private mRecipientBCCAddress As String
Private mSubject As String
Private mBody As String
Private mSMTPHost As String
Private mSMTPUserName As String
Private mSMTPPassword As String
Private mReadyToSend As Boolean
Private mLastTryToSend As Date
Private mSentDateTime As Date
Private mCAttachments As Collection
Private mAttachmentsSourceFiles As String

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

Public Property Get DateTime() As Date
    DateTime = mDateTime
End Property

Public Property Let DateTime(ByVal Value As Date)
    If Value <> mDateTime Then
        mIsDirty = True
    End If
    mDateTime = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get SenderDisplayName() As String
    SenderDisplayName = mSenderDisplayName
End Property

Public Property Let SenderDisplayName(ByVal Value As String)
    If Value <> mSenderDisplayName Then
        mIsDirty = True
    End If
    mSenderDisplayName = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get SenderAddress() As String
    SenderAddress = mSenderAddress
End Property

Public Property Let SenderAddress(ByVal Value As String)
    If Value <> mSenderAddress Then
        mIsDirty = True
    End If
    mSenderAddress = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RecipientToDisplayName() As String
    RecipientToDisplayName = mRecipientToDisplayName
End Property

Public Property Let RecipientToDisplayName(ByVal Value As String)
    If Value <> mRecipientToDisplayName Then
        mIsDirty = True
    End If
    mRecipientToDisplayName = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RecipientToAddress() As String
    RecipientToAddress = mRecipientToAddress
End Property

Public Property Let RecipientToAddress(ByVal Value As String)
    If Value <> mRecipientToAddress Then
        mIsDirty = True
    End If
    mRecipientToAddress = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RecipientCCDisplayName() As String
    RecipientCCDisplayName = mRecipientCCDisplayName
End Property

Public Property Let RecipientCCDisplayName(ByVal Value As String)
    If Value <> mRecipientCCDisplayName Then
        mIsDirty = True
    End If
    mRecipientCCDisplayName = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RecipientCCAddress() As String
    RecipientCCAddress = mRecipientCCAddress
End Property

Public Property Let RecipientCCAddress(ByVal Value As String)
    If Value <> mRecipientCCAddress Then
        mIsDirty = True
    End If
    mRecipientCCAddress = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RecipientBCCDisplayName() As String
    RecipientBCCDisplayName = mRecipientBCCDisplayName
End Property

Public Property Let RecipientBCCDisplayName(ByVal Value As String)
    If Value <> mRecipientBCCDisplayName Then
        mIsDirty = True
    End If
    mRecipientBCCDisplayName = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RecipientBCCAddress() As String
    RecipientBCCAddress = mRecipientBCCAddress
End Property

Public Property Let RecipientBCCAddress(ByVal Value As String)
    If Value <> mRecipientBCCAddress Then
        mIsDirty = True
    End If
    mRecipientBCCAddress = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Subject() As String
    Subject = mSubject
End Property

Public Property Let Subject(ByVal Value As String)
    If Value <> mSubject Then
        mIsDirty = True
    End If
    mSubject = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Body() As String
    Body = mBody
End Property

Public Property Let Body(ByVal Value As String)
    If Value <> mBody Then
        mIsDirty = True
    End If
    mBody = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get SMTPHost() As String
    SMTPHost = mSMTPHost
End Property

Public Property Let SMTPHost(ByVal Value As String)
    If Value <> mSMTPHost Then
        mIsDirty = True
    End If
    mSMTPHost = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get SMTPUserName() As String
    SMTPUserName = mSMTPUserName
End Property

Public Property Let SMTPUserName(ByVal Value As String)
    If Value <> mSMTPUserName Then
        mIsDirty = True
    End If
    mSMTPUserName = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get SMTPPassword() As String
    SMTPPassword = mSMTPPassword
End Property

Public Property Let SMTPPassword(ByVal Value As String)
    If Value <> mSMTPPassword Then
        mIsDirty = True
    End If
    mSMTPPassword = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ReadyToSend() As Boolean
    ReadyToSend = mReadyToSend
End Property

Public Property Let ReadyToSend(ByVal Value As Boolean)
    If Value <> mReadyToSend Then
        mIsDirty = True
    End If
    mReadyToSend = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get LastTryToSend() As Date
    LastTryToSend = mLastTryToSend
End Property

Public Property Let LastTryToSend(ByVal Value As Date)
    If Value <> mLastTryToSend Then
        mIsDirty = True
    End If
    mLastTryToSend = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get SentDateTime() As Date
    SentDateTime = mSentDateTime
End Property

Public Property Let SentDateTime(ByVal Value As Date)
    If Value <> mSentDateTime Then
        mIsDirty = True
    End If
    mSentDateTime = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get CAttachments() As Collection
    Set CAttachments = mCAttachments
End Property

Public Property Set CAttachments(ByVal Value As Collection)
    Set mCAttachments = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get AttachmentsSourceFiles() As String
    AttachmentsSourceFiles = mAttachmentsSourceFiles
End Property

Public Property Let AttachmentsSourceFiles(ByVal Value As String)
    If Value <> mAttachmentsSourceFiles Then
        mIsDirty = True
    End If
    mAttachmentsSourceFiles = Value
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
    mrecData.Source = "SELECT * FROM EmailMessage WHERE MessageID = " & MessageID
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").Value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.EmailMessage.OpenRecordset", "Error al Abrir la Tabla EmailMessage."
    If Not mrecData Is Nothing Then
        Set mrecData = Nothing
    End If
End Function

Public Function Load() As Boolean
    Dim recAttachment As ADODB.Recordset
    Dim EmailAttachment As EmailAttachment
    Dim DES As CSC_Encryption_DES
    
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

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
        
        mDateTime = mrecData("DateTime").Value
        mSenderDisplayName = mrecData("SenderDisplayName").Value
        mSenderAddress = mrecData("SenderAddress").Value
        mRecipientToDisplayName = mrecData("RecipientToDisplayName").Value & ""
        mRecipientToAddress = mrecData("RecipientToAddress").Value & ""
        mRecipientCCDisplayName = mrecData("RecipientCCDisplayName").Value & ""
        mRecipientCCAddress = mrecData("RecipientCCAddress").Value & ""
        mRecipientBCCDisplayName = mrecData("RecipientBCCDisplayName").Value & ""
        mRecipientBCCAddress = mrecData("RecipientBCCAddress").Value & ""
        mSubject = mrecData("Subject").Value & ""
        mBody = mrecData("Body").Value & ""
        mSMTPHost = mrecData("SMTPHost").Value & ""
        mSMTPUserName = mrecData("SMTPUserName").Value & ""
        mSMTPPassword = mrecData("SMTPPassword").Value & ""
        If mSMTPPassword <> "" Then
            Set DES = New CSC_Encryption_DES
            mSMTPPassword = DES.DecryptString(mSMTPPassword, PASSWORD_ENCRYPTION_KEY)
            Set DES = Nothing
        End If
        mReadyToSend = mrecData("ReadyToSend").Value
        mLastTryToSend = IIf(IsNull(mrecData("LastTryToSend").Value), DATE_TIME_FIELD_NULL_VALUE, mrecData("LastTryToSend").Value)
        mSentDateTime = IIf(IsNull(mrecData("SentDateTime").Value), DATE_TIME_FIELD_NULL_VALUE, mrecData("SentDateTime").Value)
    
        Set recAttachment = New ADODB.Recordset
        Set recAttachment.ActiveConnection = pDatabase.Connection
        recAttachment.CursorType = adOpenForwardOnly
        recAttachment.LockType = adLockReadOnly
        recAttachment.Source = "SELECT AttachmentID FROM EmailAttachment WHERE MessageID = " & MessageID
        recAttachment.Open , , , , adCmdText
        
        Do While Not recAttachment.EOF
            Set EmailAttachment = New EmailAttachment
            EmailAttachment.MessageID = mMessageID
            EmailAttachment.AttachmentID = recAttachment("AttachmentID").Value
            Call EmailAttachment.Load
            mCAttachments.Add EmailAttachment, KEY_STRINGER & EmailAttachment.AttachmentID
            mAttachmentsSourceFiles = mAttachmentsSourceFiles & IIf(mAttachmentsSourceFiles = "", "", ";") & EmailAttachment.FileSourcePath
            Set EmailAttachment = Nothing
            
            recAttachment.MoveNext
        Loop
        
        recAttachment.Close
        Set recAttachment = Nothing
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.EmailMessage.Load", "Error al obtener los datos de el EmailMessage." & vbCr & vbCr & "MessageID: " & MessageID
    If Not recAttachment Is Nothing Then
        If recAttachment.State = adStateOpen Then
            recAttachment.Close
        End If
        Set recAttachment = Nothing
    End If
End Function

Public Function Add() As Boolean
    Dim recMaxID As ADODB.Recordset
    Dim DES As CSC_Encryption_DES
    
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
    recMaxID.Source = "SELECT MAX(MessageID) AS MaxID FROM EmailMessage"
    recMaxID.Open , , , , adCmdText
    mMessageID = Val(recMaxID("MaxID").Value & "") + 1
    recMaxID.Close
    Set recMaxID = Nothing
    
    mrecData.AddNew
    mrecData("MessageID").Value = mMessageID
    mrecData("DateTime").Value = mDateTime
    mrecData("SenderDisplayName").Value = mSenderDisplayName
    mrecData("SenderAddress").Value = mSenderAddress
    mrecData("RecipientToDisplayName").Value = IIf(Trim(mRecipientToDisplayName) = "", Null, mRecipientToDisplayName)
    mrecData("RecipientToAddress").Value = IIf(Trim(mRecipientToAddress) = "", Null, mRecipientToAddress)
    mrecData("RecipientCCDisplayName").Value = IIf(Trim(mRecipientCCDisplayName) = "", Null, mRecipientCCDisplayName)
    mrecData("RecipientCCAddress").Value = IIf(Trim(mRecipientCCAddress) = "", Null, mRecipientCCAddress)
    mrecData("RecipientBCCDisplayName").Value = IIf(Trim(mRecipientBCCDisplayName) = "", Null, mRecipientBCCDisplayName)
    mrecData("RecipientBCCAddress").Value = IIf(Trim(mRecipientBCCAddress) = "", Null, mRecipientBCCAddress)
    mrecData("Subject").Value = IIf(Trim(mSubject) = "", Null, mSubject)
    mrecData("Body").Value = IIf(Trim(mBody) = "", Null, mBody)
    mrecData("SMTPHost").Value = IIf(Trim(mSMTPHost) = "", Null, mSMTPHost)
    mrecData("SMTPUserName").Value = IIf(Trim(mSMTPUserName) = "", Null, mSMTPUserName)
    If Trim(mSMTPPassword) = "" Then
        mrecData("SMTPPassword").Value = Null
    Else
        Set DES = New CSC_Encryption_DES
        mrecData("SMTPPassword").Value = DES.EncryptString(mSMTPPassword, PASSWORD_ENCRYPTION_KEY, False)
        Set DES = Nothing
    End If
    mrecData("ReadyToSend").Value = mReadyToSend
    mrecData("LastTryToSend").Value = IIf(mLastTryToSend = DATE_TIME_FIELD_NULL_VALUE, Null, mLastTryToSend)
    mrecData("SentDateTime").Value = IIf(mSentDateTime = DATE_TIME_FIELD_NULL_VALUE, Null, mSentDateTime)
    mrecData.Update
    
    If RefreshList Then
        'RefreshList_RefreshEmailMessage MessageID
    End If
    
    Screen.MousePointer = vbDefault
    Add = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un EmailMessage con el mismo Nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.EmailMessage.Add", "Error al Agregar el EmailMessage." & vbCr & vbCr & "MessageID: " & MessageID
        End If
    Else
        ShowErrorMessage "Classes.EmailMessage.Add", "Error al Agregar el EmailMessage." & vbCr & vbCr & "MessageID: " & MessageID
    End If
    If Not recMaxID Is Nothing Then
        If recMaxID.State = adStateOpen Then
            recMaxID.Close
        End If
        Set recMaxID = Nothing
    End If
End Function

Public Function Update() As Boolean
    Dim DES As CSC_Encryption_DES
    
    If Not mIsDirty Then
        Update = True
        Exit Function
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    mrecData("DateTime").Value = mDateTime
    mrecData("SenderDisplayName").Value = mSenderDisplayName
    mrecData("SenderAddress").Value = mSenderAddress
    mrecData("RecipientToDisplayName").Value = IIf(Trim(mRecipientToDisplayName) = "", Null, mRecipientToDisplayName)
    mrecData("RecipientToAddress").Value = IIf(Trim(mRecipientToAddress) = "", Null, mRecipientToAddress)
    mrecData("RecipientCCDisplayName").Value = IIf(Trim(mRecipientCCDisplayName) = "", Null, mRecipientCCDisplayName)
    mrecData("RecipientCCAddress").Value = IIf(Trim(mRecipientCCAddress) = "", Null, mRecipientCCAddress)
    mrecData("RecipientBCCDisplayName").Value = IIf(Trim(mRecipientBCCDisplayName) = "", Null, mRecipientBCCDisplayName)
    mrecData("RecipientBCCAddress").Value = IIf(Trim(mRecipientBCCAddress) = "", Null, mRecipientBCCAddress)
    mrecData("Subject").Value = IIf(Trim(mSubject) = "", Null, mSubject)
    mrecData("Body").Value = IIf(Trim(mBody) = "", Null, mBody)
    mrecData("SMTPHost").Value = IIf(Trim(mSMTPHost) = "", Null, mSMTPHost)
    mrecData("SMTPUserName").Value = IIf(Trim(mSMTPUserName) = "", Null, mSMTPUserName)
    If Trim(mSMTPPassword) = "" Then
        mrecData("SMTPPassword").Value = Null
    Else
        Set DES = New CSC_Encryption_DES
        mrecData("SMTPPassword").Value = DES.EncryptString(mSMTPPassword, PASSWORD_ENCRYPTION_KEY, False)
        Set DES = Nothing
    End If
    mrecData("ReadyToSend").Value = mReadyToSend
    mrecData("LastTryToSend").Value = IIf(mLastTryToSend = DATE_TIME_FIELD_NULL_VALUE, Null, mLastTryToSend)
    mrecData("SentDateTime").Value = IIf(mSentDateTime = DATE_TIME_FIELD_NULL_VALUE, Null, mSentDateTime)
    mrecData.Update
        
    If RefreshList Then
        'RefreshList_RefreshEmailMessage MessageID
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un EmailMessage con el mismo Nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.EmailMessage.Update", "Error al Actualizar el EmailMessage." & vbCr & vbCr & "MessageID: " & MessageID
        End If
    Else
        ShowErrorMessage "Classes.EmailMessage.Update", "Error al Actualizar el EmailMessage." & vbCr & vbCr & "MessageID: " & MessageID
    End If
End Function

Public Function Delete() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    mrecData.Delete
    
    If RefreshList Then
        'RefreshList_RefreshEmailMessage 0
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede Eliminar el EmailMessage debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.EmailMessage.Delete", "Error al Eliminar el EmailMessage." & vbCr & vbCr & "MessageID: " & MessageID
        End If
    Else
        ShowErrorMessage "Classes.EmailMessage.Delete", "Error al Eliminar el EmailMessage." & vbCr & vbCr & "MessageID: " & MessageID
    End If
End Function

Private Sub InitializeValues()
    mDateTime = DATE_TIME_FIELD_NULL_VALUE
    mSenderDisplayName = ""
    mSenderAddress = ""
    mRecipientToDisplayName = ""
    mRecipientToAddress = ""
    mRecipientCCDisplayName = ""
    mRecipientCCAddress = ""
    mRecipientBCCDisplayName = ""
    mRecipientBCCAddress = ""
    mSubject = ""
    mBody = ""
    mSMTPHost = ""
    mSMTPUserName = ""
    mSMTPPassword = ""
    mReadyToSend = False
    mLastTryToSend = DATE_TIME_FIELD_NULL_VALUE
    mSentDateTime = DATE_TIME_FIELD_NULL_VALUE
    Set mCAttachments = New Collection
    mAttachmentsSourceFiles = ""

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
    Set mCAttachments = Nothing
End Sub
