VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' C_Error%.cls (c)  Serious Solutions
'
' Programmer:  Stephen S. Coleman
'
' Date: July 16, 2004
'
' Comments:
'This is actually a big deal for me. I've been using this type of error logging for a while now and love it.
' I only add to the top of procedures, not top and bottom.  Is there anything unsafe about it?
' It makes adding\removing error logging\trapping very easy.
' If the Gurus out there can tell me if my error log system is Bad,Unsafe, or really just crap.
' I would like comments on the below and the ErrorTrap Sub

' I add this to a module or form at the top of a Module or Class procedure

'On Error GoTo Error_Handler '--Error Trap------------------------------------------------
'If Err.Number Then
'Error_Handler:
'     ErrC.ErrorTrap "Insert Module\Class Name", "Insert Procedure Name", Err.Number, Err.Description, Err.Source
'     Resume Next
'End If '--Error Trap---------------------------------------------------------------------------

' I add this to a module or form at the top of a Form procedure

'On Error GoTo Error_Handler '--Error Trap------------------------------------------------
'If Err.Number Then
'Error_Handler:
'     ErrC.ErrorTrap Me.Name, "Insert Procedure Name", Err.Number, Err.Description, Err.Source
'     Resume Next
'End If '--Error Trap---------------------------------------------------------------------------

' Description: Easy Error Trapping
'
' Modifications:  SSC - Created at 2:11 PM
'-------------------------------------------------------------------------------

Option Explicit

'FOR EventLOGS**********************************
Public Enum EventType
    EVENTLOG_SUCCESS = 0
    EVENTLOG_ERROR_TYPE = 1
    EVENTLOG_WARNING_TYPE = 2
    EVENTLOG_INFORMATION_TYPE = 4
    EVENTLOG_AUDIT_SUCCESS = 8
    EVENTLOG_AUDIT_FAILURE = 10
End Enum
'----------------------------------------------------
'Private Declarations Section
'----------------------------------------------------
Private Type NTEventType
    intEventID      As Integer
    intEventType    As Integer
    strMessage      As String
End Type

Private Const m_ClassName = "clsNTEventLog"
Private mlngEventLog                As Long             'Handle of the NT Event log
Private mbolNTEventsLoaded          As Boolean          'Have we loaded NT Messaging?
Private mlngNTEventMessageCount     As Long             'NT Event message count
Private mtypNteventMessages()       As NTEventType      'Array of NT Messages
Private mstrNTEventMachine          As String           'Destination of event "" = local machine
Private mstrSource                  As String
Private mstrVersionNumber           As String
Private Const NTEVENT_CATEGORYID = 0

Private Declare Sub NTEvent_MemCpy Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function NTEvent_RegisterSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Private Declare Function NTEvent_DeregisterSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
Private Declare Function NTEvent_ReportEvent Lib "advapi32.dll" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Integer, ByVal wCategory As Integer, ByVal dwEventID As Long, ByVal lpUserSid As Any, ByVal wNumStrings As Integer, ByVal dwDataSize As Long, plpStrings As Long, lpRawData As Any) As Boolean
Private Declare Function NTEvent_GetLastError Lib "kernel32" Alias "GetLastError" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'FOR EventLOGS***************************************************

Public ErrEmail As Boolean
Public ErrPost As Boolean
Public ErrDebug As Boolean
Public ErrOutFile As Boolean
Public ErrEventLog As Boolean
Public ErrKillApp As Boolean
Public ErrResume As Boolean
Private M_OutputErrPath As String
Private M_EmailErrAdmin As String

Private Sub Class_Initialize()

    ErrDebug = False
    ErrOutFile = False
    ErrEmail = False
    ErrPost = False
    ErrEventLog = False
    ErrKillApp = False
    ErrResume = True
    M_EmailErrAdmin = "Myemail@mYemail.com"
    M_OutputErrPath = App.Path & "\" & App.Title & ".log"

End Sub

Public Sub createerror()

  Dim x As Variant

    On Error GoTo exiterrorlog
    x = Error.rasie
exiterrorlog:

End Sub

Public Property Get EmailErrAdmin() As Variant

    EmailErrAdmin = M_EmailErrAdmin

End Property

Public Property Let EmailErrAdmin(ByVal vNewValue As Variant)

    M_EmailErrAdmin = EmailErrAdmin

End Property

'-----ErrorTrap----------------------------------------------------------
' Parameters:
'    ErrForm      - O:   (String)
'    ErrProcedure - O:   (String)
'    ErrNum       - O:   (Long)
'    ErrDes       - O:   (String)
'    ErrSource    - O:   (String)
'
' Returns:
'   None
' Comments: Main function for errors
' Author:    Stephen S. Coleman
' Date:      July 16, 2004
'-----ErrorTrap----------------------------------------------------------
Public Sub ErrorTrap(ErrForm As String, ErrProcedure As String, ErrNum As Long, ErrDes As String, ErrSource As String)

    On Error GoTo ExitErrTrap
    'CaptureScreen ClientArea, "C:\Program Files\Software Management Tool\error.bmp", Overwrite
    If ErrDebug Then
        Debug.Print ErrSource & "\" & ErrForm & "\" & ErrProcedure & ":Error#" & ErrNum & ": " & ErrDes
    End If
    If ErrOutFile Then
        WriteErrToFile ErrSource & "\" & ErrForm & "\" & ErrProcedure & ":Error#" & ErrNum & ": " & ErrDes
    End If
    If ErrEmail Then
        If MsgBox(App.Title & " has encountered an error.  Would you like to send your " & App.Title & " admin an email", vbYesNo) = 6 Then
            notesmail M_EmailErrAdmin, App.Title & " SMT Error report", "This email was sent by an internal error in application " & App.Title, M_OutputErrPath
        End If
    End If
    If ErrPost Then
        MsgBox ErrSource & "\" & ErrForm & "\" & ErrProcedure & ":Error#" & ErrNum & ": " & ErrDes
    End If
    If ErrEventLog Then
        NTEventWrite ErrNum, ErrSource & "\" & ErrForm & "\" & ErrProcedure & ":Error#" & ErrNum & ": " & ErrDes
    End If
    If Not ErrResume Then
ExitErrTrap:
        ' Unloads all forms if error trap fails
        UnloadAll
    End If

End Sub

Private Sub notesmail(txtTo As String, txtSubject As String, txtBody As String, sAttachment As String)

  ' **********************************************************************
  ' * Comments         : Sending a mail message to Lotus Notes
  ' *
  ' * declare objects for notes objects
  ' *
  ' **********************************************************************

  Dim session As Object, db As Object, doc As Object, rtobject As Object
  Dim rtitem As Object

  Dim User As String, Server As String, mailfile As String
  Dim object As Object
  Dim Body As Object
  Dim oItem As Object

    On Error GoTo exiterrorlog

    ' doc.CopyTo = strCC
    ' 'Blind Copies the User
    ' doc.BlindCopyTo = strusername

    ' get a handle to a notes session
    Set session = CreateObject("notes.notessession")

    User = session.UserName
    Server = session.GetEnvironmentString("MailServer", True)
    mailfile = session.GetEnvironmentString("MailFile", True)
    Set db = session.GetDatabase(Server, mailfile)
    Set doc = db.CreateDocument()
    Set oItem = doc.CreateRichTextItem("BODY")
    doc.Form = "Memo"
    '    doc.SendTo = "david henderson@sehe.com"
    doc.sendto = txtTo
    '    doc.Subject = "notes mail created from vb"
    doc.Subject = txtSubject
    '   the following lines are only needed when attaching files
    doc.Body = txtBody
    Call oItem.EmbedObject(1454, "", sAttachment)

    Call doc.Send(False)
    '   clean up after ourselves
    Set session = Nothing
    Set db = Nothing
    Set doc = Nothing
    'The above code (and also the code by "waty") only works when sending to one address in either the SendTo or CC fields.  If you want to send to more than one person, you need to put each address into an array first, then pass the array to the  SendTo or CC fields.
exiterrorlog:

End Sub

' Public Function NTEventAddMessage(lngEventID As Long,
'                                   lngEventType As EventType,
'                                   strMessage As String) As Long
'
' [in] lngEventID        =  A long integer value indicating the event id that is to be used when registering the event with the Evetn Log
' [in] lngEventType      =  This is a numeric value that is retrieved from a public enum which has the following format
'                           EVENTLOG_SUCCESS          = 0
'                           EVENTLOG_ERROR_TYPE       = 1
'                           EVENTLOG_WARNING_TYPE     = 2
'                           EVENTLOG_INFORMATION_TYPE = 4
'                           EVENTLOG_AUDIT_SUCCESS    = 8
'                           EVENTLOG_AUDIT_FAILURE    = 10
'
' [in] strMessage        =  The event error description that is to be writtent to the event log
' [out] NTEventAddMessage  =  A long integer value that defines the error / success code 0=failed 1=success
'
'
' Such That:
' A Type array is built of NT Event log messages that are to be written to the NT event Log
'
'
' Author             : Roger Taylor
' Date               : 13th Dec 2000
' Last Modified Date :
'
' Revision history:
' - None
'
'--------------------------------------------------------------------------------------------------------
'
Public Function NTEventAddMessage(lngEventID As Long, lngEventType As EventType, _
                                  strMessage As String) As Long

    NTEventAddMessage = 0
    On Error GoTo Error

    'Increment the message count by one and resize the data
    'array for the NT Event messages
    mlngNTEventMessageCount = mlngNTEventMessageCount + 1
    ReDim Preserve mtypNteventMessages(1 To mlngNTEventMessageCount)

    'Add the event information...
    mtypNteventMessages(mlngNTEventMessageCount).intEventID = lngEventID
    mtypNteventMessages(mlngNTEventMessageCount).intEventType = lngEventType
    mtypNteventMessages(mlngNTEventMessageCount).strMessage = strMessage

    'Success
    NTEventAddMessage = 1
Error:

End Function

'--------------------------------------------------------------------------------------------------------
'Private Sub NTEventGetMessage(lngEventID As Long,
'                              strMessage As String,
'                              lngEventType As Long)
'
' [in] lngEventID        =  A long integer value indicating the event id that is to be used when registering the event with the Evetn Log
' [in] strMessage        =  The event error description that is to be writtent to the event log
' [in] lngEventType      =  This is a numeric value that is retrieved from a public enum which has the following format
'                           EVENTLOG_SUCCESS          = 0
'                           EVENTLOG_ERROR_TYPE       = 1
'                           EVENTLOG_WARNING_TYPE     = 2
'                           EVENTLOG_INFORMATION_TYPE = 4
'                           EVENTLOG_AUDIT_SUCCESS    = 8
'                           EVENTLOG_AUDIT_FAILURE    = 10
'
' Such That:
' The NT event is retrieved from the event array as a call from NTEventWrite
'
' Author             : Roger Taylor
' Date               : 13th Dec 2000
' Last Modified Date :
'
' Revision history:
' - None
'--------------------------------------------------------------------------------------------------------
'
Private Sub NTEventGetMessage(lngEventID As Long, strMessage As String, lngEventType As Long)

  Dim lngCount            As Long         'Counter'

    On Error GoTo Error_Handler

    'Loop through all the messages and select the correct
    'message and event type for the event id
    For lngCount = 1 To mlngNTEventMessageCount
        If lngEventID = mtypNteventMessages(lngCount).intEventID Then
            'Assign message and id and exit sub
            lngEventType = mtypNteventMessages(lngCount).intEventType
            strMessage = mtypNteventMessages(lngCount).strMessage
            Exit Sub
        End If
    Next lngCount

    'Did not find an entry for the event id so return a
    'standard message and info type
    strMessage = "Undefined Event"
    lngEventType = EVENTLOG_INFORMATION_TYPE

Exit Sub

Error_Handler:

End Sub

'--------------------------------------------------------------------------------------------------------
'Public Function NTEventRegisterLog() As Long
'
' [out] NTEventRegisterLog =  A long integer value that defines the error / success code
'                                1=success 0=failed 3=registration failed
'
'
' Such That:
' The creator of the NT event message is registered with the event log as an event log user.
'
' Author             : Roger Taylor
' Date               : 13th Dec 2000
' Last Modified Date :
'
' Revision history:
' - None
'
'--------------------------------------------------------------------------------------------------------

Public Function NTEventRegisterLog(strAppName As String, _
                                   Optional strDestinationMachineName As String = "", _
                                   Optional VersionNumber As String = "N/A") As Long

    On Error GoTo Error
    NTEventRegisterLog = 0

    'Check to see if NT Events have already been loaded
    'Steve
    ' If mbolNTEventsLoaded = True Then
    If mbolNTEventsLoaded Then
        Exit Function
    End If
    mstrNTEventMachine = strDestinationMachineName
    mstrSource = strAppName
    mstrVersionNumber = VersionNumber

    ' Register this app. as new source, mhEventLog is handle
    mlngEventLog = NTEvent_RegisterSource(mstrNTEventMachine, mstrSource)

    If mlngEventLog = Null Then
        NTEventRegisterLog = 3 'Registration Failed
        Exit Function
    End If

    'Success reported
    NTEventRegisterLog = 1

    'We have now connected to the NT Event manager
    mbolNTEventsLoaded = True
Error:

End Function

'--------------------------------------------------------------------------------------------------------
'Public Function NTEventWrite(lngEventID As Long,
'                             strAdditional As String) As Long
'
' [in] lngEventID    = A long integer value indicating the event id that is to be used when registering the event with the Evetn Log
' [in] strAdditional = A string which defines additional text that will describe the event that occured
' [out] NTEventWrite =  A long integer value that defines the error / success code
'                       0=failure 1=success
'
' Such That:
' An Event is written to the windows NT 4 Event log
'
'
'
' Author             : Roger Taylor
' Date               : 13th Dec 2000
' Last Modified Date :
'
' Revision history:
' - None
'
'--------------------------------------------------------------------------------------------------------
'
Public Function NTEventWrite(lngEventID As Long, Optional strAdditional As String = "") As Long

  Dim strMessage      As String          'Error message
  Dim lngEventType    As Long            'Error type
  Dim lngMsg          As Long            'Handle used for memory space allocation
  Dim lngLen          As Long            'String length
  Dim lngRet          As Long            'Function return

    On Error GoTo Error
    NTEventWrite = 0
    'Configure additional text
    strAdditional = Trim$(strAdditional)
    If Len(strAdditional) > 0 Then
        strAdditional = ": " & strAdditional
    End If

    'Now retrieve the error message from our array of messages
    Call NTEventGetMessage(lngEventID, strMessage, lngEventType)

    'Configure text of message and set up length
    strMessage = vbCrLf & vbCrLf & "App: " & mstrSource & ", Ver: " & mstrVersionNumber & ", Date: " & Format$(Now, "dd/mmm/yyyy") & vbCrLf & strMessage & strAdditional
    lngLen = Len(strMessage) + 1

    'Now allocate an amount of global memory for our error message
    lngMsg = GlobalAlloc(&H40, lngLen)

    'Copy message string into memory space obtained from GlobalAlloc
    NTEvent_MemCpy ByVal lngMsg, ByVal strMessage, lngLen

    'Now actually write the message to the log...
    lngRet = NTEvent_ReportEvent(mlngEventLog, lngEventType, NTEVENT_CATEGORYID, lngEventID, 0&, 1, lngLen, lngMsg, lngMsg)

    'Now free off our global memory
    Call GlobalFree(lngMsg)

    NTEventWrite = 1

Error:

End Function

Public Property Get OutputErrPath() As String

    OutputErrPath = M_OutputErrPath

End Property

Public Property Let OutputErrPath(ByVal vNewValue As String)

    M_OutputErrPath = OutputErrPath

End Property

Public Sub UnloadAll()

  Dim Cnt As Integer

    On Error GoTo Error_Handler '--Error Trap------------------------------------------------------
    If Err.Number Then
Error_Handler:
        ErrorTrap "C_Error%.cls", "UnloadAll", Err.Number, Err.Description, Err.Source
        Resume Next
    End If '--Error Trap---------------------------------------------------------------------------

    Do While Forms.Count > 1
        Cnt = Cnt + 1
        If Cnt > 3000 Then
            Exit Do
        End If
        Unload Forms(1)
    Loop
    Reset
    End

End Sub

Private Sub WriteErrToFile(Message As String)

  'Write transaction log to the log file

  Dim strErrorLog As String
  Dim sMessage As String

    On Error GoTo exiterrorlog
    strErrorLog = FreeFile()
    Open M_OutputErrPath For Append As strErrorLog
    sMessage = Message
    Print #strErrorLog, Now & Chr$(9) & sMessage
    Close #strErrorLog
exiterrorlog:

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2004-Jul-16 14:36) 85 + 407 = 492 Lines
