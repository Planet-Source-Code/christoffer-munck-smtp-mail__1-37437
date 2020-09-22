VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSMTPmailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private strServ As String
Private strTo As String
Private strFrom As String
Private strSubj As String
Private strMsg As String
Private strFromName As String
Private blnHTML As Boolean
Private strErrors As String
Private intPort As Integer
Private WithEvents objWinsock As Winsock
Attribute objWinsock.VB_VarHelpID = -1

Public Property Let Server(ByVal strTemp As String)
    strServ = strTemp
End Property

Public Property Let SendTo(ByVal strTemp As String)
    strTo = strTemp
End Property

Public Property Let MessageSubject(ByVal strTemp As String)
    strSubj = strTemp
End Property

Public Property Let MessageText(ByVal strTemp As String)
    strMsg = strTemp
End Property

Public Property Let MailFrom(ByVal strTemp As String)
    strFrom = strTemp
End Property


Public Property Get Server() As String
    Server = strServ
End Property

Public Property Get SendTo() As String
    SendTo = strTo
End Property

Public Property Get MessageSubject() As String
    MessageSubject = strSubj
End Property

Public Property Get MessageText() As String
    MessageText = strMsg
End Property

Public Property Get MailFrom() As String
    MailFrom = strFrom
End Property



Public Property Get HTML() As Boolean
    HTML = blnHTML
End Property

Public Property Let HTML(isHTML As Boolean)
    blnHTML = isHTML
End Property

Public Property Let ServerPort(intPort As Integer)
    intPort = ServerPort
End Property

Public Property Get ServerPort() As Integer
    ServerPort = intPort
End Property
Public Sub Send()

    On Error Resume Next

    Dim strContentType As String
    Dim strMessage As String
    Dim msgIn As String
    strMessage = ""
    If blnHTML = True Then
        strContentType = "html;"
        Else
        strContentType = "plain;"
    End If

    With objWinsock
    
        .Connect strServ, intPort
    
        Wait 0.5
        If .State = 7 Then
            strMessage = strMessage & "HELO " & .LocalHostName & vbCrLf
            .SendData strMessage
            strMessage = ""
            
            Wait 0.5
            
            strMessage = strMessage & "MAIL FROM:<" & strFrom & ">" & vbCrLf
            .SendData strMessage
            strMessage = ""
            .GetData msgIn
            If Not funcCheckError(msgIn) = True Then
                Err.Raise Left(msgIn, 3), , strErrors
                Exit Sub
            End If
            Wait 0.5
            
            strMessage = "RCPT TO:<" + strTo + ">" + vbCrLf
            .SendData strMessage
            strMessage = ""
            Wait 0.5
            
            strMessage = strMessage & "DATA" & vbCrLf
            .SendData strMessage
            strMessage = ""
            .GetData msgIn
            If Not funcCheckError(msgIn) = True Then
                Err.Raise Left(msgIn, 3), , strErrors
                Exit Sub
            End If
            
            
            Wait 0.5
            strMessage = strMessage & "From: <" & strFrom & ">" & vbCrLf
            .SendData strMessage
            strMessage = ""
             .GetData msgIn
            If Not funcCheckError(msgIn) = True Then
                Err.Raise Left(msgIn, 3), , strErrors
                Exit Sub
            End If
            strMessage = strMessage & "To: " & strTo & vbCrLf
             .SendData strMessage
            strMessage = ""
             .GetData msgIn
            If Not funcCheckError(msgIn) = True Then
                Err.Raise Left(msgIn, 3), , strErrors
                Exit Sub
            End If
            strMessage = strMessage & "Subject: " & strSubj & vbCrLf
                        .SendData strMessage
            strMessage = ""
             .GetData msgIn
            If Not funcCheckError(msgIn) = True Then
                Err.Raise Left(msgIn, 3), , strErrors
                Exit Sub
            End If
            strMessage = strMessage & "X-Mailer: 2pro" & vbCrLf
                        .SendData strMessage
            strMessage = ""
             .GetData msgIn
            If Not funcCheckError(msgIn) = True Then
                Err.Raise Left(msgIn, 3), , strErrors
                Exit Sub
            End If
            
            strMessage = strMessage & "Mime-Version: 1.0" & vbCrLf
                        .SendData strMessage
            strMessage = ""
             .GetData msgIn
            If Not funcCheckError(msgIn) = True Then
                Err.Raise Left(msgIn, 3), , strErrors
                Exit Sub
            End If
            
            strMessage = strMessage & "Content-Type: text/" & strContentType & vbTab & "charset=us-ascii" & vbCrLf & vbCrLf
            .SendData strMessage
            strMessage = ""
             .GetData msgIn
            If Not funcCheckError(msgIn) = True Then
                Err.Raise Left(msgIn, 3), , strErrors
                Exit Sub
            End If
            
            strMessage = strMessage & strMsg
            .SendData strMessage
            strMessage = ""
            
            .SendData vbCrLf & "." & vbCrLf
            .GetData msgIn
            Wait 0.5
            
            If Not funcCheckError(msgIn) = True Then
                Err.Raise Left(msgIn, 3), , strErrors
                Exit Sub
            End If
            
            
            .SendData "QUIT"
            
            Wait 0.5
            .GetData msgIn
        
            If Not funcCheckError(msgIn) = True Then
                Err.Raise Left(msgIn, 3), , strErrors
                Exit Sub
            End If
            
        
        End If
        
        .Close
        
    End With
    
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, , Err.Description
        Exit Sub
    End If
    
End Sub


Public Property Get MailFromName() As String
    MailFromName = strFromName
End Property

Public Property Let MailFromName(ByVal strName As String)
    strFromName = strName
End Property

'## ERROR HANDLING (can be improved)
Private Function funcCheckError(strServerMessage As String) As Boolean

    If strServerMessage = "" Then
    funcCheckError = True
    Exit Function
    End If
    
    Dim intErrorNo As Integer
    Dim strTemp
    strTemp = Left(strServerMessage, 3)
    intErrorNo = Int(strTemp)
    
    Select Case intErrorNo
    
        Case 220
            funcCheckError = True
            
        Case 250
            funcCheckError = True
        Case 354
            funcCheckError = True
        Case Else
            funcCheckError = False
            strErrors = "Server returned following message:" & vbCrLf & strServerMessage
    End Select
    
    

End Function

Public Sub Wait(WaitTime)

    Dim StartTime As Double
    
    StartTime = Timer
    
    
    Do While Timer < StartTime + WaitTime
        If Timer > 86395 Or Timer = 0 Then Exit Do
        DoEvents
    Loop
    
End Sub

Private Sub Class_Initialize()
    '## CREATE WINSOCK OBJECT
    Set objWinsock = New Winsock
    '## SET DEFAULT VALUES
    strTo = "cmu@mondo.dk"
    strFrom = "cmu@mondo.dk"
    intPort = 25
    blnHTML = True
    strServ = objWinsock.LocalIP
    
End Sub

Private Sub Class_Terminate()
    Set objWinsock = Nothing
End Sub
