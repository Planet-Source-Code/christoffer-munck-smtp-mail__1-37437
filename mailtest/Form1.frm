VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objMail As New clsSMTPmailer

Private Sub Form_Load()

    With objMail
    .HTML = True
    .MailFrom = "cmu@mondo.dk"
    .SendTo = "cmu@mondo.dk"
    .Server = "10.1.1.10"
    .MessageSubject = " Hello there"
    .MessageText = "this is the body"
    .Send
    End With
    
End Sub
