VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Index"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   3975
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   480
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Error!"
      Default         =   -1  'True
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Error Number"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************************************************
'Error Index
'Coded by: Dustin Davis
'Bootleg Software Inc.
'http://www.warpnet.org/bsi
'
'This code is intended to help all VB coders who would like more help with errors and error
'handling. This program is exelent for learning how to define errors, handle them and so on
'I have tried to define as many errors as i could, but could only define a few of them. If
'you can define more, please 'ADD' your name to this code, and upload it. Please try tio keep
'it in the same format as i've done. This way we can all share it as a valuable resource for
'handling errors!
'*******************************************************************************************

'Note: If the description is "Application Defined or Object defined error"
'then it is unknown and will come from a Control you've added to your form

Private Sub Command1_Click()
On Error GoTo HandleIt
'Define unknown error number Please Add if you can!
If Text1.Text = "" Then
    MsgBox "You must enter an Error number", vbExclamation, "HEY!"
    Exit Sub
ElseIf Text1.Text = 40020 Then
    Err.Raise Text1.Text, "Winsock Control", "invalid operation at current state. This occurs when a socket is not connected and you try to send data"
ElseIf Text1.Text = 32755 Then
    Err.Raise Text1.Text, "Common Dialog Control", "Cancel Pressed"
ElseIf Text1.Text = 40006 Then
    Err.Raise Text1.Text, "Winsock Control", "Error on Data Arival"
ElseIf Text1.Text = 10048 Then
    Err.Raise Text1.Text, "Winsock Control", "Address already in Use This happens when you try to connect to a port or listen to a port that is in use by something else"
ElseIf Text1.Text = 1002 Then
    Err.Raise Text1.Text, "Image / Scan Control", "no image specified: cannot get horizontal resolution of the specified image. Happens when you try to exit the scan control without scanning a picture"
ElseIf Text1.Text = 1007 Then
    Err.Raise Text1.Text, "Image Control", "selection rectangle is requierd: unable to zoom to the selection. Happens when you dont select an area"
ElseIf Text1.Text = 0 Then
    AddText Text2, vbCrLf & "0 - Unknown Source, Cannot Define! Happens when something wants to be an ass!!"
    Exit Sub
Else 'this will only bring up system errors like runtime error, compile error, Disk errors, etc.
    Err.Raise Text1.Text
End If
HandleIt:
    AddText Text2, vbCrLf & Err.Number & " - " & Err.Description & vbCrLf & "Source of Error: " & Err.Source
End Sub

Function AddText(textcontrol As Object, text2add As String)

'code obtained from planet-source-code.com

    On Error GoTo errhandlr
    tmptxt$ = textcontrol.Text 'just in Case of an accident
    textcontrol.SelStart = Len(textcontrol.Text) ' move the "cursor" to the End of the text file
    textcontrol.SelLength = 0 ' highlight nothing (this becomes the selected text)
    textcontrol.SelText = text2add ' set the selected text ot text2add
    AddText = 1
    GoTo quitt ' goto the End of the Sub
    'error handlers
errhandlr:


    If Err.Number <> 438 Then 'check the Error number and restore the
        textcontrol.Text = tmptxt$ 'original text If the control supports it
    End If

    AddText = 0
    GoTo quitt
quitt:
    tmptxt$ = ""
End Function

Private Sub Command2_Click()
Dim errnum As Long
On Error GoTo WriteErr
errnum = 870
Do
    Err.Raise errnum
    errnum = errnum + 1
    DoEvents
Loop Until errnum >= 1200
WriteErr:
AddText Text2, vbCrLf & Err.Number & "-" & Err.Description
Resume Next
End Sub
