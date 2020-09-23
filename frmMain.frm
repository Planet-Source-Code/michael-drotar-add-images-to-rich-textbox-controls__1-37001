VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chat Icons"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbNames 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame fraIcons 
      Caption         =   " The Chat Icons "
      Height          =   1935
      Left            =   3240
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Image imgIcon 
         Height          =   195
         Index           =   4
         Left            =   1560
         Picture         =   "frmMain.frx":0032
         Tag             =   ":( :-("
         Top             =   360
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   195
         Index           =   3
         Left            =   1200
         Picture         =   "frmMain.frx":00B9
         Tag             =   ";) ;-)"
         Top             =   360
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   210
         Index           =   2
         Left            =   840
         Picture         =   "frmMain.frx":013E
         Tag             =   ":P :-P"
         Top             =   360
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "frmMain.frx":01D7
         Tag             =   ":D :-D"
         Top             =   360
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   195
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":0271
         Tag             =   ":) :-)"
         Top             =   360
         Width           =   285
      End
   End
   Begin VB.PictureBox picBuffer 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   5055
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":02FD
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const USERNAME = "Bob"                          'Your username

Const WM_PASTE = &H302
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    


Private Sub ChatMessage(ByVal sUser As String, ByVal sMessage As String)
    Dim lImagePos As Long
    Dim lStartMessage As Long
    Dim i As Integer, iCC As Integer
    Dim CharCombo() As String
    Dim ClipboardContents As Variant
    Dim bClipHasImage As Boolean
    
    bClipHasImage = Clipboard.GetFormat(vbCFBitmap) 'If there's an image in the clipboard
    If bClipHasImage Then picBuffer.Picture = Clipboard.GetData 'Store it to picBuffer
    
    With rtbChat
        .Locked = False                 'Must be unlocked for SendMessage() to work
        .SelStart = Len(.Text)          'Move cursor to the end to begin the new message
        
        .SelBold = True                 'Username in bold
                                        'If the user is you, use Red.. otherwise, blue
        .SelColor = IIf(sUser = USERNAME, vbRed, vbBlue)
        .SelText = sUser & ": "         'Show the user namme
        
        lStartMessage = Len(.Text) - 1  'Where the new message begins (search starts here
                                        '   for the icons)
        
        .SelBold = False                'Message text is not bold
        .SelColor = vbBlack             'Back to basic black
        .SelText = sMessage & vbCrLf    'Show the message with a linebreak
    End With
    

    For i = 0 To imgIcon.Count - 1                  'Loop through each icon
        CharCombo = Split(imgIcon(i).Tag, " ")      'Get the valid character combinations
                                                    '   which should be delimited by spaces
                                                    '   in the .Tag property
                                                    
        For iCC = 0 To UBound(CharCombo)            'Loop through those character combos
        
                                                    'Find where the char combo starts
            lImagePos = InStr(lStartMessage, rtbChat.Text, CharCombo(iCC))


            While lImagePos > 0                     'While the char combo is present
                rtbChat.SelStart = lImagePos - 1
                rtbChat.SelLength = Len(CharCombo(iCC)) 'Clear the char combo text
                rtbChat.SelText = ""
                
                Clipboard.Clear                             'Clear the clipboard (required)
                Clipboard.SetData imgIcon(i).Picture        'Set the icon in it
                SendMessage rtbChat.hWnd, WM_PASTE, 0, 0    'Paste it to the rtbChat
                
                                                    'Find any more of that same icon
                lImagePos = InStr(lImagePos, rtbChat.Text, CharCombo(iCC))
            Wend
        Next iCC
    Next i
    
    rtbChat.Locked = True                           'Lock the chat back up
    
    If bClipHasImage Then
        Clipboard.SetData picBuffer.Picture         'Put the old clipboard contents back
    Else
        Clipboard.Clear 'If there were none, then clear it.  There's no use in leaving
    End If              '   an icon sitting in there
    
    rtbChat.SelStart = Len(rtbChat.Text)            'Move the cursor to the end
End Sub

Private Sub Form_Activate()
    txtMessage.SetFocus             'Once the form is loaded, set the focus to txtMessage
End Sub                             'Note: .SetFocus cannot be used during Form_Load()

Private Sub Form_Load()
    cmbNames.ListIndex = 0          'Start with the first name selected
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then                  'If they pressed the Enter/Return key
        KeyAscii = 0                                'Stops the annoying beep
        
        If Len(txtMessage.Text) Then                    'If a message is entered
            ChatMessage cmbNames.Text, txtMessage.Text  'Send it to the chat
            txtMessage.Text = ""                        'And clear it away
        End If
    End If
End Sub
