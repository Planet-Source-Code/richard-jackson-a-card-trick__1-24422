VERSION 5.00
Begin VB.Form CardTrick 
   BackColor       =   &H0000C000&
   Caption         =   "Card Trick!"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Column 3"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Column 2"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Column 1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   $"CardTrick.frx":0000
      Height          =   2655
      Left            =   5040
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   1455
      Left            =   5520
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   14
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   13
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   12
      Left            =   720
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   11
      Left            =   720
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   10
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   9
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   8
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   7
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   6
      Left            =   720
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   5
      Left            =   720
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   4
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   3
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   2
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   1
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   0
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "CardTrick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cardHold(1 To 15) As Integer
Dim card(1 To 15) As Integer
Dim clicked As Integer

Private Sub Command1_Click()
    
    'Add 1 to clicked to keep up with number of times the command
    'buttons have been clicked
    
    clicked = clicked + 1
    
        'pick up column 2 first
        
        cardHold(1) = card(2)
        cardHold(2) = card(5)
        cardHold(3) = card(8)
        cardHold(4) = card(11)
        cardHold(5) = card(14)
    
        'pick up column 1 next
        cardHold(6) = card(1)
        cardHold(7) = card(6)
        cardHold(8) = card(7)
        cardHold(9) = card(12)
        cardHold(10) = card(13)
    
        'pick up column 3 last
        cardHold(11) = card(3)
        cardHold(12) = card(4)
        cardHold(13) = card(9)
        cardHold(14) = card(10)
        cardHold(15) = card(15)
    
    'check to see if this is the third time clicked
    
    If clicked < 3 Then
        
        'if not third time, reassign the order of cards, and redisplay cards
        'in new order
        
        For i = 1 To 15
            card(i) = cardHold(i)
            Image1(i - 1).Picture = LoadPicture(App.Path & "\" & Trim(Str(card(i))) & ".JPG")
        Next i
    
    Else
        
        'if third time, reassign the order, but do not redisplay
        
        For i = 1 To 15
            card(i) = cardHold(i)
        Next i
        
        'show the user the card they had selected
        
        Image2.Picture = LoadPicture(App.Path & "\" & Trim(Str(card(8))) & ".jpg")
    
        're-initialize clicked
        
        Call Form_Load
        clicked = 0
    
    End If
    
End Sub

Private Sub Command2_Click()
    
    clicked = clicked + 1
    
        cardHold(1) = card(1)
        cardHold(2) = card(6)
        cardHold(3) = card(7)
        cardHold(4) = card(12)
        cardHold(5) = card(13)
    
        cardHold(6) = card(2)
        cardHold(7) = card(5)
        cardHold(8) = card(8)
        cardHold(9) = card(11)
        cardHold(10) = card(14)
    
        cardHold(11) = card(3)
        cardHold(12) = card(4)
        cardHold(13) = card(9)
        cardHold(14) = card(10)
        cardHold(15) = card(15)
    
    If clicked < 3 Then
        
        For i = 1 To 15
            card(i) = cardHold(i)
            Image1(i - 1).Picture = LoadPicture(App.Path & "\" & Trim(Str(card(i))) & ".JPG")
        Next i
    
    Else
    
        For i = 1 To 15
            card(i) = cardHold(i)
        Next i
        
        Image2.Picture = LoadPicture(App.Path & "\" & Trim(Str(card(8))) & ".jpg")
    
        Call Form_Load
        clicked = 0
                
    End If
 
End Sub

Private Sub Command3_Click()

    clicked = clicked + 1
    
        cardHold(1) = card(2)
        cardHold(2) = card(5)
        cardHold(3) = card(8)
        cardHold(4) = card(11)
        cardHold(5) = card(14)
    
        cardHold(6) = card(3)
        cardHold(7) = card(4)
        cardHold(8) = card(9)
        cardHold(9) = card(10)
        cardHold(10) = card(15)
    
        cardHold(11) = card(1)
        cardHold(12) = card(6)
        cardHold(13) = card(7)
        cardHold(14) = card(12)
        cardHold(15) = card(13)
            
    If clicked < 3 Then
        
        For i = 1 To 15
            card(i) = cardHold(i)
            Image1(i - 1).Picture = LoadPicture(App.Path & "\" & Trim(Str(card(i))) & ".JPG")
        Next i

    Else
        
        For i = 1 To 15
            card(i) = cardHold(i)
        Next i
        
        Image2.Picture = LoadPicture(App.Path & "\" & Trim(Str(card(8))) & ".jpg")
    
        Call Form_Load
        clicked = 0
        
    End If
    
End Sub

Private Sub Form_Load()

    'Center the form on the screen
    
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    Randomize Timer
    
    'Select 15 random cards from a deck of 52
    
    For Outer = 1 To 15
        Do
            uniqueCard = True
            cardNumber = Int(Rnd * 52) + 1
            For Inner = 1 To Outer
                If cardNumber = card(Inner) Then
                    uniqueCard = False
                End If
            Next Inner
        Loop Until uniqueCard
        card(Outer) = cardNumber
    Next Outer
        
    'Load the cards into the image control array...
    'Named the graphic files of the individual cards by number,
    'such as... 1.jpg 2.jpg 3.jpg ... so on, and so on
    
    For i = 1 To 15
        Image1(i - 1).Picture = LoadPicture(App.Path & "\" & Trim(Str(card(i))) & ".JPG")
    Next i
    
End Sub

