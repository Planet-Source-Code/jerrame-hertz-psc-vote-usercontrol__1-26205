VERSION 5.00
Begin VB.UserControl VoteControl 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8310
   ScaleHeight     =   3030
   ScaleWidth      =   8310
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   240
      ScaleHeight     =   2055
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   600
      Width           =   7815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Poor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Below Average"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Average"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Good"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Excellent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Rate It!"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "coding contest!)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   4680
         MouseIcon       =   "VoteControl.ctx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "What do you think of this code(in the Intermediate catagory)?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(The codewith your highest vote will win this month's"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "See Voting Log"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   0
         MouseIcon       =   "VoteControl.ctx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Your Vote!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "VoteControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private selected As Integer
'Default Property Values:
Const m_def_CodeID = 0
'Property Variables:
Dim m_CodeID As Integer



Private Sub Command1_Click()
    OpenWebsite ("http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=" & Label3.Caption & "&optCodeRatingValue=" & selected)
End Sub

Private Sub Label2_Click(Index As Integer)
    Select Case Index
        Case 2
            OpenWebsite ("http://www.planet-source-code.com/vb/contest/contest.asp?lngWId=1")
        Case 3
            OpenWebsite ("http://www.planet-source-code.com/vb/scripts/voting/VoteLog.asp?txtCodeId=" & Label3.Caption & "&txtCodeName=&intUserRatingTotal=&intNumOfUserRatings=&lngWid=1")
                        ' http://www.planet-source-code.com/vb/scripts/voting/VoteLog.asp?txtCodeId=" & Label3.Caption & "txtCodeName=Skinnable_Meter%20UserControl&intUserRatingTotal=&intNumOfUserRatings=&lngWid=1
    End Select
End Sub

Private Sub Option1_Click(Index As Integer)

    selected = Index
End Sub

Private Sub UserControl_Initialize()
    selected = 5
    Option1(5).Value = True
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get CodeID() As Integer
Attribute CodeID.VB_Description = "The codeID you wish to send your vote to."
    CodeID = m_CodeID
    Label3.Caption = CodeID
End Property

Public Property Let CodeID(ByVal New_CodeID As Integer)
    m_CodeID = New_CodeID
    PropertyChanged "CodeID"
    Label3.Caption = CodeID
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_CodeID = m_def_CodeID
    Label3.Caption = CodeID
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_CodeID = PropBag.ReadProperty("CodeID", m_def_CodeID)
        Label3.Caption = CodeID

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CodeID", m_CodeID, m_def_CodeID)
    Label3.Caption = CodeID
End Sub

