VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PSC Vote Box UserControl"
   ClientHeight    =   3375
   ClientLeft      =   195
   ClientTop       =   1575
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   8730
   Begin PSCVoteControl.VoteControl VoteControl1 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5106
      CodeID          =   26205
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'#b3c4ff
End Sub

Public Function LastDay(Year As Integer, Month As Integer)
    LastDay = DateSerial(Year, Month + 1, 0)
    LastDay = Format(LastDay, "dd")
End Function

