VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Different Menus"
   ClientHeight    =   7575
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   5505
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   3555
      TabIndex        =   6
      Top             =   3120
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "Status Bar"
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.ComboBox Combo3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2640
      List            =   "Form1.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   0
      Width           =   3135
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Form1.frx":003C
      Left            =   1200
      List            =   "Form1.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   4575
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Form1.frx":0059
      Left            =   0
      List            =   "Form1.frx":0063
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Status Bar"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   7320
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "For more information contact the Author by email (bfuzz@mbox.com.au) or visit www.goodproject.bizland.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "This example shows you how to have combo boxs as menus rather then the normal windows menus."
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "Exit" Then
End
End If
End Sub

Private Sub Combo2_Click()
If Combo2.Text = "Changer" Then
Picture1.Visible = True
Combo2.Text = "View"
End If
End Sub

Private Sub Combo3_Click()
If Combo3.Text = "Show Status Bar" Then
Label2.Visible = True
Combo3.Text = "Options"
End If
If Combo3.Text = "Hide Status Bar" Then
Label2.Visible = False
Combo3.Text = "Options"
End If
'Change Menu name back


End Sub

Private Sub Command1_Click()
Label2.Caption = Text1.Text
Picture1.Visible = False
End Sub

Private Sub Form_Load()
Combo1 = "File"
Combo2 = "View"
Combo3 = "Options"
End Sub
