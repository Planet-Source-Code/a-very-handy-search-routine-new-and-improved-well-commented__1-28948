VERSION 5.00
Begin VB.Form frmFound 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Item Found!"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFound 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   750
      TabIndex        =   0
      Top             =   75
      Width           =   3540
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   1200
         TabIndex        =   8
         Top             =   1125
         Width           =   105
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   1200
         TabIndex        =   7
         Top             =   825
         Width           =   105
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   525
         Width           =   105
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Top             =   225
         Width           =   105
      End
      Begin VB.Label lblFound 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   240
         Index           =   3
         Left            =   225
         TabIndex        =   4
         Top             =   1160
         Width           =   840
      End
      Begin VB.Label lblFound 
         BackStyle       =   0  'Transparent
         Caption         =   "Part No"
         Height          =   240
         Index           =   2
         Left            =   225
         TabIndex        =   3
         Top             =   885
         Width           =   840
      End
      Begin VB.Label lblFound 
         BackStyle       =   0  'Transparent
         Caption         =   "Model No"
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   2
         Top             =   580
         Width           =   840
      End
      Begin VB.Label lblFound 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   300
         Width           =   840
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   75
      Picture         =   "frmFound.frx":0000
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frmFound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Dim lx As Long
lblItem(0) = mItem
For lx = 1 To 3
    lblItem(lx) = mItem.SubItems(lx)
Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set mItem = Nothing
End Sub
