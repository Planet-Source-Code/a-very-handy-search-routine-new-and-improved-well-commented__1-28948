VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRunSearch 
   Caption         =   "Run Search Example"
   ClientHeight    =   5655
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   6120
   Icon            =   "frmRunSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSearching 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   1725
      Top             =   450
   End
   Begin MSComctlLib.ListView lvReturn 
      Height          =   4515
      Left            =   75
      TabIndex        =   3
      Top             =   450
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7964
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "imglRunSearch"
      SmallIcons      =   "imglRunSearch"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item Name"
         Object.Width           =   3951
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Model No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Part No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Location"
         Object.Width           =   1129
      EndProperty
   End
   Begin MSComctlLib.ImageList imglRunSearch 
      Left            =   3450
      Top             =   4350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunSearch.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRunSearch 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      Height          =   465
      Left            =   75
      TabIndex        =   0
      Top             =   5100
      Width           =   5940
   End
   Begin VB.Label lblCurrent 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   75
      TabIndex        =   2
      Top             =   75
      Width           =   690
   End
   Begin VB.Label lblRecCount 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5775
      TabIndex        =   1
      Top             =   75
      Width           =   225
   End
   Begin VB.Menu mnuQuickSearchParent 
      Caption         =   "QuickSearch"
      Visible         =   0   'False
      Begin VB.Menu mnuSearch 
         Caption         =   "&Find It!"
         Index           =   0
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Item Name"
         Index           =   2
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Part No"
         Index           =   3
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Model No"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmRunSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbExample As Database
Const strDB_Path = "\Example.mdb"

Private Function RunSearch(FieldToSearch As String, List As MSComctlLib.ListView) As Boolean
Dim rsReturn As Recordset
Dim strWhatToFind As String
Dim strTemp As String
Dim TwiceAround As Boolean

On Error GoTo ErrorHandler
Searching = True
'I find that a very nice pair-up in interface and behind the scenes action is DAO in conjunction with a ListView (or
'even a treeview). To some this may already be apparent, to others perhaps not.  Copy this code in, using the correct
'DAO references and the correct ListView, pass the three parameters in, and it should work like a charm.

'I usually call this function from a friendly little popup menu with the fields to search listed therin.  I tag it onto the
'right-click as well as CTRL+F on the Form_Keyup events and my users love it.

'Basically what RunSearch does is works symbioctically with a ListView (a NEW ListView, as in MSComctlLib) optionally a current
'recordset to locate the first instance of the closest match, if any, to the criteria the user asked for.

'The 1st arg, FieldToSearch, is for the column that you want to search through, as in Serial No, or Part No, etc.
'This will then be used against BOTH the table/RS and the columnheader of the assed ListView.  The 2nd arg is for the ListView.

'The 3rd argument is for the recordset / sql statment.  I pass in a string so that it can either be a name of a table, or an
'actual SQL Select statement.  Either way is fine.  Depending on your sql statement, if you go that way, you may have to play with

'it to pass it in and concatenate it correctly.

KeyPreview = False
strWhatToFind = InputBox("What " & FieldToSearch & " would you like to search for?", FieldToSearch & " Search")
If strWhatToFind <> "" Then
    strTemp = "] ="
    Do Until TwiceAround
    'The reason I loop it 'TwiceAround' is, if you skip to the bottom there where the loop is, I throw on an asterick the second time.
    'This attempts to locate an exact match the first time on just what was asked for -- if it fails it comes back around a 2nd time'
    'to find the closest match.  This will work with initial wildcard given by the user as well.
        rsExample.FindFirst "[" & FieldToSearch & strTemp & " '" & strWhatToFind & "'"
        If Not rsExample.NoMatch Then
            'Now use the ItemID -- which if you look at the RefreshMe sub you'll see I've put those unique
            'values into each Listitem as its key, preceeded with "||", simply because it has to be alpha
            'numeric -- to locate our mystery man.
            strTemp = "||" & rsExample("ItemID")
            Set mItem = lvReturn.ListItems(strTemp)
            mItem.EnsureVisible
            mItem.Selected = True
            tmrSearching.Enabled = True
            xBeep
            frmFound.Show
            Exit Do
            'OUILA!  So simple and pretty it hurts.  Impress the boss with that....
            'Another way you can do it, especially if you want to locate / isolate more
            'than one record -- very common -- is to use the filter instead. Findfirst
            'just finds the first SINGLE instance of your criteria.  Use filter, as in:
            
            'rsExample.Filter = "[" & FieldToSearch & strTemp & " '" & strWhatToFind & "'"
            'Set rsReturn = rsExample.OpenRecordset
            
            'That simple.  Using that method will give you a pool of all records matching
            'the criteria you gave it.  Good for "closest match" type search engines.
                        
                        'You can also change the .selected property
                        'of the matched listitem to true.  I also like
                        'to temporarily enable the .EntireRowSelect
                        'property to make it really apparent.  Other
                        'ideas include .selected = true;
                        'List.Remove;
                        'RunSearch.Ghosted = true .... the list is nearly endless (pardone the pun).
                        
                        'As far as the recordset, do with it
                        'what you will -- delete it, edit
                        'it, get the other fields from it, etc.
                        
                        'Good luck and have fun.
                        
                    
        End If
        If Right$(strWhatToFind, 1) <> "*" Then strWhatToFind = strWhatToFind & "*" Else TwiceAround = True
        strTemp = "] LIKE"
    Loop
    If rsExample.NoMatch Then MsgBox "Unfortunetly the Item was not found.", vbInformation, App.Title & " - Item not found"
End If
Searching = False
KeyPreview = True
Exit Function
ErrorHandler:
Searching = False
MsgBox Error
End Function











Private Sub cmdRunSearch_Click()
MsgBox "Did you like it?" & vbCr & vbCr & "Did you learn somthing?" & vbCr & vbCr & "Did it work?" & vbCr & vbCr & "Thanks for checking out my code.  Always remember that both you AND " & vbCr & "the user are supposed to have fun.  Working is easy, anybody can work -- enjoying " & vbCr & "your work is the challenge." & vbCr & vbCr & "- KWM", vbQuestion, "RunSearch Example"
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    RefreshMe
ElseIf KeyCode = 70 And Shift = 2 Then
    PopupMenu mnuQuickSearchParent, , , , mnuSearch(0)
End If
End Sub

Private Sub Form_Load()
RefreshMe
End Sub


Private Sub lvReturn_ItemClick(ByVal Item As MSComctlLib.ListItem)
If tmrSearching.Enabled Then Exit Sub
Set mItem = Item
rsExample.FindFirst "ItemID = " & Mid$(mItem.Key, 3)
lvReturn.FullRowSelect = False
lblCurrent = "Item " & mItem.Index
End Sub

Private Sub mnuSearch_Click(Index As Integer)
Dim strField As String
' DON'T FORGET TO TAKE INTO ACCOUNT AMPERSANDS IN THE MENU CAPTIONS
strField = Mid$(mnuSearch(Index).Caption, 2)
''''''''''
''''''''''
If Index > 1 Then RunSearch strField, lvReturn
End Sub

Private Sub lvReturn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Me.PopupMenu mnuQuickSearchParent, , , , mnuSearch(0)
End Sub

Private Sub RefreshMe()
Dim lx&
Dim strItem As String
MousePointer = 11
If ConnectedToDB(dbExample, App.Path & strDB_Path) Then
    Set rsExample = dbExample.OpenRecordset("SELECT * FROM [tblExample];")
    
    With lvReturn
        .ListItems.Clear
        .Sorted = False
        If rsExample.RecordCount > 0 Then rsExample.MoveFirst
        Do Until rsExample.EOF
            strItem = IIf(IsNull(rsExample(.ColumnHeaders((1)))), "", rsExample(.ColumnHeaders(1)))
            
            'Add the item
            Set mItem = .ListItems.Add(, "||" & rsExample("ItemID"), strItem, 1, 1)
            'Do the subitems
            For lx = 1 To .ColumnHeaders.Count - 1
                If IsNull(rsExample(.ColumnHeaders(lx + 1))) Then mItem.SubItems(lx) = "" Else mItem.SubItems(lx) = rsExample(.ColumnHeaders(lx + 1))
            Next
            
            'Refresh for appearance
            If .ListItems.Count = 45 Then
                .Refresh
                DoEvents
            End If
            rsExample.MoveNext
            
            'Warn for excessive returns
            If lx = 10000 Then
                If MsgBox("Excessive returns found for this query.  " & vbCr & "Currently at record number: " & vbCr & vbCr & Format(lx, "#,##0") & vbCr & vbCr & "Continue adding returns to list?", vbYesNo + vbQuestion, App.Title & " Excessively Large Return!") = vbNo Then Exit Do
            End If
            lx = lx + 1
        Loop
        rsExample.MoveFirst
        Set mItem = .ListItems(1)
        .Refresh
    End With
    lblRecCount = rsExample.RecordCount & " Records"
    MousePointer = 99
End If
End Sub

Private Sub tmrSearching_Timer()
'Keep the selection on the item we asked for, until we're done.
On Error GoTo ct
If Not mItem.Selected Then
    mItem.Selected = True
    mItem.EnsureVisible
ct: tmrSearching.Enabled = False
End If

End Sub
