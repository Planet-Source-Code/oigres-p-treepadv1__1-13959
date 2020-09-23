VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Treepad-V1 "
   ClientHeight    =   4320
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdNodeAction 
      Height          =   360
      Index           =   3
      Left            =   2520
      Picture         =   "treepadV1Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "delete node"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton cmdNodeAction 
      Height          =   360
      Index           =   2
      Left            =   2160
      Picture         =   "treepadV1Form1.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "add child node"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton cmdNodeAction 
      Height          =   360
      Index           =   1
      Left            =   1800
      Picture         =   "treepadV1Form1.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "add node after"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton cmdNodeAction 
      Height          =   360
      Index           =   0
      Left            =   1440
      Picture         =   "treepadV1Form1.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "add node before"
      Top             =   0
      Width           =   360
   End
   Begin VB.PictureBox picSplitter 
      Height          =   3495
      Left            =   3120
      ScaleHeight     =   3435
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   480
      Width           =   90
   End
   Begin VB.TextBox Text1 
      Height          =   3495
      HideSelection   =   0   'False
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6165
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuInsert 
         Caption         =   "&Insert"
         Begin VB.Menu mnuBefore 
            Caption         =   "&Before"
         End
         Begin VB.Menu mnuAfter 
            Caption         =   "&After"
         End
         Begin VB.Menu mnuchild 
            Caption         =   "&Child"
         End
      End
      Begin VB.Menu mnuDeleteNode 
         Caption         =   "&Delete Node"
      End
      Begin VB.Menu mnuEditNodeText 
         Caption         =   "&EditNodeText"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuFindNode 
         Caption         =   "Find &Node"
      End
      Begin VB.Menu mnuFindInText 
         Caption         =   "Find In &Text"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelpDetails 
         Caption         =   "&HelpDetails"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'recreation of Treepad by Henk Kagedoorn
'Henk's Version 2.6.9c allows you to have/move a node
'on the same level as the root (an error ?)(it saves it strangely- childofroot)
'If a tree structure has a root (only 1 root) and 0 or children
'then mine is okay(it still might have bugs).
'
'<Treepad by Sergio Perciballi>
'email: my easteregg
'limitations: big files are slow to load(treeview doesn't like >integer (32676) nodes)
'I accept no responsibilty for errors or damage from this program
'use at own risk
Dim bMouseDown As Boolean
Dim bNeedToSave As Boolean 'file flag
Dim bInEditmode As Boolean, strEditText As String
Dim bLoadInProgress  As Boolean
Dim loadedFilename As String
Dim intIndentLevel As Long, ff As Integer
Dim lastfocus As String
' The source control
Dim SourceTreeView As TreeView
' The source Node object
Dim SourceNode As node
' The state of Shift key during the drag-and-drop operation
Dim ShiftState As Integer

Private Sub mnuAfter_Click()
    cmdNodeAction_Click 1
End Sub

Private Sub mnuBefore_Click()
    cmdNodeAction_Click 0

End Sub

Private Sub mnuchild_Click()
    cmdNodeAction_Click 2
End Sub

Private Sub mnuDeleteNode_Click()
    cmdNodeAction_Click 3
End Sub

Private Sub mnuEditNodeText_Click()
    EditNodeText
End Sub

Private Sub mnuFindNode_Click()
    'search tree for node text
    Dim node As node, strSearchFor, index As Long

    strSearchFor = VBA.InputBox$("Search For", "Node Text Search")
    If strSearchFor <> "" Then
        MsgBox "you input " & strSearchFor
        For index = 1 To TreeView1.Nodes.Count

            If TreeView1.Nodes.Item(index).Text = strSearchFor Then
                Set node = TreeView1.Nodes.Item(index)
                node.Selected = True
                node.EnsureVisible
                bLoadInProgress = True
                Text1.Text = node.Tag
                bLoadInProgress = False
                Exit For
            End If

        Next index
    End If
    Set node = Nothing


End Sub

Private Sub mnuFindInText_Click()
    'search tree for node text
    Dim node As node, strSearchFor, index As Long
    Dim textpos As Long
    If TreeView1.Nodes.Count = 0 Then Exit Sub
    strSearchFor = VBA.InputBox$("Search For Text", "Text Search")
    If strSearchFor <> "" Then
        MsgBox "you input " & strSearchFor
        'nodes collection is 1 based
        For index = 1 To TreeView1.Nodes.Count
            textpos = InStr(1, TreeView1.Nodes.Item(index).Tag, strSearchFor)
            If textpos > 0 Then
                Set node = TreeView1.Nodes.Item(index)
                node.Selected = True
                node.EnsureVisible
                bLoadInProgress = True
                Text1.Text = node.Tag
                'select text -selstart is 0 based
                Text1.SelStart = textpos - 1
                Text1.SelLength = Len(strSearchFor)
                Text1.SetFocus
                bLoadInProgress = False
                Exit For
            End If

        Next index
    End If
    Set node = Nothing
End Sub

Private Sub mnuHelpDetails_Click()
    Dim strMsg As String
    strMsg = "Drag nodes with right mouse button to move node(s)" & vbCrLf
    strMsg = strMsg & "Press control and drag with right mouse button to copy node(s)"
    MsgBox strMsg
End Sub

Private Sub mnuSaveAs_Click()
    'give new file name
    '
    Dim breturn As Boolean, strLastFile As String
    Dim strPath As String, pos As Long

    If loadedFilename <> "" Then
        pos = VBA.InStrRev(loadedFilename, "\") ', 1)
        strPath = Mid$(loadedFilename, 1, pos - 1)
        strLastFile = Mid$(loadedFilename, pos + 1)
        'MsgBox strpath
        'MsgBox strlastfile

        breturn = SaveTextControl(Text1, cdl1, strLastFile)

        If breturn = True Then
            'MsgBox "Return good"
            bNeedToSave = False
        Else
            'cancelled save dialog or error
            '''MsgBox "Error saving file"
        End If
    End If
End Sub

Private Sub Text1_GotFocus()
lastfocus = "Text1"
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    If NewString = vbNullString Then
        'no need to save; no change;no new string
    Else
        bNeedToSave = True

    End If
End Sub

Private Sub TreeView1_GotFocus()
lastfocus = "Treeview1"
End Sub

Private Sub TreeView1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        EditNodeText
    Else
        'test debug- number of vbcrlf
        If KeyCode = vbKeyF7 Then
            Dim node As node, pos As Long, cnt As Long
            Set node = TreeView1.SelectedItem
            If node Is Nothing Then
            Else
                'get number of vbcrlf
                pos = 1: cnt = 0

                pos = InStr(pos, node.Tag, vbCrLf)
                While pos > 0
                    cnt = cnt + 1
                    pos = InStr(pos + 1, node.Tag, vbCrLf)
                Wend
                MsgBox "Count of vbcrlf = " & cnt
            End If
        End If
        'But what does it say?--------------------------
        If KeyCode = 8.5 * 7 * 2 And ((Shift And vbAltMask) > 0 And (Shift And vbCtrlMask) > 0) Then

            MsgBox retCodeStr("b_ffi&siob[p_`ioh^nb__[mn_l_aaical_m:jimng[mn_l(]i(oe")
        End If
    End If
End Sub

Private Sub EditNodeText()
    Dim node As node
    Set node = TreeView1.SelectedItem
        If node Is Nothing Then
        'do nothing
        Else
            TreeView1.StartLabelEdit

        End If
    Set node = Nothing
End Sub

'massive bug detected ; using form scalemode vbpixels
'causes drag and drop in treeview to not work correctly
'Routine adapted From  Programming Microsoft Visual Basic 6.0
'By Francesco Balena - (buy that book!)
Private Sub TreeView1_MouseDown( _
            Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Check whether we are starting a drag operation.
    If Button <> 2 Then
        'Debug.Print "button <> 2"
        Exit Sub
    Else
        'Debug.Print "Okay"
    End If
    ' Set the Node being dragged, or exit if there is none.
    Set SourceNode = TreeView1.HitTest(x, y)
    If SourceNode Is Nothing Then
        'Debug.Print "no node mdown"
        Exit Sub
    Else
        'Debug.Print SourceNode.Text
    End If
    ' Save values for later.
    Set SourceTreeView = TreeView1
    ShiftState = Shift
    ' Start the drag operation.
    TreeView1.OLEDrag
End Sub
'Routine adapted From  Programming Microsoft Visual Basic 6.0
'By Francesco Balena - (buy that book!)
Private Sub TreeView1_OLEStartDrag( _
            Data As MSComctlLib.DataObject, AllowedEffects As Long)
    ' Pass the Key property of the Node being dragged.
    ' (This value is not used; we can actually pass anything.)
    '''If SourceNode Is Nothing Then Exit Sub
    Data.SetData SourceNode.Key
    If ShiftState And vbCtrlMask Then
        AllowedEffects = vbDropEffectCopy
    Else
        AllowedEffects = vbDropEffectMove
    End If
End Sub
'Routine adapted From  Programming Microsoft Visual Basic 6.0
'By Francesco Balena - (buy that book!)
Private Sub TreeView1_OLEDragOver( _
            Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, _
            Shift As Integer, x As Single, y As Single, State As Integer)
    ' Highlight the Node the mouse is over.
    '''Debug.Print x; y
    If TreeView1.HitTest(x, y) Is Nothing Then
    Else
        Debug.Print TreeView1.HitTest(x, y).Text
    End If
    Set TreeView1.DropHighlight = TreeView1.HitTest(x, y)
End Sub
'Routine adapted From  Programming Microsoft Visual Basic 6.0
'By Francesco Balena - (buy that book!)
Private Sub TreeView1_OLEDragDrop( _
            Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, _
            Shift As Integer, x As Single, y As Single)
    Dim dest As node, nd As node
    ' Get the target Node.
    Set dest = TreeView1.DropHighlight
    '    MsgBox dest.Index
    If dest Is Nothing Then
       
        Exit Sub
    Else
        'MsgBox "destination " & dest.Text
        ' Check that the destination isn't a descendant of the source
        ' Node.
        If SourceTreeView Is TreeView1 Then
            Set nd = dest
            Do
                If nd Is SourceNode Then
                    MsgBox "Unable to drag Nodes here", vbExclamation
                    Set TreeView1.DropHighlight = Nothing

                    Exit Sub
                End If
                Set nd = nd.Parent
            Loop Until nd Is Nothing
        End If
        Set nd = TreeView1.Nodes.Add(dest.index, tvwChild, , _
                SourceNode.Text, SourceNode.Image)
        'copy tags for treeview file text
        nd.Tag = SourceNode.Tag
        ''MsgBox "nd --- = " & nd.Text & ":" & nd.Tag
    End If
    'nd.ExpandedImage = 2:
    nd.Expanded = True

    ' Copy the subtree from source to target control.
    CopySubTree SourceTreeView, SourceNode, TreeView1, nd
    'select
    nd.Selected = True
    Text1.Text = nd.Tag
    bNeedToSave = True
    ' If this is a move operation, delete the source subtree.
    If Effect = vbDropEffectMove Then
        SourceTreeView.Nodes.Remove SourceNode.index
    End If
    Set TreeView1.DropHighlight = Nothing
End Sub
'Routine adapted From  Programming Microsoft Visual Basic 6.0
'By Francesco Balena - (buy that book!)
Sub CopySubTree(SourceTV As TreeView, sourceND As node, _
            DestTV As TreeView, destND As node)
    ' Copy or move all children of a Node to another Node.
    Dim i As Long, so As node, de As node
    If sourceND.Children = 0 Then Exit Sub
    'MsgBox "---0.so = " & sourceND.Text & ": " & destND.Text

    Set so = sourceND.Child
    ''MsgBox "1.so = " & so.Text

    For i = 1 To sourceND.Children
        ' Add a Node in the destination TreeView control.

        Set de = DestTV.Nodes.Add(destND, tvwChild, , so.Text, _
                so.Image, so.SelectedImage)
        de.Tag = so.Tag
        de.EnsureVisible
        ''MsgBox "2.dee = " & de.Text & " s0== " & so.Text
        '''de.Tag =
        de.ExpandedImage = so.ExpandedImage
        ' Now add all the children of this Node in a recursive manner.
        ''MsgBox "3.before so= " & so.Text & " de = " & de.Text
        CopySubTree SourceTV, so, DestTV, de
        ''MsgBox "4.after so= " & so.Text & " de = " & de.Text
        ' Get a reference to the next sibling.
        Set so = so.Next
        ''     MsgBox "copy " & so.Text
    Next
    ''''MsgBox "5.last  so= " & so.Text & " de = " & de.Text
End Sub

Private Sub cmdNodeAction_Click(index As Integer)
    '4 cmd buttons
    Dim node As node
    Set node = TreeView1.SelectedItem
    If node Is Nothing Then
        MsgBox "nothing selected"
        Set node = Nothing
        Exit Sub

    End If
    'if root node selected and we try to add node after/before or delete root node
    'then exit sub
    If node.index = 1 And (index = 0 Or index = 1 Or index = 3) Then
        Beep
        TreeView1.SetFocus
        Set node = Nothing
        Exit Sub
    End If


    'we get here so we can add/delete nodes
    Select Case index
    Case 0 ' add node before
        Set node = TreeView1.Nodes.Add(node, tvwPrevious, , "New Subject")

    Case 1 'add node after

        Set node = TreeView1.Nodes.Add(node, tvwNext, , "New Subject")

    Case 2 'add child node
        Set node = TreeView1.Nodes.Add(node, tvwChild, , "New Subject")

    Case 3 'delete node

        TreeView1.Nodes.Remove node.index
        TreeView1.SetFocus
        bNeedToSave = True
        Set node = Nothing
        Exit Sub
    End Select
    '
    TreeView1.SetFocus
    bLoadInProgress = True
    node.Selected = True
    Text1.Text = ""
    bLoadInProgress = False
    TreeView1.StartLabelEdit
    bNeedToSave = True

    Set node = Nothing

End Sub



Private Sub Form_Load()
    '
    Dim strAnswer As String
    'initialise global variables
    bMouseDown = False
    bNeedToSave = False
    bLoadInProgress = False
    loadedFilename = ""
    Me.Show 'show all form before we focus on anything
    'see if already stored file previously in registry
    strAnswer = GetSetting("TreepadV1", "LastFile", "Path")
    If strAnswer = "" Then
        ''MsgBox "no last file"
    Else
        '
        'see if fileexists
        If FileExists(strAnswer) = True Then
            'open load file
            ''MsgBox "open file on load"
            
            If LoadTreeViewFromFile(strAnswer, TreeView1) = True Then
            Else
                MsgBox "Error loading treeview from file"
            End If
           
           ' loadedFilename = strAnswer
        End If

    End If

    '''Call doTreenodes
End Sub
'
Sub doTreenodes()
    Dim node As node
    Set node = TreeView1.Nodes.Add(, , "Root", "Root")
    Set node = TreeView1.Nodes.Add(node, tvwChild, , "child1")
    Set node = TreeView1.Nodes.Add(node, tvwChild, , "child2")
    Set node = TreeView1.Nodes.Add(node, tvwChild, , "child3")
    Set node = Nothing
End Sub
Private Sub Form_Resize()
    'adjust controls -
    'stops minimize error
    On Error Resume Next
    picSplitter.Top = 25 * Screen.TwipsPerPixelY
    picSplitter.Height = Form1.Height - 80 * Screen.TwipsPerPixelY '(Form1.Height / Screen.TwipsPerPixelY) - 80
    '
    Text1.Visible = False
    TreeView1.Visible = False
    Text1.Top = picSplitter.Top
    Text1.Height = picSplitter.Height
    Text1.Left = picSplitter.Left + picSplitter.Width
    'Text1.Width = (Form1.Width / Screen.TwipsPerPixelX) - picSplitter.Left - picSplitter.Width - 10
    Text1.Width = Form1.Width - picSplitter.Left - picSplitter.Width - 15 * Screen.TwipsPerPixelX
    '
    TreeView1.Top = picSplitter.Top
    TreeView1.Height = picSplitter.Height
    TreeView1.Width = picSplitter.Left - TreeView1.Left  '(Form1.Width / Screen.TwipsPerPixelX)

    Text1.Visible = True
    TreeView1.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim iReturn As Integer, bRet As Boolean
    '
    If bNeedToSave = True Then
        iReturn = MsgBox("Do you want to save this file?", vbYesNoCancel, "Save File")
        If iReturn = vbCancel Then
            MsgBox "Cancelled"
            'cancel=true means stop unloading form
            Cancel = True
            Exit Sub
        End If
        If iReturn = vbYes Then
            'save to file

            MsgBox "save to file"
            'see if we have a filename
            If loadedFilename <> "" Then
                setupSaveFile loadedFilename
            Else
                'no filename so get the commondialogbox save
                bRet = SaveTextControl(Text1, cdl1, "Save File Name")
                If bRet = True Then
                    Cancel = False 'do not stop unloading form
                Else
                    'cancelled save dialog-
                    Cancel = True
                End If
            End If


        End If
        If iReturn = vbNo Then
            'we are leaving without a save
            ''MsgBox "Leave no save"
            Cancel = False

        End If
    End If

End Sub

Private Sub mnuAbout_Click()
    MsgBox "TreepadV1 By Sergio Perciballi Dec/2000" & vbCrLf & "Based on Treepad By Henk Hagedoorn"
End Sub

Private Sub mnuEdit_Click()
    '
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNew_Click()
    Dim node As node, iRet As Integer
    '
    'save last file ?
    If bNeedToSave = True Then
        'get this file saved    'goto save routine
        If loadedFilename <> "" Then
            'save previous file
            iRet = MsgBox("Do you want to save this file?", vbYesNoCancel, "Save previous file")
            If iRet = vbYes Then
                If FileExists(loadedFilename) = True Then
                    setupSaveFile loadedFilename
                End If
            End If
            If iRet = vbCancel Then
                Exit Sub
            End If
        Else

        End If
    Else
    End If
    'clear tree
    TreeView1.Nodes.Clear
    Text1.Text = ""
    loadedFilename = ""
    bNeedToSave = True
    Set node = TreeView1.Nodes.Add(, , "Root", "New File")
    node.Selected = True
    TreeView1.StartLabelEdit
    Set node = Nothing
End Sub
'calls setupSaveFile->parsetree2
'openfile->loadtreeviewfromfile
'savetextcontrol->setupSaveFile->parsetree2
Private Sub mnuOpen_Click()
    Dim iRet As Long
    'how do we know we have file to save ?
    If bNeedToSave = True Then
        'do you want to save last one ?
        If loadedFilename <> "" Then
            'should ask if wants to save previous
            iRet = MsgBox("Save this file before opening another?", vbYesNoCancel, "Save file")
            If iRet = vbYes Then
                'save previous file
                setupSaveFile loadedFilename
                'then open dialog
                openFile
            End If
            If iRet = vbNo Then
                openFile
            End If
            If iRet = vbCancel Then
                Exit Sub
            End If

        Else
            '
            SaveTextControl Text1, cdl1, "Choose A File Name"
        End If
    Else 'bneedtosave=false
        openFile
    End If

End Sub
Private Sub openFile()
    On Error GoTo mnuOpenError
    cdl1.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
    cdl1.Filter = "Treepad Files (*.tpf)|*.tpf"
    cdl1.DefaultExt = "tpf"
    'cdl1.InitDir = "c:\"
    cdl1.ShowOpen
    MsgBox "you chose " & cdl1.Filename
    If (cdl1.Flags And cdlOFNExtensionDifferent) Then
        MsgBox "File not correct extension"
        Exit Sub
    Else
        MsgBox "File is correct extension"
        'open load file
        ''LoadTreeViewFromFile cdl1.Filename, TreeView1
        If LoadTreeViewFromFile(cdl1.Filename, TreeView1) = True Then
            loadedFilename = cdl1.Filename
            SaveSetting "TreepadV1", "LastFile", "Path", loadedFilename
        Else
            MsgBox "Error loading treeview from file"
            'loadedfilename=
        End If


    End If

    Exit Sub
mnuOpenError:
    If Err.Number = cdlCancel Then
        MsgBox "Operation Cancelled"
        '
    Else
        MsgBox Err.Description & ":" & Err.Number
    End If

End Sub

'calls setSavefile->parsetree2
'savetextcontrol->setupsavefile->parsetree2
Private Sub mnuSave_Click()
    Dim breturn As Boolean
    '
    If bNeedToSave = True Then
        If loadedFilename <> "" Then
            '
            setupSaveFile loadedFilename
        Else
            'bring up commondialog save to choose name

            breturn = SaveTextControl(Text1, cdl1, "savefilename")

            If breturn = True Then
                'MsgBox "Return good"
                bNeedToSave = False
            Else
                'or cancelled
                'MsgBox "Error saving file"
            End If
        End If
    End If
End Sub

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = True
    picSplitter.BackColor = vbRed
    
End Sub

Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSplitter.MousePointer = vbSizeWE
    DoEvents
    If bMouseDown = True Then

        picSplitter.Visible = False
        picSplitter.Left = picSplitter.Left + x ' / Screen.TwipsPerPixelX)
        picSplitter.Top = TreeView1.Top ' .Top
        picSplitter.Visible = True
    End If

End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = False
    picSplitter.Visible = False
    Text1.Visible = False
    TreeView1.Visible = False
    picSplitter.BackColor = vbButtonFace
    '
    'Text1.Top = picSplitter.Top
    'Text1.Height = picSplitter.Height
    If picSplitter.Left > Form1.Width - 50 * Screen.TwipsPerPixelX Then
        picSplitter.Left = (Form1.Width - 50 * Screen.TwipsPerPixelX)
    End If
    If picSplitter.Left < 50 * Screen.TwipsPerPixelX Then
        picSplitter.Left = 50 * Screen.TwipsPerPixelX
    End If
    Text1.Left = picSplitter.Left + picSplitter.Width
    Text1.Width = Form1.Width - picSplitter.Left - picSplitter.Width - 15 * Screen.TwipsPerPixelX
    'Text1.Width = (Form1.Width / Screen.TwipsPerPixelX) - picSplitter.Left - picSplitter.Width - 10

    '
    'TreeView1.Top = picSplitter.Top
    'TreeView1.Height = picSplitter.Height
    TreeView1.Width = picSplitter.Left - TreeView1.Left  '(Form1.Width / Screen.TwipsPerPixelX)
    picSplitter.Visible = True
    Text1.Visible = True
    TreeView1.Visible = True
    ChooseSetFocus
    'TreeView1.SetFocus


End Sub

Private Sub ChooseSetFocus()
If lastfocus = "Treeview1" Then
TreeView1.SetFocus
End If
If lastfocus = "Text1" Then
Text1.SetFocus
End If
End Sub
Private Sub Text1_Change()
    Dim node As node
    Set node = TreeView1.SelectedItem
    If node Is Nothing Then Exit Sub
    'stop text change event when loading file
    If bLoadInProgress = True Then
    Else
        node.Tag = Text1.Text

        bNeedToSave = True
    End If
    Set node = Nothing
End Sub

'Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
'    '
'    'Dim bineditmode
'    bInEditmode = True
'    strEditText = TreeView1.SelectedItem.Text
'
'    'MsgBox "beflabeledit"
'
'
'End Sub

'Private Sub TreeView1_KeyPress(KeyAscii As Integer)
'
'    'If KeyAscii = vbKeyReturn Then
'    'If bInEditmode Then
'    '''MsgBox "key return"
'    'bInEditmode = False
'    '
'    'End If
'    '
'    'End If
'
'    'If KeyAscii = vbKeyEscape Then
'    'If bInEditmode Then
'    'TreeView1.SelectedItem.Text = strEditText
'    'bInEditmode = False
'    'End If
'    'End If
'    'If KeyAscii = vbKeyF2 Then
'    '''MsgBox "Edit mode"
'    'End If
'End Sub

'Private Sub TreeView1_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF2 Then
'        'MsgBox "Edit mode"
'        TreeView1.StartLabelEdit
'    End If
'
'
'End Sub
'returns true/false on loading treeview from file
'Routine adapted from www.vb-helper.com
Private Function LoadTreeViewFromFile(ByVal file_name As String, ByVal trv As TreeView) As Boolean
    Dim fnum As Integer
    Dim text_level As String, dummy As String, nodetext As String
    Dim level As Integer
    Dim tree_nodes() As node
    Dim num_nodes As Integer
    Dim maintext As String, inputline As String, previousline As String
   On Error GoTo loadErr
    bLoadInProgress = True
    fnum = FreeFile
    Open file_name For Input As fnum

    trv.Nodes.Clear
    Do While Not EOF(fnum)
        ' <node>
        Line Input #fnum, dummy
        'node text
        Line Input #fnum, nodetext
        'level of indent
        Line Input #fnum, text_level
        'main text goes into node.tag
        maintext = "": inputline = "": previousline = ""
        Do While Not EOF(fnum)
            '<node>
            'nodetext
            'level
            '[node.tag]
            '</node>1z2b3k
            ''previousline = inputline
            Line Input #fnum, inputline
            'If inputline = "</node>1z2b3k" And previousline = vbNullString Then
            If inputline = "</node>1z2b3k" Then
                Exit Do
            Else
                maintext = maintext & (inputline & vbCrLf)
            End If

            '</node>1z2b3k' last line
            ''Line Input #fnum, dummy
        Loop 'Until inputline = "</node>1z2b3k"
        'find rightmost specialChar (chr$(31))

        '        If Right$(maintext, 1) = vbCrLf Then
        '            maintext = Left$(maintext, Len(maintext) - 1)
        '        End If
        Dim pos As Long
        pos = InStr(1, maintext, Chr$(31))
        If pos > 0 Then
            maintext = Left$(maintext, pos - 1)
        End If

        level = Val(text_level) + 1

        ' Make room for the new node.
        If level > num_nodes Then
            num_nodes = level
            ReDim Preserve tree_nodes(1 To num_nodes)
        End If

        ' Add the new node.
        If level = 1 Then
            Set tree_nodes(level) = trv.Nodes.Add(, , , nodetext)
            tree_nodes(level).Tag = maintext
            'tree_nodes(level).Expanded = True
        Else
            Set tree_nodes(level) = trv.Nodes.Add(tree_nodes(level - 1), tvwChild, , nodetext)
            tree_nodes(level).Tag = maintext
            'tree_nodes(level).Expanded = True
            'tons faster without this
            ''tree_nodes(level).EnsureVisible
        End If
    Loop

    Close fnum
    Dim node As node

    TreeView1.SetFocus
    Set node = TreeView1.Nodes.Item(1).Root

    If node Is Nothing Then
        MsgBox "nothing select"
    Else
        node.Selected = True
        Text1.Text = node.Tag
    End If
    bLoadInProgress = False
    loadedFilename = file_name
    Form1.Caption = "Treepad-V1: " & file_name
    Set node = Nothing
    
    LoadTreeViewFromFile = True
    Exit Function
loadErr:
    MsgBox Err.Description & ":" & Err.Number

    Close fnum
    'function return
    LoadTreeViewFromFile = False
    TreeView1.Nodes.Clear
    Form1.Caption = "Treepad-V1: "
    bLoadInProgress = False

End Function
' Returns False if the Save command has been canceled,
' True otherwise. adapted from Programming ms vb6 by Balena
Function SaveTextControl(TB As Control, CD As CommonDialog, _
            Filename As String) As Boolean
    Dim filenum As Integer
    On Error GoTo ExitNow

    CD.Filter = "Treepad Files (*.tpf)"
    CD.FilterIndex = 1
    CD.DefaultExt = "tpf"
    CD.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or _
            cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
    CD.DialogTitle = "Select the destination file "
    CD.Filename = Filename
    ' Exit if user presses Cancel.
    CD.CancelError = True
    CD.ShowSave
    Filename = CD.Filename
    setupSaveFile CD.Filename
    Form1.Caption = "Treepad-V1: " & CD.Filename
    '
    'update registry lastfile entry
    SaveSetting "TreepadV1", "LastFile", "Path", Filename
    
    ' Signal success.
    SaveTextControl = True
ExitNow:

End Function
'called by form_unload;savetextcontrol;
Private Sub setupSaveFile(f As String)
    Dim objnode As node
    ff = FreeFile
   
    Open f For Output As #ff
   
    intIndentLevel = 0

    ' Set the root node
    Set objnode = TreeView1.Nodes(1)

    ' Call the print routine with the root node
    ParseTree2 objnode ', ff

    ' Close the text file
    Close #ff

End Sub

''tree save routine -recursive ;adapted from MSDN sample
Sub ParseTree2(objnode As node)
    Dim intIndentSpaces, mystring As String
    intIndentSpaces = 4
    ' Print the node that was passed in and
    ' account for the node's level
    ''Print #ff, Space(intIndentLevel * intIndentSpaces) & objNode.Text
    Print #ff, "<node>"
    '    Print #ff, Space(intIndentLevel * intIndentSpaces) & objNode.Text
    Print #ff, objnode.Text
    Print #ff, CStr(intIndentLevel)
    'stored data - avoid adding blank line for no data
    mystring = objnode.Tag
    If mystring = vbNullString Then
    Else

        'help to stop accumulated returns put in by print
        'add special character to stop vbcrlf acumulation problem
        'with line input - seems to work
        Print #ff, mystring & Chr$(31)

    End If
    '''end of node record
    Print #ff, "</node>1z2b3k"

    objnode.Selected = True
    ''MsgBox objnode.Text & "level " & intIndentLevel
    ''MsgBox objnode.Tag
    ' Check to see if the current node has children
    If objnode.Children > 0 Then
        ' Increase the indent if children exist
        intIndentLevel = intIndentLevel + 1
        ' Pass the first child node to the print routine
        ParseTree2 objnode.Child ', fnum

    End If
    ' Set the next node to print
    Set objnode = objnode.Next
    ' As long as we have not reached the last node in
    ' a branch, continue to call the print routine
    'If TypeName(objnode) <> "Nothing" Then
    If Not (objnode Is Nothing) Then
        '        objNode.Selected = True
        '    MsgBox objNode.Text
        ParseTree2 objnode
    Else
        ' If the last node of a branch was reached,
        ' decrease the indentation counter
        intIndentLevel = intIndentLevel - 1
    End If
End Sub


Private Sub TreeView1_NodeClick(ByVal node As MSComctlLib.node)
    bLoadInProgress = True
    Text1.Text = node.Tag
    bLoadInProgress = False
End Sub

Private Function FileExists(sPath As String) As Boolean
    'fileexists = Dir$(sPath) <> vbNullString
    If Dir$(sPath) <> vbNullString Then
        FileExists = True
    Else
        FileExists = False
    End If

End Function
'
Private Function retCodeStr(str As String) As String
    Dim x As Long, tmp As String, offset As Long
    offset = 6
    For x = 1 To Len(str)
        retCodeStr = retCodeStr & Chr$(Asc(Mid$(str, x, 1)) + offset)

    Next x


End Function



