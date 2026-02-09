VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmStoreEmail 
   Caption         =   "SYSTEM"
   ClientHeight    =   12690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14010
   OleObjectBlob   =   "FrmStoreEmail.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmStoreEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =========================
' UserForm: FrmStoreEmail
' Purpose:
' - Suggest folders based on conversation history.
' - If multiple target folders exist, let user choose.
' - If none found, list all eligible folders from the same store as the selected item.
' - Provide a live filter TextBox to quickly find a folder by name/path.
' =========================

Private mSelectedItem As Object              ' Selected Outlook item (usually MailItem)
Private mSelectedStore As Outlook.store      ' Store of selected item
Private mAllFolderPaths As Collection        ' Internal source of truth for ListBox content

' -------------------------
' CONFIG: folders to ignore
' Add or remove names here easily.
' NOTE: These are compared to the FIRST LEVEL folder name (under the store root).
' -------------------------
Private Function IgnoredFirstLevelFolders() As Variant
    IgnoredFirstLevelFolders = Array( _
        "Inbox", _
        "Drafts", _
        "Sent Items", _
        "Calendar", _
        "Auto Replies", _
        "Deleted Items", _
        "Junk Email", _
        "Outbox", _
        "RSS Feeds", _
        "Sync Issues", _
        "Sync Issues (This computer only)", _
        "Calender (This computer only)", _
        "Contacts (This computer only)", _
        "Journal (This computer only)", _
        "Notes (This computer only)", _
        "Tasks (This computer only)", _
        "Conversation Action Settings Sync Issues (This computer only)", _
        "Quick Step Settings (This computer only)", _
        "Archive" _
    )
End Function

Private Sub UserForm_Initialize()
    Me.ListBox1.Clear
    Me.txtFolderFilter.Value = ""
    Set mAllFolderPaths = New Collection
End Sub

Public Function PrepareChoices() As Long
    ' Prepares folder choices based on selected item and store.
    ' Returns the number of available folder choices.

    Me.ListBox1.Clear
    Me.txtFolderFilter.Value = ""
    Set mAllFolderPaths = New Collection

    If Not InitSelectedItemAndStore() Then
        PrepareChoices = 0
        Exit Function
    End If

    BuildFolderChoices

    PrepareChoices = Me.ListBox1.ListCount
End Function

Public Function GetSingleChoicePath() As String
    If Me.ListBox1.ListCount = 1 Then
        GetSingleChoicePath = Me.ListBox1.Column(0, 0)
    Else
        GetSingleChoicePath = ""
    End If
End Function

Public Sub MoveToPath(ByVal folderPath As String)
    MoveSelectionToFolder folderPath
End Sub

Private Sub UserForm_Activate()
    ' Put cursor directly in the filter box for fast typing.
    On Error Resume Next
    Me.txtFolderFilter.SetFocus
    On Error GoTo 0
End Sub

Private Function InitSelectedItemAndStore() As Boolean
    ' Gets the item selected in Outlook Explorer and its store.
    On Error GoTo CleanFail

    Dim host As Outlook.Application
    Dim parentFolder As Outlook.folder

    Set host = ThisOutlookSession.Application

    If host.ActiveExplorer Is Nothing Then GoTo CleanFail
    If host.ActiveExplorer.Selection.Count = 0 Then GoTo CleanFail

    Set mSelectedItem = host.ActiveExplorer.Selection.Item(1)
    Set parentFolder = mSelectedItem.Parent
    Set mSelectedStore = parentFolder.store

    InitSelectedItemAndStore = True
    Exit Function

CleanFail:
    InitSelectedItemAndStore = False
End Function

Private Sub BuildFolderChoices()
    ' Main logic:
    ' 1) Try to find target folders based on conversation items.
    ' 2) If none found, load all eligible folders from the same store.

    If TryLoadConversationFolders() Then
        Exit Sub
    End If

    LoadAllEligibleFoldersFromSelectedStore
End Sub

Private Function TryLoadConversationFolders() As Boolean
    ' Returns True if at least one folder was added based on the conversation.
    ' Otherwise returns False.

    On Error GoTo FailSafe

    If Not mSelectedStore.IsConversationEnabled Then
        TryLoadConversationFolders = False
        Exit Function
    End If

    Dim conv As Outlook.Conversation
    Set conv = mSelectedItem.GetConversation

    If conv Is Nothing Then
        TryLoadConversationFolders = False
        Exit Function
    End If

    Dim roots As Outlook.SimpleItems
    Set roots = conv.GetRootItems

    Dim it As Object
    For Each it In roots
        AddItemParentFolderIfEligible it
        WalkConversationChildren it, conv
    Next it

    TryLoadConversationFolders = (Me.ListBox1.ListCount > 0)
    Exit Function

FailSafe:
    ' If anything goes wrong, do not block the user: fallback will take over.
    TryLoadConversationFolders = False
End Function

Private Sub WalkConversationChildren(ByVal anItem As Object, ByVal conv As Outlook.Conversation)
    ' Recursively walks conversation children and adds eligible folders.

    Dim kids As Outlook.SimpleItems
    Set kids = conv.GetChildren(anItem)

    If kids Is Nothing Then Exit Sub
    If kids.Count = 0 Then Exit Sub

    Dim it As Object
    For Each it In kids
        AddItemParentFolderIfEligible it
        WalkConversationChildren it, conv
    Next it
End Sub

Private Sub AddItemParentFolderIfEligible(ByVal it As Object)
    ' Adds the parent folder path of a conversation item to the list if eligible and unique.

    On Error GoTo SafeExit

    If Not (TypeOf it Is Outlook.MailItem Or TypeOf it Is Outlook.AppointmentItem Or TypeOf it Is Outlook.MeetingItem) Then
        GoTo SafeExit
    End If

    Dim fld As Outlook.folder
    Set fld = it.Parent

    ' Only keep folders from the same store as the selected item
    If fld.store.StoreID <> mSelectedStore.StoreID Then GoTo SafeExit

    Dim folderPathEncoded As String
    folderPathEncoded = Replace(fld.folderPath, "%2F", "/")

    If IsFolderEligible(folderPathEncoded) Then
        AddFolderToListBoxUnique folderPathEncoded
    End If

SafeExit:
End Sub

Private Sub LoadAllEligibleFoldersFromSelectedStore()
    ' Enumerates all folders in the selected store and adds eligible ones.

    On Error GoTo SafeExit

    Me.ListBox1.Clear
    Set mAllFolderPaths = New Collection

    Dim root As Outlook.folder
    Set root = mSelectedStore.GetRootFolder

    AddFoldersRecursive root

    Exit Sub

SafeExit:
End Sub

Private Sub AddFoldersRecursive(ByVal f As Outlook.folder)
    ' Recursively adds eligible folders from a store root.

    On Error GoTo SafeExit

    Dim folderPathEncoded As String
    folderPathEncoded = Replace(f.folderPath, "%2F", "/")

    If IsFolderEligible(folderPathEncoded) Then
        AddFolderToListBoxUnique folderPathEncoded
    End If

    Dim subF As Outlook.folder
    For Each subF In f.Folders
        AddFoldersRecursive subF
    Next subF

SafeExit:
End Sub

Private Function IsFolderEligible(ByVal folderPathEncoded As String) As Boolean
    ' Applies exclusion rules:
    ' - Exclude any folder path containing [Gmail]
    ' - Exclude if the first level folder is in the ignore list
    ' - Exclude if first level contains "Shared Data" (your previous rule)

    IsFolderEligible = False

    If InStr(1, folderPathEncoded, "[Gmail]", vbTextCompare) > 0 Then Exit Function

    Dim firstLevel As String
    firstLevel = GetFirstLevelFolder(folderPathEncoded)

    If firstLevel = "" Then Exit Function

    If IsInIgnoredFirstLevel(firstLevel) Then Exit Function
    If InStr(1, firstLevel, "Shared Data", vbTextCompare) > 0 Then Exit Function

    IsFolderEligible = True
End Function

Private Function GetFirstLevelFolder(ByVal folderPathEncoded As String) As String
    ' Converts "\\Account\Folder\Sub" to firstLevel = "Folder"
    On Error GoTo SafeExit

    Dim sfld As String
    sfld = Mid(folderPathEncoded, InStr(3, folderPathEncoded, "\") + 1)

    If Len(sfld) = 0 Then GoTo SafeExit

    GetFirstLevelFolder = Split(sfld, "\")(0)
    Exit Function

SafeExit:
    GetFirstLevelFolder = ""
End Function

Private Function IsInIgnoredFirstLevel(ByVal firstLevel As String) As Boolean
    ' Checks if firstLevel folder is in the ignore array.

    Dim v As Variant
    For Each v In IgnoredFirstLevelFolders()
        If StrComp(CStr(v), firstLevel, vbTextCompare) = 0 Then
            IsInIgnoredFirstLevel = True
            Exit Function
        End If
    Next v

    IsInIgnoredFirstLevel = False
End Function

Private Sub AddFolderToListBoxUnique(ByVal folderPathEncoded As String)
    ' Adds folder path to internal collection (unique) and ListBox.

    ' Unique by key
    On Error Resume Next
    mAllFolderPaths.Add folderPathEncoded, folderPathEncoded
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Me.ListBox1.AddItem folderPathEncoded
End Sub

Private Sub txtFolderFilter_Change()
    ' Live filtering of the list.
    ApplyFolderFilter Me.txtFolderFilter.Text
End Sub

Private Sub ApplyFolderFilter(ByVal filterText As String)
    ' Filters ListBox content using mAllFolderPaths as the source of truth.

    Me.ListBox1.Clear

    Dim filterLower As String
    filterLower = LCase(filterText)

    Dim v As Variant
    For Each v In mAllFolderPaths
        If filterLower = "" Then
            Me.ListBox1.AddItem CStr(v)
        ElseIf InStr(1, LCase(CStr(v)), filterLower, vbTextCompare) > 0 Then
            Me.ListBox1.AddItem CStr(v)
        End If
    Next v
End Sub

Private Sub ListBox1_Click()
    ' User picked a folder, move selection and close.

    If Me.ListBox1.ListIndex < 0 Then Exit Sub

    MoveSelectionToFolder Me.ListBox1.Value
    Unload Me
End Sub

Private Sub MoveSelectionToFolder(ByVal inputFolderPath As String)
    ' Moves all selected items in Explorer to the destination folder.
    ' Adds robustness against 80040109 "message has been changed".

    Dim destFolder As Outlook.folder
    Set destFolder = GetFolderByPath(inputFolderPath)

    If destFolder Is Nothing Then
        MsgBox "Destination folder not found: " & inputFolderPath, vbExclamation
        Exit Sub
    End If

    Dim sel As Outlook.Selection
    Set sel = Application.ActiveExplorer.Selection

    If sel Is Nothing Or sel.Count = 0 Then Exit Sub

    ' Freeze selection by storing EntryIDs
    Dim ids() As String
    Dim i As Long
    ReDim ids(1 To sel.Count)

    For i = 1 To sel.Count
        On Error Resume Next
        ids(i) = sel.Item(i).entryID
        On Error GoTo 0
    Next i

    Dim ns As Outlook.NameSpace
    Set ns = Application.Session

    Dim objItem As Object

    For i = LBound(ids) To UBound(ids)

        If Len(ids(i)) = 0 Then GoTo NextItem

        On Error Resume Next
        Set objItem = ns.GetItemFromID(ids(i))
        On Error GoTo 0

        If objItem Is Nothing Then GoTo NextItem

        ' Only attempt move when the item has a folder parent
        If TypeName(objItem.Parent) <> "MAPIFolder" Then GoTo NextItem

        ' Skip if destination is current folder
        If objItem.Parent Is destFolder Then
            Debug.Print "Skipped (same folder): " & SafeConversationTopic(objItem)
            GoTo NextItem
        End If

        ' First attempt
        If Not TrySaveAndMove(objItem, destFolder) Then
            Debug.Print "Failed to move after retry: " & SafeConversationTopic(objItem)
        End If

NextItem:
        Set objItem = Nothing
    Next i

End Sub

Private Function TrySaveAndMove(ByVal objItem As Object, ByVal destFolder As Outlook.folder) As Boolean
    ' Returns True if moved successfully.
    ' Handles 80040109 with retries.

    Dim attempt As Long

    For attempt = 1 To 2

        On Error GoTo ErrHandler

        objItem.Save
        DoEvents

        objItem.Move destFolder

        TrySaveAndMove = True
        Exit Function

ErrHandler:
        ' -2147221239 = 80040109 "message has been changed"
        If Err.Number = -2147221239 Then
            Err.Clear
            DoEvents
            ' retry loop continues
        Else
            ' other error, do not retry
            TrySaveAndMove = False
            Exit Function
        End If

    Next attempt

    ' If we reach here, retries were exhausted
    TrySaveAndMove = False
End Function

Private Function SafeConversationTopic(ByVal it As Object) As String
    On Error GoTo SafeExit
    SafeConversationTopic = it.ConversationTopic
    Exit Function
SafeExit:
    SafeConversationTopic = "(no topic)"
End Function

Private Function GetFolderByPath(ByVal folderPath As String) As Outlook.folder
    ' Converts a folder path like "\\Mailbox - Name\Folder\Sub" to a Folder object.

    On Error GoTo GetFolder_Error

    Dim foldersArray As Variant
    Dim i As Long
    Dim testFolder As Outlook.folder

    If Left(folderPath, 2) = "\\" Then
        folderPath = Mid(folderPath, 3)
    End If

    foldersArray = Split(folderPath, "\")
    If UBound(foldersArray) < 0 Then GoTo GetFolder_Error

    Set testFolder = Application.Session.Folders.Item(CStr(foldersArray(0)))
    If testFolder Is Nothing Then GoTo GetFolder_Error

    For i = 1 To UBound(foldersArray)
        Set testFolder = testFolder.Folders.Item(CStr(foldersArray(i)))
        If testFolder Is Nothing Then GoTo GetFolder_Error
    Next i

    Set GetFolderByPath = testFolder
    Exit Function

GetFolder_Error:
    Set GetFolderByPath = Nothing
End Function


