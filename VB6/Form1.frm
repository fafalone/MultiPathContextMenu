VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi-Path Context Menu Demo"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Menu"
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Files (1 per line)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Ok, ok, since it was only 5 minutes of work (and another 5 debugging) I back-ported this to VB6
'/*
'Multi-path Context Menu Demo
'v1 0#
'Show an IContextMenu for files across multiple paths (and drives!)
'by Jon Johnson
'https://github.com/fafalone/MultiPathContextMenu
'
'This method takes advantage of two new features in Windows 7+, Libraries, and Search
'Folders. Libraries were created for the express purpose of combining multiple paths as
'one, so it's a natural fit. Unlike some other methods for this, using Libraries helps
'ensure that it works smoothly even when files are spread across drive letters. We're not
'backed by an Explorer window here, so we need a way of getting only the folders and
'files we need. For that we hook it up with search.
'
'First, the search scope is set: We take the set of full paths and create a de-duplicated
'list of folders, then add them to a new Shell Library object (purely virtual, it's not
'creating a .library-ms file).
'
'Then, we use the SearchFolderItemFactory class and create a condition for it that matches
'only our exact files-- while this is a shell search, you can search by PROPERTYKEY, and the
'PKEY_ItemPathDisplay key is a string containing the full file path, so we can match exactly
'what we want but not mix up e.g. if files with the same name exist in 2+ folders but only
'one was requested.
'
'Finally, that gives us a result as an IShellItem representing a folder containing our files.
'And only our files. So we enumerate all the items, get pidls for them, then create an
'IShellItemArray that 's based on the search folder, so the pidls are all single level and
'work for a context menu. All that's left is to query it for IContextMenu and display!
'
'If you know a better method, that displays the complete context menu you'd get in a real
'Library, by all means share. I tried many other methods; DEFCONTEXTMENU omitted most items
'even if the proper registry keys were opened, for example. Multiple people mention using
'an IShellFolder implementation, but never any details or source.
'
'
'**Requirements**
'Windows 7+
'This code depends on my WinDevLib package.
'
'**Changelog**
'v1.0.2 - Bug fix: Customer owner and coords removed.
'v1.0 (17 Jun 2025) - Initial release.
'
'*/

'Replace WDL helpers/APIs not in oleexp:
#If (TWINBASIC = 0) Then
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Point) As BOOL
Private Declare Function ILClone Lib "shell32" (ByVal pidl As LongPtr) As LongPtr
Private Declare Function CreatePopupMenu Lib "user32" () As LongPtr
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As LongPtr) As BOOL

Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal uFlags As TPM_wFlags, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As LongPtr, lpRC As Any) As Long
Private Enum TPM_wFlags
    TPM_LEFTBUTTON = &H0
    TPM_RECURSE = &H1
    TPM_RIGHTBUTTON = &H2
    TPM_LEFTALIGN = &H0
    TPM_CENTERALIGN = &H4
    TPM_RIGHTALIGN = &H8
    TPM_TOPALIGN = &H0
    TPM_VCENTERALIGN = &H10
    TPM_BOTTOMALIGN = &H20

    TPM_HORIZONTAL = &H0         ' Horz alignment matters more
    TPM_VERTICAL = &H40            ' Vert alignment matters more
    TPM_NONOTIFY = &H80           ' Don't send any notification msgs
    TPM_RETURNCMD = &H100

    TPM_HORPOSANIMATION = &H400
    TPM_HORNEGANIMATION = &H800
    TPM_VERPOSANIMATION = &H1000
    TPM_VERNEGANIMATION = &H2000
    TPM_NOANIMATION = &H4000
    TPM_LAYOUTRTL = &H8000&
    'Win7+:
    TPM_WORKAREA = &H10000
End Enum

    Private Sub FreeIDListArray(ppidls() As LongPtr, cItems As Long)
        Dim i As Long
        For i = 0 To UBound(ppidls)
            CoTaskMemFree ppidls(i)
        Next i
    End Sub
#End If
    Private Sub Command1_Click() 'Handles Command1.Click
        If Text1.Text = "" Then
            MsgBox "Error: No files entered.", vbCritical Or vbOKOnly, App.Title
            Exit Sub
        End If
        Dim sFiles() As String: sFiles = Split(Text1.Text, vbCrLf)
        Dim i As Long
        For i = 0 To UBound(sFiles)
            If PathFileExists(sFiles(i)) = 0 Then
                MsgBox "Error: All files must exist; " & sFiles(i) & " was not found.", vbCritical Or vbOKOnly, App.Title
                Exit Sub
            End If
        Next
        MultiPathContextMenu sFiles, Me.hwnd
    End Sub
    
    Private Function MultiPathContextMenu(sFiles() As String, ByVal hOwner As LongPtr, Optional ByVal ptX As Long = -1, Optional ByVal ptY As Long = -1, Optional ByVal dwFlags As QueryContextMenuFlags = CMF_EXPLORE) As Long
        Dim pSearchFact As ISearchFolderItemFactory
        Set pSearchFact = New SearchFolderItemFactory
        Dim piaScope As IShellItemArray
        Dim hr As Long
        If CreateSearchScope(sFiles, piaScope) = S_OK Then
            pSearchFact.SetScope piaScope
            pSearchFact.SetDisplayName StrPtr("TempResults")
            Dim pCond As ICondition
            If GetCondition(sFiles, pCond) = S_OK Then
                pSearchFact.SetCondition pCond
                Dim siRes As IShellItem, pidlRes As LongPtr
                Dim pEnum As IEnumShellItems, siChild As IShellItem
                pSearchFact.GetShellItem IID_IShellItem, siRes
                If (siRes Is Nothing) = False Then
                    Dim pidlFQ() As LongPtr, pidlRel() As LongPtr, nPidl As Long, pidlTmp As LongPtr
                    SHGetIDListFromObject siRes, pidlRes
                    siRes.BindToHandler 0, BHID_EnumItems, IID_IEnumShellItems, pEnum
                    If (pEnum Is Nothing) = False Then
                        Dim pc As Long
                        Do While pEnum.Next(1, siChild, pc) = S_OK
                            ReDim Preserve pidlFQ(nPidl)
                            ReDim Preserve pidlRel(nPidl)
                            SHGetIDListFromObject siChild, pidlTmp
                            pidlFQ(nPidl) = ILClone(pidlTmp)
                            pidlRel(nPidl) = ILFindLastID(pidlFQ(nPidl))
                            CoTaskMemFree pidlTmp
                            nPidl = nPidl + 1
                        Loop
                        Dim ppsia As IShellItemArray
                        Dim pCtx As IContextMenu
                        SHCreateShellItemArray pidlRes, Nothing, UBound(pidlRel) + 1, VarPtr(pidlRel(0)), ppsia
                        ppsia.BindToHandler 0, BHID_SFUIObject, IID_IContextMenu, pCtx
                        hr = DisplayContextMenu(pCtx, hOwner, ptX, ptY, dwFlags)
                        FreeIDListArray pidlFQ, UBound(pidlFQ) + 1
                        Set pCtx = Nothing
                        Set ppsia = Nothing
                        Set pEnum = Nothing
                    Else
                        Debug.Print "MultiPathContextMenu::Couldn't get folder enumerator."
                    End If
                    CoTaskMemFree pidlRes
                End If
            End If
            Set pCond = Nothing
            Set siRes = Nothing
            Set pSearchFact = Nothing
        Else
            Debug.Print "MultiPathContextMenu::Couldn't create scope."
        End If
        Set piaScope = Nothing
    End Function
    Private Function DisplayContextMenu(ByVal pCtx As IContextMenu, ByVal hOwner As LongPtr, Optional ByVal ptX As Long = -1, Optional ByVal ptY As Long = -1, Optional ByVal dwFlags As QueryContextMenuFlags = CMF_EXPLORE) As Long
        If (pCtx Is Nothing) = False Then
            Debug.Print "Got context menu"
            Dim hMenu As LongPtr: hMenu = CreatePopupMenu()
            pCtx.QueryContextMenu hMenu, 0, 1, &H7FFF&, dwFlags
            If (ptX = -1) Or (ptY = -1) Then
                Dim pt As Point
                GetCursorPos pt
                ptX = pt.x: ptY = pt.y
            End If
            Dim idCmd As Long: idCmd = TrackPopupMenu(hMenu, TPM_LEFTBUTTON Or TPM_RIGHTBUTTON Or _
                                    TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_HORIZONTAL Or TPM_RETURNCMD, _
                                    ptX, ptY, 0&, hOwner, 0&)
            Debug.Print "Command=" & idCmd
            If idCmd Then
                Dim cmi As CMINVOKECOMMANDINFO
                With cmi
                    .cbSize = LenB(cmi)
                    .hwnd = hOwner
                    .lpVerb = idCmd - 1 ' MAKEINTRESOURCE(idCmd-1);
                    .nShow = SW_SHOWNORMAL
                End With
                pCtx.InvokeCommand VarPtr(cmi)
            End If
            DestroyMenu hMenu
        End If
    End Function
    Private Function CreateSearchLibrary(pObC As IObjectCollection) As Long
    Set pObC = Nothing
    Dim pLib As IShellLibrary
    Set pLib = New ShellLibrary
    If (pLib Is Nothing) = False Then
        CreateSearchLibrary = pLib.GetFolders(LFF_ALLITEMS, IID_IObjectCollection, pObC)
    Else
        Debug.Print "CreateSearchLibrary->Failed to create ShellLibrary"
    End If
    End Function
    Private Function GetFoldersForFiles(sFiles() As String, sFolders() As String) As Long
        'Get a list of the folders our files are in, making sure to add each path only once.
        Dim sFolder As String
        Dim bAdded As Boolean
        Dim nFolders As Long
        Dim i As Long
        ReDim sFolders(0)
        For i = 0 To UBound(sFiles)
            sFolder = Left$(sFiles(i), InStrRev(sFiles(i), "\") - 1)
            If (Len(sFolder) = 2) Then
            If (Right$(sFolder, 1) = ":") Then
                sFolder = sFolder & "\"
            End If
            End If
            bAdded = False
            Dim j As Long
            For j = 0 To UBound(sFolders)
                If LCase$(sFolders(j)) = LCase$(sFolder) Then
                    bAdded = True: Exit For
                End If
            Next
            If bAdded = False Then
                ReDim Preserve sFolders(nFolders)
                sFolders(nFolders) = sFolder
                nFolders = nFolders + 1
            End If
        Next
        GetFoldersForFiles = nFolders
    End Function
    Private Function CreateSearchScope(sFiles() As String, ppia As IShellItemArray) As Long
    On Error GoTo e0
    Set ppia = Nothing
    Dim pObjects As IObjectCollection
    Dim hr As Long
    Dim sFolders() As String
    Dim nFolders As Long: nFolders = GetFoldersForFiles(sFiles, sFolders)
    If nFolders Then
        Dim sia() As IShellItem
        ReDim sia(nFolders - 1)
        Dim i As Long
        For i = 0 To UBound(sFolders)
            SHCreateItemFromParsingName StrPtr(sFolders(i)), Nothing, IID_IShellItem, sia(i)
        Next
        If CreateSearchLibrary(pObjects) = S_OK Then
            Dim j As Long
            For j = 0 To UBound(sia)
                pObjects.AddObject ObjPtr(sia(j))
            Next
            Set ppia = pObjects
            Set pObjects = Nothing
        End If
    End If
    CreateSearchScope = S_OK
e0:
    Debug.Print "Error in CreateSearchScope: 0x" & Hex$(Err.Number) '& ", " & GetSystemErrorString(Err.Number)
    CreateSearchScope = Err.Number
    End Function
    Private Function GetCondition(sFiles() As String, ppCondition As ICondition) As Long
    'Get a search ICondition object that matches only our exact files.
    Set ppCondition = Nothing
    GetCondition = -1
    Dim pFact As IConditionFactory2
    Set pFact = New ConditionFactory
    Dim pFile() As ICondition
    Dim nCds As Long: nCds = UBound(sFiles) + 1
    If (pFact Is Nothing) = False Then
        Dim nCOP As CONDITION_OPERATION: nCOP = COP_EQUAL 'COP_VALUE_CONTAINS
        ReDim pFile(UBound(sFiles))
        Dim i As Long
        For i = 0 To UBound(sFiles)
            pFact.CreateStringLeaf PKEY_ItemPathDisplay, nCOP, StrPtr(sFiles(i)), 0&, CONDITION_CREATION_DEFAULT, IID_ICondition, pFile(i)
        Next
        If nCds = 1 Then
            'Only one condition, don't need an array
            Set ppCondition = pFile(0)
        Else
            pFact.CreateCompoundFromArray CT_OR_CONDITION, pFile(0), nCds, CONDITION_CREATION_DEFAULT, IID_ICondition, ppCondition
        End If
        If (ppCondition Is Nothing) = False Then GetCondition = S_OK
    
        Set pFact = Nothing
    Else
        Debug.Print "GetCondition->Failed to create factory."
    End If
    
    End Function
