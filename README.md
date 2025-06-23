# MultiPathContextMenu v1.3
Show an IContextMenu for files across multiple paths (and drives!)

![image](https://github.com/user-attachments/assets/53e467ad-db0e-4841-b3ef-fca0e74c89bc)

This method takes advantage of two new features in Windows 7+, Libraries, and Search
Folders. Libraries were created for the express purpose of combining multiple paths as 
one, so it's a natural fit. Unlike some other methods for this, using Libraries helps
ensure that it works smoothly even when files are spread across drive letters. We're not 
backed by an Explorer window here, so we need a way of getting only the folders and 
files we need. For that we hook it up with search.

First, the search scope is set: We take the set of full paths and create a de-duplicated
list of folders, then add them to a new Shell Library object (purely virtual, it's not
creating a .library-ms file). 

Then, we use the SearchFolderItemFactory class and create a condition for it that matches
only our exact files-- while this is a shell search, you can search by PROPERTYKEY, and the
PKEY_ItemPathDisplay key is a string containing the full file path, so we can match exactly 
what we want but not mix up e.g. if files with the same name exist in 2+ folders but only 
one was requested.

Finally, that gives us a result as an IShellItem representing a folder containing our files.
And only our files. So we enumerate all the items, get pidls for them, then create an 
IShellItemArray that's based on the search folder, so the pidls are all single level and
work for a context menu. All that's left is to query it for IContextMenu and display!

If you know a better method, that displays the complete context menu you'd get in a real 
Library, by all means share. I tried many other methods; DEFCONTEXTMENU omitted most items
even if the proper registry keys were opened, for example. Multiple people mention using
an IShellFolder implementation, but never any details or source.


**Requirements**
- Windows 7+
- twinBASIC
- Windows Development Library for twinBASIC (References->Available packages).
 
VB6 port:
- [oleexp.tlb](https://www.vbforums.com/showthread.php?786079) with addons mIID.bas and mPKEY.bas (included with download)
- **NOTE:** VB6 port is v1.0 only, it's now behind the v1.1+ updates of the main tB version.

**Changelog**
- v1.3 (23 Jun 2025) - The original version froze if you pass a folder in the drive root, e.g.
                     C:\folder. Numerous methods using documented features were unsuccessful
                     in preventing this, so the standard method now resorts to minimal use of
                     the undocumented ISearchFolderItemFactoryPriv interface. That's not present
                     on Win10+, so we have a fallback for ISearchFolderItemFactoryPrivEx which
                     contains the same method we're interested it. It doesn't use any scope
                     factory or scope APIs, so should be compatible with all current Windows.
                     The original method is included if you want to continue exploring solutions
                     that don't rely on undocumented magic.
- v1.2 (22 Jun 2025) - Demonstration of 2 similar more efficient methods of setting the search
                     scope, using undocumented interfaces and APIs, see SHUndoc.twin for 
                     details. These do not work on Windows 7, and on 8 only Undoc v1 works. For most apps, I recommend sticking to the original
                     method, which uses all documented interfaces/APIs.
- v1.1 (18 Jun 2025) - Support IContextMenu3/2 HandleMenuMsg routing; fix custom owner/coord use
- v1.0 (17 Jun 2025) - Initial release.

**New in v1.2: Undocumented APIs**

These let us skip the whole complicated routine with IShellLibrary and IObjectArray, which seems to speed up the menu appearance noticably. `SHCreateScopeFromIDListsEx` is the fastest and most efficient; from the call stack and symbol files, I could determine that the Scope Factory interfaces were actually just a wrapper for the APIs; unusual, as it's typically the other way around.

```vba
 [InterfaceId("18455d05-d8f8-47f0-ba4c-c3aaf9c7035f")]
  [OleAutomation(False)]
  Interface ISearchFolderItemFactoryPriv Extends IUnknown
      Sub SetScopeWithDepth(ByVal scope As IShellItemArray, ByVal depth As SCOPE_ITEM_DEPTH)
  End Interface
  [InterfaceId("BD59C2F9-F763-400D-A76E-028C35D047B8")]
  [OleAutomation(False)]
  Interface ISearchFolderItemFactoryPrivEx Extends IUnknown
      Sub SetScopeWithDepth(ByVal scope As IShellItemArray, ByVal depth As SCOPE_ITEM_DEPTH)
      Sub SetScopeDirect(ByVal scope As IScope)
  End Interface
  [InterfaceId("54410B83-6787-4418-9735-5AAAABE83A9A")]
  [OleAutomation(False)]
  Interface IScopeFactory Extends IUnknown
      Sub CreateScope(ByRef riid As UUID, ByRef ppv As Any)
      Sub CreateScopeFromShellItemArray(ByVal si As IShellItemArray, ByRef riid As UUID, ByRef ppv As any)
      Sub CreateScopeFromIDLists(ByVal cidl As Long, apidl As LongPtr, ByVal type As SCOPE_ITEM_TYPE, ByVal depth As SCOPE_ITEM_DEPTH, ByVal flags As SCOPE_ITEM_FLAGS, ByRef riid As UUID, ByRef ppv As any)
      Sub CreateScopeItemFromIDList(ByVal pidl As LongPtr, ByVal type As SCOPE_ITEM_TYPE, ByVal depth As SCOPE_ITEM_DEPTH, ByVal flags As SCOPE_ITEM_FLAGS, ByRef id2 As UUID, ByRef ppv As any)
      Sub CreateScopeItemFromKnownFolder(ByRef id1 As UUID, ByVal type As SCOPE_ITEM_TYPE, ByVal depth As SCOPE_ITEM_DEPTH, ByVal flags As SCOPE_ITEM_FLAGS, ByRef id2 As UUID, ByRef ppv As any)
      Sub CreateScopeItemFromShellItem(ByVal si As IShellItem, ByVal type As SCOPE_ITEM_TYPE, ByVal depth As SCOPE_ITEM_DEPTH, ByVal flags As SCOPE_ITEM_FLAGS, ByRef id2 As UUID, ByRef ppv As any)
  End Interface
  [CoClassId("6746C347-576B-4F73-9012-CDFEEA251BC4")]
  [Description("CLSID_ScopeFactory")]
  CoClass ScopeFactory
      [Default] Interface IScopeFactory
  End CoClass
 
    
    Public Declare PtrSafe Function SHCreateScope Lib "Windows.Storage.Search.dll" (ByRef riid As UUID, ByRef ppv As LongPtr) As Long
    Public Declare PtrSafe Function SHCreateScopeFromShellItemArray Lib "Windows.Storage.Search.dll" (ByVal si As IShellItemArray, ByRef riid As UUID, ByRef ppv As Any) As Long
    Public Declare PtrSafe Function SHCreateScopeFromIDListsEx Lib "Windows.Storage.Search.dll" (ByVal cidl As Long, apidl As LongPtr, ByVal type As SCOPE_ITEM_TYPE, ByVal depth As SCOPE_ITEM_DEPTH, ByVal flags As SCOPE_ITEM_FLAGS, ByRef riid As UUID, ByRef ppv As Any) As Long
    Public Declare PtrSafe Function SHCreateScopeItemFromIDList Lib "Windows.Storage.Search.dll" (ByVal pidl As LongPtr, ByVal type As SCOPE_ITEM_TYPE, ByVal depth As SCOPE_ITEM_DEPTH, ByVal flags As SCOPE_ITEM_FLAGS, ByRef id2 As UUID, ByRef ppv As Any) As Long
    Public Declare PtrSafe Function SHCreateScopeItemFromKnownFolder Lib "Windows.Storage.Search.dll" (ByRef id1 As UUID, ByVal type As SCOPE_ITEM_TYPE, ByVal depth As SCOPE_ITEM_DEPTH, ByVal flags As SCOPE_ITEM_FLAGS, ByRef id2 As UUID, ByRef ppv As Any) As Long
    Public Declare PtrSafe Function SHCreateScopeItemFromShellItem Lib "Windows.Storage.Search.dll" (ByVal si As IShellItem, ByVal type As SCOPE_ITEM_TYPE, ByVal depth As SCOPE_ITEM_DEPTH, ByVal flags As SCOPE_ITEM_FLAGS, ByRef id2 As UUID, ByRef ppv As Any) As Long
```
