# MultiPathContextMenu
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
Windows 7+
This code depends on my WinDevLib package.

**Changelog**
v1.0 (17 Jun 2025) - Initial release.

