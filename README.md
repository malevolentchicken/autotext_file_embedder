# autotext_file_embedder
A collection of some scripts to assist with embedding a file into a Word Document as a series of AutoText entries.

**Embedding**
- Use the csharp utility to prepare/split the file you wish to embed. 
- Create a macro enabled word doc, insert and run the insert_autotext() function. Run this function to read in the split txt files and insert as autotext entries.

**Executing**

- Insert write_autotext_to_file() into the rest of your macro and modify as needed. Also insert the StrRev() function or inline.

**Notes**

To remove the autotext entries use the delete_autotextentries() function. Make sure to remove delete_autotextentries() and insert_autotext() from your final doc.

A bit clunky but has proven useful. Tested these scripts by embedding a crafted MSBuild xml file into a .doc, write it out, and execute via wmi.

Cheers
