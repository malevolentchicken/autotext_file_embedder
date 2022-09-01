/*
    Utility to assist with embedding a file into a Office Doc as AutoText entries.
    Reads in a file.
    Converts to a char array.
    Reverses array.
    Writes out array split into 254 chars per file. Filename appended with section number.

*/

using System;
using System.IO;
using System.Text;

class Test {
    
    private static void SplitByLength(string str, int maxLength) {

        int count = 1;
        for (int index = 0; index < str.Length; index += maxLength) {   
            System.IO.File.WriteAllText("C:\\dev\\AUTOTEXTEMBED" + count.ToString() + ".txt", str.Substring(index, Math.Min(maxLength, str.Length - index)));
            count++;
        }
    }
    
    public static void Main() {
        
        // Can be any filetype
        string path = @"C:\\dev\\embed_me.txt";

        Byte[] bytes = File.ReadAllBytes(path);        

        string byte_string = String.Join(",", bytes);

        var resultstring = new string(byte_string.ToCharArray().Reverse().ToArray());

        // 254 because ActiveDocument.AttachedTemplate.AutoTextEntries(text_entry).Value caps at 255 chars.
        SplitByLength(resultstring, 254);
    }
}

