# Kabatech-Program
The C16/Kabatech machines use a MS-DOS based file (.clc) that must be recreated by this software to work on the machine itself.  This is handled by being able to import and export the .clc files into and out of the excel spreadsheet.

The .clc File
It is important that the layout, structure and text within the .clc are perfect for the program to function correctly on the c16 machines.  If there are errors in this file then it is unlikely that the MS-DOS system will accept the program as what it is.  The .clc files can be opened using Microsoft Notepad.  The following diagram shows some important characteristics of the .clc file.

![Capture](https://user-images.githubusercontent.com/34693504/138613528-5805a97b-e1f8-4617-b978-6f1c3a87ae15.JPG)

1) The program always begins with "01", despite the first column being a 3 digit row count
2) The columns of the file are created using a method known as a fixed width table,  this means that each row has exactly the same number of character in it with white space being padded out with spaces (" ").  For example column 2 must always have 15 characters so entries in here might be "line off marker" or "start          ", notice there are 10 spaces after start to create the total 15.
3) Each column is also delimited with the multiple chacter string "  ³ " (this is 2 spaces folowed a special symbol 3 in superscript, and finally another space).  This means that these characters are always found between columns forming the dotted line boundaries visible in the diagram.  These characters are not included in the character count of each fixed width column.
4) The first column is used as a row counter by the C16 machines and therefore must always be a 3 digit number with preceding zeros that increments by 1 every row.  If numbers are missing or repeated then the program will not load on the c16 machines.
5) The last line on the must always be completely blank, but there must only be 1 blank line.  This includes no spaces, tabs or other invisible characters.
