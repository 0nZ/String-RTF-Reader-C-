# String RTF Reader C#
Simple RTF string converter.
<br />
Developed to meet a demand where I needed to export several lines of text in RTF format from a database and, later, read this spreadsheet (exported from DB) and transform the column that contained the RTF string into a text document in *.doce *.rtf formats.
<br /><br />
The system works like this:
<br />
1 - Load a spreadsheet that contains 2 columns, namely "TITULO" and "RTF";
<br />
2 - The system will generate the .rtf files within the "rtf_files" fold and the .doc files within the "doc_files" fold, the file names will be those informed in the "TITULO" column of the spreadsheet and the contents of the files will be the "RTF" column of the spreadsheet.
<br /><br />
In the repository you can find an example spreadsheet (.xlsx).
<br /><br />
The system has already been published, just download and run the application: bin/Release/net8.0-windows/publish/WpfApp1.exe - The output directories for the files will be created automatically.
<br /><br />
Use however you want :)
<br /><br />
![alt text](https://github.com/0nZ/String-RTF-Reader-C-/blob/main/screen.png?raw=true)
