# E2W
This app was created using WPF forms.

This app searches Word document for special markers, that point to a cell in an Excel table, and replace these markers with value from that cell.
Marker must follow this structure: <#(Worksheet number)#(Cell name)> -
for example: <#3#B23> - this marker will point to cell B23 in 3rd worksheet of Excel table.
Before running this app all instances of opened Word and Excel files should be closed or it might not work correctly. 
Also Word and Excel need to be installed on the pc before running this app.

MainWindow.xaml contains XAML user interface code.
(wordpathbutton) User presses top button that opens OpenFileDialog (that only displays files with .doc and .docx extensions) to choose Word document, that he wishes to modify.
(excelpathbutton) Second button serves the same purpose, but for opening Excel tables.
(executebutton) After both files have been chosen, the 3rd button , that starts modification process, becomes available.
Before pressing 3rd button user can choose where the endfile will be located. 3 options here (radiobuttons with a groupname "outfile"):
1) (WordRadio) save the endfile to the same path as chosen word-file;
2) (ExcelRadio) save the endfile to the same path as chosen excel-file;
3) (PathRadio) manually choose path, where endfile will be saved;
3rd option has a textbox (OutfilePathText), where user can manually type needed path, or open FolderBrowserDialog by pressing OutfilePathButton to choose needed direction.
This textbox is unavalible if PathRadio is not checked, however, should the user press OutfilePathButton and choose the needed direction, 
OutfilePathText will become available and PathRadio will become checked.

MainWindow.xaml.cs contains logic behind the user interface.
Boolean variables availabilityTop, availabilityBottom and availabilityTotal are needed to determine whether both files have been chosen in order to activate modification button.
Variables ofd and fbd are needed to open OpenFileDialog and FolderBrowserDialog.
Variables wordpathfolder and excelpathfolder are needed to save paths to Word and Excel files without the filename itself.
Methods wordpathbutton_Click and excelpathbutton_Click will open OpenFileDialog, filtered to their respectful file extensions, to choose files for modification.
After choosing needed file its path will be inserted into corresponding textbox (wordpath, excelpath), it's folder's path will be written into wordpathfolder or excelpathfolder
and availability variable for this button will become true. Then method checkTextBox will check if the other availability variable is also true, and if it is -
executebutton will become available.
Before pressing executebutton, user can choose from three options of where to save the endfile. 
If he picks option to manually choose save location, the OutfilePathText textbox will become available.
Upon clicking execute button, application will check which of three saving options have been chosen. 
Then it will create a new Replacer(this class and its contents is described in the next section)-type object and give it path to Word file, path to Excel file and - 
depending on which saving option was chosen - path to Word file folder, Excel file folder or custom folder as parameters.
It will then run this object through FindAndReplace method and show user message if modification was completed succesfully or not.

Replacer class contains code that reads and modifies files.
In order to interact with Word and Excel files, this app uses OpenXml libraries.
Replacer class consists of three fields - wordFilePath, excelFilePath, outfilepathfolder first two contain paths to files that will be processed
and last contains path where endfile will be saved - and one method named FindAndReplace.
When running FindAndReplace, the app reads byte data from chosen Word document, which is then written into a memory stream.
Then that stream is used to create a WordprocessingDocument where all the changes will be made.
Similarily to word doc, then we open an excel spreadsheet using SpreadsheetDocument object to read the data that we will transfer to out final document.
To search for markers in text, a new Regex regular expression is created that will match needed text (<#\d+#[A-Z]+\d+>).
To scan a word doc for markers we have to go through each Paragraph's Run element and additionally through each Text element of these Run elements.
All matches to a regular expressions within a Text element are saved into new MatchCollection object.
A cycle starts here that, for every found match will find corresponding cell in excel table and replace this match with the value of that cell.
It works by finding in each match information about needed sheet and cell using regular expressions.
Sheet number is written into new int object and cell reference is written into a new string object.
Using LINQ we Find needed sheet and by going down its tree of objects we can find a needed cerll value.
If cell value is of numeric type we can simply retrieve its value, but if its of type string or bool then we need to do a data conversion.
(for now lets assume that values in excel are either numeric or string - without boolean values)
The cycle is then repeated for all detected regular expressions matches.