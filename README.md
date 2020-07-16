# E2W
This app was created using WPF forms.

This app searches Word document for special markers, that point to a cell in an Excel table, and replace these markers with value from that cell.
Marker must follow this structure: <#(Worksheet number)#(Cell name)> -
for example: <#3#B23> - this marker will point to cell B23 in 3rd worksheet of Excel table.

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
Before pressing executebutton, user can choose from three options of where to save the endfile. If he picks option to manually choose save location, 
the OutfilePathText textbox will become available
