# ImageFactory 1.0.1
Program for making fast image with PDF, SVG or PNG extension from SVG template and XLSX data.

Prgoram requirements: Python 3.0+ and Inkscape.

1.First Create or take SVG from Inkscape (Other program will be not checked).

To change text all your text fields need to start with %VAR_ and end with % (e.g %VAR_Name%).
To change picture you will need add picture as LINK(URL) not Embed it.

2. Xlsx file
Need to have one additional row for table header (I only use Number with export option).
Next row should contain names that you placed with %VAR_(e.g Name)%.
Pictures need to be same name as picture file name without extension (e.g Flower).
Than you need to fill rest row with data that you want to replace.

3. Open ImageFactory

Overwrite files checkbox - user can choose to make new files or overwrite existing files.

Extension are three
- SVG
   fastest work without providing DPI and inkscape.exe
- PNG
  Only this extension need DPI value to works, for me works well on 500 but feel free to use other.
- PDF

DPI - is only needed to formating PNG, this is done by Inkscape program.

File name - can be use on first column %VAR_<something>% or different as you like.
- If you use different name all next files will be incremented.
- If file exists it will be incremented.

Template Directory - Is path with file extension to your SVG template.
Data Directory - Path with file name abd extebsuib to your XLSX.
Output Directory - Path to folder where files will be saved.
Inkscape Directory - path to Inkscape.exe.
! On MacOS application/Inkscape doesn't work. You need path like this "/Applications/Inkscape.app/Contents/MacOS/inkscape" but i suggest find by browse button.

Process button - should start program, it's done when you get window with text "Done".

Worth mention:
- SVG template broke when you change picture directory. You need manualy link them in inkscape.
- You can add column to xlsx named "Quantity" to make more copy of same image in one process.
 
