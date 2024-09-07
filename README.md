/******************************************************************************* 
Script Name: exportPageRangesToPDF.jsx 
Description: This InDesign script exports page ranges to separate PDFs (and manages files a bit). 
               
  
 Features: 
* export multiple page ranges to individual PDFs. 
* specify export-folder and PDF-preset to use. 
* manually input page ranges a labels/identifiers for each range. 
* retrieve a list of page ranges from former exported PDFs from that directory 
* retrieve a list of page ranges from the active documents sections and section markers. 
* retrieve the section(s) you are currently viewing for a quick versioned export of what you are working on. 
* handles odd/even page starts in PDF-view-settings (cover sheet YES/NO) 
* automatic versioning for exported PDFs, unless overridden. 
* saves settings to a document script label, will start with last settings and position. 
 
User input schema: 
Each line you enter in the text area should follow this format. The label in brackets is optional, don't use commas: 
Page or PageRange(label) 
 
Example: 
1 
5-7(Introduction) 
8-15 
16 
17-20(Chapter 2) 
 
The exported PDF files will be named according to this schema: 
[DocumentName]_pages_[PageRange]_[Label-from-brackets-if-given]_v[Version].pdf 

Example: 
/MyDocument_pages_5-7_Introduction_v1.pdf  

The version-number will auto increment if it finds an existing pdf file with the same name, i.e. from _v1 to _v2. 
In case you want to overwrite the last pdf file version you can suppress the version incrementation by adding a minus to the end of a line.  
 
Example: 
1- 
5-7(Introduction)- 
17-20 

If "MyDocument_pages_1_v1.pdf" and "MyDocument_pages_5-7_Introduction_v1.pdf" already exists it will overwrite those PDFs during export of those page ranges and will not create version _v2 pdfs. The other page range of 17-20 will be exported with version-incrementation to the next higher version. Of course if there was no prior pdf of any of those files, it will create v1-files. 
 
Author: [Stephan Möbius] 
Original idea: [Jongware] 
Multi language L10N-function: Marc Autret, equalizer.js, 2012. 
******************************************************************************/ 
