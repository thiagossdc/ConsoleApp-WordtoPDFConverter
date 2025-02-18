#Word to PDF Converter:
This is a simple console application that converts Microsoft Word .docx files into .pdf files using the Microsoft Word Interop API.

#Functionality:
The program reads a .docx file specified by the user.
Converts the file to .pdf format.
Saves the .pdf file in the same or a specified directory.
Technologies used
C#: Programming language used.
.NET 8: Framework used to build the application.
Microsoft.Office.Interop.Word: COM library that allows you to interact with Microsoft Word to export documents to PDF.

#Prerequisites:
Microsoft Word must be installed on your computer, as the program uses the Word Interop API.
.NET 8 SDK installed on your machine.
--------------------------------------------

git clone https://github.com/user/ConsoleApp-WordtoPDFConverter.git
cd ConsoleApp-WordtoPDFConverter

dotnet restore
dotnet build
dotnet run

#insert the relative path

Enter the path of the Word file (.docx): C:\Documents\Example\document.docx

#Output

Conversion complete! File saved at: C:\Documents\Example\document.pdf

