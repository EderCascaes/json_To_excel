// See https://aka.ms/new-console-template for more information
using System.ComponentModel;
using Aspose.Cells;
using System.IO;
using Aspose.Cells.Utility;





Console.WriteLine("Json to Excel");
Console.WriteLine("Process started");

// Instantiate the license at the beginning of the program to avoid trial version restrictions
//License JsonToExcelLicense = new License();
//JsonToExcelLicense.SetLicense("Aspose.Cells.lic");

// Create a style to format the json fields title in the output workbook 
CellsFactory factory = new CellsFactory();
Style jsonTitleStyle = factory.CreateStyle();
jsonTitleStyle.HorizontalAlignment = TextAlignmentType.Center;
jsonTitleStyle.Font.Color = System.Drawing.Color.BlueViolet;
jsonTitleStyle.Font.IsBold = true;

// Declare and define the layout of the data imported from JSON to Excel
JsonLayoutOptions jsonLayoutOptions = new JsonLayoutOptions();
jsonLayoutOptions.TitleStyle = jsonTitleStyle;
jsonLayoutOptions.ArrayAsTable = true;

// Initialize an empty workbook to import JSON data
Workbook emptyWbForJsonData = new Workbook();

// Get reference to the worksheet where data is to be imported
Worksheet targetWorksheet = emptyWbForJsonData.Worksheets[0];



Console.WriteLine("Reading .json file");
// Read the Json file into a string variable that will be used to import date
//local padrão para os arquivo
//tambem podesse configurar o local dos arquivos
//ex: File.ReadAllText("C:\\Users\\re046322\\documentos\\cushman.json");
//.....\JsonToExcel\JsonToExcel\bin\Debug\net7.0
string inputJsonString = File.ReadAllText("cushman.json");



Console.WriteLine("Writing excel file...");
// Call the ImportData function to import JSON data into the worksheet
JsonUtility.ImportData(inputJsonString, targetWorksheet.Cells, 3, 5, jsonLayoutOptions);

Console.WriteLine("Saving Excel file...");
// Save Excel file
emptyWbForJsonData.Save("cushman.xlsx");

Console.WriteLine("Process finished");