<Query Kind="Program">
  <Reference>&lt;RuntimeDirectory&gt;\System.Web.Extensions.dll</Reference>
  <GACReference>Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</GACReference>
  <Namespace>Microsoft.Office.Interop.Excel</Namespace>
  <Namespace>System.Web.Script.Serialization</Namespace>
</Query>

const string ExcelFilePath = @"C:\Test\ExcelHyperlinkFetcher\Test.xlsx";

void Main()
{
	Application excelApp;
	Workbook excelWorkbook;
	
	excelApp = new ApplicationClass();
	excelApp.Visible = true;
	excelWorkbook = excelApp.Workbooks.Open(ExcelFilePath);
	
	IterateCells(excelWorkbook);
	
	Console.Read();
	excelApp.Quit();
}

void IterateCells(Workbook workbook)
{
	foreach(Worksheet worksheet in workbook.Sheets)
	{
		IterateCells(worksheet);
	}
}

void IterateCells(Worksheet worksheet)
{
	Hyperlinks links = worksheet.UsedRange.Hyperlinks;
	
	for(var i = 1; i <= links.Count; i++)
	{
		Hyperlink link = links[i];
		
		var json = new JavaScriptSerializer()
			.Serialize(
				new { 
						WorksheetName = worksheet.Name, 
						LinkCellAddress = link.Range.Address,
						LinkCellRow = link.Range.Row,
						LinkCellColumn = link.Range.Column,
						LinkAddress = link.Address,
						LinkTextToDisplay = link.TextToDisplay
					});
		
		Console.WriteLine(json);
	}
}