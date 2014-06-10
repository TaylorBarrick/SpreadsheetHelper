SpreadsheetHelper
=================

.NET Spreadsheet Object Wrapper for SpreadsheetLight

This assembly will provide attribute based mapping and formatting of classes to spreadsheets using Vincent Tan's Spreadsheet Lite http://spreadsheetlight.com/

The primary class Spreadsheet is a wrapper for the SLDocument class.  It exposes the following methods:

	Spreadsheet.CreateAndAppendWorksheet<T>
	Spreadsheet.Save
	
By default a class without display attributes will include all public properties as columns in the spreadsheet.

Use of the following attributes will alter the display of the columns:
	
	System.ComponentModel.DisplayName -- Changes the column name of the field in the worksheet.
	SpreadsheetHelper.DisplayNoWrap -- Turns off text wrapping on the column.
	SpreadsheetHelper.DisplayWidth -- Supplies the width of the column.
	SpreadsheetHelper.DisplayHide -- Excludes the field from the worksheet.
	
The Hyperlink class constructs a hyperlink and allows for a field to link to internal or external sources.


Installation
-----------

    NuGet PM> Install-Package SpreadsheetHelper

Usage
-----
	
	Given a POCO with or without attributes:
	
		public class POCO
		{
        		[DisplayWidth(50)]
        		public string LongString { get; set; }
    
        		[DisplayWidth(10), DisplayName("IssueDate")]
        		public DateTime BadlyNamed { get; set; }
        		
        		[DisplayHide]
        		public int ExcludedProperty { get; set; }
    		}

	We can construct the following:

		List<POCO> tcs = RetrievePOCOs();
		Spreadsheet doc = new Spreadsheet();
		doc.CreateAndAppendWorksheet<POCO>(tcs);
		doc.Save("somefilename.xlsx");
