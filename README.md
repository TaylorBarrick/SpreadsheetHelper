SpreadsheetHelper
=================

.NET Spreadsheet Object Wrapper for SpreadsheetLight

[![Build status](https://ci.appveyor.com/api/github/webhook?id=8b84sd3ogmtxxxlt)](https://ci.appveyor.com/project/TaylorBarrick/spreadsheethelper)

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


License
-----
The MIT License (MIT)

Copyright (c) 2015 Taylor Barrick &lt;TaylorBarrick@gmail.com&gt;

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
