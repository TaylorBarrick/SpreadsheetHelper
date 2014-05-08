using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Globalization;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Runtime.InteropServices;
using SpreadsheetLight;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics.CodeAnalysis;

namespace SpreadsheetHelper
{
    public class Spreadsheet : IDisposable
    {
        private SLDocument doc;
        public Spreadsheet() { doc = new SLDocument(); }
        private string FirstSheet = "";
        public void CreateAndAppendWorksheet<T>(IEnumerable<T> records, string tabName = "", bool makeTable = true)
        {
            Type t = typeof(T);
            if (tabName == "")
                tabName = String.Format("{0}List", t.Name);
            if (doc.SelectWorksheet("Sheet1"))
                doc.RenameWorksheet("Sheet1", tabName);
            else
                doc.AddWorksheet(tabName);
            if (FirstSheet == "")
                FirstSheet = tabName;

            List<PropertyInfo> properties = GetOrderedProperties(t);
            List<int> hideColumns = new List<int>();
            int propertyCount = properties.Count;
            Type[] propertyTypes = new Type[propertyCount];
            DisplayFormatAttribute[] formatAttributes = new DisplayFormatAttribute[propertyCount];
            DisplayNoWrap[] wrapAttributes = new DisplayNoWrap[propertyCount];
            DisplayNameAttribute[] nameAttributes = new DisplayNameAttribute[propertyCount];
            DisplayWidth[] widthAttributes = new DisplayWidth[propertyCount];
            for (int k = 0; k < propertyCount; k++)
            {
                PropertyInfo p = properties[k];
                if (p.GetCustomAttributes(typeof(DisplayHide), true).Cast<DisplayHide>().FirstOrDefault() != null)
                    hideColumns.Add(k);
                formatAttributes[k] = p.GetCustomAttributes(typeof(DisplayFormatAttribute), true).Cast<DisplayFormatAttribute>().FirstOrDefault();
                wrapAttributes[k] = p.GetCustomAttributes(typeof(DisplayNoWrap), true).Cast<DisplayNoWrap>().FirstOrDefault();
                nameAttributes[k] = p.GetCustomAttributes(typeof(DisplayNameAttribute), true).Cast<DisplayNameAttribute>().FirstOrDefault();
                widthAttributes[k] = p.GetCustomAttributes(typeof(DisplayWidth), true).Cast<DisplayWidth>().FirstOrDefault();
                propertyTypes[k] = p.PropertyType;
            }
            
            CreateHeader(properties, hideColumns, nameAttributes);
            CreateRows<T>(records, properties, hideColumns, formatAttributes, propertyTypes);
            FormatColumns(properties, records, makeTable, hideColumns, wrapAttributes, widthAttributes);
        }

        private List<PropertyInfo> GetOrderedProperties(Type t)
        {
            PropertyInfo[] propArray = t.GetProperties();
            var orderedProps = propArray
                .Select(p => new { p, Atts = p.GetCustomAttributes(typeof(DisplayAttribute), inherit: true) })
                .Where(p => p.Atts.Length != 0)
                .OrderBy(p => ((DisplayAttribute)p.Atts[0]).Order)
                .Select(p => p.p)
                .ToList();
            var unOrderedProps = propArray
                .Select(p => new { p, Atts = p.GetCustomAttributes(typeof(DisplayAttribute), inherit: true) })
                .Where(p => p.Atts.Length == 0)
                .Select(p => p.p)
                .ToList();
            List<PropertyInfo> properties = new List<PropertyInfo>();
            properties.AddRange(orderedProps);
            properties.AddRange(unOrderedProps);
            return properties;
        }

        [ExcludeFromCodeCoverageAttribute]   
        public static string MimeTypeName { get { return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; } }
        public void Save(string fileName)
        {
            if (FirstSheet != "")
                doc.SelectWorksheet(FirstSheet);
            doc.SaveAs(fileName);
        }
        public void Save(Stream stream)
        {
            if (FirstSheet != "")
                doc.SelectWorksheet(FirstSheet);
            doc.SaveAs(stream);
            stream.Position = 0;
        }

        private void CreateHeader(List<PropertyInfo> properties, IEnumerable<int> hideColumns, DisplayNameAttribute[] nameAttributes)
        {
            int columnIndex = 0;
            for (int i = 0; i < properties.Count; i++)
            {
                if (!hideColumns.Contains(i))
                {
                    if (nameAttributes[i] == null)
                        doc.SetCellValue(1, columnIndex + 1, properties[i].Name);
                    else
                        doc.SetCellValue(1, columnIndex + 1, nameAttributes[i].DisplayName);
                    columnIndex++;
                }
            }
        }

        private void FormatColumns<T>(List<PropertyInfo> properties, IEnumerable<T> records, bool makeTable, IEnumerable<int> hideColumns, DisplayNoWrap[] wrapAttributes, DisplayWidth[] widthAttributes)
        {
            int skippedColumns = 0;
            int columnIndex = 0;
            for (int k = 0; k < properties.Count; k++)
            {
                if (!hideColumns.Contains(k))
                {
                    SLStyle style = new SLStyle();
                    if (wrapAttributes[k] != null)
                        style.SetWrapText(false);
                    else
                        style.SetWrapText(true);
                    var widthAttribute = widthAttributes[k];
                    if (widthAttribute != null)
                        doc.SetColumnWidth(columnIndex + 1, widthAttribute.Width);
                    doc.SetColumnStyle(columnIndex + 1, style);
                    columnIndex++;
                }
                else
                    skippedColumns++;
            }
            if (makeTable)
            {
                SLTable tbl = doc.CreateTable(1, 1, records.Count() + 1, properties.Count() - skippedColumns);
                tbl.SetTableStyle(SLTableStyleTypeValues.Light1);
                doc.InsertTable(tbl);
            }
        }

        private void CreateRows<T>(IEnumerable<T> records, List<PropertyInfo> properties, IEnumerable<int> hideColumns, DisplayFormatAttribute[] formatAttributes, Type[] propertyTypes)
        {
            int columnIndex = 0;
            int rowIndex = 0;
            SLStyle nullStyle = new SLStyle();
            nullStyle.Fill.SetPatternType(PatternValues.LightTrellis);
            for (int i = 0; i < records.Count(); i++)
            {
                List<T> list = records.ToList();
                columnIndex = 0;
                if (list.ToList()[i] == null)
                {
                    doc.SetRowStyle(rowIndex+2, nullStyle);
                    rowIndex++;
                }
                else
                {
                    for (int j = 0; j < properties.Count; j++)
                    {
                        if (!hideColumns.Contains(j))
                        {
                            int x = rowIndex + 2, y = columnIndex + 1;
                            var value = properties[j].GetValue(list[i], null);
                            var valueType = propertyTypes[j];
                            double result;
                            if (value != null)
                            {
                                if (properties[j].PropertyType == typeof(DateTime))
                                {
                                    if ((DateTime)value != DateTime.MinValue)
                                    {
                                        if (formatAttributes[j] != null)
                                            doc.SetCellValue(x, y, ((DateTime)value).ToString(formatAttributes[j].DataFormatString));
                                        else
                                            doc.SetCellValue(x, y, ((DateTime)value).ToString("M/dd"));
                                    }
                                }
                                else if (properties[j].PropertyType == typeof(Hyperlink))
                                {

                                    Hyperlink hyperLink = (Hyperlink)value;
                                    switch (hyperLink.Type)
                                    {
                                        case Hyperlink.HyperLinkType.External:
                                            doc.InsertHyperlink(x, y, SLHyperlinkTypeValues.Url, ((Hyperlink)value).Link, ((Hyperlink)value).Text, "");
                                            break;
                                        case Hyperlink.HyperLinkType.Internal:
                                            doc.InsertHyperlink(x, y, SLHyperlinkTypeValues.InternalDocumentLink, ((Hyperlink)value).Link, ((Hyperlink)value).Text, "");
                                            break;
                                        default:
                                            throw new NotSupportedException("This hyperlink type has not been implemented yet.");
                                    }
                                }
                                else
                                {
                                    string resultString = Convert.ToString(value);
                                    if (double.TryParse(resultString, out result))
                                    {
                                        if (formatAttributes[j] != null)
                                            doc.SetCellValueNumeric(x, y, result.ToString(formatAttributes[j].DataFormatString));
                                        else
                                            doc.SetCellValueNumeric(x, y, result.ToString("0.0"));
                                    }
                                    else
                                    {
                                        doc.SetCellValue(x, y, resultString);
                                    }
                                }
                            }
                            columnIndex++;
                        }
                    }
                    rowIndex++;
                }
            }
        }

        #region IDisposable
        [ExcludeFromCodeCoverageAttribute]   
        ~Spreadsheet() { Dispose(false); }

        // Dispose() calls Dispose(true)
        [ExcludeFromCodeCoverageAttribute]   
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        [ExcludeFromCodeCoverageAttribute]   
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // free managed resources
                if (doc != null)
                {
                    doc.Dispose();
                    doc = null;
                }
            }
        }
        #endregion
    }

    #region "Display Attributes"
    [ExcludeFromCodeCoverageAttribute]
    public class DisplayNoWrap : Attribute { }

    [ExcludeFromCodeCoverageAttribute]
    public class DisplayHide : Attribute { }

    [ExcludeFromCodeCoverageAttribute]    
    public class DisplayWidth : Attribute
    {
        public DisplayWidth(int width) { this.Width = width; }
        public int Width { get; set; }
    }
    #endregion

    public class Hyperlink
    {
        public string Text { get; set; }
        public string Link { get; set; }
        public HyperLinkType Type { get; set; }
        public enum HyperLinkType
        {
            External, Internal, Email, FilePath
        }
    }
}

