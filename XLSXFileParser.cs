using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    class XLSXFileParser : Parsers
    {
        //string filePath;
        public string gettext(string filePath){
            StringBuilder b = new StringBuilder() ;
            b.Append("");

            using (SpreadsheetDocument d = SpreadsheetDocument.Open(filePath, false))
            {
                // Load the shared strings table.
                SharedStringTablePart stringTable =
                 d.WorkbookPart.GetPartsOfType<SharedStringTablePart>()
                 .FirstOrDefault();
                if (stringTable == null) System.Diagnostics.Debug.WriteLine("Null string table");
                foreach (WorksheetPart part in d.WorkbookPart.WorksheetParts)
                {
                    foreach (SheetData sheet in part.Worksheet.Elements<SheetData>())
                    {
                        bool added = false;
                        foreach (Row r in sheet.Elements<Row>())
                        {
                            foreach (Cell c in r.Elements<Cell>())
                            {
                                if (c.DataType != null)
                                {
                                    string v = c.CellValue.Text;
                                    if (v != null && c.DataType.Value == CellValues.SharedString)
                                    {
                                        var tableEntry = stringTable.SharedStringTable.ElementAt(int.Parse(v));
                                        if (tableEntry != null)
                                        {
                                            v = tableEntry.InnerText;
                                        }
                                    }
                                    if (v != null)
                                    {
                                        if (added) b.Append('\t');
                                        b.Append(v);
                                        added = true;
                                    }
                                }
                            }
                            if (added) b.AppendLine();
                        }
                    }
                }
            }
            return b.ToString();
        }
    }
}
