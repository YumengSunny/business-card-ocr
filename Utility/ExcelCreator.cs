using System;
using System.Data;
using System.Text;
using System.Xml;

namespace BCardReader.Utility
{
    class ExcelCreator
    {
        /// <summary>       
        /// Create one Excel-XML-Document with SpreadsheetML from a DataTable
        /// </summary>        
        /// <param name="dataSource">Datasource which would be exported in Excel</param>
        /// <param name="fileName">Name of exported file</param>
        public static void Create(DataTable dtSource, string strFileName)
        {
            // Create XMLWriter
            XmlTextWriter xtwWriter = new XmlTextWriter(strFileName, Encoding.UTF8);

            //Format the output file for reading easier
            xtwWriter.Formatting = Formatting.Indented;

            // <?xml version="1.0"?>
            xtwWriter.WriteStartDocument();

            // <?mso-application progid="Excel.Sheet"?>
            xtwWriter.WriteProcessingInstruction("mso-application", "progid=\"Excel.Sheet\"");

            // <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet >"
            xtwWriter.WriteStartElement("Workbook", "urn:schemas-microsoft-com:office:spreadsheet");

            //Write definition of namespace
            xtwWriter.WriteAttributeString("xmlns", "o", null, "urn:schemas-microsoft-com:office:office");
            xtwWriter.WriteAttributeString("xmlns", "x", null, "urn:schemas-microsoft-com:office:excel");
            xtwWriter.WriteAttributeString("xmlns", "ss", null, "urn:schemas-microsoft-com:office:spreadsheet");
            xtwWriter.WriteAttributeString("xmlns", "html", null, "http://www.w3.org/TR/REC-html40");

            // <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
            xtwWriter.WriteStartElement("DocumentProperties", "urn:schemas-microsoft-com:office:office");

            // Write document properties
            xtwWriter.WriteElementString("Author", Environment.UserName);
            xtwWriter.WriteElementString("LastAuthor", Environment.UserName);
            xtwWriter.WriteElementString("Created", DateTime.Now.ToString("u") + "Z");
            xtwWriter.WriteElementString("Company", "Unknown");
            xtwWriter.WriteElementString("Version", "11.8122");

            // </DocumentProperties>
            xtwWriter.WriteEndElement();

            // <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
            xtwWriter.WriteStartElement("ExcelWorkbook", "urn:schemas-microsoft-com:office:excel");

            // Write settings of workbook
            xtwWriter.WriteElementString("WindowHeight", "13170");
            xtwWriter.WriteElementString("WindowWidth", "17580");
            xtwWriter.WriteElementString("WindowTopX", "120");
            xtwWriter.WriteElementString("WindowTopY", "60");
            xtwWriter.WriteElementString("ProtectStructure", "False");
            xtwWriter.WriteElementString("ProtectWindows", "False");

            // </ExcelWorkbook>
            xtwWriter.WriteEndElement();

            // <Styles>
            xtwWriter.WriteStartElement("Styles");

            // <Style ss:ID="Default" ss:Name="Normal">
            xtwWriter.WriteStartElement("Style");
            xtwWriter.WriteAttributeString("ss", "ID", null, "Default");
            xtwWriter.WriteAttributeString("ss", "Name", null, "Normal");

            // <Alignment ss:Vertical="Bottom"/>
            xtwWriter.WriteStartElement("Alignment");
            xtwWriter.WriteAttributeString("ss", "Vertical", null, "Bottom");
            xtwWriter.WriteEndElement();

            // Write null on the other properties
            xtwWriter.WriteElementString("Borders", null);
            xtwWriter.WriteElementString("Font", null);
            xtwWriter.WriteElementString("Interior", null);
            xtwWriter.WriteElementString("NumberFormat", null);
            xtwWriter.WriteElementString("Protection", null);

            // </Style>
            xtwWriter.WriteEndElement();

            // </Styles>
            xtwWriter.WriteEndElement();

            // <Worksheet ss:Name="xxx">
            xtwWriter.WriteStartElement("Worksheet");
            xtwWriter.WriteAttributeString("ss", "Name", null, dtSource.TableName);

            // <Table ss:ExpandedColumnCount="2" ss:ExpandedRowCount="3" x:FullColumns="1" x:FullRows="1" ss:DefaultColumnWidth="60">
            xtwWriter.WriteStartElement("Table");
            xtwWriter.WriteAttributeString("ss", "ExpandedColumnCount", null, dtSource.Columns.Count.ToString());
            xtwWriter.WriteAttributeString("ss", "ExpandedRowCount", null, dtSource.Rows.Count.ToString());
            xtwWriter.WriteAttributeString("x", "FullColumns", null, "1");
            xtwWriter.WriteAttributeString("x", "FullRows", null, "1");
            xtwWriter.WriteAttributeString("ss", "DefaultColumnWidth", null, "60");

            // Run through all rows of data source
            foreach (DataRow row in dtSource.Rows)
            {
                // <Row>
                xtwWriter.WriteStartElement("Row");

                // Run through all cell of current rows
                foreach (object cellValue in row.ItemArray)
                {
                    // <Cell>
                    xtwWriter.WriteStartElement("Cell");

                    // <Data ss:Type="String">xxx</Data>
                    xtwWriter.WriteStartElement("Data");
                    xtwWriter.WriteAttributeString("ss", "Type", null, "String");

                    // Write content of cell
                    xtwWriter.WriteValue(cellValue);

                    // </Data>
                    xtwWriter.WriteEndElement();

                    // </Cell>
                    xtwWriter.WriteEndElement();
                }
                // </Row>
                xtwWriter.WriteEndElement();
            }
            // </Table>
            xtwWriter.WriteEndElement();

            // <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
            xtwWriter.WriteStartElement("WorksheetOptions", "urn:schemas-microsoft-com:office:excel");

            // Write settings of page
            xtwWriter.WriteStartElement("PageSetup");
            xtwWriter.WriteStartElement("Header");
            xtwWriter.WriteAttributeString("x", "Margin", null, "0.4921259845");
            xtwWriter.WriteEndElement();
            xtwWriter.WriteStartElement("Footer");
            xtwWriter.WriteAttributeString("x", "Margin", null, "0.4921259845");
            xtwWriter.WriteEndElement();
            xtwWriter.WriteStartElement("PageMargins");
            xtwWriter.WriteAttributeString("x", "Bottom", null, "0.984251969");
            xtwWriter.WriteAttributeString("x", "Left", null, "0.78740157499999996");
            xtwWriter.WriteAttributeString("x", "Right", null, "0.78740157499999996");
            xtwWriter.WriteAttributeString("x", "Top", null, "0.984251969");
            xtwWriter.WriteEndElement();
            xtwWriter.WriteEndElement();

            // <Selected/>
            xtwWriter.WriteElementString("Selected", null);

            // <Panes>
            xtwWriter.WriteStartElement("Panes");

            // <Pane>
            xtwWriter.WriteStartElement("Pane");

            // Write settings of active field
            xtwWriter.WriteElementString("Number", "1");
            xtwWriter.WriteElementString("ActiveRow", "1");
            xtwWriter.WriteElementString("ActiveCol", "1");

            // </Pane>
            xtwWriter.WriteEndElement();

            // </Panes>
            xtwWriter.WriteEndElement();

            // <ProtectObjects>False</ProtectObjects>
            xtwWriter.WriteElementString("ProtectObjects", "False");

            // <ProtectScenarios>False</ProtectScenarios>
            xtwWriter.WriteElementString("ProtectScenarios", "False");

            // </WorksheetOptions>
            xtwWriter.WriteEndElement();

            // </Worksheet>
            xtwWriter.WriteEndElement();

            // </Workbook>
            xtwWriter.WriteEndElement();

            // Write file on hard disk
            xtwWriter.Flush();
            xtwWriter.Close();
        }
    }
}
