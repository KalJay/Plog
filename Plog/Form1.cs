using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Plog
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Global.f1 = this;
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.ExcelFile == "")
            {
                DialogResult = MessageBox.Show("No Spreadsheet document selected! Click 'Select' and then 'Open' to load a Spreadsheet.", "Error: No Spreadsheet selected!", MessageBoxButtons.OK);
            } else
            {
                DateTime currentTime = DateTime.Now;
                string area = comboBox1.Text;
                string name = textBox1.Text;
                try
                {
                    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(Properties.Settings.Default.ExcelFile, true))
                    {
                        IEnumerable<Sheet> sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == "Data");
                        WorksheetPart wsPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(sheets?.First().Id.Value);

                        SharedStringTablePart sharedStrings = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                        Worksheet worksheet = wsPart.Worksheet;



                        InsertData(area, name, currentTime, wsPart, sharedStrings);
                        spreadSheet.Save();
                        spreadSheet.Close();
                    }
                    textBox1.Text = "";
                } catch (Exception ex)
                {
                    MessageBox.Show("Error attempting to save to the spreadsheet! Exception: " + ex.Message);
                }
            }

        }

        private void InsertData(string Area, string Name, DateTime dateTime, WorksheetPart wsPart, SharedStringTablePart shareStringPart)
        {
            int lastRowIndex = int.Parse(wsPart.Worksheet.Descendants<Row>().LastOrDefault().RowIndex);

            SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            Row row;
            row = new Row() { RowIndex = Convert.ToUInt32(lastRowIndex + 1) };
            sheetData.Append(row);


            Cell timeCell = new Cell() { CellReference = "A" + (lastRowIndex + 1), CellValue = new CellValue(dateTime.ToOADate().ToString()), DataType = new EnumValue<CellValues>(CellValues.Number) };
            timeCell.StyleIndex = new UInt32Value { Value = (UInt32)4 };
            row.InsertBefore(timeCell, null);

            int index = InsertSharedStringItem(Name, shareStringPart);
            Cell nameCell = new Cell() { CellReference = "B" + (lastRowIndex + 1), CellValue = new CellValue(index.ToString()), DataType = new EnumValue<CellValues>(CellValues.SharedString) };
            row.InsertAfter(nameCell, timeCell);

            index = InsertSharedStringItem(Area, shareStringPart);
            Cell areaCell = new Cell() { CellReference = "C" + (lastRowIndex + 1), CellValue = new CellValue(index.ToString()), DataType = new EnumValue<CellValues>(CellValues.SharedString) };
            row.InsertAfter(areaCell, nameCell);
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
            SuspendLayout();
        }
    }
    public static class Global
    {
        public static Form2 f2;
        public static Form1 f1;
    }
}
