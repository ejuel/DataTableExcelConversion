using ClosedXML.Excel;
using System.Data;

namespace ClosedXmlTools
{
    public class XmlTools
    {
        private readonly string workingDir;

        public XmlTools(string workingDir)
        {
            Console.WriteLine("test");
            this.workingDir = workingDir;
        }
        public bool LoadXml(string fileName, out DataTable? dataTable, string worksheetName = "")
        {
            try
            {
                XLWorkbook test = new XLWorkbook($"{workingDir}\\{fileName}");
                IXLWorksheet worksheet;
                if (string.IsNullOrEmpty(worksheetName))
                {
                    //No name provided, default to first worksheet
                    worksheet = test.Worksheet(1);
                }
                else if (!test.TryGetWorksheet(worksheetName, out worksheet))
                {
                    Console.WriteLine($"No worksheet named {worksheetName} found in {workingDir}\\{fileName}");
                    dataTable = null;
                    return false;
                }
                dataTable = ConvertWorksheet(worksheet, "asd");
                Console.WriteLine("test");

                if (dataTable.Columns.Count == 0 && dataTable.Rows.Count == 0)
                {
                    dataTable = null;
                    return false; //No data
                }
                else return true;
            }
            catch (IOException ex)
            {
                if (ex.ToString().Contains("it is being used by another process"))
                {
                    Console.WriteLine("failed to load file. It is being used elsewhere...");
                }
                dataTable = null;
                return false;
            }

        }

        public DataTable ConvertWorksheet(IXLWorksheet worksheet, string FirstRowValue = "")
        {
            //Create a new DataTable.
            DataTable dt = new DataTable();

            //Loop through the Worksheet rows.
            bool firstRow = false;
            bool firstRowWasFound = false;
            foreach (IXLRow row in worksheet.Rows())
            {
                if (!firstRowWasFound
                    && (string.Equals(row.Cell(1).Value, FirstRowValue) || string.IsNullOrWhiteSpace(FirstRowValue)))
                {
                    firstRowWasFound = true;
                    firstRow = true;
                }

                //Use the first row to add columns to DataTable.
                if (firstRow)
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        dt.Columns.Add(cell.Value.ToString());
                    }
                    firstRow = false;
                }
                else if (firstRowWasFound && row.FirstCellUsed() is not null)
                {
                    dt.Rows.Add();
                    int i = 0;
                    foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                    {
                        dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                        i++;
                    }
                }
            }
            return dt;
        }
    }
}