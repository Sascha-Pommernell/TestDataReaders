namespace TestDataReaders.ExcelReader
{
    public class ExcelDataReader
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out IntPtr ProcessId);
        public static IEnumerable<TestCaseData> ReadExcel(String fileName, string sheetName)
        {
            Excel.Application excelApp = new Excel.Application();
            IntPtr hwnd = new IntPtr(excelApp.Hwnd);
            IntPtr processId;
            List<string> iList = new List<string>();
            var testCase = new List<TestCaseData>();

            try
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(fileName);
                Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[sheetName];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int columnCount = excelRange.Columns.Count;
                object[,] valueArray = (object[,])excelRange.Value[Excel.XlRangeValueDataType.xlRangeValueDefault];
                ArrayList testData = new ArrayList();

                for (int i = 2; i <= rowCount; i++)
                {
                    String[] testDataArry = new string[columnCount];

                    for (int k = 1; k <= columnCount; k++)
                    {
                        testDataArry[k - 1] = valueArray[i, k].ToString();
                    }
                    testData.Add(new TestCaseData(testDataArry));
                }

                excelWorkbook.Close();
                excelApp.Quit();

                if (excelWorksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorksheet);
                }

                if (excelRange != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelRange);
                }

                if (excelWorkbook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                }

                IntPtr program = GetWindowThreadProcessId(hwnd, out processId);
                Process process = Process.GetProcessById(processId.ToInt32());
                process.Kill();
                GC.Collect();
            }
            catch(Exception)
            {
                IntPtr program = GetWindowThreadProcessId(hwnd, out processId);
                Process process = Process.GetProcessById((processId.ToInt32()));
                process.Kill();
                GC.Collect();
                iList.Add("Invalid Datafile or sheetName. Please Check!");
            }

            if(testCase != null)
                foreach(TestCaseData testCaseData in testCase)
                {
                    yield return testCaseData;
                }
        }
    }
}
