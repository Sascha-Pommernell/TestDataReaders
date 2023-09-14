namespace TestDataReaders.CSVReader
{
    public class CSVDataReader
    {
        public static IEnumerable<TestCaseData> ReadCSV(string filename)
        {
            string filePath = filename;
            List<string> iList = new List<string>();
            var testCase = new List<TestCaseData>();
            StreamReader streamReader = null!;

            try
            {
                if(File.Exists(filePath))
                {
                    streamReader = new StreamReader(File.OpenRead(filePath));
                    List<string> listA = new List<string>();

                    while(!streamReader.EndOfStream)
                    {
                        var line = streamReader.ReadLine();
                        var values = line.Split(';');
                        testCase.Add(new TestCaseData(values));
                    }
                }
                else
                {
                    iList.Add("Invalid Datafile! Please Ceck!");
                }
            }
            catch (Exception ex)
            {
                iList.Add(ex.ToString());
            }

            if(testCase != null)
            {
                foreach(TestCaseData testCaseData in testCase)
                {
                    yield return testCaseData;
                }
            }
        }
    }
}
