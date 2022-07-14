// See https://aka.ms/new-console-template for more information
using ExcelDataReader;
using Newtonsoft.Json;

Console.WriteLine("Start App");

try
{
    var filePath = @"path-to\\CodeBook.xlsx";

    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

    using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
    {
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            // 2. Use the AsDataSet extension method
            var result = reader.AsDataSet();
            reader.Close();

            // The result of each spreadsheet is in result.Tables
            if (result != null && result.Tables.Count > 0)
            {
                var lgaRecords = result.Tables[6];
                for (int i = 0; i < lgaRecords.Rows.Count; i++)
                {
                    if (i == 0) continue;

                    var row = lgaRecords.Rows[i];

                    var lgaCode = row[0];
                    var lga = row[1];
                    var stateCode = row[2];

                    Console.WriteLine("lgaCode: {0}", Convert.ToString(lgaCode));
                    Console.WriteLine("lgaName: {0}", Convert.ToString(lga));
                    Console.WriteLine("stateCode: {0}", Convert.ToString(stateCode));

                }
            }
        }
    }
}
catch (Exception ex)
{
    Console.WriteLine("{0} Last exception caught.", ex.Message);
}
