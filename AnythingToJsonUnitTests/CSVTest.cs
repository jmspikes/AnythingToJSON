using System.Data;
using System.Text.Json;
using AnythingToJson;
using Xunit;
using JsonException = Newtonsoft.Json.JsonException;

namespace CSVToJsonUnitTests
{
    public class CSVTest
    {
        [Fact]
        public void CSVToJsonIntegrationTest()
        {
            AnythingToJsonConverter converter = new AnythingToJsonConverter();
            
            var noQuotesTest = converter.ConvertCsvFromFilePath(@"CSV/noquotedata.csv").FirstOrDefault();
            var basicInputWithSymbols = converter.ConvertCsvFromFilePath(@"CSV/zillow.csv").FirstOrDefault();
            var semiColonInput = converter.ConvertCsvFromFilePath(@"CSV/semi-colon-test.csv").FirstOrDefault();
            var spaceInput = converter.ConvertCsvFromFilePath(@"CSV/space-test.csv").FirstOrDefault();
            var basicInput = converter.ConvertCsvFromFilePath(@"CSV/airtravel.csv", hasHeaderLine:false).FirstOrDefault();
            Assert.True(ValidJson(basicInputWithSymbols));
            Assert.True(ValidJson(semiColonInput));
            Assert.True(ValidJson(spaceInput));
            Assert.True(ValidJson(basicInput));
        }

        [Fact]
        public void ExcelToJsonIntegrationTest()
        {
            AnythingToJsonConverter converter = new AnythingToJsonConverter();
            var financialDataSet = converter.ConvertExcelFromFilePath(@"Excel/Financial Sample.xlsx").FirstOrDefault();
            var plainFile = converter.ConvertExcelFromFilePath(@"Excel/Plain File.xlsx").FirstOrDefault();
            var plainFileWithFormatting = converter.ConvertExcelFromFilePath(@"Excel/Plain File With Formatting.xlsx").FirstOrDefault();
            Assert.True(ValidJson(financialDataSet));
            Assert.True(ValidJson(plainFile));
            Assert.True(ValidJson(plainFileWithFormatting));
        }

        [Fact]
        public void DataTableToJsonIntegrationTest()
        {
            AnythingToJsonConverter converter = new AnythingToJsonConverter();
            DataTable table = new DataTable("Test Table");
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Age", typeof(int));
            table.Rows.Add(1, "John Doe", 30);
            table.Rows.Add(2, "Jane Smith", 25);
            table.Rows.Add(3, "Timothy James", 40);

            var basicDataTable = converter.ConvertDataTable(table).FirstOrDefault();
            Assert.True(ValidJson(basicDataTable));
        }

        private bool ValidJson(string json)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(json) || json.Equals("{}"))
                {
                    return false;
                }

                using var document = JsonDocument.Parse(json);
                return true;
            }
            catch (JsonException e)
            {
                return false;
            }
        }
    }
}