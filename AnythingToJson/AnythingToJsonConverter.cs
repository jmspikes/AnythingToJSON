using System.Data;
using System.IO;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Text.Json;
using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;
using JsonSerializer = System.Text.Json.JsonSerializer;

namespace AnythingToJson
{

    public class AnythingToJsonConverter
    {
        private bool _ignoreErrors = false;

        /// <summary>
        /// Initializes a new instance of <see cref="AnythingToJsonConverter"/> to be used to parse Excel files, CSV files, and DataTable objects. <br></br>If you want to turn off exception bubbling and simply return an empty JSON object on error, recreate constructor instance with the "ignoreErrors" flag as true. 
        /// </summary>
        public AnythingToJsonConverter()
        {

        }

        /// <summary>
        /// Initializes a new instance of <see cref="AnythingToJsonConverter"/> to be used to parse Excel files, CSV files, and DataTable objects. If an exception is encountered with "ignoreErrors == true" it will return an empty JSON object instead of bubbling exception to calling code. 
        /// </summary>
        public AnythingToJsonConverter(bool ignoreErrors)
        {
            _ignoreErrors = ignoreErrors;
        }

        /// <summary>
        /// <para>Use this to generate a JSON object from a symbol delimited based dataset's <see cref="MemoryStream"/>. Method expects a valid <see cref="MemoryStream"/> to be passed and will read its contents, transform input into JSON, and return that as a string. </para>
        /// </summary>
        /// <remarks><para>You can pass any memory stream to this method provided the data is delimited by some symbol. Method will infer what delimiter is used if none is provided based on the frequency of symbols used in the input data.</para>
        /// <para>It is recommended to provide a delimiter when known to avoid potential parsing mistakes.</para>
        /// </remarks>
        /// <param name="stream"><para>Memory stream to parse.</para>
        /// <example>
        /// Example of how to form expected stream:
        ///<code>
        ///byte[] data = File.ReadAllBytes(path);
        ///using var memoryStream = new MemoryStream(data);
        ///</code>
        ///</example></param>
        /// <param name="delimiter">(Optional) Delimiter for parsing logic to use to separate data into chunks. If none is provided code will attempt a best guess.</param>
        /// <param name="hasHeaderLine">(Optional) Use this flag when you have CSVs that are missing a header line. JSON keys will be generated programatically instead.</param>
        /// <returns>Formatted JSON string.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public IEnumerable<string> ConvertCsvFromMemoryStream(MemoryStream stream, string delimiter = null, bool hasHeaderLine = true) 
        {
            try
            {
                if (stream == null)
                {
                    throw new ArgumentNullException();
                }
                stream.Position = 0L;
                using var reader = new StreamReader((Stream)stream);
                using var csvReader = new TextFieldParser((TextReader)reader);
                TextFieldParser csvReaderToUse = csvReader;
                bool hasQuotes = false;
                var dataWithoutHeaders = new List<string>();
                if (delimiter == null)
                {
                    List<string> stringList = new List<string>();
                    reader.DiscardBufferedData();
                    reader.BaseStream.Position = 0L;
                    while (!reader.EndOfStream)
                    {
                        stringList.Add(reader.ReadLine());
                    }
                    // when reading a form file it will have header information, strip that out
                    if (stringList.FirstOrDefault() != null)
                    {
                        if (stringList.First().Contains("-----") && stringList.Last().Contains("-----"))
                        {
                            var skippedHeaders = false;
                            stringList.Remove(stringList.Last());
                            foreach (var item in stringList)
                            {
                                if (!string.IsNullOrWhiteSpace(item) && !skippedHeaders)
                                {
                                    continue;
                                }
                                skippedHeaders = true;
                                if (!string.IsNullOrWhiteSpace(item))
                                {
                                    dataWithoutHeaders.Add(item);
                                }

                            }
                            stringList = dataWithoutHeaders;
                        }
                    }
                    var str1 = stringList.Count >= 2 ? stringList[0] : throw new ArgumentException("No data found in provided file, please check provided file and retry.");
                    var str2 = stringList[1].Replace("\"", "");
                    Dictionary<char, int> source = new Dictionary<char, int>();
                    for (int index = 0; index < str2.Length - 1; ++index)
                    {
                        if (char.IsLetterOrDigit(str2[index]) && !char.IsLetterOrDigit(str2[index + 1]))
                        {
                            if (source.ContainsKey(str2[index + 1]))
                                source[str2[index + 1]]++;
                            else
                                source.Add(str2[index + 1], 1);
                        }
                    }
                    delimiter = source.OrderByDescending<KeyValuePair<char, int>, int>((Func<KeyValuePair<char, int>, int>)(x => x.Value)).Take<KeyValuePair<char, int>>(1).First<KeyValuePair<char, int>>().Key.ToString();
                    hasQuotes = str1.Contains("\"");
                }

                if (dataWithoutHeaders.Count > 0)
                {
                    var updatedReader = new StringReader(string.Join("\n", dataWithoutHeaders));
                    csvReaderToUse = new TextFieldParser(updatedReader);
                }
                var dataTable = ReadFileAsDataTable(csvReaderToUse, delimiter, hasQuotes, hasHeaderLine);
                var json = JsonSerializer.Serialize(ToDictionary(dataTable), new JsonSerializerOptions() { WriteIndented = true });
                return new List<string>() { json };
            }
            catch (ArgumentNullException ane)
            {
                if (_ignoreErrors)
                {
                    return new List<string>() { "{}" };
                }
                var message = $"[AnythingToJson] Null memory stream provided, please check calling code to ensure memory stream is valid and retry.";
                throw new ArgumentNullException(message);
            }
            catch (Exception e)
            {
                if (_ignoreErrors)
                {
                    return new List<string>() { "{}" };
                }
                var message = $"[AnythingToJson] Encountered general error while parsing, please address and retry. If problem is with the library please submit an issue at https://github.com/jmspikes/AnythingToJSON/issues so the root cause can be fixed, thanks!\nError:\n{e}";
                throw new Exception(message);
            }
            return new List<string>() { "{}" };
        }
        /// <summary>
        /// <para>Use this to generate a JSON object from a symbol delimited based dataset's <see cref="File"/> path. Method expects a valid path to be passed and will read its contents, transform input into JSON, and return that as a string. </para>
        /// </summary>
        /// <param name="path">Valid file location path to read data from.</param>
        /// <param name="delimiter">(Optional) Delimiter for parsing logic to use to separate data into chunks. If none is provided code will attempt a best guess.</param>
        /// <param name="hasHeaderLine">(Optional) Use this flag when you have CSVs that are missing a header line. JSON keys will be generated programatically instead.</param>
        /// <returns>Formatted JSON string.</returns>
        public IEnumerable<string> ConvertCsvFromFilePath(string path, string delimiter = null, bool hasHeaderLine = true)
        {
            try
            {
                // convert provided file to stream then process
                byte[] data = File.ReadAllBytes(path);
                using var memoryStream = new MemoryStream(data);
                var json = ConvertCsvFromMemoryStream(memoryStream, delimiter, hasHeaderLine);
                return json;
            }
            catch (Exception e)
            {
                if (_ignoreErrors)
                {
                    return new List<string>() { "{}" };
                }
                Console.WriteLine($"[AnythingToJson] Encountered error while parsing, returning empty json object. Error:\n{e}");
            }
            return new List<string>() { "{}" };
        }

        /// <summary>
        /// <para>Use this to generate a JSON object from a symbol delimited based dataset's <see cref="MemoryStream"/>. Method expects a valid <see cref="MemoryStream"/> to be passed and will read its contents, transform input into JSON, and return that as a string. </para>
        /// <para>NOTE: This method will be less reliable as it's a general method that accepts any memory stream and attempts to read/convert to JSON.</para>
        /// </summary>
        /// <param name="path">Valid file location path to read data from.</param>
        /// <param name="delimiter">(Optional) Delimiter for parsing logic to use to separate data into chunks. If none is provided code will attempt a best guess.</param>
        /// <param name="hasHeaderLine">Use this flag to indicate whether there is a header line or not. When there is no header the JSON key will be a number.</param>
        /// <returns>Formatted JSON string.</returns>
        public IEnumerable<string> ConvertFromMemoryStream(MemoryStream stream, string delimiter = null, bool hasHeaderLine = true)
        {
            try
            {
                var json = ConvertCsvFromMemoryStream(stream, delimiter, hasHeaderLine);
                return json;
            }
            catch (Exception e)
            {
                if (_ignoreErrors)
                {
                    return new List<string>() { "{}" };
                }
                Console.WriteLine($"[AnythingToJson] Encountered error while parsing, returning empty json object. Error:\n{e}");
            }
            return new List<string>() { "{}" };
        }

        /// <summary>
        /// <para>Use this to generate a JSON object from an Excel file's <see cref="MemoryStream"/>. Method expects a valid <see cref="MemoryStream"/> to be passed and will read its contents, transform input into JSON, and return that as a string.</para>
        /// </summary>
        /// <remarks><para><b>Note:</b> Because this library leverages <seealso href="https://docs.closedxml.io/en/latest/">ClosedXML</seealso> to read Excel files, only Excel 2007+ (.xlsx, .xlsm) formats will be supported.</para>
        /// </remarks>
        /// <param name="stream"><para>Memory stream to parse data from.</para>
        /// <example>
        /// Example of how to form expected stream:
        ///<code>
        ///byte[] data = File.ReadAllBytes(path);
        ///using var memoryStream = new MemoryStream(data);
        ///</code>
        ///</example></param>
        /// <returns>Formatted JSON string.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public IEnumerable<string> ConvertExcelFromMemoryStream(MemoryStream stream)
        {
            try
            {
                if (stream == null)
                {
                    throw new ArgumentNullException();
                }
                var parsedJson = new List<string>();
                using var workbook = new XLWorkbook(stream);
                // iterate all worksheets to get data to parse
                foreach (var worksheet in workbook.Worksheets)
                {
                    var csvData = new StringBuilder();
                    // Iterate over the rows and columns
                    foreach (var row in worksheet.RangeUsed().Rows())
                    {
                        bool firstColumn = true;
                        foreach (var cell in row.Cells())
                        {
                            if (!firstColumn)
                            {
                                csvData.Append(",");
                            }
                            else
                            {
                                firstColumn = false;
                            }
                            var cellValue = cell.Value.ToString();

                            // Escape if necessary (ie if the value contains a comma)
                            if (cellValue.Contains(",") || cellValue.Contains("\""))
                            {
                                cellValue = $"\"{cellValue.Replace("\"", "\"\"")}\"";
                            }
                            csvData.Append(cellValue);
                        }
                        csvData.AppendLine();
                    }
                    // convert data to stream to send to parser
                    var dataAsBytes = Encoding.UTF8.GetBytes(csvData.ToString());
                    using var memoryStream = new MemoryStream(dataAsBytes);
                    var json = ConvertCsvFromMemoryStream(memoryStream);
                    parsedJson.Add(json.First());
                }
                return parsedJson;
            }
            catch (ArgumentNullException ane)
            {
                if (_ignoreErrors)
                {
                    return new List<string>() { "{}" };
                }
                var message = $"[AnythingToJson] Null memory stream provided, please check calling code to ensure memory stream is valid and retry.";
                throw new ArgumentNullException(message);
            }
            catch (Exception e)
            {
                if (_ignoreErrors)
                {
                    return new List<string>() { "{}" };
                }
                var message = $"[AnythingToJson] Encountered general error while parsing, please address and retry. If problem is with the library please submit an issue at https://github.com/jmspikes/AnythingToJSON/issues so the root cause can be fixed, thanks!\nError:\n{e}";
                throw new Exception(message);
            }
            return new List<string>() { "{}" };
        }

        /// <summary>
        /// <para>Use this to generate a JSON object from an Excel <see cref="File"/> given provided path. Method expects a valid path to be passed and will read its contents, transform input into JSON, and return that as a string.</para>
        /// </summary>
        /// <remarks><para><b>Note:</b> Because this library leverages <seealso href="https://docs.closedxml.io/en/latest/">ClosedXML</seealso> to read Excel files, only Excel 2007+ (.xlsx, .xlsm) formats will be supported.</para>
        /// </remarks>
        /// <param name="path">Valid file location path to read data from.</param>
        /// <returns>Formatted JSON string.</returns>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        public IEnumerable<string> ConvertExcelFromFilePath(string path)
        {
            try
            {
                // convert provided file to stream then process
                byte[] data = File.ReadAllBytes(path);
                using var memoryStream = new MemoryStream(data);
                var json = ConvertExcelFromMemoryStream(memoryStream);
                return json;
            }
            catch (DirectoryNotFoundException dnfe)
            {
                if (_ignoreErrors)
                {
                    return new List<string>() { "{}" };
                }
                var message = $"[AnythingToJson] Could not find the provided directory, is the file at the location provided?\nPath given: \n{path}\n{dnfe}";
                throw new DirectoryNotFoundException(message);
            }
            catch (FileNotFoundException fnfe)
            {
                if (_ignoreErrors)
                {
                    return new List<string>() { "{}" };
                }
                var message = $"[AnythingToJson] Could not find the provided directory, is the file at the location provided?\nPath given: \n{path}\n{fnfe}";
                throw new FileNotFoundException(message);
            }
            catch (Exception e)
            {
                if (_ignoreErrors)
                {
                    return new List<string>() { "{}" };
                }
                var message = $"[AnythingToJson] Encountered general error while parsing, please address and retry. If problem is with the library please submit an issue at https://github.com/jmspikes/AnythingToJSON/issues so the root cause can be fixed, thanks!\nError:\n{e}";
                throw new Exception(message);
            }
            return new List<string>() { "{}" };
        }
        /// <summary>
        /// <para>Use this to generate a JSON object from a given <see cref="DataTable"/>. Method expects a valid <see cref="DataTable"/> to be passed and will read its contents, transform input into JSON, and return that as a string.</para>
        /// </summary>
        /// <param name="dataTable"><para><see cref="DataTable"/> to read from.</para></param>
        /// <returns>Formatted JSON string.</returns>
        public IEnumerable<string> ConvertDataTable(DataTable dataTable)
        {
            try
            {
                var json = JsonSerializer.Serialize(ToDictionary(dataTable), new JsonSerializerOptions() { WriteIndented = true });
                return new List<string>() { json };
            }
            catch (Exception e)
            {
                if (_ignoreErrors)
                {
                    return new List<string>() { "{}" };
                }
                var message = $"[AnythingToJson] Encountered general error while parsing, please address and retry. If problem is with the library please submit an issue at https://github.com/jmspikes/AnythingToJSON/issues so the root cause can be fixed, thanks!\nError:\n{e}";
                throw new Exception(message);
            }
            return new List<string>() { "{}" };
        }

        /// <summary>
        /// <para>This method is more or less just a wrapper for calling: <code>JsonSerializer.Serialize(Your Object)</code></para><para>Would recommend just using <see cref="JsonSerializer"/> directly but this method provided for "completeness" sake for the <i>Anything</i> part of the name.</para>
        /// </summary>
        /// <typeparam name="T">Generic class type to use.</typeparam>
        /// <param name="item">Object of given <see cref="Type"/> to serialize.</param>
        /// <returns></returns>
        public static IEnumerable<string> ConvertThis<T>(T item)
        {
            try
            {
                return new List<string>() { JsonSerializer.Serialize<T>(item, new JsonSerializerOptions() { WriteIndented = true }) };
            }
            catch (Exception e)
            {
                var message = $"[AnythingToJson] Encountered general error while parsing, please address and retry. If problem is with the library please submit an issue at https://github.com/jmspikes/AnythingToJSON/issues so the root cause can be fixed, thanks!\nError:\n{e}";
                throw new Exception(message);
            }
            return new List<string>() { "{}" };
        }

        private DataTable ReadFileAsDataTable(TextFieldParser csvReader, string parsedDelimiter, bool hasQuotes, bool hasHeaderLine)
        {
            DataTable csvData = new();
            csvReader.SetDelimiters(new string[] { parsedDelimiter });
            csvReader.HasFieldsEnclosedInQuotes = hasQuotes;
            bool tableCreated = false;
            while (!tableCreated)
            {
                var colFields = csvReader.ReadFields();
                string[] preserveLine = colFields.ToArray();
                if (!hasHeaderLine)
                {
                    for (int i = 0; i < colFields.Count(); i++)
                    {
                        colFields[i] = i.ToString();
                    }
                }
                foreach (string column in colFields)
                {
                    DataColumn dataColumn = new(column)
                    {
                        AllowDBNull = true
                    };
                    csvData.Columns.Add(dataColumn);
                }
                // when no header is provided we'll need to add back the already read line
                if (!hasHeaderLine)
                {
                    csvData.Rows.Add(preserveLine);
                }
                tableCreated = true;
            }
            while (!csvReader.EndOfData)
            {
                csvData.Rows.Add(csvReader.ReadFields());
            }
            return csvData;
        }

        private IEnumerable<Dictionary<string, object>> ToDictionary(DataTable table)
        {
            string[] columns = table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();
            IEnumerable<Dictionary<string, object>> result = table.Rows.Cast<DataRow>().Select(dr => columns.ToDictionary(c => c, c => dr[c]));
            return result;
        }
    }
}