# AnythingToJSON
AnythingToJSON is a lightweight C# library useful for converting symbol delimited data to JSON. Using just a memory stream or a provided file path, AnythingToJSON supports converting the following: 
- Excel Files (.xlsx, .xlsm)
- CSVs (of any variety)
- [C# DataTables](https://learn.microsoft.com/en-us/dotnet/api/system.data.datatable?view=net-8.0)

# How to Use
Install via NuGet:  
`$ dotnet add package AnythingToJson`  
To use, simply create a new instance and pick the applicable method. When no delimiter is provided for CSV conversions, code will attempt to heuristically determine what delimiter is.  
Available methods and constructors are:  
```
AnythingToJsonConverter() <--- default constructor  
AnythingToJsonConverter(ignore errors) <--- bool flag to return empty JSON object "{}" if an error occurs  
ConvertCsvFromFilePath(path, delimiter, has header line)  
ConvertCsvFromMemoryStream(memory stream, delimiter, has header line)  
ConvertExcelFromPath(path)  
ConvertExcelFromMemoryStream(memory stream)  
ConvertFromMemoryStream(memory stream, delimiter, has header line)  
ConvertDataTable(datatable)  
ConvertThis(object)  
```
# Code Samples
Below are a few examples of how to use the library.  
Use the __File__ methods when converting a file locally.  
Use the __Memory Stream__ methods if you want to convert `multipart/form-data` within an endpoint or the like.  
Use the __DataTable__ method if you want to convert a DataTable object to JSON.  
### From file:
~~~
CSV:
    var converter = new AnythingToJsonConverter();
    var csvFromPath = converter.ConvertCsvFromFilePath(@"CSV/semi-colon-example.csv");
Excel:
    var converter = new AnythingToJsonConverter();
    var excelFromPath = converter.ConvertExcelFromFilePath(@"Excel/excel-example.xlsx");
~~~
### Memory Stream:
~~~
CSV:
    byte[] data = File.ReadAllBytes(@"CSV/semi-colon-example.csv");
    using var memoryStream = new MemoryStream(data);
    var csvFromMemoryStream = converter.ConvertCsvFromMemoryStream(memoryStream);
Excel:
    byte[] data = File.ReadAllBytes(@"Excel/excel-example.xlsx");
    using var memoryStream = new MemoryStream(data);
    var excelFromMemoryStream = converter.ConvertExcelFromMemoryStream(memoryStream);
~~~
### DataTable:
~~~
    AnythingToJsonConverter converter = new AnythingToJsonConverter();
    DataTable table = new DataTable("Table Example");
    table.Columns.Add("ID", typeof(int));
    table.Columns.Add("Name", typeof(string));
    table.Columns.Add("Age", typeof(int));
    table.Rows.Add(1, "John Doe", 30);
    table.Rows.Add(2, "Jane Smith", 25);
    table.Rows.Add(3, "Timothy James", 40);
    var dataTableJson = converter.ConvertDataTable(table);
~~~  
Example response from library:
```json
[
  {
    "ID": 1,
    "Name": "John Doe",
    "Age": 30
  },
  {
    "ID": 2,
    "Name": "Jane Smith",
    "Age": 25
  },
  {
    "ID": 3,
    "Name": "Timothy James",
    "Age": 40
  }
]
```  
If you encounter any issues or have suggestions, please feel free to post them at the [AnythingToJSON's GitHub](https://github.com/jmspikes/AnythingToJSON/issues)  
Happy coding and thank you for using this library!  
