# ExcelDataReaderWriter
A simple C# library example for reading and writing Microsoft Excel .xlsx, .xls, .csv, and .xlsb files.

**ExcelDataReaderWriter** is a simple and lightweight C# library for reading and writing `.xlsx`, `.xls`, `.csv`, and `.xlsb` files using popular packages like EPPlus and NPOI.

## 📦 Required NuGet Packages

Before using this library, install the following NuGet packages:

```bash
Install-Package EPPlus
Install-Package NPOI
Install-Package System.Data.OleDb
```

## 📖 Usage Examples

### ✅ Reading Files

```csharp
// Read an XLSX file
var data = ExcelHelper.ReadExcel("file.xlsx");

// Read a specific sheet from an XLSX file
var sheetData = ExcelHelper.ReadExcel("file.xlsx", "Sheet1");

// Read an XLS file
var xlsData = ExcelHelper.ReadExcel("file.xls");

// Read a CSV file
var csvData = ExcelHelper.ReadExcel("file.csv");

// Read an XLSB file
var xlsbData = ExcelHelper.ReadExcel("file.xlsb");
```

### ✍️ Writing Excel Files

```csharp
// Create a new DataTable
var dt = new DataTable();
dt.Columns.Add("First Name");
dt.Columns.Add("Last Name");
dt.Columns.Add("Age");

// Add sample data
dt.Rows.Add("John", "Doe", 30);
dt.Rows.Add("Jack", "Doe", 25);

// Save as XLSX
ExcelHelper.WriteExcel(dt, "output.xlsx", "Personnel");

// Save as XLS
ExcelHelper.WriteExcel(dt, "output.xls", "Personnel");

// Save as CSV
ExcelHelper.WriteExcel(dt, "output.csv");

// Save as XLSB
ExcelHelper.WriteExcel(dt, "output.xlsb", "Personnel");
```

---

Please feel free to fork the repository and submit pull requests to the develop branch, and don't hesitate to customize or expand the helper functions based on your specific needs.
