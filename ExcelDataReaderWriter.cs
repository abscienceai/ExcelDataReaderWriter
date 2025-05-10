using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeOpenXml;

namespace LightweightExcelLibrary
{
    public class ExcelHelper
    {
        public static DataTable ReadExcel(string filePath, string sheetName = null)
        {
            var extension = Path.GetExtension(filePath).ToLower();

            return extension switch
            {
                ".xlsx" => ReadXlsx(filePath, sheetName),
                ".xls" => ReadXls(filePath, sheetName),
                ".csv" => ReadCsv(filePath),
                ".xlsb" => ReadXlsb(filePath, sheetName),
                _ => throw new NotSupportedException("Unsupported file format")
            };
        }

        public static void WriteExcel(DataTable data, string filePath, string sheetName = "Sheet1")
        {
            var extension = Path.GetExtension(filePath).ToLower();

            switch (extension)
            {
                case ".xlsx":
                    WriteXlsx(data, filePath, sheetName);
                    break;
                case ".xls":
                    WriteXls(data, filePath, sheetName);
                    break;
                case ".csv":
                    WriteCsv(data, filePath);
                    break;
                case ".xlsb":
                    WriteXlsb(data, filePath, sheetName);
                    break;
                default:
                    throw new NotSupportedException("Unsupported file format");
            }
        }

        #region Private Read Methods

        private static DataTable ReadXlsx(string filePath, string sheetName)
        {
            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = sheetName == null
                ? package.Workbook.Worksheets[0]
                : package.Workbook.Worksheets[sheetName];

            if (worksheet == null)
                throw new Exception("Page not found");

            var dt = new DataTable();

            // Add titles
            foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                dt.Columns.Add(firstRowCell.Text);
            }

            // Add data
            for (var rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
            {
                var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                var row = dt.NewRow();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
                dt.Rows.Add(row);
            }

            return dt;
        }

        private static DataTable ReadXls(string filePath, string sheetName)
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var workbook = new HSSFWorkbook(stream);
            var sheet = sheetName == null
                ? workbook.GetSheetAt(0)
                : workbook.GetSheet(sheetName);

            if (sheet == null)
                throw new Exception("Page not found");

            var dt = new DataTable();

            // Add titles
            var headerRow = sheet.GetRow(0);
            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                dt.Columns.Add(headerRow.GetCell(i).ToString());
            }

            // Add data
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null) continue;

                var dataRow = dt.NewRow();
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    dataRow[j] = row.GetCell(j)?.ToString();
                }
                dt.Rows.Add(dataRow);
            }

            return dt;
        }

        private static DataTable ReadCsv(string filePath, char delimiter = ',')
        {
            var dt = new DataTable();
            using var reader = new StreamReader(filePath);

            // Read the titles
            var headers = reader.ReadLine()?.Split(delimiter);
            if (headers == null)
                throw new Exception("CSV file is empty");

            foreach (var header in headers)
            {
                dt.Columns.Add(header);
            }

            // Read data
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (string.IsNullOrEmpty(line)) continue;

                var values = line.Split(delimiter);
                var row = dt.NewRow();
                for (int i = 0; i < Math.Min(values.Length, dt.Columns.Count); i++)
                {
                    row[i] = values[i];
                }
                dt.Rows.Add(row);
            }

            return dt;
        }

        private static DataTable ReadXlsb(string filePath, string sheetName)
        {
            var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            var dt = new DataTable();

            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                var sheet = sheetName ?? GetFirstSheetName(conn);
                var cmd = new OleDbCommand($"SELECT * FROM [{sheet}]", conn);
                using var adapter = new OleDbDataAdapter(cmd);
                adapter.Fill(dt);
            }

            return dt;
        }

        private static string GetFirstSheetName(OleDbConnection conn)
        {
            var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (dt == null || dt.Rows.Count == 0)
                throw new Exception("Page not found");

            return dt.Rows[0]["TABLE_NAME"].ToString();
        }

        #endregion

        #region Private Write Methods

        private static void WriteXlsx(DataTable data, string filePath, string sheetName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(sheetName);

            // Write the titles
            for (int i = 0; i < data.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = data.Columns[i].ColumnName;
            }

            // Write data
            for (int i = 0; i < data.Rows.Count; i++)
            {
                for (int j = 0; j < data.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1].Value = data.Rows[i][j];
                }
            }

            package.SaveAs(new FileInfo(filePath));
        }

        private static void WriteXls(DataTable data, string filePath, string sheetName)
        {
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet(sheetName);

            // Write the titles
            var headerRow = sheet.CreateRow(0);
            for (int i = 0; i < data.Columns.Count; i++)
            {
                headerRow.CreateCell(i).SetCellValue(data.Columns[i].ColumnName);
            }

            // Write data
            for (int i = 0; i < data.Rows.Count; i++)
            {
                var row = sheet.CreateRow(i + 1);
                for (int j = 0; j < data.Columns.Count; j++)
                {
                    row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                }
            }

            using var fileData = new FileStream(filePath, FileMode.Create);
            workbook.Write(fileData);
        }

        private static void WriteCsv(DataTable data, string filePath, char delimiter = ',')
        {
            using var writer = new StreamWriter(filePath);

            // Write the titles
            writer.WriteLine(string.Join(delimiter, data.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));

            // Write data
            foreach (DataRow row in data.Rows)
            {
                writer.WriteLine(string.Join(delimiter, row.ItemArray));
            }
        }

        private static void WriteXlsb(DataTable data, string filePath, string sheetName)
        {
            var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties=\"Excel 12.0;HDR=YES\"";

            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                // Create a table
                var createTable = new OleDbCommand($"CREATE TABLE [{sheetName}] ({GenerateColumnDefinitions(data)})", conn);
                createTable.ExecuteNonQuery();

                // Add data
                foreach (DataRow row in data.Rows)
                {
                    var columns = string.Join(", ", data.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
                    var values = string.Join(", ", row.ItemArray.Select(v => $"'{v}'"));
                    var insert = new OleDbCommand($"INSERT INTO [{sheetName}] ({columns}) VALUES ({values})", conn);
                    insert.ExecuteNonQuery();
                }
            }
        }

        private static string GenerateColumnDefinitions(DataTable data)
        {
            var definitions = new List<string>();
            foreach (DataColumn column in data.Columns)
            {
                definitions.Add($"[{column.ColumnName}] VARCHAR(255)");
            }
            return string.Join(", ", definitions);
        }

        #endregion
    }
}