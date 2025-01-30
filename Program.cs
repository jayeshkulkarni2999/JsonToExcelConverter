using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace JsonToExcelConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "JsonFile", "SBOM.json");
            string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelFile", "ConvertedFile.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            string jsonContent = File.ReadAllText(jsonFilePath);
            JsonModel json = JsonConvert.DeserializeObject<JsonModel>(jsonContent);
            DataTable dataTable = GetDatatable(json.components);
            dataTable = RemoveEmptyRows(dataTable);
            string message = ExportDataTableToExcel(dataTable, excelFilePath);
            if(message == "Error")
            {
                Console.WriteLine("Console Application executed!!!");
            }
            else
            {
                Console.WriteLine("Console Failed!!!");
            }
        }

        public static DataTable GetDatatable(List<Components> components)
        {
            DataTable dt = new DataTable();
            try
            {
                
                dt.Columns.Add("LicenseName", typeof(string));
                dt.Columns.Add("LicenseURL", typeof(string));
                dt.Columns.Add("purl", typeof(string));
                dt.Columns.Add("version", typeof(string));

                foreach (var data in components)
                {
                    DataRow dr = dt.NewRow();

                    if (data.licenses.Count == 0)
                    {
                        dr["LicenseName"] = string.Empty;
                        dr["LicenseURL"] = string.Empty;
                        dr["purl"] = data.purl?.ToString() ?? string.Empty;
                        dr["version"] = data.version?.ToString() ?? string.Empty;
                        dt.Rows.Add(dr);
                    }
                    else
                    {
                        foreach (var licenseData in data.licenses)
                        {
                            dr["LicenseName"] = licenseData.license.name?.ToString() ?? string.Empty;
                            dr["LicenseURL"] = licenseData.license.url?.ToString() ?? string.Empty;
                            dr["purl"] = data.purl?.ToString() ?? string.Empty;
                            dr["version"] = data.version?.ToString() ?? string.Empty;
                            dt.Rows.Add(dr);
                        }
                    }
                }

                return dt;
            }
            catch(Exception ex)
            {
                Console.WriteLine("Exception at GetDatatable " + ex.Message + "-" + ex.StackTrace);
                return dt;
            }
        }

        static DataTable RemoveEmptyRows(DataTable dt)
        {
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                DataRow row = dt.Rows[i];

                bool isEmptyRow = true;
                foreach (DataColumn column in dt.Columns)
                {
                    if (row[column] != DBNull.Value && !string.IsNullOrEmpty(row[column].ToString()))
                    {
                        isEmptyRow = false;
                        break;
                    }
                }

                // If the row is empty, delete it
                if (isEmptyRow)
                {
                    dt.Rows.Remove(row);
                }
            }
            return dt;
        }


        static string ExportDataTableToExcel(DataTable dt, string filePath)
        {
            string resMessage = string.Empty;
            try
            {
                // Ensure EPPlus uses the ExcelPackage type (it is free in the newer versions)
                using (var package = new ExcelPackage())
                {
                    // Create an Excel worksheet from the DataTable
                    var worksheet = package.Workbook.Worksheets.Add("JsonToExcel");

                    // Load DataTable into the worksheet, starting at row 1 and column 1
                    worksheet.Cells["A1"].LoadFromDataTable(dt, PrintHeaders: true);

                    // Save the Excel file to the specified path
                    FileInfo fi = new FileInfo(filePath);
                    package.SaveAs(fi);
                }
                return "File exported Successfully!!!";
            }
            catch(Exception ex)
            {
                Console.WriteLine("Exception at ExportDataTableToExcel " + ex.Message + "-" + ex.StackTrace);
                return "Error";
            }
        }

        public class JsonModel
        {
            public string specVersion { get; set; }
            public List<Components> components { get; set; } = new List<Components>();

        }

        public class Components
        {
            public string version { get; set; }
            public string purl { get; set; }
            public List<Licenses> licenses { get; set; } = new List<Licenses>();
        }

        public class Licenses
        {
            public License license { get; set; } = new License();
        }

        public class License
        {
            public string name { get; set; }
            public string url { get; set; }
        }
    }
}
