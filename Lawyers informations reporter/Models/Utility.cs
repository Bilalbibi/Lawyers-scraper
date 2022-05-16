using MetroFramework.Controls;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace scrapingTemplateV51.Models
{
    public static class Utility
    {
        public static string ConnectionString = "Data Source=system.db;Version=3;";
        public static string SimpleDateFormat = "dd/MM/yyyy HH:mm:ss";

        public static void SaveCookies(CookieContainer cookieContainer, string url)
        {
            try
            {
                var cookies = new List<Cookie>();
                foreach (Cookie cookie in cookieContainer.GetCookies(new Uri(url)))
                    cookies.Add(new Cookie { Name = cookie.Name, Value = cookie.Value });
                File.WriteAllText("ses", JsonConvert.SerializeObject(cookies));
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        public static string BetweenStrings(string text, string start, string end)
        {
            var p1 = text.IndexOf(start, StringComparison.Ordinal) + start.Length;
            var p2 = text.IndexOf(end, p1, StringComparison.Ordinal);
            if (end == "") return (text.Substring(p1));
            else return text.Substring(p1, p2 - p1);
        }

        public static CookieContainer LoadCookies(string url)
        {
            var cookieContainer = new CookieContainer();
            try
            {
                var myCookies = JsonConvert.DeserializeObject<List<Cookie>>(File.ReadAllText("ses"));
                foreach (var myCookie in myCookies)
                    cookieContainer.Add(new Uri(url), new Cookie(myCookie.Name, myCookie.Value));
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return cookieContainer;
        }

    }
    public static class SaveToExcel
    {
        public static async void Save<T2>(this List<T2> objects, string path, string sheetName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var excelPkg = new ExcelPackage(new FileInfo(path));

            var sheet = excelPkg.Workbook.Worksheets.Add(sheetName);
            sheet.Protection.IsProtected = false;
            sheet.Protection.AllowSelectLockedCells = false;
            sheet.Row(1).Height = 20;
            sheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Row(1).Style.Font.Bold = true;
            sheet.Row(1).Style.Font.Size = 8;
            var col = 1;
            var colums = new Dictionary<string, string>();
            foreach (var propertyInfo in typeof(T2).GetProperties())
            {
                var propertyInfoName = propertyInfo.Name;
                var coulumName = propertyInfoName;
                for (int i = 1; i < propertyInfoName.Length - 1; i++)
                {
                    if (char.IsUpper(propertyInfoName[i]) && !char.IsUpper(propertyInfoName[i + 1]))
                        coulumName = coulumName.Replace(propertyInfoName[i] + "", " " + propertyInfoName[i]);
                }
                colums.Add(coulumName, propertyInfoName);
                sheet.Cells[1, col].Value = coulumName;
                col++;
            }

            var colNbr = typeof(T2).GetProperties().Count();
            var columnLetter = ExcelCellAddress.GetColumnLetter(colNbr);
            var range = sheet.Cells[$"A1:{columnLetter}{objects.Count + 1}"];
            var tab = sheet.Tables.Add(range, "");
            tab.TableStyle = TableStyles.Medium2;
            sheet.Cells.Style.Font.Size = 12;
            var row = 2;
            foreach (var obj in objects)
            {
                for (int i = 1; i <= sheet.Dimension.End.Column; i++)
                {
                    var colName = (string)sheet.Cells[1, i].Value;
                    var prop = colums[colName];
                    var value = (obj.GetType().GetProperty(prop))?.GetValue(obj, null);
                    sheet.Cells[row, i].Value = value;
                }
                row++;
            }
            for (int i = 1; i <= sheet.Dimension.End.Column; i++)
                sheet.Column(i).AutoFit();
            await excelPkg.SaveAsync();
        }
    }
    public static class Read
    {
        private static Regex _rgx = new Regex("[^a-zA-Z0-9 -]");
        public static List<List<T>> ReadAllSheetsFromExcel<T>(this string path) where T : new()
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            if (!File.Exists(path))
            {
                throw new Exception("File don't exist");
            }
            var list2 = new List<List<T>>();
            ExcelWorksheets excelWorksheets = new ExcelPackage(new FileInfo(path)).Workbook.Worksheets;
            PropertyInfo[] properties = typeof(T).GetProperties();
            foreach (var propertie in properties)
            {
                var x = propertie.Name;
            }
            List<string> list = new List<string>();
            foreach (var excelWorksheet in excelWorksheets)
            {
                var list3 = new List<T>();
                for (int i = 1; i <= excelWorksheet.Dimension.End.Column; i++)
                {
                    string input = excelWorksheet.Cells[1, i].Value?.ToString();
                    list.Add(_rgx.Replace(input, "").ToLower().Replace(" ", ""));
                }


                for (int j = 2; j <= excelWorksheet.Dimension.End.Row; j++)
                {
                    T val = new T();
                    for (int k = 1; k <= excelWorksheet.Dimension.End.Column; k++)
                    {
                        string name = list[k - 1];
                       
                        PropertyInfo propertyInfo = Enumerable.FirstOrDefault(properties, (PropertyInfo y) => y.Name.ToLower().Equals(name));
                        if (propertyInfo == null)
                        {
                            continue;
                        }

                        if (propertyInfo.PropertyType == typeof(DateTime))
                        {
                            try
                            {
                                DateTime dateTime = DateTime.Parse(excelWorksheet.Cells[j, k].Value?.ToString());
                                propertyInfo.SetValue(val, dateTime);
                            }
                            catch (Exception)
                            {
                            }
                        }
                        else if (propertyInfo.PropertyType == typeof(int))
                        {
                            if (int.TryParse(excelWorksheet.Cells[j, k].Value?.ToString() ?? "0", out int result))
                            {
                                propertyInfo.SetValue(val, result);
                            }
                        }
                        else if (propertyInfo.PropertyType == typeof(double))
                        {
                            if (double.TryParse(excelWorksheet.Cells[j, k].Value?.ToString() ?? "0.0", out double result2))
                            {
                                propertyInfo.SetValue(val, result2);
                            }
                        }
                        else if (propertyInfo.PropertyType == typeof(decimal))
                        {
                            if (decimal.TryParse(excelWorksheet.Cells[j, k].Value?.ToString() ?? "0.0", out decimal result3))
                            {
                                propertyInfo.SetValue(val, result3);
                            }
                        }
                        else if (propertyInfo.PropertyType == typeof(bool))
                        {
                            if (bool.TryParse(excelWorksheet.Cells[j, k].Value?.ToString().ToLower() ?? "false", out bool result4))
                            {
                                propertyInfo.SetValue(val, result4);
                            }
                        }
                        else
                        {
                            propertyInfo.SetValue(val, excelWorksheet.Cells[j, k].Value?.ToString());
                        }
                    }

                    list3.Add(val);
                }
                list2.Add(list3);
            }

            return list2;
        }
    }
}
