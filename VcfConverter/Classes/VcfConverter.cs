using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using MixERP.Net.VCards;
using MixERP.Net.VCards.Models;
using MixERP.Net.VCards.Serializer;
using MixERP.Net.VCards.Types;

namespace VcfConverter.Classes
{
    public class VcfConverter
    {
        private readonly string _filePath;
        private readonly Type _typeOfEnumerable = typeof(IEnumerable);
        public VcfConverter(string filePath)
        {
            _filePath = filePath;
        }

        public static string EncodeQuotedPrintable(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            var builder = new StringBuilder();

            var bytes = Encoding.UTF8.GetBytes(value);
            foreach (var v in bytes)
            {
                // The following are not required to be encoded:
                // - Tab (ASCII 9)
                // - Space (ASCII 32)
                // - Characters 33 to 126, except for the equal sign (61).

                if ((v == 9) || ((v >= 32) && (v <= 60)) || ((v >= 62) && (v <= 126)))
                {
                    builder.Append(System.Convert.ToChar(v));
                }
                else
                {
                    builder.Append('=');
                    builder.Append(v.ToString("X2"));
                }
            }

            var lastChar = builder[^1];
            if (char.IsWhiteSpace(lastChar))
            {
                builder.Remove(builder.Length - 1, 1);
                builder.Append('=');
                builder.Append(((int)lastChar).ToString("X2"));
            }

            return builder.ToString();
        }
        public static string DecodeQuotedPrintable(string input, string charSet)
        {
            Encoding enc;

            try
            {
                enc = Encoding.GetEncoding(charSet);
            }
            catch
            {
                enc = new UTF8Encoding();
            }

            var occurrences = new Regex(@"(=[0-9A-Z]{2}){1,}", RegexOptions.Multiline);
            var matches = occurrences.Matches(input);

            foreach (Match match in matches)
            {
                try
                {
                    byte[] b = new byte[match.Groups[0].Value.Length / 3];
                    for (int i = 0; i < match.Groups[0].Value.Length / 3; i++)
                    {
                        b[i] = byte.Parse(match.Groups[0].Value.Substring(i * 3 + 1, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
                    }
                    char[] hexChar = enc.GetChars(b);
                    input = input.Replace(match.Groups[0].Value, new string(hexChar));
                }
                catch
                {; }
            }
            input = input.Replace("?=", "").Replace("\r\n", "");

            return input;
        }
        private List<Dictionary<string, List<dynamic>>> GetValuesFromVcfFile()
        {
            var ignoredProps = new List<string> { "FormattedName" };
            var excelRowsData = new List<Dictionary<string, List<dynamic>>>();
            var vCardProps = typeof(VCard).GetProperties().ToList();
            // تصحیح مقادیر
            var fileContent = File.ReadAllText(_filePath);
            fileContent = Regex.Replace(fileContent, @"=\r\n", "", RegexOptions.IgnoreCase);
            var vCards = Deserializer.GetVCards(fileContent).ToList();
            foreach (var vCard in vCards)
            {
                var excelRowData = new Dictionary<string, List<dynamic>>();
                foreach (var pi in vCardProps)
                {
                    var defaultValue = pi.PropertyType.IsValueType ? Activator.CreateInstance(pi.PropertyType) : null;
                    var propValue = pi.GetValue(vCard);
                    if (propValue == null || propValue.Equals(defaultValue) || ignoredProps.Contains(pi.Name)) continue;
                    var propName = pi.Name;
                    if (pi.PropertyType == typeof(string))
                    {
                        var propValueStr = DecodeQuotedPrintable(propValue.ToString(), "utf-8");
                        if (string.IsNullOrWhiteSpace(propValueStr)) continue;
                        excelRowData.Add(propName, new List<dynamic> { propValueStr });
                    }
                    else if (_typeOfEnumerable.IsAssignableFrom(pi.PropertyType))
                    {
                        var values = new List<dynamic>();
                        var propValueList = (IList)propValue;
                        foreach (var v in propValueList)
                        {
                            var decodedValue = "";
                            switch (v)
                            {
                                case Telephone telephone:
                                    decodedValue = DecodeQuotedPrintable(telephone.Number, "utf-8");
                                    break;
                                case Email email:
                                    decodedValue = DecodeQuotedPrintable(email.EmailAddress, "utf-8");
                                    break;
                                case Address address:
                                    decodedValue = DecodeQuotedPrintable(address.Street, "utf-8");
                                    break;
                            }
                            if (!string.IsNullOrWhiteSpace(decodedValue))
                                values.Add(decodedValue);
                        }
                        if (values.Any())
                            excelRowData.Add(propName, values);
                    }
                }
                excelRowsData.Add(excelRowData);
            }
            return excelRowsData;
        }
        private static byte[] ConvertToExcel(List<Dictionary<string, List<dynamic>>> rowsData)
        {
            var excelColumns = new List<ExcelCellModel>();
            var excelRowsData = new List<List<dynamic>>();
            var arrayLength = rowsData.Max(q => q.Count) + rowsData.Select(q => q.Values).Max(q => q.Count);
            // ستون های اکسل
            foreach (var rowDataDic in rowsData)
            {
                foreach (var (key, _) in rowDataDic)
                {
                    var excelColumnIndex = excelColumns.FindIndex(q => q.Title.Equals(key, StringComparison.InvariantCultureIgnoreCase));
                    if (excelColumnIndex != -1) continue;
                    var excelColumn = new ExcelCellModel { Title = key, Width = 25 };
                    excelColumns.Add(excelColumn);
                }
            }
            excelColumns = excelColumns.OrderBy(q => q.Title).ToList();
            foreach (var rowDataDic in rowsData)
            {
                var excelRowData = new dynamic[arrayLength];
                foreach (var (key, value) in rowDataDic)
                    excelRowData[excelColumns.FindIndex(q => q.Title.Equals(key, StringComparison.InvariantCultureIgnoreCase))] = string.Join(Environment.NewLine, value);
                excelRowsData.Add(excelRowData.ToList());
            }
            return ExcelHelper.GetExcelFile("Contacts", false, new ExcelFileModel
            {
                Cells = excelColumns,
                Rows = excelRowsData
            });
        }
        private byte[] ConvertToVcf(Action<string> logger)
        {
            var vCardStringBuilder = new StringBuilder();
            var vCardProps = typeof(VCard).GetProperties().ToList();
            var excelRowsData = ExcelHelper.GetExcelRowsData(_filePath);
            foreach (var dictionary in excelRowsData)
            {
                var vCard = new VCard { Version = VCardVersion.V2_1 };
                foreach (var (propName, value) in dictionary)
                {
                    var pi = vCardProps.Find(q => q.Name.Equals(propName));
                    if (pi == null)
                    {
                        logger.Invoke($"{propName} is not valid in VCard");
                        continue;
                    }
                    var values = (value as string)?.Split("\n", StringSplitOptions.RemoveEmptyEntries).Distinct().Select(q => q.Replace("\r", "").Trim()).ToList();
                    if (values == null || !values.Any()) continue;
                    switch (propName)
                    {
                        case nameof(VCard.Addresses):
                            vCard.Addresses = values.Select(v => new Address { Street = v }).ToList();
                            break;
                        case nameof(VCard.Emails):
                            vCard.Emails = values.Select(v => new Email { EmailAddress = v }).ToList();
                            break;
                        case nameof(VCard.Telephones):
                            vCard.Telephones = values.Select(v => new Telephone { Number = v }).ToList();
                            break;
                        default:
                            pi.SetValue(vCard, value);
                            break;
                    }
                }
                vCardStringBuilder.Append(vCard.Serialize());
            }
            return Encoding.UTF8.GetBytes(vCardStringBuilder.ToString());
        }

        public VcfConverterResult Convert(Action<string> logger)
        {
            VcfConverterFileFormat fileFormat;
            byte[] fileContent;
            var fileExtension = Path.GetExtension(_filePath);
            switch (fileExtension)
            {
                case ".xlsx":
                    fileFormat = VcfConverterFileFormat.Vcf;
                    fileContent = ConvertToVcf(logger);
                    break;
                case ".vcf":
                    fileFormat = VcfConverterFileFormat.Excel;
                    var values = GetValuesFromVcfFile();
                    fileContent = ConvertToExcel(values);
                    break;
                default:
                    throw new Exception($"File Extension {fileExtension} is not supported");
            }

            return new VcfConverterResult
            {
                FileFormat = fileFormat,
                FileContent = fileContent
            };
        }
    }
}
