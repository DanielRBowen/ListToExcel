![tag](https://dev6.blob.core.windows.net/blog-images/tag-list-excel.png)
# ListToExcel
Export List of T to Excel and Parse Excel to List of T with validation and highlighting options


Please let me know if you find any problems or have any questions.

```C#
namespace ExcelHelper.ExamplesAndTests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using ClosedXML.Excel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Tynamix.ObjectFiller;

    [TestClass]
    public class Examples
    {
        [TestMethod]
        public void RoundTrip()
        {
            var excelHelper = new ExcelHelper();
            var list = new Filler<WidgetSauce>().Create(100).ToList();
            var fileName = "WidgetSauces.xlsx";
            using (var excelResult = excelHelper.ListToExcel(list))
            {
                excelResult.SaveAs(fileName);//look in the bin for this :)
            }

            Func<WidgetSauce, List<string>> validator = sauce =>
            {
                var result = new List<string>();

                if (sauce.WidgetPrice > 1234567483)
                {
                    result.Add("This widget is way to expensive.");
                }

                if (sauce.WidgetPrice < 0)
                {
                    result.Add("I'm not gonna give you money for my widget sauce.");
                }

                return result;
            };

            using (var fileStream = File.Open(fileName, FileMode.Open))
            {
                using (var excelParsed = excelHelper.ParseExcel(excelStream: fileStream, validateT: validator))
                {
                    Assert.IsNotNull(excelParsed);
                    Assert.IsNotNull(excelParsed.ValidList);
                    excelParsed.XlWorkbook.SaveAs($"Parsed_{fileName}");
                }
            }
        }

        [TestMethod]
        public void RoundTripExtended()
        {
            var excelHelper = new ExcelHelperExtend();

            var list = new Filler<WidgetSauce>().Create(100).ToList();
            var fileName = "WidgetSaucesEx.xlsx";
            using (var excelResult = excelHelper.ListToExcel(list))
            {
                excelResult.SaveAs(fileName);//look in the bin for this :)
            }

            Func<WidgetSauce, List<string>> validator = sauce =>
            {
                var result = new List<string>();

                if (sauce.WidgetPrice > 1234567483)
                {
                    result.Add("This widget is way to expensive.");
                }

                if (sauce.WidgetPrice < 0)
                {
                    result.Add("I'm not gonna give you money for my widget sauce.");
                }

                return result;
            };

            using (var fileStream = File.Open(fileName, FileMode.Open))
            {
                using (var excelParsed = excelHelper.ParseExcel(excelStream: fileStream, validateT: validator))
                {
                    Assert.IsNotNull(excelParsed);
                    Assert.IsNotNull(excelParsed.ValidList);
                    excelParsed.XlWorkbook.SaveAs($"Parsed_{fileName}");
                }
            }
        }

        [TestMethod]
        public void BuildExcelTemplate()
        {
            var xl = new ExcelHelper();
            using (var excelResult = xl.BuildExcelTemplate<WidgetSauce>())
            {
                excelResult.SaveAs("Template.xlsx");
            }
        }
    }
    public class WidgetSauce
    {
        public Guid WidgetGuid { get; set; }
        [ExcelColumnName("Widget /Id#")]//when the columns names are crazy and can't be parsed directy to C#. Works in both directions to list and parse.
        public int WidgetId { get; set; }
        public int? SomeNumber { get; set; }
        public string WidgetName { get; set; }
        public decimal WidgetPrice { get; set; }
    }

    public class ExcelHelperExtend : ExcelHelper
    {
        public override void FillCellBackground(ref IXLCell cell, bool isValid)
        {
            cell.Style.Fill.BackgroundColor = isValid ? XLColor.PinkOrange : XLColor.Green;
        }

        public override void FillRowBackgroundWithValidationMessage(ref IXLRow row, bool isValid, List<string> validationMessages)
        {
            row.Style.Fill.BackgroundColor = isValid ? XLColor.Green : XLColor.PinkOrange;
            if (!isValid && validationMessages != null && validationMessages.Any())//only add messages to invalid rows
            {
                var messageCell = row.Cell(row.LastCellUsed().CellRight().WorksheetColumn().ColumnNumber());
                messageCell.Value = String.Join(">>>>>", validationMessages);
            }
        }
    }   
}
```

Web API use example:
```C#
/// <summary>
/// Gets the users as an excel spreadsheet
/// Uses https://www.nuget.org/packages/ListToExcel.Fork/1.0.0
/// https://github.com/DanielRBowen/ListToExcel
/// Application Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
/// </summary>
/// <returns></returns>
[HttpGet, ActionName("GetUsersAsExcel")]
public async Task<IActionResult> GetUsersAsExcel()
{
    try
    {
        var users = await QueryUsersViewModels().ToListAsync();
        var excelHelper = new ExcelHelper.ExcelHelper();
        using var workbook = excelHelper.ListToExcel(users);
        using var memoryStream = new MemoryStream();
        workbook.SaveAs(memoryStream);
        var bytes = memoryStream.ToArray();
        return Ok(bytes);
    }
    catch (Exception ex)
    {
        _logger.LogError(ex, ex.Message);
        return StatusCode(500, ex);
    }
}
``
