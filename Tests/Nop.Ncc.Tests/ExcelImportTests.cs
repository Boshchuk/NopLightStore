using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Schema;
using NUnit.Framework;
//using OfficeOpenXml;
//using OfficeOpenXml.Drawing;

namespace Nop.Ncc.Tests
{
    [TestFixture]
    public class ExcelImportTests
    {
        //[Test]
        //public void TestImport()
        //{
        //    //const string excelFileName = "small.xlsx";
        //    const string excelFileName = "test2.xlsx";
        //    var importManager = new NccImportManager();

        //    using (var stream = new FileStream(excelFileName, FileMode.Open))
        //    {
        //        var results = importManager.GetProductsProductDatas(stream);
                
        //        Assert.Less(0,results.Count, "Now products");
        //    }
        //}

        //[Test]
        //public void GetImagesTest()
        //{
        //    var excelFileName = "Test2.xlsx";

        //    using (var stream = new FileStream(excelFileName, FileMode.Open))
        //    {
        //        using (var xlPackage = new ExcelPackage(stream))
        //        {
        //            // get the first worksheet in the workbook
        //            ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();
        //            Assert.IsNotNull(worksheet, "No worksheet found");

        //            if (Directory.Exists("pict"))
        //            {
        //                Directory.Delete("pict",true);
        //            }
        //            Directory.CreateDirectory("pict");
                    

        //            foreach (ExcelDrawing excelDrawing in worksheet.Drawings)
        //            {
        //               string pictureName = excelDrawing .Name;// Access Picture

        //               ExcelPicture pict = excelDrawing as ExcelPicture;


        //                var c = pict.From.Column;
        //                var r = pict.From.Row;

        //               pict.Image.Save(string.Format("pict/{0}_r{1}_c{2}.jpeg",pictureName, r,c)); 
        //            }
                        

        //        }
        //    }
        //}
    }
}
