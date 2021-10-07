using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using kudelinSUKA;
using OfficeOpenXml;
using System.Linq;

namespace kudelinSUKA_CumСОЛЬ
{
    class Program
    {



        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(@"C:\Users\Egor\Desktop\JijaPapka\Jija.xlsx");
            using (var package = new ExcelPackage(file))
            {
                if (package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "My Sheet") == null)
                {
                    var sheet = package.Workbook.Worksheets.Add("My Sheet");
                    sheet.Cells["A1"].Value = "Hello World!";
                }


                // Save to file
                package.Save();
            }





        }
    }
}
