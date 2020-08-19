using OpenDoc.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OpenDoc.ExcelTests
{
    class Program
    {
        static void Main(string[] args)
        {
            var model = new ParameterDataXBarRangeViewModel()
            {
                XBarRangeDataList = new List<XBarRangeViewModel>()
            };
            model.XBarRangeDataList.Add(new XBarRangeViewModel
            {
                GroupID = 10,
                Average = 9
            });
            model.XBarRangeDataList.Add(new XBarRangeViewModel
            {
                GroupID = 13,
                Average = 5
            });
            model.XBarRangeDataList.Add(new XBarRangeViewModel
            {
                GroupID = 20,
                Average = 4
            });
            model.ParameterComputeData = new ParameterComputeViewModel();
            try
            {
                FileInfo file = new System.IO.FileInfo("template2.xml");
                var template = new ExcelTemplate(model, new XmlConfig(file));
                using (var stream = File.Create("doc.xlsx"))
                {
                    var bytes = template.Generate();
                    stream.Write(bytes, 0, bytes.Length);
                }
                Process.Start("doc.xlsx");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
            }
            Console.ReadKey();
        }
    }
}
