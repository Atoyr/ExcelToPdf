using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.IO;

namespace ExcelToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() == 0)
            {
                // TODO write Help
                return;
            }

            string directory = string.Empty;
            string outputPath = string.Empty;
            string inputPath = string.Empty;
            var sheetnames = new List<string>();

            string nextParamKey = string.Empty;
            foreach (var arg in args)
            {
                if (string.IsNullOrEmpty(nextParamKey))
                {
                    if (arg[0] == '-')
                    {
                        var f = arg.Substring(1).ToLower();
                        switch (f)
                        {
                            case "d":
                                nextParamKey = "directory";
                                break;
                            case "o":
                                nextParamKey = "output";
                                break;
                            case "i":
                                nextParamKey = "input";
                                break;
                            default:
                                Console.WriteLine("flag not found");
                                return;
                        }
                    }
                    else
                    {
                        sheetnames.Add(arg);
                    }
                }
                else
                {
                    switch (nextParamKey)
                    {
                        case "directory":
                            directory = arg;
                            break;
                        case "output":
                            outputPath = arg;
                            break;
                        case "input":
                            inputPath = arg;
                            break;
                    }
                    nextParamKey = string.Empty;
                }
            }

            if (string.IsNullOrEmpty(outputPath) || !Directory.Exists(outputPath))
            {
                Console.WriteLine("directory not found");
                return;
            }
            outputPath = Path.GetFullPath(outputPath);


            var app = new Excel.Application();
            app.Visible = false;
            if (!string.IsNullOrEmpty(directory))
            {
                foreach (var filePath in System.IO.Directory.GetFiles(directory, "*.xls", System.IO.SearchOption.AllDirectories))
                {
                    try
                    {
                        var f = Path.GetFullPath(filePath);
                        if (f[0] == '~') continue;
                        if (sheetnames.Count() == 0)
                        {
                            exportPdf(app, outputPath, f);
                        }
                        else
                        {
                            exportPdf(app, outputPath, f, sheetnames);
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
            }
            if (!string.IsNullOrEmpty(inputPath))
            {
                inputPath = Path.GetFullPath(inputPath);
                try
                {
                    if (sheetnames.Count() == 0)
                    {
                        exportPdf(app, outputPath, inputPath);
                    }
                    else
                    {
                        exportPdf(app, outputPath, inputPath, sheetnames);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }


        private static void exportPdf(Excel.Application app, string outputpath, string inputpath, IList<string> sheetnames)
        {
            if (!File.Exists(inputpath)) throw new Exception("FIle Not Found");
            Excel.Workbook book = app.Workbooks.Open(inputpath);
            string fileName = Path.GetFileNameWithoutExtension(inputpath);

            foreach (var sheetname in sheetnames)
            {
                var index = getSheetIndex(sheetname, book.Sheets);
                Console.WriteLine(Path.Combine(outputpath, fileName + "_" + sheetname + ".pdf"));
                if (index > 0)
                {
                    (book.Sheets[index] as Excel.Worksheet).ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
                    Path.Combine(outputpath, fileName + "_" + sheetname + ".pdf"),
                    Excel.XlFixedFormatQuality.xlQualityStandard,
                    true,
                    false,
                    Type.Missing,
                    Type.Missing,
                    false,
                    Type.Missing);
                }
            }
            book.Close(false);
        }

        private static void exportPdf(Excel.Application app, string outputpath, string inputpath)
        {
            if (!File.Exists(inputpath)) throw new Exception("FIle Not Found");
            Excel.Workbook book = app.Workbooks.Open(inputpath);
            string fileName = Path.GetFileNameWithoutExtension(inputpath);
            book.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
            Path.Combine(outputpath, fileName + ".pdf"),
            Excel.XlFixedFormatQuality.xlQualityStandard,
            true,
            false,
            Type.Missing,
            Type.Missing,
            false,
            Type.Missing);

            book.Close(false);
        }

        // 指定されたワークシート名のインデックスを返すメソッド
        private static int getSheetIndex(string sheetName, Excel.Sheets shs)
        {
            int i = 0;
            foreach (Excel.Worksheet sh in shs)
            {
                if (sheetName == sh.Name)
                {
                    return i + 1;
                }
                i += 1;
            }
            return 0;
        }
    }

}
