// See https://aka.ms/new-console-template for more information

using System.Xml.Linq;
using System.IO;
using System.Collections;

using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Unicode;


using utils;
using System.ComponentModel;
using OfficeOpenXml;
using System.Transactions;

using Microsoft.Office.Interop.Excel;

// See https://aka.ms/new-console-template for more information

ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

///Входной каталог
string xlsdir = ".\\";
///Имя выходного Json
string outJsonName = "";

bool noFileOutput = false;

///Вывод в стандартный поток
bool toStd = false;

for (int i = 0; i<args.Length; ++i)
{
    switch(args[i])
    {
        case "-?":
            Console.WriteLine();
            Console.WriteLine("gparce [OPTIONS]");
            Console.WriteLine();
            Console.WriteLine("OPTIONS:");
            Console.WriteLine("   -i DIR - директория входных файлов, xls со списками групп, def - ./");
            Console.WriteLine("   -o FILE - выходной файл (JSON), def - пусто");
            Console.WriteLine("   -s - вывод в стандартный поток");
            Console.WriteLine();

            return;
            break;
        case "-i": xlsdir= args[i+1]; break;
        case "-o": outJsonName = args[i+1]; break;
        case "-s": toStd = true; break;

    }
}


try
{
    var fl = Directory.GetFiles(xlsdir, "*.xls*");

    GroupList groups = new GroupList();
    foreach (var f in fl)
    {
        FileInfo fi = new FileInfo(f);
        //if (fi.Extension == ".xls")
        //{
        //    Application a = new Application();
        //    _Workbook wb = a.Workbooks.Add(fi.FullName);
        //    wb.SaveAs2(fi.Name, XlFileFormat.xlOpenXMLWorkbook);
        //    a.Quit();
        //}
        ExcelPackage ex = new ExcelPackage(fi.FullName);

        foreach (var sheet in ex.Workbook.Worksheets)
        {

            List<Rect> rects = new List<Rect>();

            for (int row = 1; row < 100; row++)
            {
                for (int col = 1; col < 100; col++)
                {
                    //Проверка, выхдит ли текущая ячейка в исключенный диапазон
                    if (rects.Select(x => x.IsThere(row, col)).Any(x => x == true))
                        continue;
                    Console.WriteLine($"{sheet.Name} {row} {col}");
                    Group group = new Group();
                    //Console.WriteLine("===");
                    //ищем единицу на листе Экселя
                    //затем, если под ней есть столбик 1 2 3 ... - то это список группы
                    //считываем его пока последовательность е закончится
                    if (sheet.Cells[row,col].Text == "1")
                    {
                        Rect r = new Rect { Row1 = row - 1, Col1 = col, Col2 = col + 3, Row2 = row };

                        for (int i = 0; i < 99; ++i)
                        {
                            var val = sheet.Cells[row + i, col].Text;
                            if (val == null) break;
                            if (!Int32.TryParse(val, out var num))
                                continue;
                            if (num - 1 == i)
                            {
                                r.Row2 = i;
                                var sn = sheet.Cells[row + i, col + 1]?.Value?.ToString() ?? "";
                                var n = sheet.Cells[row + i, col + 2]?.Value?.ToString() ?? "";
                                var ln = sheet.Cells[row + i, col + 3]?.Value?.ToString() ?? "";

                                Student student = new Student
                                {
                                    Surname = sn,
                                    Name = n,
                                    Lastname = ln                                   ,
                                };

                                //debug
                                //Console.WriteLine(student.Surname + " " + ss.Cells[row + i, col + 1].Font.Bold);
                                if ( sheet.Cells[row + i, col + 1].Style.Font.Bold)   
                                    student.IsHeadman = true;
                                group.Students.Add(student);

                                
                            }
                            else
                                break;
                        }
                        //добавляем текущий диапазон, в котором список группы, в запрещенный
                        //для оптимизации  обхода ячеек
                        rects.Add(r);

                        if (group.Students.Count > 1)
                        {

                            group.Name = sheet.Cells[row - 1, col + 1].Text;
                            groups.Groups.Add(group);  
                        }

                    }
                }
            }
        }
       

    }

    JsonSerializerOptions options = new JsonSerializerOptions();
    options.WriteIndented = true;
    options.Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic);
    options.DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingDefault;


    string jsonString = JsonSerializer.Serialize(groups, options);


    if (outJsonName != "")
        File.WriteAllText(outJsonName, jsonString);

    

    if (toStd)
        Console.WriteLine(jsonString);
}
catch (Exception ex)
{
    Console.WriteLine(ex.ToString());
}
finally
{
    
}


/// <summary>
/// Прямоульник на листе Эксселя. 
/// </summary>
public class Rect
{
    /// <summary>
    /// Верхняя граница, включительно
    /// </summary>
    public int Row1 { set; get; }
    /// <summary>
    /// Левая граница, включительно
    /// </summary>
    public int Col1 { set; get; }
    /// <summary>
    /// Нижняя строка, включительно
    /// </summary>
    public int Row2 { set; get; }
    /// <summary>
    /// Правая граница включительно
    /// </summary>
    public int Col2 { set; get; }

    /// <summary>
    /// Попадает ли ячейка row col в прямоугольник?
    /// </summary>
    /// <param name="row"></param>
    /// <param name="col"></param>
    /// <returns>Да/нет</returns>
    public bool IsThere(int row, int col)
    {
        return row >= Row1 && row <= Row2 && col >=Col1 && col <= Col2;
    }
};


