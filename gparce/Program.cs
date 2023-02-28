// See https://aka.ms/new-console-template for more information


using Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
using System.IO;
using System.Collections;

using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Unicode;

using utils;

// See https://aka.ms/new-console-template for more information

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





Application oXL;
_Workbook oWB;
try
{
    oXL = new Application();
    oXL.Visible = false;
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);

    return;
}
finally
{
   
}



try
{
    var fl = Directory.GetFiles(xlsdir, "*.xls*");

    GroupList groups = new GroupList();
    foreach (var f in fl)
    {
        FileInfo fi = new FileInfo(f);
        oWB = oXL.Workbooks.Add(fi.FullName);

        foreach (var sheet in oXL.Sheets)
        {
            _Worksheet ss = sheet as _Worksheet;
            List<Rect> rects = new List<Rect>();

            for (int row = 1; row < ss.UsedRange.Rows.Count; row++)
            {
                for (int col = 1; col < ss.UsedRange.Columns.Count; col++)
                {
                    //Проверка, выхдит ли текущая ячейка в исключенный диапазон
                    if (rects.Select(x => x.IsThere(row, col)).Any(x => x == true))
                        continue;

                    Group group = new Group();
                    //Console.WriteLine("===");
                    //ищем единицу на листе Экселя
                    //затем, если под ней есть столбик 1 2 3 ... - то это список группы
                    //считываем его пока последовательность е закончится
                    if (ss.Cells[row, col].Text == "1")
                    {
                        Rect r = new Rect { Row1 = row - 1, Col1 = col, Col2 = col + 3, Row2 = row };

                        for (int i = 0; i < 99; ++i)
                        {
                            var val = ss.Cells[row + i, col].Value;
                            if (val == null) break;
                            int num = (int)val;
                            if (num - 1 == i)
                            {
                                r.Row2 = i;
                                Student student = new Student
                                {
                                    Surname = ss.Cells[row + i, col + 1].Value ?? "",
                                    Name = ss.Cells[row + i, col + 2].Value ?? "",
                                    Lastname = ss.Cells[row + i, col + 3].Value ?? "",
                                };

                                //debug
                                //Console.WriteLine(student.Surname + " " + ss.Cells[row + i, col + 1].Font.Bold);

                                if ( Convert.ToString(ss.Cells[row + i, col + 1].Font.Bold) == "True")   
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

                            group.Name = ss.Cells[row - 1, col + 1].Text;
                            groups.Groups.Add(group);  
                        }

                    }
                }
            }
        }
        oWB.Close();

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
    oXL.Quit();
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


