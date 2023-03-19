// See https://aka.ms/new-console-template for more information

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
//using Microsoft.Office.Interop.Excel;


using NPOI.XSSF.UserModel;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
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

for (int i = 0; i < args.Length; ++i)
{
    switch (args[i])
    {
        case "-?":
            Console.WriteLine();
            Console.WriteLine("gparce [OPTIONS]");
            Console.WriteLine();
            Console.WriteLine("OPTIONS:");
            Console.WriteLine("   -i DIR - директория входных файлов, xls[x] со списками групп, def - ./");
            Console.WriteLine("   -o FILE - выходной файл (JSON), def - пусто");
            Console.WriteLine("   -s - вывод в стандартный поток");
            Console.WriteLine();

            return;
            break;
        case "-i": xlsdir = args[++i]; break;
        case "-o": outJsonName = args[++i]; break;
        case "-s": toStd = true; break;

    }
}

try
{
    var fl = Directory.GetFiles(xlsdir, "*.xls*");

    GroupList groups = new GroupList();

    //foreach (var f in fl)
    Parallel.ForEach(fl, f =>
    {
        FileInfo fi = new FileInfo(f);
        NPOI.SS.UserModel.IWorkbook wb;

        using FileStream file = new FileStream(fi.FullName, FileMode.Open, FileAccess.Read);
        if (fi.Extension == ".xls")
        {

            wb = new HSSFWorkbook(file);
        }
        else
        {
            wb = new XSSFWorkbook(file);
        }
        file.Close();


        for (int _i = 0; _i < wb.NumberOfSheets; ++_i)
        {
            var ws = wb.GetSheetAt(_i);


            List<Rect> rects = new List<Rect>();

            for (int row = 0; row < ws.LastRowNum; row++)
            {
                var _r = ws.GetRow(row);
                if (_r == null) continue;
                for (int col = 0; col < _r.LastCellNum; col++)
                {
                    ICell cell = _r.GetCell(col);


                    //Проверка, выхдит ли текущая ячейка в исключенный диапазон
                    if (rects.Select(x => x.IsThere(row, col)).Any(x => x == true))
                        continue;
                    //Console.WriteLine($"{wb.GetSheetName(_i)} {row} {col}");

                    //Console.WriteLine("===");
                    //ищем единицу на листе Экселя
                    //затем, если под ней есть столбик 1 2 3 ... - то это список группы
                    //считываем его пока последовательность е закончится


                    if (cell?.ToString() == "1")
                    {
                        utils.Group group = new utils.Group();

                        Rect r = new Rect { Row1 = row - 1, Col1 = col, Col2 = col + 3, Row2 = row };

                        for (int i = 0; i < 99; ++i)
                        {
                            var val = ws.GetRow(row + i)?.GetCell(col)?.ToString();

                            if (val == null) break;
                            if (!Int32.TryParse(val, out var num))
                                continue;
                            if (num - 1 == i)
                            {
                                r.Row2 = i;

                                var sn = ws.GetRow(row + i).GetCell(col + 1).ToString();
                                var n = ws.GetRow(row + i).GetCell(col + 2).ToString();
                                var ln = ws.GetRow(row + i).GetCell(col + 3).ToString();

                                Student student = new Student
                                {
                                    Surname = sn,
                                    Name = n,
                                    Lastname = ln,
                                };

                                //debug
                                //Console.WriteLine(student.Surname + " " + ss.Cells[row + i, col + 1].Font.Bold);
                                if (ws.GetRow(row + i).GetCell(col + 1).CellStyle.GetFont(wb).IsBold)
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

                            group.Name = ws.GetRow(row - 1).GetCell(col + 1).ToString();
                            groups.Groups.Add(group);
                        }

                    }
                }
            }
        }


    });

    groups.Groups.Sort((a, b) => a.Name.CompareTo(b.Name));

    JsonSerializerOptions options = new JsonSerializerOptions();
    options.WriteIndented = true;
    options.Encoder = JavaScriptEncoder.Create(UnicodeRanges.All);
    options.DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingDefault;

    string jsonString = JsonSerializer.Serialize(groups, options);

    if (outJsonName != "")
        File.WriteAllText(outJsonName, jsonString);

    if (toStd)
    {
        Console.WriteLine(jsonString);
    }

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
        return row >= Row1 && row <= Row2 && col >= Col1 && col <= Col2;
    }
};


