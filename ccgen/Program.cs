using utils;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text.Json;
using System.Text.Json.Serialization;
using NPOI.HPSF;
using NPOI.Util;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Unicode;
using NPOI.SS.Util;
using MathNet.Numerics;

Console.InputEncoding = Encoding.Unicode;
Console.OutputEncoding = Encoding.Unicode;

string inputJson = "";
string outFileName = "";
string paternFile = "";
string _parameters = "A1;A2;B2;1";

FileInfo fi = new FileInfo("pattern.xlsx");
bool singleFile = false;


for (int i = 0; i < args.Length; ++i)
{
    switch (args[i])
    {
        case "-?":
            Console.WriteLine();
            Console.WriteLine("ccgen [OPTIONS]");
            Console.WriteLine();
            Console.WriteLine("OPTIONS:");
            Console.WriteLine("   -i FILE - входной файл");
            Console.WriteLine("   -o PATH - путь до файла или директории с результатом");
            Console.WriteLine("   -g - вывод в один файл");
            Console.WriteLine("   -p FILE - файл шаблона с группами");
            Console.WriteLine("   -s SETTINGS_STRTING - строка с настройками:");
            Console.WriteLine("      CELL_GROUP;CELL_LIST;CELL_STUDENT;MODE");
            Console.WriteLine("      CELL_LIST - ячейка, c которой начинается нумерация студентов");
            Console.WriteLine("      CELL_GROUP - ячейка, куда пишется название группы");
            Console.WriteLine("      CELL_STUDENT - ячейка с которой начинается списокгруппы");
            Console.WriteLine("      MODE - Режим:");
            Console.WriteLine("         1 - ФИО занимает 3 ячейки");
            Console.WriteLine("         2 - ФИО занимает 1 ячейку");
            Console.WriteLine("         3 - ФИО занимает 1 ячейку и записывается в формате Фамилия И. О.");
            Console.WriteLine();
            return;
            break;
        case "-i": inputJson = args[++i]; break;
        case "-o": outFileName = args[++i]; break;
        case "-p": paternFile = args[++i]; break;
        case "-g": singleFile = true; break;
        case "-s": _parameters = args[++i]; break;

    }
}

FileInfo out_fi = new FileInfo(outFileName);

if (!Console.IsInputRedirected && inputJson == "")
{
    Console.WriteLine("Для программы нет входных данных");
    return;
}

if (Console.IsInputRedirected && inputJson != "")
{
    Console.WriteLine("Неоднозначный ввод");
    return;
}

var parameters = _parameters.Split(";");
if (parameters.Length != 4)
{
    Console.WriteLine($"Ошбика в строке настройки {_parameters}");
    return;
}
int mode;

if (!Int32.TryParse(parameters[3],out mode) || mode > 3 )
{
    Console.WriteLine("Неизвестный режим");
    return;
}

//string json = File.ReadAllText(groupsjson);
string json = "";
string line = "";

if (Console.IsInputRedirected)
{
    var sr = new StringBuilder();

    while ((line = Console.ReadLine()) != null && line != "")
    {
        sr.AppendLine(line);
    }
    json = sr.ToString();
}
else
{
    File.ReadAllText(inputJson, Encoding.UTF8);
}



JsonSerializerOptions options = new JsonSerializerOptions
{
    Encoder = JavaScriptEncoder.Create(),
};

var grlist = JsonSerializer.Deserialize<GroupList>(json, options);

Console.WriteLine(json);

NPOI.SS.UserModel.IWorkbook wb = null;

if (paternFile == "")
{
    wb = new XSSFWorkbook();
    wb.CreateSheet();
}
else
{
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
}



var paternSheet = wb.GetSheetAt(0);



IWorkbook? newGr = null;
ICellStyle? BoldStyle = null;
IFont? font = null;

if (singleFile)
{
    newGr = new XSSFWorkbook();
    font = newGr.CreateFont();
    font.IsBold = true;

}


foreach (Group gr in grlist.Groups)
{
    if (!singleFile)
    { 
        newGr = new XSSFWorkbook();
        font = newGr.CreateFont();
        font.IsBold = true;
    }
    paternSheet.CopyTo(newGr, gr.Name, true, true);
    var newSheet = newGr.GetSheetAt(newGr.NumberOfSheets-1);
     
   

    if (parameters[0]!= "")
    {
        CellReference groups_start_cell = new CellReference(parameters[0]);


        while (groups_start_cell.Row > newSheet.LastRowNum)
            newSheet.CreateRow(newSheet.LastRowNum + 1);

        newSheet.GetRow(groups_start_cell.Row).GetCell(groups_start_cell.Col, MissingCellPolicy.CREATE_NULL_AS_BLANK).SetCellValue(gr.Name);


    }




    for (int i=0; i<gr.Students.Count; ++i)
    {


        if (parameters[1] != "")
        {
            CellReference numbers_start_cell = new CellReference(parameters[1]);

            while (numbers_start_cell.Row + i > newSheet.LastRowNum)
                newSheet.CreateRow(newSheet.LastRowNum + 1);

            ICell number_cell = newSheet.GetRow(numbers_start_cell.Row + i).GetCell(numbers_start_cell.Col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            number_cell.SetCellValue(i + 1);

            if (gr.Students[i].IsHeadman)
            {
                var style = newGr.CreateCellStyle();
                try
                {
                    style.CloneStyleFrom(number_cell.CellStyle);
                }
                catch (Exception e) { }

                style.SetFont(font);
                number_cell.CellStyle = style;
            }

        }

        if (parameters[2] != "")
        {
            CellReference students_start_cell = new CellReference(parameters[2]);

            while (students_start_cell.Row + i > newSheet.LastRowNum)
                newSheet.CreateRow(newSheet.LastRowNum + 1);

            switch (mode)
            {
                case 1:
                    {
                        ICell f_cell = newSheet.GetRow(students_start_cell.Row + i).GetCell(students_start_cell.Col + 0, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        ICell n_cell = newSheet.GetRow(students_start_cell.Row + i).GetCell(students_start_cell.Col + 1, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        ICell o_cell = newSheet.GetRow(students_start_cell.Row + i).GetCell(students_start_cell.Col + 2, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                        f_cell.SetCellValue(gr.Students[i].Surname);
                        n_cell.SetCellValue(gr.Students[i].Name);
                        o_cell.SetCellValue(gr.Students[i].Lastname);

                        if (gr.Students[i].IsHeadman)
                        {
                            var style = newGr.CreateCellStyle();
                            try
                            {
                                style.CloneStyleFrom(f_cell.CellStyle);
                            }
                            catch(Exception e) { }
                            
                            style.SetFont(font);
                            f_cell.CellStyle = style;


                            style = newGr.CreateCellStyle();
                            try
                            {
                                style.CloneStyleFrom(n_cell.CellStyle);
                            }
                            catch (Exception e) { }
                           
                            style.SetFont(font);
                            n_cell.CellStyle = style;

                            style = newGr.CreateCellStyle();
                            try
                            {
                                style.CloneStyleFrom(o_cell.CellStyle);
                            }
                            catch (Exception e) { }
                            
                            style.SetFont(font);
                            o_cell.CellStyle = style;
                        }

                        break;
                    }
                case 2:
                    {
                        var fio = gr.Students[i].Surname + " ";

                        if (gr.Students[i].Name != "")
                            fio += gr.Students[i].Name + " ";

                        if (gr.Students[i].Lastname != "")
                            fio += gr.Students[i].Lastname;

                        ICell fio_cell = newSheet.GetRow(students_start_cell.Row + i).GetCell(students_start_cell.Col);
                        fio_cell.SetCellValue(fio);

                        if (gr.Students[i].IsHeadman)
                        {
                            var style = newGr.CreateCellStyle();
                            try
                            {
                                style.CloneStyleFrom(fio_cell.CellStyle);
                            }
                            catch (Exception e) { }

                            style.SetFont(font);
                            fio_cell.CellStyle = style;
                        }
                        break;
                    }
                case 3:
                    {
                        var fio = gr.Students[i].Surname + " ";

                        if (gr.Students[i].Name != "")
                            fio += gr.Students[i].Name[0] + ".";

                        if (gr.Students[i].Lastname != "")
                            fio += gr.Students[i].Lastname[0] + ".";

                        ICell fio_cell = newSheet.GetRow(students_start_cell.Row + i).GetCell(students_start_cell.Col);
                        fio_cell.SetCellValue(fio);

                        if (gr.Students[i].IsHeadman)
                        {
                            var style = newGr.CreateCellStyle();
                            try
                            {
                                style.CloneStyleFrom(fio_cell.CellStyle);
                            }
                            catch (Exception e) { }

                            style.SetFont(font);
                            fio_cell.CellStyle = style;
                        }
                        break;
                    }
            }
            
        }
    }

    if (!singleFile)
    {
        Directory.CreateDirectory($"{outFileName}");
        FileStream xfile = new FileStream($"{outFileName}\\{gr.Name}.xlsx", FileMode.Create, System.IO.FileAccess.Write);
        newGr.Write(xfile, false);
        xfile.Close();
    }

   
    
}

if (singleFile)
{
    FileStream xfile = new FileStream(outFileName, FileMode.Create, System.IO.FileAccess.Write);
    newGr.Write(xfile, false);
    xfile.Close();
}



//Console.ReadKey();