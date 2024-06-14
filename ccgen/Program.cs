using utils;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text;
using System.Text.Encodings.Web;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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
            Console.WriteLine("      CELL_GROUP - ячейка, куда пишется название группы");
            Console.WriteLine("      CELL_LIST - ячейка, c которой начинается нумерация студентов");
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
    json = File.ReadAllText(inputJson, Encoding.UTF8);
}



JsonSerializerOptions options = new JsonSerializerOptions
{
    Encoder = JavaScriptEncoder.Create(),
};

var grlist = JsonSerializer.Deserialize<GroupList>(json, options);

Console.WriteLine(json);

NPOI.SS.UserModel.IWorkbook wb = null;

ExcelPackage? package  = null;
ExcelWorksheet? paternSheet = null;

if (paternFile == "")
{
    package = new ExcelPackage(new FileInfo(paternFile));
    paternSheet = package.Workbook.Worksheets[0];

}
else
{
    package = new ExcelPackage(new FileInfo(paternFile));
    paternSheet = package.Workbook.Worksheets[0];
}

List<ExcelPackage?> packages = new List<ExcelPackage?>();
if (singleFile)
{
    packages.Add(new ExcelPackage());
}

foreach (Group gr in grlist.Groups)
{
    if (!singleFile)
    {
        packages.Add(new ExcelPackage());

    }

    var curSheet = packages.Last()?.Workbook.Worksheets.Add(gr.Name, paternSheet);

    if (parameters[0]!= "")
    {
        curSheet.Cells[parameters[0]].Value = gr.Name;
    }


    for (int i=0; i<gr.Students.Count; ++i)
    {

        if (parameters[1] != "")
        {
            curSheet.Cells[curSheet.Cells[parameters[1]].Start.Row + i, curSheet.Cells[parameters[1]].Start.Column].Value = i + 1;

            if (gr.Students[i].IsHeadman)
            {
                curSheet.Cells[curSheet.Cells[parameters[1]].Start.Row + i, curSheet.Cells[parameters[1]].Start.Column].Style.Font.Bold = true;
            }
        }


        if (parameters[2] != "")
        {
            switch (mode)
            {
                case 1:
                    {

                        curSheet.Cells[curSheet.Cells[parameters[2]].Start.Row + i, curSheet.Cells[parameters[2]].Start.Column + 0].Value = gr.Students[i].Surname;
                        curSheet.Cells[curSheet.Cells[parameters[2]].Start.Row + i, curSheet.Cells[parameters[2]].Start.Column + 1].Value = gr.Students[i].Name;
                        curSheet.Cells[curSheet.Cells[parameters[2]].Start.Row + i, curSheet.Cells[parameters[2]].Start.Column + 2].Value = gr.Students[i].Lastname;
                    
                        if (gr.Students[i].IsHeadman)
                        {
                            curSheet.Cells[curSheet.Cells[parameters[2]].Start.Row + i, curSheet.Cells[parameters[2]].Start.Column + 0].Style.Font.Bold = true;
                            curSheet.Cells[curSheet.Cells[parameters[2]].Start.Row + i, curSheet.Cells[parameters[2]].Start.Column + 1].Style.Font.Bold = true;
                            curSheet.Cells[curSheet.Cells[parameters[2]].Start.Row + i, curSheet.Cells[parameters[2]].Start.Column + 2].Style.Font.Bold = true;
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

                        curSheet.Cells[curSheet.Cells[parameters[2]].Start.Row + i, curSheet.Cells[parameters[2]].Start.Column + 0].Value = fio;
                        

                        if (gr.Students[i].IsHeadman)
                        {
                            curSheet.Cells[curSheet.Cells[parameters[2]].Start.Row + i, curSheet.Cells[parameters[2]].Start.Column + 0].Style.Font.Bold = true;
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

                        curSheet.Cells[curSheet.Cells[parameters[2]].Start.Row + i, curSheet.Cells[parameters[2]].Start.Column + 0].Value = fio;
                        

                        if (gr.Students[i].IsHeadman)
                        {
                            curSheet.Cells[curSheet.Cells[parameters[2]].Start.Row + i, curSheet.Cells[parameters[2]].Start.Column + 0].Style.Font.Bold = true;
                        }
                        break;
                    }
            }
            
        }
    }

    //if (!singleFile)
    //{
    //    Directory.CreateDirectory($"{outFileName}");
    //    FileStream xfile = new FileStream($"{outFileName}\\{gr.Name}.xlsx", FileMode.Create, System.IO.FileAccess.Write);
    //    newGr.Write(xfile, false);
    //    xfile.Close();
    //}
}

//if (singleFile)
//{
//    FileStream xfile = new FileStream(outFileName, FileMode.Create, System.IO.FileAccess.Write);
//    newGr.Write(xfile, false);
//    xfile.Close();
//}

Directory.CreateDirectory($"{outFileName}");
foreach (var p in packages)
{
    p.SaveAs($"{outFileName}\\{p.Workbook.Worksheets[0].Name}.xlsx");
}

//Console.ReadKey();