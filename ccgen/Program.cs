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

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);


string inputJson = "";
string outFileName = "";
string paternFile = "";

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
            Console.WriteLine();

            return;
            break;
        case "-i": inputJson = args[++i]; break;
        case "-o": outFileName = args[++i]; break;
        case "-p": paternFile = args[++i]; break;
        case "-g": singleFile = true; break;

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

    newSheet.GetRow(2).GetCell(5).SetCellValue(gr.Name);
    for (int i=0; i<gr.Students.Count; ++i)
    {
        var fio = gr.Students[i].Surname + " ";

        if (gr.Students[i].Name != "")
            fio += gr.Students[i].Name[0] + ".";

       if (gr.Students[i].Lastname != "")
            fio += gr.Students[i].Lastname[0] + ".";

        newSheet.GetRow(6 + i).GetCell(1).SetCellValue(fio);
        newSheet.GetRow(6 + i).GetCell(0).SetCellValue(i+1);  
        
        //починить коректное делание жирным
        if (gr.Students[i].IsHeadman)
        {
            
            var _st1 = newSheet.GetRow(6 + i).GetCell(0).CellStyle;
            var _st2 = newSheet.GetRow(6 + i).GetCell(1).CellStyle;

            var  s1 = newGr.CreateCellStyle();
            var  s2 = newGr.CreateCellStyle();

            s1.BorderBottom = _st1.BorderBottom;    
            s1.BorderLeft = _st1.BorderLeft;
            s1.BorderTop = _st1.BorderTop;
            s1.BorderRight = _st1.BorderRight;

            s2.BorderBottom = _st2.BorderBottom;
            s2.BorderLeft = _st2.BorderLeft;
            s2.BorderTop = _st2.BorderTop;
            s2.BorderRight = _st2.BorderRight;

            s1.SetFont(font);
            s2.SetFont(font);


            newSheet.GetRow(6 + i).GetCell(1).CellStyle = s2;
            newSheet.GetRow(6 + i).GetCell(0).CellStyle = s1;
        }
    }

    if (!singleFile)
    {
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