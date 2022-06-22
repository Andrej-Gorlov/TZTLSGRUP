// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;
using System.Collections.Concurrent;
using TZTLSGRUP;

try
{
    Helper helper = new ();
    var filename = helper.ShowDialog();
    Console.WriteLine(filename);
    var list = new ConcurrentBag<FinExample>();

    using (var package = new ExcelPackage(new FileInfo(filename)))
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
        var rowCount = worksheet.Dimension.Rows;
        for (int row = 2; row <= rowCount; row++)
        {
            if (double.Parse(worksheet.Cells[row, 13].Value.ToString().Trim()) > 100000)
            {
                FinExample finExample = new()
                {
                    Id = int.Parse(worksheet.Cells[row, 1].Value.ToString().Trim()),
                    Product = worksheet.Cells[row, 4].Text.ToString().Trim(),
                    Country = worksheet.Cells[row, 3].Text.ToString().Trim(),
                    Date = DateOnly.Parse(worksheet.Cells[row, 14].Text.ToString().Trim()),
                    Profit = decimal.Parse(worksheet.Cells[row, 13].Value.ToString().Trim())
                };
                list.Add(finExample);
            }
        }
    }
    foreach (var item in list.OrderBy(x => x.Id))
    {
        Console.WriteLine($" Id: {item.Id}\n" +
            $" Product: {item.Product}\n" +
            $" Country: {item.Country}\n" +
            $" Date:{item.Date}\n" +
            $" Profit: {item.Profit}" );
    }
    Console.WriteLine("\n \t Выберите в каком формате сохранить данные," +
        "\n \t где 1 - json, 2 - csv, 0 - выйти из программы.");
    
    var fins = list.OrderBy(x => x.Id).ToList();
    Console.WriteLine(helper.SaveFile(ref fins));
}
catch (Exception ex)
{

    Console.WriteLine(ex);
}







//[DllImport("comdlg32.dll", SetLastError = true, CharSet = CharSet.Auto)]
//static extern bool GetOpenFileName(ref OpenFileName ofn);

// static string ShowDialog()
//{
//    var ofn = new OpenFileName();
//    ofn.lStructSize = Marshal.SizeOf(ofn);
//    // Define Filter for your extensions (Excel, ...)
//    ofn.lpstrFilter = "Excel Files (*.xlsx)\0*.xlsx\0All Files (*.*)\0*.*\0";
//    ofn.lpstrFile = new string(new char[256]);
//    ofn.nMaxFile = ofn.lpstrFile.Length;
//    ofn.lpstrFileTitle = new string(new char[64]);
//    ofn.nMaxFileTitle = ofn.lpstrFileTitle.Length;
//    ofn.lpstrTitle = "Open File Dialog...";
//    if (GetOpenFileName(ref ofn))
//        return ofn.lpstrFile;
//    return string.Empty;
//}

//[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
//struct OpenFileName
//{
//    public int lStructSize;
//    public IntPtr hwndOwner;
//    public IntPtr hInstance;
//    public string lpstrFilter;
//    public string lpstrCustomFilter;
//    public int nMaxCustFilter;
//    public int nFilterIndex;
//    public string lpstrFile;
//    public int nMaxFile;
//    public string lpstrFileTitle;
//    public int nMaxFileTitle;
//    public string lpstrInitialDir;
//    public string lpstrTitle;
//    public int Flags;
//    public short nFileOffset;
//    public short nFileExtension;
//    public string lpstrDefExt;
//    public IntPtr lCustData;
//    public IntPtr lpfnHook;
//    public string lpTemplateName;
//    public IntPtr pvReserved;
//    public int dwReserved;
//    public int flagsEx;
//}