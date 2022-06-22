using CsvHelper;
using CsvHelper.Configuration;
using Newtonsoft.Json;
using System.Globalization;
using System.Runtime.InteropServices;

namespace TZTLSGRUP
{
    internal class Helper
    {
        internal string SaveFile<T>(ref List<T> fins)
        {
            try
            {
                while (true)
                {
                    string? value = Console.ReadLine();
                    if (value is "")
                    {
                        Console.WriteLine("\n Пожалуйста, выберите в каком формате сохранить результат json или csv. \n");
                    }
                    else if (int.Parse(value) is 0)
                    {
                        return "Вы вышли!";
                    }
                    else if (int.Parse(value) is 1)
                    {
                        string path = @"JsonFile";
                        DirectoryInfo dirInfo = new DirectoryInfo(path);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        string json = JsonConvert.SerializeObject(fins.ToArray());
                        File.WriteAllText($@"JsonFile\FinExample {DateTime.Now.ToString("M/d/yyyy")}", json);
                        return "Файл сохранен в формате json";
                    }
                    else if (int.Parse(value) is 2)
                    {
                        string path = @"CsvFile";
                        DirectoryInfo dirInfo = new DirectoryInfo(path);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        var cfg = new CsvConfiguration(CultureInfo.InvariantCulture);
                        cfg.Delimiter = ";"; 
                        using (var sw = new StreamWriter($@"CsvFile\FinExample {DateTime.Now.ToString("M / d / yyyy")}.csv"))
                        using (var csv = new CsvWriter(sw, cfg))
                        {
                            csv.WriteRecords(fins);
                        }
                        return "Файл сохранен в формате csv";
                    }
                }
            }
            catch (Exception ex)
            {
                return $"Некорректный ввод значения.\n" +
                    $"Дополнительная информация об ошибке: {ex}";
            }
        }

        [DllImport("comdlg32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern bool GetOpenFileName(ref OpenFileName ofn);
        internal string ShowDialog()
        {
            var ofn = new OpenFileName();
            ofn.lStructSize = Marshal.SizeOf(ofn);
            // Define Filter for your extensions (Excel, ...)
            ofn.lpstrFilter = "Excel Files (*.xlsx)\0*.xlsx\0All Files (*.*)\0*.*\0";
            ofn.lpstrFile = new string(new char[256]);
            ofn.nMaxFile = ofn.lpstrFile.Length;
            ofn.lpstrFileTitle = new string(new char[64]);
            ofn.nMaxFileTitle = ofn.lpstrFileTitle.Length;
            ofn.lpstrTitle = "Open File Dialog...";
            if (GetOpenFileName(ref ofn))
                return ofn.lpstrFile;
            return string.Empty;
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct OpenFileName
        {
            public int lStructSize;
            public IntPtr hwndOwner;
            public IntPtr hInstance;
            public string lpstrFilter;
            public string lpstrCustomFilter;
            public int nMaxCustFilter;
            public int nFilterIndex;
            public string lpstrFile;
            public int nMaxFile;
            public string lpstrFileTitle;
            public int nMaxFileTitle;
            public string lpstrInitialDir;
            public string lpstrTitle;
            public int Flags;
            public short nFileOffset;
            public short nFileExtension;
            public string lpstrDefExt;
            public IntPtr lCustData;
            public IntPtr lpfnHook;
            public string lpTemplateName;
            public IntPtr pvReserved;
            public int dwReserved;
            public int flagsEx;
        }
    }
   
}
