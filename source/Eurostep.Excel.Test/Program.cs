namespace Eurostep.Excel.Test
{
    internal class Program
    {
        private Program()
        {
        }

        private static void Main(string[] args)
        {
            try
            {
                Program p = new Program();
                //p.CreateExcelWithHeader();
                p.CreateExcelWithTable();
                //p.CreateExcelWithHeaderAndTable();
            }
            catch (Exception e)
            {
                WriteLine($"{e.GetType().Name}: '{e.Message}'");
                WriteLine("Hit enter...", ConsoleColor.Yellow);
                Console.ReadLine();
            }
            finally
            {
                WriteLine("Done", ConsoleColor.Green);
            }
        }

        private static void WriteLine(string? message, ConsoleColor foregroundColor = ConsoleColor.Gray)
        {
            Console.ForegroundColor = foregroundColor;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.Gray;
        }

        //private void CreateExcelWithHeader()
        //{
        //    using var file = GetFile(@"TestWithHeader.xlsx");
        //    using var excel = SheetWriter.GetClient(file, "Test");
        //    excel.AddHeaders().New("Col 1", 10, excel.GetLightBlueHeaderStyle()).New("Col 2", 50, excel.GetLightGreenHeaderStyle()).Build();
        //    excel.AddRow().New("A").New("B").Build();
        //    excel.AddRow().New("C").New("D").Build();
        //    excel.EndSheet(false);
        //    excel.Close();
        //    file.Flush();
        //}

        private void CreateExcelWithTable()
        {
            using FileStream file = GetFile(@"TestWithTable.xlsx");
            using ISheetWriter excel = SheetWriter.GetClient(file, "Test");
            excel.AddHeaders().New("Col 1", 40).New("Col 2", 10).Build();
            excel.AddRow().New("A").New("B").Build();
            excel.AddRow().New("C").New("D").Build();
            excel.EndSheet(true);
            excel.Close();
            file.Flush();
        }

        //private void CreateExcelWithHeaderAndTable()
        //{
        //    using var file = GetFile(@"TestWithHeaderAndTable.xlsx");
        //    using var excel = SheetWriter.GetClient(file, "Test");
        //    excel.AddHeaders().New("Col 1", 20, excel.GetLightBlueHeaderStyle()).New("Col 2", 20, excel.GetLightGreenHeaderStyle()).Build();
        //    excel.AddRow().New("A").New("B").Build();
        //    excel.AddRow().New("C").New("D").Build();
        //    excel.EndSheet(true);
        //    excel.Close();
        //    file.Flush();
        //}

        private FileStream GetFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (File.Exists(filePath)) File.Delete(filePath);
            return new FileStream(filePath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read);
        }
    }
}