using System;

namespace ExcelSample
{
    using Excel = Microsoft.Office.Interop.Excel;

    class Program
    {
        static void Main(string[] args)
        {
            //abrindo uma instância da aplicação
            var application = new Excel.Application();

            //adicionando um workbook
            var workBook = application.Workbooks.Add();

            //criando uma worksheet
            var primeiraWorkSheet = workBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            primeiraWorkSheet.Name = "Benefícios";

            var segundaWorkSheet = workBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            segundaWorkSheet.Name = "Comissão";

            Console.WriteLine("Tudo ok");
            Console.ReadKey();
        }
    }
}
