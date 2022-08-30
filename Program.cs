// See https://aka.ms/new-console-template for more information
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop;


namespace WordToExcel

{
    class Program
    { 

        static void Main(string[] args)
        {
            //Create Doc
            string docPath = @"D:\OFFICE APPLICATION TEST.docx";
            Application app = new Application();
            Document doc = app.Documents.Open(docPath);

            //Get all words
            string allWords = doc.Content.Text;
            doc.Close();
            app.Quit();
           
            var xlApp = new Microsoft.Office.Interop.Excel.Application();

            var xlWorkBook = xlApp.Workbooks.Add();
            var xlWorkSheet = xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[2, 2] = allWords;
                       
            xlWorkBook.SaveAs(Filename: @"D:\csharp-Excel.xls", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, AccessMode: Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive);
            xlWorkBook.Close(SaveChanges: true);
            xlApp.Quit();
        }

    }

    }

