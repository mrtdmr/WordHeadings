using Microsoft.Office.Interop.Word;
using System;

namespace WordHeadings
{
    class Program
    {
        private static string paragraphText = "";
        static void Main(string[] args)
        {
            Application application = new Application();
            Document document = null;
            try
            {
                document = application.Documents.Open("D:\\TAEKBGYS-006 BGYS Risk Degerlendirme Metodolojisi.docx");
                ToNextHeaderText(document.Paragraphs);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                document.Close();
            }

            Console.ReadLine();
        }
        static void ToNextHeaderText(Paragraphs paragraphs)
        {
            Style style = null;
            foreach (Paragraph paragraph in paragraphs)
            {

                paragraphText = "";
                style = paragraph.get_Style();
                if (
                    style.NameLocal == "Heading 1" ||
                    style.NameLocal == "Heading 2" ||
                    style.NameLocal == "Heading 3" ||
                    style.NameLocal == "Heading 4" ||
                    style.NameLocal == "Heading 5" ||
                    style.NameLocal == "Heading 6" ||
                    style.NameLocal == "Heading 7" ||
                    style.NameLocal == "Heading 8" ||
                    style.NameLocal == "Heading 9")
                {
                    Console.WriteLine("\n\n"+paragraph.Range.Text + "\n\n");
                }
                else
                {
                    paragraphText += paragraph.Range.Text;
                    Console.WriteLine(paragraphText);
                    continue;
                }
            }
        }
    }
}
