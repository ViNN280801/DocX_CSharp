using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed;
using Xceed.Words.NET;

namespace WorkWithDOCX
{
    class Program
    {
        public static void Paragraphs()
        {
            using (var document = DocX.Create("XceedExample.docx"))
            {
                document.InsertParagraph("Formatted paragraphs").FontSize(15d).SpacingAfter(30);

                var p = document.InsertParagraph();

                p.Append("This is a simple formatted red bold paragraph")
                    .Font(new Xceed.Words.NET.Font("Times New Roman"))
                    .FontSize(25)
                    .Bold()
                    .Append(" containing a blue italic text.")
                    .Font(new Xceed.Words.NET.Font("Arial"))
                    .SpacingAfter(40);

                document.Save();
            }
        }
        static void Main(string[] args)
        {
            Paragraphs();
        }
    }
}