using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Word_Graf.Models;

namespace Word_Graf.Context
{
    public class OwnerContext : Owner
    {
        public OwnerContext(string FirstName, string LastName, string SurName, int NumberRoom)
            : base(FirstName, LastName, SurName, NumberRoom) { }

        public static List<OwnerContext> AllOwners()
        {
            List<OwnerContext> allOwenrs = new List<OwnerContext>();
            allOwenrs.Add(new OwnerContext("test", "test", "test", 1));
            allOwenrs.Add(new OwnerContext("test2", "test2", "test2", 2));
            allOwenrs.Add(new OwnerContext("test3", "test3", "test3", 3));
            allOwenrs.Add(new OwnerContext("test4", "test4", "test4", 4));
            allOwenrs.Add(new OwnerContext("test5", "test5", "test5", 5));
            allOwenrs.Add(new OwnerContext("test6", "test6", "test6", 6));
            allOwenrs.Add(new OwnerContext("test7", "test7", "test7", 7));
            allOwenrs.Add(new OwnerContext("test8", "test8", "test8", 8));
            allOwenrs.Add(new OwnerContext("test9", "test9", "test9", 9));
            allOwenrs.Add(new OwnerContext("test10", "test10", "test10", 10));
            allOwenrs.Add(new OwnerContext("test11", "test11", "test11", 11));
            allOwenrs.Add(new OwnerContext("test12", "test12", "test12", 12));
            allOwenrs.Add(new OwnerContext("test13", "test13", "test13", 13));
            allOwenrs.Add(new OwnerContext("test14", "test14", "test14", 14));
            allOwenrs.Add(new OwnerContext("test15", "test15", "test15", 15));
            allOwenrs.Add(new OwnerContext("test16", "test16", "test16", 16));
            allOwenrs.Add(new OwnerContext("test17", "test17", "test17", 17));
            allOwenrs.Add(new OwnerContext("test18", "test18", "test18", 18));
            allOwenrs.Add(new OwnerContext("test19", "test19", "test19", 19));
            return allOwenrs;
        }
        public static void Report(string fileName)
        {
            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add();

            Word.Paragraph paraHeader = doc.Paragraphs.Add();
            paraHeader.Range.Font.Size = 16;
            paraHeader.Range.Text = "Список жильцов дома";
            paraHeader.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paraHeader.Range.ParagraphFormat.SpaceAfter = 0;
            paraHeader.Range.Font.Bold = 1;
            paraHeader.Range.InsertParagraphAfter();

            Word.Paragraph paraAddress = doc.Paragraphs.Add();
            paraAddress.Range.Font.Size = 14;
            paraAddress.Range.Text = "по адресу: г. Пермь, ул. Луначарского, д. 24";
            paraAddress.Range.ParagraphFormat.SpaceAfter = 20;
            paraAddress.Range.Font.Bold = 0;
            paraAddress.Range.InsertParagraphAfter();

            Word.Paragraph paraCount = doc.Paragraphs.Add();
            paraCount.Range.Font.Size = 14;
            paraCount.Range.Text = $"Всего жильцов: {AllOwners().Count}";
            paraCount.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paraCount.Range.ParagraphFormat.SpaceAfter = 0;
            paraCount.Range.InsertParagraphAfter();

            Word.Paragraph tableParagraph = doc.Paragraphs.Add();
            Word.Table paymentsTable = doc.Tables.Add(tableParagraph.Range, AllOwners().Count + 1, 4);
            paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            Cell("№", paymentsTable.Cell(1, 1).Range);
            Cell("Фамилия", paymentsTable.Cell(1, 2).Range);
            Cell("Имя", paymentsTable.Cell(1, 3).Range);
            Cell("Отчество", paymentsTable.Cell(1, 4).Range);

            for (int i = 0; i < AllOwners().Count; i++)
            {
                OwnerContext owner = AllOwners()[i];

                Cell((i + 1).ToString(), paymentsTable.Cell(1 + 1 + i, 1).Range);
                Cell(owner.LastName, paymentsTable.Cell(1 + 1 + i, 2).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.FirstName, paymentsTable.Cell(1 + 1 + i, 3).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.SurName, paymentsTable.Cell(1 + 1 + i, 4).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
            }

            doc.SaveAs2(fileName);
            doc.Close();
            app.Quit();
        }
        public static void Cell(string Text, Word.Range Cell,
            Word.WdParagraphAlignment Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter)
        {
            Cell.Text = Text;
            Cell.ParagraphFormat.Alignment = Alignment;
        }
    }
}
