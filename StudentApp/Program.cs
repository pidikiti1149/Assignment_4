using System;
using SautinSoft.Document;
using System.IO;
using SautinSoft.Document.Drawing;
using ExcelLibrary.CompoundDocumentFormat;
using ExcelLibrary.SpreadSheet;
using Picture = SautinSoft.Document.Drawing.Picture;
using FTPApp.Models;
using System.Collections.Generic;

namespace StudentApp
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo docFile = new FileInfo("info.docx");
            DocumentCore docx = new DocumentCore();
            docx.Content.Start.Insert("Hello my name is<Aditya>", new CharacterFormat() { FontName = "Arial", FontColor = Color.DarkBlue, Size = 18.5 });
            docx.Save(docFile.FullName);
            AddPictures();

            AddExcel();
        }
        public static void AddPictures()
        {
            string documentPath = "info.docx";
            string pictPath = @"/Users/adityapidikiti/Downloads/myimage.jpeg"; 


            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Add a new section.
            Section s = new Section(dc);
            dc.Sections.Add(s);

            // 1. Picture with InlineLayout:

            // Create a new paragraph with picture.
            Paragraph par = new Paragraph(dc);
            s.Blocks.Add(par);
            par.ParagraphFormat.Alignment = HorizontalAlignment.Left;

            // Add some text content.
            par.Content.End.Insert("Hello my name is<Aditya> ", new CharacterFormat() { FontName = "Calibri", Size = 16.0, FontColor = Color.DarkBlue });

            // Our picture has InlineLayout - it doesn't have positioning by coordinates
            // and located as flowing content together with text (Run and other Inline elements).
            Picture pict1 = new SautinSoft.Document.Drawing.Picture(dc, InlineLayout.Inline(new Size(100, 100)), pictPath);

            // Add picture to the paragraph.
            par.Inlines.Add(pict1);

            // Add some text content.
            par.Content.End.Insert(" Picture at my friend bday party.", new CharacterFormat() { FontName = "Calibri", Size = 16.0, FontColor = Color.DarkBlue });

            // 2. Picture with FloatingLayout:
            // Floating layout means that the Picture (or Shape) is positioned by coordinates.
            Picture pict2 = new SautinSoft.Document.Drawing.Picture(dc, pictPath);
            pict2.Layout = FloatingLayout.Floating(
                new HorizontalPosition(50, LengthUnit.Millimeter, HorizontalPositionAnchor.Page),
                new VerticalPosition(70, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new Size(LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point),
                         LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point))
                         );

            // Set the wrapping style.
            (pict2.Layout as FloatingLayout).WrappingStyle = WrappingStyle.Square;

            // Add our picture into the section.
            s.Content.End.Insert(pict2.Content);

            // Save our document into DOCX format.
            dc.Save(documentPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
        public static void AddExcel()
        {
            string file = "info.xls";
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("First Sheet");
            worksheet.Cells[1, 0] = new Cell("Hello my name is <Aditya>");
            
            worksheet.Cells.ColumnWidth[1, 0] = 3000;
            workbook.Worksheets.Add(worksheet);
            workbook.Save(file);
            
            AddGuid();
            // open xls file Workbook book = Workbook.Load(file); Worksheet sheet = book.Worksheets[0];

            // traverse cells foreach (Pair, Cell> cell in sheet.Cells) { dgvCells[cell.Left.Right, cell.Left.Left].Value = cell.Right.Value; }

            // traverse rows by Index for (int rowIndex = sheet.Cells.FirstRowIndex; rowIndex <= sheet.Cells.LastRowIndex; rowIndex++) { Row row = sheet.Cells.GetRow(rowIndex); for (int colIndex = row.FirstColIndex; colIndex <= row.LastColIndex; colIndex++) { Cell cell = row.GetCell(colIndex); } } ```
        }
        public static void AddGuid()
        {
            var myguid = Guid.NewGuid();

            Student myrecord = new Student { UID = "669067a9-478f-40af-b7fd-dd4e7dd4a097", FirstName = "Aditya", LastName = "Pidikiti", StudentId = "200429757" };

            //List<Guid> guids = new List<Guid>();
            //for (int i = 0; i < 100; i++)
            //{
            //    var guid = Guid.NewGuid();
            //    guids.Add(guid);
            //}

            //foreach (var guid in guids)
            //{
            //    Console.WriteLine(guid.ToString());
            //}

            List<Student> students = new List<Student>();
            for (int i = 0; i < 100; i++)
            {
                var guid = Guid.NewGuid();
                Student student = new Student();
                student.UID = guid.ToString();
                student.FirstName = "Student";
                student.LastName = i.ToString();

                students.Add(student);

                if (i == 45)
                {
                    students.Add(myrecord);
                }
            }


            Student me = students.Find(x => x.UID == "669067a9-478f-40af-b7fd-dd4e7dd4a097");
            me.IsMe = true;

            foreach (var student in students)
            {
                Console.WriteLine(student.UID + " - " + student.ToString());
            }
        }
    }
}
   

