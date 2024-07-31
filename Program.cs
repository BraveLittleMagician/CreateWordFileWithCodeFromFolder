using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

Console.WriteLine("Введите путь к папке:");
string? folderPath = Console.ReadLine();

if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
{
    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
    string wordFilePath = Path.Combine(desktopPath, "CombinedScripts.docx");
    CreateWordDocument(wordFilePath, folderPath);
    Console.WriteLine($"Файл {wordFilePath} создан успешно.");
}
else
    Console.WriteLine("Указанная папка не существует.");

static void CreateWordDocument(string wordFilePath, string folderPath)
{
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(wordFilePath, WordprocessingDocumentType.Document))
    {
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
        mainPart.Document = new Document();
        Body body = new Body();

        SectionProperties sectionProperties = new SectionProperties();
        PageMargin pageMargin = new PageMargin()
        {
            Top    = 0,
            Right  = (UInt32Value)0U,
            Bottom = 0,
            Left   = (UInt32Value)0U,
            Header = (UInt32Value)0U,
            Footer = (UInt32Value)0U,
            Gutter = (UInt32Value)0U
        }; 
        sectionProperties.Append(pageMargin);
        body.Append(sectionProperties);

        var csFiles = Directory.GetFiles(folderPath, "*.cs", SearchOption.AllDirectories);
        foreach (var file in csFiles)
        {
            string fileContent = File.ReadAllText(file);
            AppendFormattedText(body, fileContent);

            Paragraph pageBreak = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
            body.AppendChild(pageBreak);
        }

        mainPart.Document.Append(body);
        mainPart.Document.Save();
    }
}

static void AppendFormattedText(Body body, string text)
{
    var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
    foreach (var line in lines)
    {
        Run run = new Run();
        RunProperties runProperties = new RunProperties();

        RunFonts runFonts = new RunFonts { Ascii = "Cascadia Mono SemiLight" };
        FontSize fontSize = new FontSize() { Val = "16" };
        runProperties.Append(runFonts); 
        runProperties.Append(fontSize);

        Text t = new Text(line) { Space = SpaceProcessingModeValues.Preserve };
        run.Append(runProperties);
        run.Append(t);

        Paragraph para = new Paragraph(run);
        body.AppendChild(para);
    }
}