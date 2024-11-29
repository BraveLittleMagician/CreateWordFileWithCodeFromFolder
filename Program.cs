using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

Console.WriteLine("Это приложение получает путь к папке и затем получает тип расширения текстовых файлов. После их получение программа собирает все текстовые файлы с полученным расширение (например, \"txt\") внутри указанной папки и всех её под папок и на их основе создает один единый большой текстовый файл \"CombinedText.docx\" на рабочем столе в котором хранится объединенный текст из всех полученных фалов. \nЧтобы продолжить, следуюте указанным действиям: ");
Console.WriteLine("\nВведите путь к папке: ");
string? folderPath = Console.ReadLine();

if (string.IsNullOrEmpty(folderPath))
    ThrowTextException();

folderPath = folderPath.Trim('\n', '\t', '\r', ' ');
if (string.IsNullOrEmpty(folderPath))
    ThrowTextException();

if (!Directory.Exists(folderPath))
    throw new Exception("Указанная папка не существует!");

Console.WriteLine("\nВведите расширение текстовых файлов (без точки, просто символы (например: для файла Text.txt ввод должен быть просто txt)): ");
string? extension = Console.ReadLine();

if (string.IsNullOrEmpty(extension))
    ThrowTextException();

extension = extension.Trim('\n', '\t', '\r', ' ');
if (string.IsNullOrEmpty(extension))
    ThrowTextException();

foreach (var c in extension)
{
    if (!char.IsLetterOrDigit(c))
        ThrowTextException();
}

string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
string wordFilePath = Path.Combine(desktopPath, "CombinedText.docx");
CreateWordDocument(wordFilePath, folderPath, extension);
Console.ForegroundColor = ConsoleColor.Green;
Console.WriteLine($"Файл {wordFilePath} создан успешно.");
Console.ForegroundColor = ConsoleColor.White;

static void ThrowTextException() => throw new Exception("Полученный текст пуст либо некорректен!"); 
static void CreateWordDocument(string wordFilePath, string folderPath, string extension)
{
    Console.WriteLine("\nНачато создание файла. Если файлов в папке много, то на это может потребоваться некоторое время. Ожидайте...");
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(wordFilePath, WordprocessingDocumentType.Document))
    {
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
        mainPart.Document = new Document();
        Body body = new Body();

        SectionProperties sectionProperties = new SectionProperties();
        PageMargin pageMargin = new PageMargin()
        {
            Top = 0,
            Right = (UInt32Value)0U,
            Bottom = 0,
            Left = (UInt32Value)0U,
            Header = (UInt32Value)0U,
            Footer = (UInt32Value)0U,
            Gutter = (UInt32Value)0U
        };
        sectionProperties.Append(pageMargin);
        body.Append(sectionProperties);

        var csFiles = Directory.GetFiles(folderPath, $"*.{extension}", SearchOption.AllDirectories);
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
        string sanitizedLine = SanitizeText(line);
        if (string.IsNullOrWhiteSpace(sanitizedLine)) continue;
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
static string SanitizeText(string input)
{
    if (string.IsNullOrEmpty(input))
        return string.Empty;

    return new string(input.Where(c => XmlConvert.IsXmlChar(c)).ToArray());
}