// See https://aka.ms/new-console-template for more information
using ConsoleApp3;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

if (args.Length < 2)
{
    Console.WriteLine("Введите: {Путь к docx файлу} и {Имя файла}");
}

var file = File.ReadAllBytes(args[0]);

DocumentCore.Serial = "Введите свой серийник";

using (var ms = new MemoryStream(file))
{
    var document = DocumentCore.Load(ms, LoadOptions.DocxDefault);
    AddSignatureStamps(document);
    document.Save($"{args[1]}.pdf", SaveOptions.PdfDefault);
}

void AddSignatureStamps(DocumentCore document)
{
    var builder = new DocumentBuilder(document);
    var sign = new Sign { SigningTime = DateTime.Now.ToLongDateString(),
        Employee = "Иванов Иван Иванович",
        OrganizationName = "ООО Рога и Копыта",
        SerialNumber = "228228",
        SignatureTimeStampTime = DateTime.Now.ToLongDateString(),
        ValidityPeriod = $"{DateTime.Now.ToLongDateString()} - {DateTime.Now.ToLongDateString()}"
    };

    AddSignatureStamp(builder, sign);
}

void AddSignatureStamp(DocumentBuilder builder, Sign sign)
{
    builder.CellFormat.Borders.SetBorders(MultipleBorderTypes.Bottom, BorderStyle.None, Color.Blue, 1);
    builder.Writeln();

    var table = builder.StartTable();

    builder.TableFormat.PreferredWidth = new TableWidth(LengthUnitConverter.Convert(6.5, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point);
    builder.TableFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Blue, 1);
    builder.TableFormat.Alignment = HorizontalAlignment.Center;

    builder.CellFormat.Padding = new Padding(0.5, 0.2, LengthUnit.Centimeter);
    builder.CellFormat.VerticalAlignment = VerticalAlignment.Center;
    builder.ParagraphFormat.Alignment = HorizontalAlignment.Left;

    builder.RowFormat.Height = new TableRowHeight(5f, HeightRule.Exact);
    builder.InsertCell();
    builder.CellFormat.PreferredWidth = new TableWidth(LengthUnitConverter.Convert(1.8, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point);
    builder.InsertCell();
    builder.EndRow();

    builder.RowFormat.Height = new TableRowHeight(18f, HeightRule.Exact);
    builder.InsertCell();
    builder.CharacterFormat.FontColor = Color.Blue;
    builder.CharacterFormat.Size = 8;
    builder.CharacterFormat.FontName = "Arial";
    builder.Write("Подписано: ");
    builder.InsertCell();
    builder.Write($"{sign.OrganizationName} {sign.Employee}");
    builder.EndRow();

    builder.RowFormat.Height = new TableRowHeight(18f, HeightRule.Exact);
    builder.InsertCell();
    builder.Write("Серийный номер: ");
    builder.InsertCell();
    builder.Write(sign.SerialNumber);
    builder.EndRow();

    builder.RowFormat.Height = new TableRowHeight(18f, HeightRule.Exact);
    builder.InsertCell();
    builder.Write("Действует:");
    builder.InsertCell();
    builder.Write(sign.ValidityPeriod);
    builder.EndRow();

    if (!string.IsNullOrEmpty(sign.SignatureTimeStampTime) || !string.IsNullOrEmpty(sign.SigningTime))
    {
        builder.CellFormat.Borders.SetBorders(MultipleBorderTypes.Top, BorderStyle.Dashed, Color.Blue, 1);
        builder.RowFormat.Height = new TableRowHeight(5f, HeightRule.Exact);
        builder.InsertCell();
        builder.InsertCell();
        builder.EndRow();

        builder.CellFormat.Borders.SetBorders(MultipleBorderTypes.Top, BorderStyle.None, Color.Blue, 1);

        if (!string.IsNullOrEmpty(sign.SigningTime))
        {
            builder.RowFormat.Height = new TableRowHeight(18f, HeightRule.Exact);
            builder.InsertCell();
            builder.Write("Время подписания:");
            builder.InsertCell();
            builder.Write(sign.SigningTime);
            builder.EndRow();
        }

        if (!string.IsNullOrEmpty(sign.SignatureTimeStampTime))
        {
            builder.RowFormat.Height = new TableRowHeight(18f, HeightRule.Exact);
            builder.InsertCell();
            builder.Write("Подпись подтверждена:");
            builder.InsertCell();
            builder.Write(sign.SignatureTimeStampTime);
            builder.EndRow();
        }
    }

    builder.CellFormat.Borders.SetBorders(MultipleBorderTypes.Bottom, BorderStyle.Single, Color.Blue, 1);
    builder.RowFormat.Height = new TableRowHeight(1f, HeightRule.Exact);
    builder.InsertCell();
    builder.InsertCell();
    builder.EndRow();

    foreach (TableCell cell in table.GetChildElements(true, ElementType.TableCell))
    {
        // Cell should have at least one paragraph.
        if (cell.Blocks.Count == 0)
        {
            cell.Blocks.Add(new Paragraph(cell.Document));
        }

        foreach (Paragraph paragraph in cell.GetChildElements(true, ElementType.Paragraph))
        {
            paragraph.ParagraphFormat.KeepWithNext = true;
        }
    }

    builder.EndTable();
}