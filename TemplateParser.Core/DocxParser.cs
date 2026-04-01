using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TemplateParser.Core;

public sealed class DocxParser
{
    public ParserResult? ParseDocxTemplate(string filePath, Guid templateId)
    {
        // 1. Open the word document in read mode.
        // 2. Parse the document into XML using the OpenXml library.
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
        {
            // A robust version that fails if the document is not structured properly:
            Body? body = wordDoc?.MainDocumentPart?.Document?.Body;
            
            // This line makes sure we don't crash if the body is missing.
            ArgumentNullException.ThrowIfNull(body, "Document is empty.");

            // 3. Loop through every paragraph.
            foreach (Paragraph p in body.Descendants<Paragraph>())
            {
                // 4. Extract and display the paragraph style.
                string style = p?.ParagraphProperties?.ParagraphStyleId?.Val ?? "No Style";
                Console.WriteLine(style);

                // 5. Extract and display the actual text.
                string? text = p?.InnerText;
                Console.WriteLine(text);

                // Spacing out the output for better readability.
                Console.WriteLine("--------------------------------");
            }
        }

        // We return null for now as per the "Code Notes" in your assignment.
        return null; 
    }
}