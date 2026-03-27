using System;
using DocumentFormat.OpenXml.Packaging;         // Needed for WordProcessingDocument.
using DocumentFormat.OpenXml.Wordprocessing;    // Needed for all Word schema objects (Body, Paragraph, etc.)


namespace TemplateParser.Core;

public sealed class DocxParser
{
    public ParserResult ParseDocxTemplate(string filePath, Guid templateId)
    {
        // 1. Open the word document in read mode.
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
        {
            // 2. Access the main body of the document safely
            // We use the '?' operator to handle cases where the document structure might be missing
            Body? body = wordDoc?.MainDocumentPart?.Document?.Body;

            // Use the built-in check to throw an error if the body is null
            ArgumentNullException.ThrowIfNull(body, "The document body could not be loaded.");

            // 3. Loop through every paragraph to understand the structure
            // Descendants<Paragraph>() finds every paragraph regardless of where it sits in the XML Doc
            foreach (Paragraph p in body.Descendants<Paragraph>())
            {
                // 4. Extract the Style ID (e.g., Heading1, Normal)
                // We use '??' to provide a fallback value if no style is applied
                string? style = p?.ParagraphProperties?.ParagraphStyleId?.Val ?? "No Style";

                // 5. Extract the actual text content
                string? text = p?.InnerText;

                // Print the results to the console for verification (Week 1 Goal)
                Console.WriteLine($"Style: {style}");
                Console.WriteLine($"Text: {text}");
                Console.WriteLine("--------------------------------");
            }
        }

        // From the current instructions, we will return null. 
        return null;

        // TODO (Week 1-4): Implement core DOCX parsing here.
        // Recommended responsibilities for this method:
        // 1) [Week 1] Learn DOCX structure and print paragraphs from the document.
        // 2) [Week 2] Build section hierarchy using Word heading styles.
        // 3) [Week 3] Detect tables, lists, and images as structured content nodes.
        // 4) [Week 4] Add formatting heuristics for files missing heading styles.
        // 5) [Week 2-4] Create Node instances with:
        //    - Id: new Guide for each node
        //    - TemplateId: the templateId argument
        //    - ParentId: null for root nodes, set for child nodes
        //    - Type/Title/OrderIndex/MetadataJson based on parsed content
        // 6) [Week 4] Return ParserResult with Nodes in deterministic order.
        //
        // Helper guidance [Week 3-6]:
        // - YES, create helper classes if this method gets long or hard to read.
        // - Keep helpers inside TemplateParser.Core (for example, Parsing/ or Utilities/ folders).
        // - Keep this method as the high-level orchestration entry point.
        // - In Week 6, refactor large blocks from this method into focused helper classes.
        //
        // Do not place parsing logic in the CLI project; keep it in Core.

        //throw new NotImplementedException("DOCX parsing is intentionally not implemented in this starter repository.");
    }
}
