using System;
using System.IO;
using System.Text.RegularExpressions;
using Syncfusion.Licensing;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

class Program
{
    static void Main()
    {
        SyncfusionLicenseProvider.RegisterLicense("");

        // Paths
        string root = @"D:\";
        string inDocx = Path.Combine(root, "output", "Output.docx"); // Input: Word doc with raw links
        string outDir = Path.Combine(root, "output");
        Directory.CreateDirectory(outDir);
        string outDocx = Path.Combine(outDir, "Output_Hyperlinks.docx"); // Output: With clickable links

        // Open input document
        using (FileStream fs = File.OpenRead(inDocx))
        using (WordDocument doc = new WordDocument(fs, FormatType.Docx))
        {
            // Process each paragraph to find and replace link patterns
            ProcessDocumentForHyperlinks(doc);

            // Save modified document
            using (FileStream outStream = File.Create(outDocx))
            {
                doc.Save(outStream, FormatType.Docx);
            }
        }

        Console.WriteLine($"Output with hyperlinks saved: {outDocx}");
    }

    // Process the entire document to convert link patterns to hyperlinks
    static void ProcessDocumentForHyperlinks(WordDocument doc)
    {
        // Regex to match: (https://... - Filename.pdf)
        Regex linkRegex = new Regex(@"\[](https://[^\s(]+)\s*-\s*([^\)]+)\)");

        foreach (WSection section in doc.Sections)
        {
            foreach (WParagraph paragraph in section.Paragraphs)
            {
                // Get current paragraph text for matching
                string paraText = paragraph.Text;

                // Find all matches in this paragraph
                MatchCollection matches = linkRegex.Matches(paraText);

                if (matches.Count == 0) continue; // No links, skip

                // Clear the paragraph (remove all items/runs)
                paragraph.ChildEntities.Clear();

                // Rebuild the paragraph: non-link text + hyperlinks
                int lastIndex = 0;
                foreach (Match match in matches)
                {
                    // Add text before the match
                    string beforeText = paraText.Substring(lastIndex, match.Index - lastIndex);
                    if (!string.IsNullOrEmpty(beforeText))
                    {
                        paragraph.AppendText(beforeText);
                    }

                    // Extract URL and filename
                    string url = match.Groups[1].Value.Trim();
                    string filename = match.Groups[2].Value.Trim();

                    // Append hyperlink with filename as display text
                    IWField hyperlink = paragraph.AppendHyperlink(url, filename, HyperlinkType.WebLink);
                    // Note: Hyperlinks are automatically formatted as blue and underlined by default

                    lastIndex = match.Index + match.Length;
                }

                // Add remaining text after last match
                if (lastIndex < paraText.Length)
                {
                    string afterText = paraText.Substring(lastIndex);
                    paragraph.AppendText(afterText);
                }
            }
        }
    }
}


