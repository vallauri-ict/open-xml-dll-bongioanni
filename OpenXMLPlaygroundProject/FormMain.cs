using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using System.Diagnostics;

namespace OpenXMLPlaygroundProject
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void btnSimpleWordText_Click(object sender, EventArgs e)
        {
            try
            {
                string filepath = "test.docx";
                string msg = "Hello, World!";

                using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
                {
                    // Add a main document part. 
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());

                    // String msg contains the text, "Hello, Word!"
                    run.AppendChild(new Text(msg));

                    // Add your paragraph to docx body
                    body.AppendChild(createParagraph());

                    // Add your paragraph to docx body
                    body.AppendChild(createJustification());

                    // Add your heading
                    addHeading1Style(mainPart);
                    body.AppendChild(createHeading("Ciao"));

                    // Add your table
                    Table myTable = createTable();
                    body.Append(myTable);  /// Con le tabelle bisogna fare append perchè non è un oggetto singolo

                    // Open file
                    Process.Start(filepath);
                }
                MessageBox.Show("Il documento è stato creato");
            }
            catch (Exception)
            {
                MessageBox.Show("Chiudere il documento e riprovare", "ATTENZIONE", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Paragraph createParagraph()
        {
            Paragraph p = new Paragraph();
            // Run 1
            Run r1 = new Run();
            Text t1 = new Text("Pellentesque ") { Space = SpaceProcessingModeValues.Preserve };
            // The Space attribute preserve white space before and after your text
            r1.Append(t1);
            p.Append(r1);

            // Run 2 - Bold
            Run r2 = new Run();
            RunProperties rp2 = new RunProperties();
            rp2.Bold = new Bold();
            // Always add properties first
            r2.Append(rp2);
            Text t2 = new Text("commodo ") { Space = SpaceProcessingModeValues.Preserve };
            r2.Append(t2);
            p.Append(r2);

            // Run 3
            Run r3 = new Run();
            Text t3 = new Text("rhoncus ") { Space = SpaceProcessingModeValues.Preserve };
            r3.Append(t3);
            p.Append(r3);

            // Run 4 – Italic
            Run r4 = new Run();
            RunProperties rp4 = new RunProperties();
            rp4.Italic = new Italic();
            // Always add properties first
            r4.Append(rp4);
            Text t4 = new Text("mauris") { Space = SpaceProcessingModeValues.Preserve };
            r4.Append(t4);
            p.Append(r4);

            // Run 5
            Run r5 = new Run();
            Text t5 = new Text(", sit ") { Space = SpaceProcessingModeValues.Preserve };
            r5.Append(t5);
            p.Append(r5);

            // Run 6 – Italic , bold and underlined
            Run r6 = new Run();
            RunProperties rp6 = new RunProperties();
            rp6.Italic = new Italic();
            rp6.Bold = new Bold();
            rp6.Underline = new Underline() { Val = UnderlineValues.WavyDouble };
            // Always add properties first
            r6.Append(rp6);
            Text t6 = new Text("amet ") { Space = SpaceProcessingModeValues.Preserve };
            r6.Append(t6);
            p.Append(r6);

            // Run 7
            Run r7 = new Run();
            Text t7 = new Text("faucibus arcu ") { Space = SpaceProcessingModeValues.Preserve };
            r7.Append(t7);
            p.Append(r7);

            // Run 8 – Red color
            Run r8 = new Run();
            RunProperties rp8 = new RunProperties();
            rp8.Color = new Color() { Val = "FF0000" };
            // Always add properties first
            r8.Append(rp8);
            Text t8 = new Text("porttitor ") { Space = SpaceProcessingModeValues.Preserve };
            r8.Append(t8);
            p.Append(r8);

            // Run 9
            Run r9 = new Run();
            Text t9 = new Text("pharetra. Maecenas quis erat quis eros iaculis placerat ut at mauris.") { Space = SpaceProcessingModeValues.Preserve };
            r9.Append(t9);
            p.Append(r9);
            // Return paragraph
            return p;
        }

        private Paragraph createJustification()
        {
            Paragraph p = new Paragraph();
            // Set the paragraph properties
            ParagraphProperties pp = new ParagraphProperties();
            pp.Justification = new Justification() { Val = JustificationValues.Center };
            // Add paragraph properties to your paragraph
            p.Append(pp);
            // Run
            Run r = new Run();
            Text t = new Text("Nam eu tortor ut mi euismod eleifend in ut ante. Donec a ligula ante. Sed rutrum ex quam. Nunc id mi ultricies, vestibulum sapien vel, posuere dui.") { Space = SpaceProcessingModeValues.Preserve };
            r.Append(t);
            p.Append(r);
            // Return paragraph
            return p;
        }

        private void addHeading1Style(MainDocumentPart mainPart)
        {
            // we have to set the properties
            RunProperties rPr = new RunProperties();
            Color color = new Color() { Val = "FF0000" }; // the color is red
            RunFonts rFont = new RunFonts();
            rFont.Ascii = "Arial"; // the font is Arial
            rPr.Append(color);
            rPr.Append(rFont);
            rPr.Append(new Bold()); // it is Bold
            rPr.Append(new FontSize() { Val = "28" }); //font size (in 1/72 of an inch)

            Style style = new Style();
            style.StyleId = "MyHeading1"; //this is the ID of the style
            style.Append(new Name() { Val = "My Heading 1" }); //this is the name of the new style
            style.Append(rPr); //we are adding properties previously defined

            // we have to add style that we have created to the StylePart
            StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylePart.Styles = new Styles();
            stylePart.Styles.Append(style);
            stylePart.Styles.Save(); // we save the style part
        }

        private Paragraph createHeading(string content)
        {
            Paragraph heading = new Paragraph();
            Run r = new Run();
            Text t = new Text(content);
            ParagraphProperties pp = new ParagraphProperties();
            // we set the style
            pp.ParagraphStyleId = new ParagraphStyleId() { Val = "MyHeading1" };
            // we set the alignement
            pp.Justification = new Justification() { Val = JustificationValues.Center };
            heading.Append(pp);
            r.Append(t);
            heading.Append(r);
            return heading;
        }

        private Table createTable()
        {
            Table table = new Table();
            // Set table properties
            table.AppendChild(getTableProperties());

            // Row 1
            TableRow tr1 = new TableRow();

            TableCell tc11 = new TableCell();
            Paragraph p11 = new Paragraph(new Run(new Text("A")));
            tc11.Append(p11);
            tr1.Append(tc11);

            TableCell tc12 = new TableCell();
            Paragraph p12 = new Paragraph();
            Run r12 = new Run();
            RunProperties rp12 = new RunProperties();
            rp12.Bold = new Bold();
            r12.Append(rp12);
            r12.Append(new Text("Nice"));
            p12.Append(r12);
            tc12.Append(p12);

            tr1.Append(tc12);
            table.Append(tr1);

            // Row 2
            TableRow tr2 = new TableRow();


            TableCell tc21 = new TableCell();
            Paragraph p21 = new Paragraph(new Run(new Text("Little")));
            tc21.Append(p21);
            tr2.Append(tc21);

            TableCell tc22 = new TableCell();
            Paragraph p22 = new Paragraph();
            ParagraphProperties pp22 = new ParagraphProperties();
            pp22.Justification = new Justification() { Val = JustificationValues.Center };
            p22.Append(pp22);
            p22.Append(new Run(new Text("Table")));
            tc22.Append(p22);

            tr2.Append(tc22);
            table.Append(tr2);

            // Add your table to docx body
            return table;
        }

        private TableProperties getTableProperties()
        {
            TableProperties tblProps = new TableProperties();

            return tblProps;
        }
    }
}
