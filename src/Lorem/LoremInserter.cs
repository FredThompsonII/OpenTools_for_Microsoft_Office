// ---------------------------------------------------------------------------------------------------------------
// Name:     LoremInserter.cs
// Purpose:  Puts Lorem Ipsum text into the active document, using the Lorem Ipsum settings.
//           Word: types at the cursor. Each section is an optional Heading 1 title, then a random number of
//           paragraphs between the low and high counts, each with the set number of sentences.
//           Excel: types one paragraph, with the set number of sentences, into the active cell.
//           PowerPoint: fills each empty text placeholder on the current slide. A title placeholder gets a
//           title, a subtitle gets one sentence, and any other text placeholder gets one sentence per paragraph,
//           with a random paragraph count between the low and high counts. Placeholders holding a chart, table,
//           SmartArt or picture have no text frame and are skipped, as are footers, dates and slide numbers.
//           Office objects are released in finally blocks.
// Revision: 1
// DTG:      261025030000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace OpenTools
{
    internal static class LoremInserter
    {
        public static void Insert(object application, Settings settings)
        {
            LoremText lorem = new LoremText();
            Word.Application word = application as Word.Application;
            if (word != null)
            {
                InsertWord(word, settings, lorem);
                return;
            }

            Excel.Application excel = application as Excel.Application;
            if (excel != null)
            {
                InsertExcel(excel, settings, lorem);
                return;
            }

            PowerPoint.Application powerPoint = application as PowerPoint.Application;
            if (powerPoint != null)
            {
                InsertPowerPoint(powerPoint, settings, lorem);
            }
        }

        // ----- Word -----

        private static void InsertWord(Word.Application word, Settings settings, LoremText lorem)
        {
            Word.Selection selection = null;
            try
            {
                selection = word.Selection;
                int sections = Math.Max(1, settings.GetInt(SettingNames.LoremSections));
                int sentences = Math.Max(1, settings.GetInt(SettingNames.LoremSentences));
                bool titles = settings.GetBool(SettingNames.LoremTitles);
                for (int section = 0; section < sections; section++)
                {
                    if (titles)
                    {
                        object heading = Word.WdBuiltinStyle.wdStyleHeading1;
                        selection.set_Style(ref heading);
                        selection.TypeText(lorem.Title());
                        selection.TypeParagraph();
                        object normal = Word.WdBuiltinStyle.wdStyleNormal;
                        selection.set_Style(ref normal);
                    }

                    int paragraphs = lorem.Between(settings.GetInt(SettingNames.LoremParagraphsLow),
                        settings.GetInt(SettingNames.LoremParagraphsHigh));
                    for (int i = 0; i < Math.Max(1, paragraphs); i++)
                    {
                        selection.TypeText(lorem.Paragraph(sentences));
                        selection.TypeParagraph();
                    }
                }
            }
            finally
            {
                Release(selection);
            }
        }

        // ----- Excel -----

        private static void InsertExcel(Excel.Application excel, Settings settings, LoremText lorem)
        {
            Excel.Range cell = null;
            try
            {
                cell = excel.ActiveCell;
                if (cell == null)
                {
                    return;
                }

                cell.Value2 = lorem.Paragraph(Math.Max(1, settings.GetInt(SettingNames.LoremSentences)));
            }
            finally
            {
                Release(cell);
            }
        }

        // ----- PowerPoint -----

        private static void InsertPowerPoint(PowerPoint.Application powerPoint, Settings settings, LoremText lorem)
        {
            PowerPoint.DocumentWindow window = null;
            PowerPoint.View view = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shapes shapes = null;
            try
            {
                window = powerPoint.ActiveWindow;
                view = window.View;
                slide = view.Slide as PowerPoint.Slide;
                if (slide == null)
                {
                    return;
                }

                shapes = slide.Shapes;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = shapes[i];
                        FillPlaceholder(shape, settings, lorem);
                    }
                    finally
                    {
                        Release(shape);
                    }
                }
            }
            finally
            {
                Release(shapes);
                Release(slide);
                Release(view);
                Release(window);
            }
        }

        private static void FillPlaceholder(PowerPoint.Shape shape, Settings settings, LoremText lorem)
        {
            if (shape.Type != Office.MsoShapeType.msoPlaceholder || shape.HasTextFrame != Office.MsoTriState.msoTrue)
            {
                return;
            }

            PowerPoint.PlaceholderFormat placeholder = null;
            PowerPoint.TextFrame frame = null;
            PowerPoint.TextRange range = null;
            try
            {
                placeholder = shape.PlaceholderFormat;
                PowerPoint.PpPlaceholderType kind = placeholder.Type;
                if (kind == PowerPoint.PpPlaceholderType.ppPlaceholderFooter
                    || kind == PowerPoint.PpPlaceholderType.ppPlaceholderDate
                    || kind == PowerPoint.PpPlaceholderType.ppPlaceholderSlideNumber
                    || kind == PowerPoint.PpPlaceholderType.ppPlaceholderHeader)
                {
                    return;
                }

                frame = shape.TextFrame;
                if (frame.HasText == Office.MsoTriState.msoTrue)
                {
                    return;
                }

                range = frame.TextRange;
                range.Text = TextFor(kind, settings, lorem);
            }
            finally
            {
                Release(range);
                Release(frame);
                Release(placeholder);
            }
        }

        private static string TextFor(PowerPoint.PpPlaceholderType kind, Settings settings, LoremText lorem)
        {
            if (kind == PowerPoint.PpPlaceholderType.ppPlaceholderTitle
                || kind == PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle
                || kind == PowerPoint.PpPlaceholderType.ppPlaceholderVerticalTitle)
            {
                return lorem.Title();
            }

            if (kind == PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle)
            {
                return lorem.Sentence();
            }

            int paragraphs = lorem.Between(settings.GetInt(SettingNames.LoremParagraphsLow),
                settings.GetInt(SettingNames.LoremParagraphsHigh));
            StringBuilder text = new StringBuilder();
            for (int i = 0; i < Math.Max(1, paragraphs); i++)
            {
                if (i > 0)
                {
                    text.Append('\r');
                }

                text.Append(lorem.Sentence());
            }

            return text.ToString();
        }

        private static void Release(object comObject)
        {
            try
            {
                if (comObject != null && Marshal.IsComObject(comObject))
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("LoremInserter.Release", ex);
            }
        }
    }
}
