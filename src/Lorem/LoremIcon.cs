// ---------------------------------------------------------------------------------------------------------------
// Name:     LoremIcon.cs
// Purpose:  Draws the Lorem Ipsum button picture in code at the size asked for: a page with a folded corner
//           and lines of text. No artwork file is used and none is borrowed. Mid gray lines and a white page
//           read on light and dark Office themes, on a transparent background.
// Revision: 1
// DTG:      261025040000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace OpenTools
{
    internal static class LoremIcon
    {
        private const int TextLines = 5;
        private static readonly Color PageColor = Color.White;
        private static readonly Color InkColor = Color.FromArgb(255, 96, 96, 96);

        // ----- Drawing -----

        public static Bitmap Draw(int size)
        {
            Bitmap bitmap = new Bitmap(size, size, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bitmap))
            using (SolidBrush page = new SolidBrush(PageColor))
            using (SolidBrush ink = new SolidBrush(InkColor))
            using (Pen edge = new Pen(InkColor, size / 16f))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                float left = size * 0.16f;
                float right = size * 0.84f;
                float top = size * 0.06f;
                float bottom = size * 0.94f;
                float fold = size * 0.22f;
                PointF[] outline =
                {
                    new PointF(left, top), new PointF(right - fold, top), new PointF(right, top + fold),
                    new PointF(right, bottom), new PointF(left, bottom)
                };
                g.FillPolygon(page, outline);
                g.DrawPolygon(edge, outline);
                g.DrawLine(edge, right - fold, top, right - fold, top + fold);
                g.DrawLine(edge, right - fold, top + fold, right, top + fold);

                float lineHeight = size / 16f;
                float first = top + fold + size * 0.08f;
                float gap = (bottom - first - size * 0.06f) / TextLines;
                for (int i = 0; i < TextLines; i++)
                {
                    float width = (right - left) - size * 0.2f - (i == TextLines - 1 ? size * 0.2f : 0f);
                    g.FillRectangle(ink, left + size * 0.1f, first + i * gap, width, lineHeight);
                }
            }

            return bitmap;
        }
    }
}
