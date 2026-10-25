// ---------------------------------------------------------------------------------------------------------------
// Name:     GearIcon.cs
// Purpose:  Draws the OpenTools Settings picture, a gear, in code at the size asked for. No artwork file is
//           used and none is borrowed. The gear has eight teeth and a round hole, drawn in a mid gray that
//           reads on light and dark Office themes, on a transparent background.
// Revision: 1
// DTG:      261025020000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace OpenTools
{
    internal static class GearIcon
    {
        private const int Teeth = 8;
        private const float OuterRadius = 0.48f;
        private const float RootRadius = 0.36f;
        private const float HoleRadius = 0.15f;
        private const float ToothShare = 0.45f;
        private static readonly Color GearColor = Color.FromArgb(255, 96, 96, 96);

        // ----- Drawing -----

        public static Bitmap Draw(int size)
        {
            Bitmap bitmap = new Bitmap(size, size, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bitmap))
            using (GraphicsPath path = new GraphicsPath(FillMode.Alternate))
            using (SolidBrush brush = new SolidBrush(GearColor))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                float center = size / 2f;
                path.AddPolygon(Outline(center, size));
                float hole = HoleRadius * size;
                path.AddEllipse(center - hole, center - hole, hole * 2f, hole * 2f);
                g.FillPath(brush, path);
            }

            return bitmap;
        }

        // Each tooth is a flat-topped step: up from the root circle, across the top, back down.
        private static PointF[] Outline(float center, int size)
        {
            PointF[] points = new PointF[Teeth * 4];
            double step = 2.0 * Math.PI / Teeth;
            double half = step * ToothShare / 2.0;
            float outer = OuterRadius * size;
            float root = RootRadius * size;
            for (int i = 0; i < Teeth; i++)
            {
                double middle = i * step;
                points[i * 4] = Point(center, root, middle - half * 1.4);
                points[i * 4 + 1] = Point(center, outer, middle - half);
                points[i * 4 + 2] = Point(center, outer, middle + half);
                points[i * 4 + 3] = Point(center, root, middle + half * 1.4);
            }

            return points;
        }

        private static PointF Point(float center, float radius, double angle)
        {
            return new PointF(center + radius * (float)Math.Cos(angle), center + radius * (float)Math.Sin(angle));
        }
    }
}
