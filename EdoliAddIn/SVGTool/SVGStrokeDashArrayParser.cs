

using System;
using System.Linq;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{
    public static class SVGStrokeDashArrayParser
    {
        public static void ApplyStrokeDashArray(PowerPoint.Shape shape, string dashArray)
        {
            if (string.IsNullOrWhiteSpace(dashArray))
            {
                shape.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
                return;
            }

            var values = dashArray.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries)
                                  .Select(float.Parse)
                                  .ToArray();

            if (values.Length == 0)
            {
                shape.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
                return;
            }

            // Normalize values
            float totalLength = values.Sum();
            float[] normalizedValues = values.Select(v => v / totalLength).ToArray();

            // Choose the best matching preset dash style
            MsoLineDashStyle dashStyle = ChooseBestMatchingDashStyle(normalizedValues);
            shape.Line.DashStyle = dashStyle;
        }

        private static MsoLineDashStyle ChooseBestMatchingDashStyle(float[] normalizedValues)
        {
            if (normalizedValues.Length == 1 || (normalizedValues.Length == 2 && Math.Abs(normalizedValues[0] - normalizedValues[1]) < 0.1))
            {
                return MsoLineDashStyle.msoLineDash;
            }
            else if (normalizedValues.Length == 2)
            {
                if (normalizedValues[0] > normalizedValues[1])
                {
                    return MsoLineDashStyle.msoLineLongDash;
                }
                else
                {
                    return MsoLineDashStyle.msoLineDashDot;
                }
            }
            else if (normalizedValues.Length >= 4)
            {
                return MsoLineDashStyle.msoLineDashDotDot;
            }
            else
            {
                return MsoLineDashStyle.msoLineDash;
            }
        }
    }
}