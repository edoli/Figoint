
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{
    public class SVGGradientParser
    {
        // TODO: support radial gradient
        private XDocument svgDocument;
        private Dictionary<string, XElement> gradientDefs;

        public SVGGradientParser(XElement svgRoot)
        {
            // gradientDefs = svgRoot.Descendants("linearGradient")
            //     .ToDictionary(e => e.Attribute("id").Value, e => e);
            gradientDefs = svgRoot.Descendants()
                .Where(e => e.Name.LocalName == "linearGradient" || e.Name.LocalName == "radialGradient")
                .ToDictionary(e => e.Attribute("id").Value, e => e);
        }

        public void ApplyGradient(PowerPoint.Shape shape, string gradientRef, SVGStyleParser styleParser)
        {
            if (string.IsNullOrEmpty(gradientRef) || !gradientRef.StartsWith("url(#"))
            {
                return;
            }
            
            string gradientId = gradientRef.Substring(5, gradientRef.Length - 6);
            if (!gradientDefs.TryGetValue(gradientId, out XElement gradientElement))
            {
                return;
            }

            var stops = ParseGradientStops(gradientElement, styleParser);
            if (stops.Count < 2)
            {
                return;
            }

            if (gradientElement.Name.LocalName == "linearGradient")
            {
                ApplyLinearGradient(shape, stops, gradientElement);
            }
            else if (gradientElement.Name.LocalName == "radialGradient")
            {
                ApplyRadialGradient(shape, stops, gradientElement);
            }
            
            shape.Fill.Visible = MsoTriState.msoTrue;

            shape.Fill.ForeColor.RGB = stops[0].Color.ToInt();
            shape.Fill.BackColor.RGB = stops[stops.Count - 1].Color.ToInt();

            for (int i = 0; i < stops.Count; i++)
            {
                var color = stops[i].Color;
                if (i == 0)
                {
                    shape.Fill.GradientStops[1].Transparency = color.Transparency;
                }
                else if (i == stops.Count - 1)
                {
                    shape.Fill.GradientStops[2].Transparency = color.Transparency;
                }
                else
                {
                    shape.Fill.GradientStops.Insert(color.ToInt(), stops[i].Position, color.Transparency);
                }
            }
        }

        
        private void ApplyLinearGradient(PowerPoint.Shape shape, List<GradientStop> stops, XElement gradientElement)
        {
            shape.Fill.TwoColorGradient(MsoGradientStyle.msoGradientHorizontal, 1);

            // Apply gradient angle
            float angle = ParseGradientAngle(gradientElement);
            shape.Fill.GradientAngle = angle;
        }

        private void ApplyRadialGradient(PowerPoint.Shape shape, List<GradientStop> stops, XElement gradientElement)
        {
            // HACK: currently set to rectangular gradient. circular gradient woule be better.
            shape.Fill.TwoColorGradient(MsoGradientStyle.msoGradientFromCenter, 1);
        }

        private float ParseAttributeValue(XElement element, string attributeName, float defaultValue)
        {
            string value = element.Attribute(attributeName)?.Value;
            if (string.IsNullOrEmpty(value))
                return defaultValue;

            // Remove % if present and parse
            value = value.TrimEnd('%');
            if (float.TryParse(value, out float result))
                return result / (value.EndsWith("%") ? 100f : 1f);

            return defaultValue;
        }

        private List<GradientStop> ParseGradientStops(XElement gradientElement, SVGStyleParser styleParser)
        {
            var stops = new List<GradientStop>();
            var stopElements = gradientElement.Elements()
                .Where(e => e.Name.LocalName == "stop")
                .Select(element => styleParser.ParseAndConvertStyle(element));

            foreach (var stopElement in stopElements)
            {
                float offset = float.Parse(stopElement.Attribute("offset").Value.TrimEnd('%')) / 100f;
                string colorString = stopElement.Attribute("stop-color")?.Value ?? "black";
                string opacityString = stopElement.Attribute("stop-opacity")?.Value ?? "1";

                MyColor color = MyColor.FromSVG(colorString, opacityString);

                stops.Add(new GradientStop { Position = offset, Color = color });
            }
            return stops.OrderBy(s => s.Position).ToList();
        }

        private float ParseGradientAngle(XElement gradientElement)
        {
            // Default to horizontal gradient (0 degrees in PowerPoint)
            float angle = 0;

            string x1 = gradientElement.Attribute("x1")?.Value ?? "0%";
            string y1 = gradientElement.Attribute("y1")?.Value ?? "0%";
            string x2 = gradientElement.Attribute("x2")?.Value ?? "100%";
            string y2 = gradientElement.Attribute("y2")?.Value ?? "0%";

            // Convert percentage to float (0-1 range)
            float x1f = ParsePercentage(x1);
            float y1f = ParsePercentage(y1);
            float x2f = ParsePercentage(x2);
            float y2f = ParsePercentage(y2);

            // Calculate angle
            float dx = x2f - x1f;
            float dy = y2f - y1f;
            angle = (float)(Math.Atan2(dy, dx) * (180 / Math.PI));

            // Adjust angle to match PowerPoint's gradient angle system
            angle = (angle + 90) % 360;

            return angle;
        }

        private float ParsePercentage(string value)
        {
            if (value.EndsWith("%"))
                return float.Parse(value.TrimEnd('%')) / 100f;
            return float.Parse(value);
        }

        private class GradientStop
        {
            public float Position { get; set; }
            public MyColor Color { get; set; }
        }
    }
}