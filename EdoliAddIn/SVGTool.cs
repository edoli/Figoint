using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Numerics;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{

    public class SVGtoPPTParser
    {
        private PowerPoint.Slide slide;
        private ShapeInfoQuadTree quadTree;

        public static void AddSVGFigureFromClipboard()
        {
            var svgCode = Clipboard.GetText();

            AddSVGFigure(svgCode);
        }

        public static void AddSVGFigure(string svgCode)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var parser = new SVGtoPPTParser(slide);
            parser.ParseAndDraw(svgCode);
        }

        public SVGtoPPTParser(PowerPoint.Slide slide)
        {
            this.slide = slide;
            this.quadTree = new ShapeInfoQuadTree();
        }

        public void ParseAndDraw(string svgCode)
        {
            svgCode = svgCode.Replace("&", "&amp;");
            XElement svg = XElement.Parse(svgCode);
            foreach (XElement element in svg.Elements())
            {
                ShapeInfo shapeInfo = null;
                switch (element.Name.LocalName.ToLower())
                {
                    case "rect":
                        shapeInfo = DrawRectangle(element);
                        break;
                    case "circle":
                        shapeInfo = DrawCircle(element);
                        break;
                    case "line":
                        shapeInfo = DrawLine(element);
                        break;
                    case "path":
                        shapeInfo = DrawPath(element);
                        break;
                    case "text":
                        shapeInfo = DrawText(element);
                        break;
                }

                if (shapeInfo != null)
                {
                    ApplyBasicStyles(shapeInfo.Shape, element);
                    quadTree.Insert(shapeInfo);
                }
            }
        }

        private ShapeInfo DrawRectangle(XElement rect)
        {
            float x = float.Parse(rect.Attribute("x")?.Value ?? "0");
            float y = float.Parse(rect.Attribute("y")?.Value ?? "0");
            float width = float.Parse(rect.Attribute("width").Value);
            float height = float.Parse(rect.Attribute("height").Value);

            var shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, width, height);
            return new ShapeInfo { Shape = shape, CenterX = x + width / 2, CenterY = y + height / 2 };
        }

        private ShapeInfo DrawCircle(XElement circle)
        {
            float cx = float.Parse(circle.Attribute("cx").Value);
            float cy = float.Parse(circle.Attribute("cy").Value);
            float r = float.Parse(circle.Attribute("r").Value);

            var shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, cx - r, cy - r, r * 2, r * 2);
            return new ShapeInfo { Shape = shape, CenterX = cx, CenterY = cy };
        }

        private ShapeInfo DrawLine(XElement line)
        {
            float x1 = float.Parse(line.Attribute("x1").Value);
            float y1 = float.Parse(line.Attribute("y1").Value);
            float x2 = float.Parse(line.Attribute("x2").Value);
            float y2 = float.Parse(line.Attribute("y2").Value);

            PowerPoint.Shape shape = slide.Shapes.AddLine(x1, y1, x2, y2);
            return new ShapeInfo { Shape = shape, CenterX = (x1 + x2) / 2, CenterY = (y1 + y2) / 2 };
        }

        private ShapeInfo DrawPath(XElement pathElement)
        {
            string d = pathElement.Attribute("d").Value;
            List<Vector2> points = ParsePathToPoints(d);
            int numPoints = points.Count;

            if (numPoints < 2) return null; // 최소 2개의 점이 필요합니다.

            // PowerPoint의 좌표계에 맞게 점들을 변환합니다.
            float minX = float.MaxValue, minY = float.MaxValue;
            float maxX = float.MinValue, maxY = float.MinValue;
            float sumX = 0, sumY = 0;

            foreach (var point in points)
            {
                minX = Math.Min(minX, point.X);
                minY = Math.Min(minY, point.Y);
                maxX = Math.Max(maxX, point.X);
                maxY = Math.Max(maxY, point.Y);

                sumX += point.X;
                sumY += point.Y;
            }

            float width = maxX - minX;
            float height = maxY - minY;
            float centerX = sumX / numPoints;
            float centerY = sumY / numPoints;

            // Polyline 생성
            float[,] pointsArray = new float[numPoints, 2];
            for (int i = 0; i < numPoints; i++)
            {
                pointsArray[i, 0] = points[i].X - minX;
                pointsArray[i, 1] = points[i].Y - minY;
            }

            // Polyline 생성
            PowerPoint.Shape shape = slide.Shapes.AddPolyline(pointsArray);

            // 위치와 크기 조정
            shape.Left = minX;
            shape.Top = minY;
            shape.Width = width;
            shape.Height = height;

            ApplyMarkers(shape, pathElement);            
            return new ShapeInfo { Shape = shape, CenterX = centerX, CenterY = centerY };

        }

        private void ApplyMarkers(PowerPoint.Shape shape, XElement pathElement)
        {
            string markerStart = pathElement.Attribute("marker-start")?.Value;
            string markerEnd = pathElement.Attribute("marker-end")?.Value;

            if (!string.IsNullOrEmpty(markerStart) && markerStart.Contains("arrow"))
            {
                shape.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
            }

            if (!string.IsNullOrEmpty(markerEnd) && markerEnd.Contains("arrow"))
            {
                shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
            }
        }

        private List<Vector2> ParsePathToPoints(string d)
        {
            List<Vector2> points = new List<Vector2>();
            float currentX = 0, currentY = 0;

            string[] commands = Regex.Split(d, @"(?=[MmLlHhVvCcSsQqTtAaZz])");

            foreach (string cmd in commands)
            {
                if (string.IsNullOrWhiteSpace(cmd)) continue;

                char command = cmd[0];
                string[] parameters = cmd.Substring(1).Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);

                switch (command)
                {
                    case 'M':
                    case 'm':
                        MoveTo(ref currentX, ref currentY, parameters, command == 'm', points);
                        break;
                    case 'L':
                    case 'l':
                        LineTo(ref currentX, ref currentY, parameters, command == 'l', points);
                        break;
                    case 'H':
                    case 'h':
                        HorizontalLineTo(ref currentX, ref currentY, parameters, command == 'h', points);
                        break;
                    case 'V':
                    case 'v':
                        VerticalLineTo(ref currentX, ref currentY, parameters, command == 'v', points);
                        break;
                    case 'Z':
                    case 'z':
                        // 시작점으로 돌아가는 것은 Polyline에서 자동으로 처리되지 않으므로,
                        // 필요하다면 시작점을 다시 추가할 수 있습니다.
                        if (points.Count > 0)
                            points.Add(points[0]);
                        break;
                        // 여기에 곡선 명령어 (C, c, S, s, Q, q, T, t, A, a)를 추가할 수 있습니다.
                }
            }

            return points;
        }

        private void MoveTo(ref float currentX, ref float currentY, string[] parameters, bool isRelative, List<Vector2> points)
        {
            float x = float.Parse(parameters[0]);
            float y = float.Parse(parameters[1]);

            if (isRelative)
            {
                currentX += x;
                currentY += y;
            }
            else
            {
                currentX = x;
                currentY = y;
            }

            points.Add(new Vector2(currentX, currentY));
        }

        private void LineTo(ref float currentX, ref float currentY, string[] parameters, bool isRelative, List<Vector2> points)
        {
            float x = float.Parse(parameters[0]);
            float y = float.Parse(parameters[1]);

            if (isRelative)
            {
                currentX += x;
                currentY += y;
            }
            else
            {
                currentX = x;
                currentY = y;
            }

            points.Add(new Vector2(currentX, currentY));
        }

        private void HorizontalLineTo(ref float currentX, ref float currentY, string[] parameters, bool isRelative, List<Vector2> points)
        {
            float x = float.Parse(parameters[0]);

            if (isRelative)
            {
                currentX += x;
            }
            else
            {
                currentX = x;
            }

            points.Add(new Vector2(currentX, currentY));
        }

        private void VerticalLineTo(ref float currentX, ref float currentY, string[] parameters, bool isRelative, List<Vector2> points)
        {
            float y = float.Parse(parameters[0]);

            if (isRelative)
            {
                currentY += y;
            }
            else
            {
                currentY = y;
            }

            points.Add(new Vector2(currentX, currentY));
        }

        private void ApplyBasicStyles(PowerPoint.Shape shape, XElement element)
        {
            string stroke = element.Attribute("stroke")?.Value;
            if (!string.IsNullOrEmpty(stroke))
            {
                shape.Line.ForeColor.RGB = ColorTranslator.FromHtml(stroke).ToRGB();
            }

            string fill = element.Attribute("fill")?.Value;
            if (!string.IsNullOrEmpty(fill))
            {
                shape.Fill.ForeColor.RGB = ColorTranslator.FromHtml(fill).ToRGB();
            }
            else
            {
                shape.Fill.Visible = MsoTriState.msoFalse;
            }

            string strokeWidth = element.Attribute("stroke-width")?.Value;
            if (!string.IsNullOrEmpty(strokeWidth))
            {
                shape.Line.Weight = float.Parse(strokeWidth);
            }
        }

        private ShapeInfo DrawText(XElement textElement)
        {
            float x = float.Parse(textElement.Attribute("x")?.Value ?? "0");
            float y = float.Parse(textElement.Attribute("y")?.Value ?? "0");
            string text = textElement.Value;

            ShapeInfo nearestShape = quadTree.FindNearest(x, y);

            if (nearestShape != null && IsCloseEnough(x, y, nearestShape))
            {
                // 도형 내부에 텍스트 추가
                nearestShape.Shape.TextFrame.TextRange.Text = text;
                PowerPoint.TextFrame2 textFrame = nearestShape.Shape.TextFrame2;
                textFrame.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                textFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                ApplyTextStyles(textFrame, textElement);
                return null;
            }
            else
            {
                PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal, x, y, 1, 1);

                var textFrame = textBox.TextFrame2;
                textBox.TextFrame2.TextRange.Text = text;

                textFrame.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
                textFrame.WordWrap = MsoTriState.msoFalse;
                AdjustTextPosition(textBox, x, y, textElement.Attribute("text-anchor")?.Value);
                ApplyTextStyles(textFrame, textElement);
                return new ShapeInfo { Shape = textBox, CenterX = x, CenterY = y };
            }
        }


        private void ApplyTextStyles(PowerPoint.TextFrame2 textFrame, XElement textElement)
        {
            var textRange = textFrame.TextRange;

            string fontFamily = textElement.Attribute("font-family")?.Value;
            if (!string.IsNullOrEmpty(fontFamily))
            {
                textRange.Font.Name = fontFamily.Trim('\'', '"');
            }

            string fontSize = textElement.Attribute("font-size")?.Value;
            if (!string.IsNullOrEmpty(fontSize))
            {
                if (fontSize.EndsWith("px"))
                {
                    fontSize = fontSize.Substring(0, fontSize.Length - 2);
                }
                textRange.Font.Size = float.Parse(fontSize);
            }

            string fill = textElement.Attribute("fill")?.Value;
            if (!string.IsNullOrEmpty(fill))
            {
                textRange.Font.Fill.ForeColor.RGB = ColorTranslator.FromHtml(fill).ToRGB();
            }

            string fontWeight = textElement.Attribute("font-weight")?.Value;
            if (!string.IsNullOrEmpty(fontWeight))
            {
                textRange.Font.Bold = fontWeight == "bold" ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }

            string fontStyle = textElement.Attribute("font-style")?.Value;
            if (!string.IsNullOrEmpty(fontStyle))
            {
                textRange.Font.Italic = fontStyle == "italic" ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }

            string textDecoration = textElement.Attribute("text-decoration")?.Value;
            if (!string.IsNullOrEmpty(textDecoration))
            {
                if (textDecoration.Contains("underline"))
                {
                    textRange.Font.UnderlineStyle = MsoTextUnderlineType.msoUnderlineSingleLine;
                }
                else
                {
                    textRange.Font.UnderlineStyle = MsoTextUnderlineType.msoNoUnderline;
                }
            }

            string textAnchor = textElement.Attribute("text-anchor")?.Value;
            if (!string.IsNullOrEmpty(textAnchor))
            {
                switch (textAnchor)
                {
                    case "start":
                        textRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft;
                        break;
                    case "middle":
                        textRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                        break;
                    case "end":
                        textRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignRight;
                        break;
                }
            }
        }

        private void AdjustTextPosition(PowerPoint.Shape textBox, float x, float y, string textAnchor)
        {
            switch (textAnchor)
            {
                case "start":
                    textBox.Left = x;
                    break;
                case "middle":
                    textBox.Left = x - (textBox.Width / 2);
                    break;
                case "end":
                    textBox.Left = x - textBox.Width;
                    break;
                default:  // 기본값은 "start"로 처리
                    textBox.Left = x;
                    break;
            }

            // SVG의 y 좌표는 텍스트의 기준선을 나타내므로, 텍스트 높이의 약 70%를 빼줌
            // 이 값은 대략적인 것으로, 폰트에 따라 조정이 필요할 수 있음
            textBox.Top = y - (textBox.Height * 0.7f);
        }

        private bool IsCloseEnough(float x, float y, ShapeInfo shape)
        {
            // 이 임계값은 조정할 수 있습니다.
            float threshold = Math.Min(shape.Shape.Width, shape.Shape.Height) / 2;
            return Math.Abs(x - shape.CenterX) < threshold && Math.Abs(y - shape.CenterY) < threshold;
        }
    }
}