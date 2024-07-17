using Microsoft.Office.Core;
using Microsoft.Scripting.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{
    public class SVGtoPPTParser
    {
        private readonly PowerPoint.Slide slide;
        private ShapeInfoQuadTree currentQuadTree;
        private List<ShapeInfo> currentGroup;
        private SVGGradientParser gradientParser;
        private SVGStyleParser styleParser;

        private static readonly string[] textShapeNames = new string[] {
            "rect", "circle", "ellipse", "polygon"
        };

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
            var rootShape = parser.ParseAndDraw(svgCode);
            rootShape.Select();
        }

        public SVGtoPPTParser(PowerPoint.Slide slide)
        {
            this.slide = slide;
        }

        public PowerPoint.Shape ParseAndDraw(string svgCode)
        {
            svgCode = svgCode.Replace("&", "&amp;");

            currentGroup = null;
            currentQuadTree = null;
            
            XElement svg = XElement.Parse(svgCode);
            styleParser = new SVGStyleParser(svg);
            gradientParser = new SVGGradientParser(svg);
            return ProcessGroup(svg).Shape;
        }

        private void ProcessElement(XElement element)
        {
            styleParser.ParseAndConvertStyle(element);

            ShapeInfo shapeInfo = null;
            string elementName = element.Name.LocalName.ToLower();
            switch (elementName)
            {
                case "g":
                    shapeInfo = ProcessGroup(element);
                    break;
                case "rect":
                    shapeInfo = DrawRectangle(element);
                    break;
                case "circle":
                    shapeInfo = DrawCircle(element);
                    break;
                case "ellipse":
                    shapeInfo = DrawEllipse(element);
                    break;
                case "line":
                    shapeInfo = DrawLine(element);
                    break;
                case "polygon":
                    shapeInfo = DrawPolygon(element);
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
                currentGroup?.Add(shapeInfo);

                if (elementName != "g" && elementName != "text")
                {
                    ApplyBasicStyles(shapeInfo.Shape, element);
                }

                if (textShapeNames.Contains(elementName))
                {
                    currentQuadTree.Insert(shapeInfo);
                }
            }
        }

        private ShapeInfo ProcessGroup(XElement groupElement)
        {
            var previousQuadTree = currentQuadTree;
            currentQuadTree = new ShapeInfoQuadTree();

            var previousGroup = currentGroup;
            currentGroup = new List<ShapeInfo>();

            PowerPoint.Shape groupedShape = null;
            float centerX = 0f;
            float centerY = 0f;

            // HACK: Is it required?
            // styleParser.ParseAndConvertStyle(groupElement);
            
            foreach (var childElement in groupElement.Elements())
            {
                ProcessElement(childElement);
            }

            if (currentGroup.Count > 1)
            {
                groupedShape = slide.Shapes.Range(currentGroup.Select(s => s.Shape.Name).ToArray()).Group();
            }
            else if (currentGroup.Count == 1)
            {
                groupedShape = currentGroup[0].Shape;
            }

            currentQuadTree = previousQuadTree;
            currentGroup = previousGroup;

            if (groupedShape == null)
            {
                return null;
            }

            ApplyGroupStyles(groupedShape, groupElement);
            
            float left = groupedShape.Left;
            float top = groupedShape.Top;
            float right = left + groupedShape.Width;
            float bottom = top + groupedShape.Height;
            centerX = (left + right) / 2;
            centerY = (top + bottom) / 2;

            return new ShapeInfo { Shape = groupedShape, CenterX = centerX, CenterY = centerY};
        }

        private ShapeInfo DrawRectangle(XElement rect)
        {
            // TODO: 둥근 모서리 지원
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

            ApplyMarkers(shape, line);
            return new ShapeInfo { Shape = shape, CenterX = (x1 + x2) / 2, CenterY = (y1 + y2) / 2 };
        }
        
        private ShapeInfo DrawEllipse(XElement ellipse)
        {
            float cx = float.Parse(ellipse.Attribute("cx")?.Value ?? "0");
            float cy = float.Parse(ellipse.Attribute("cy")?.Value ?? "0");
            float rx = float.Parse(ellipse.Attribute("rx")?.Value ?? "0");
            float ry = float.Parse(ellipse.Attribute("ry")?.Value ?? "0");

            var shape = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeOval, 
                cx - rx, cy - ry, rx * 2, ry * 2
            );

            return new ShapeInfo { Shape = shape, CenterX = cx, CenterY = cy };
        }

        private ShapeInfo DrawPolygon(XElement polygon)
        {
            string points = polygon.Attribute("points")?.Value;
            if (string.IsNullOrEmpty(points))
            {
                return null;
            }

            var pointList = points.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries)
                                  .Select(float.Parse)
                                  .ToList();

            if (pointList.Count < 4 || pointList.Count % 2 != 0)
            {
                return null;
            }

            float[,] pointArray = new float[pointList.Count / 2, 2];
            for (int i = 0; i < pointList.Count; i += 2)
            {
                pointArray[i / 2, 0] = pointList[i];
                pointArray[i / 2, 1] = pointList[i + 1];
            }

            var shape = slide.Shapes.AddPolyline(pointArray);

            // Calculate center point
            float centerX = pointList.Where((x, i) => i % 2 == 0).Average();
            float centerY = pointList.Where((y, i) => i % 2 != 0).Average();

            return new ShapeInfo { Shape = shape, CenterX = centerX, CenterY = centerY };
        }

        private ShapeInfo DrawPath(XElement pathElement)
        {
            // TODO: Curve path가 제대로 안그려지고 있음. 특히 A에 문제가 있는거 같음
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

        private List<Vector2> ParsePathToPoints(string d)
        {
            List<Vector2> points = new List<Vector2>();
            Vector2 currentPoint = new Vector2(0, 0);
            Vector2 lastCubicControlPoint = new Vector2(0, 0);
            Vector2 lastQuadraticControlPoint = new Vector2(0, 0);

            string[] commands = Regex.Split(d, @"(?=[MmLlHhVvCcSsQqTtAaZz])");

            foreach (string cmd in commands)
            {
                if (string.IsNullOrWhiteSpace(cmd)) continue;

                char command = cmd[0];
                float[] parameters = ParseParameters(cmd);

                switch (command)
                {
                    case 'M':
                    case 'm':
                        MoveTo(ref currentPoint, parameters, command == 'm', points);
                        break;
                    case 'L':
                    case 'l':
                        LineTo(ref currentPoint, parameters, command == 'l', points);
                        break;
                    case 'H':
                    case 'h':
                        HorizontalLineTo(ref currentPoint, parameters, command == 'h', points);
                        break;
                    case 'V':
                    case 'v':
                        VerticalLineTo(ref currentPoint, parameters, command == 'v', points);
                        break;
                    case 'C':
                    case 'c':
                        CubicBezierCurve(ref currentPoint, ref lastCubicControlPoint, parameters, command == 'c', points);
                        break;
                    case 'S':
                    case 's':
                        SmoothCubicBezierCurve(ref currentPoint, ref lastCubicControlPoint, parameters, command == 's', points);
                        break;
                    case 'Q':
                    case 'q':
                        QuadraticBezierCurve(ref currentPoint, ref lastQuadraticControlPoint, parameters, command == 'q', points);
                        break;
                    case 'T':
                    case 't':
                        SmoothQuadraticBezierCurve(ref currentPoint, ref lastQuadraticControlPoint, parameters, command == 't', points);
                        break;
                    case 'A':
                    case 'a':
                        EllipticalArc(ref currentPoint, parameters, command == 'a', points);
                        break;
                    case 'Z':
                    case 'z':
                        // 시작점으로 돌아가는 것은 Polyline에서 자동으로 처리되지 않으므로,
                        // 필요하다면 시작점을 다시 추가할 수 있습니다.
                        if (points.Count > 0)
                        {
                            points.Add(points[0]);
                        }
                        currentPoint = points[0];
                        break;
                }
            }

            return points;
        }
        private float[] ParseParameters(string cmd)
        {
            string parameterString = cmd.Substring(1);
            string pattern = @"[-+]?[0-9]*\.?[0-9]+(?:e[-+]?[0-9]+)?";
            MatchCollection matches = Regex.Matches(parameterString, pattern);
            return matches.Select(match => float.Parse(((Match) match).Value)).ToArray();
        }

        private void MoveTo(ref Vector2 currentPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            float x = parameters[0];
            float y = parameters[1];

            if (isRelative)
            {
                currentPoint += new Vector2(x, y);
            }
            else
            {
                currentPoint = new Vector2(x, y);
            }

            points.Add(currentPoint);
        }

        private void LineTo(ref Vector2 currentPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            float x = parameters[0];
            float y = parameters[1];

            if (isRelative)
            {
                currentPoint += new Vector2(x, y);
            }
            else
            {
                currentPoint = new Vector2(x, y);
            }

            points.Add(currentPoint);
        }

        private void HorizontalLineTo(ref Vector2 currentPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            float x = parameters[0];

            if (isRelative)
            {
                currentPoint.X += x;
            }
            else
            {
                currentPoint.X = x;
            }

            points.Add(currentPoint);
        }

        private void VerticalLineTo(ref Vector2 currentPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            float y = parameters[0];

            if (isRelative)
            {
                currentPoint.Y += y;
            }
            else
            {
                currentPoint.Y = y;
            }

            points.Add(currentPoint);
        }

        private void CubicBezierCurve(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            Vector2 control1 = ParsePoint(parameters, 0, isRelative, currentPoint);
            Vector2 control2 = ParsePoint(parameters, 2, isRelative, currentPoint);
            Vector2 end = ParsePoint(parameters, 4, isRelative, currentPoint);

            const int segments = 10; // Adjust for smoother curves
            for (int i = 1; i <= segments; i++)
            {
                float t = i / (float)segments;
                Vector2 point = CubicBezier(currentPoint, control1, control2, end, t);
                points.Add(point);
            }

            currentPoint = end;
            lastControlPoint = control2;
        }

        private void SmoothCubicBezierCurve(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            Vector2 control1 = currentPoint + (currentPoint - lastControlPoint);
            Vector2 control2 = ParsePoint(parameters, 0, isRelative, currentPoint);
            Vector2 end = ParsePoint(parameters, 2, isRelative, currentPoint);

            const int segments = 10;
            for (int i = 1; i <= segments; i++)
            {
                float t = i / (float)segments;
                Vector2 point = CubicBezier(currentPoint, control1, control2, end, t);
                points.Add(point);
            }

            currentPoint = end;
            lastControlPoint = control2;
        }

        private void QuadraticBezierCurve(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            Vector2 control = ParsePoint(parameters, 0, isRelative, currentPoint);
            Vector2 end = ParsePoint(parameters, 2, isRelative, currentPoint);

            const int segments = 10;
            for (int i = 1; i <= segments; i++)
            {
                float t = i / (float)segments;
                Vector2 point = QuadraticBezier(currentPoint, control, end, t);
                points.Add(point);
            }

            currentPoint = end;
            lastControlPoint = control;
        }

        private void SmoothQuadraticBezierCurve(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            Vector2 control = currentPoint + (currentPoint - lastControlPoint);
            Vector2 end = ParsePoint(parameters, 0, isRelative, currentPoint);

            const int segments = 10;
            for (int i = 1; i <= segments; i++)
            {
                float t = i / (float)segments;
                Vector2 point = QuadraticBezier(currentPoint, control, end, t);
                points.Add(point);
            }

            currentPoint = end;
            lastControlPoint = control;
        }

        private void EllipticalArc(ref Vector2 currentPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            float rx = Math.Abs(parameters[0]);
            float ry = Math.Abs(parameters[1]);
            float xAxisRotation = parameters[2] * (float)(Math.PI / 180);
            bool largeArcFlag = parameters[3] == 1;
            bool sweepFlag = parameters[4] == 1;
            Vector2 end = ParsePoint(parameters, 5, isRelative, currentPoint);

            if (currentPoint == end)
            {
                return;
            }

            if (rx == 0 || ry == 0)
            {
                points.Add(end);
                currentPoint = end;
                return;
            }

            EllipticalArcToBezier(currentPoint, rx, ry, xAxisRotation, largeArcFlag, sweepFlag, end, points);

            currentPoint = end;
        }

        private Vector2 CubicBezier(Vector2 start, Vector2 control1, Vector2 control2, Vector2 end, float t)
        {
            float u = 1 - t;
            float tt = t * t;
            float uu = u * u;
            float uuu = uu * u;
            float ttt = tt * t;

            Vector2 point = uuu * start;
            point += 3 * uu * t * control1;
            point += 3 * u * tt * control2;
            point += ttt * end;

            return point;
        }

        private Vector2 QuadraticBezier(Vector2 start, Vector2 control, Vector2 end, float t)
        {
            float u = 1 - t;
            float tt = t * t;
            float uu = u * u;

            Vector2 point = uu * start;
            point += 2 * u * t * control;
            point += tt * end;

            return point;
        }

        private void EllipticalArcToBezier(Vector2 start, float rx, float ry, float phi, bool largeArcFlag, bool sweepFlag, Vector2 end, List<Vector2> points)
        {
            // Implementation of elliptical arc to Bézier curves conversion
            // This is a complex algorithm, you might want to use a library or implement it separately
            // For now, we'll approximate with a line
            const int segments = 20;
            for (int i = 1; i <= segments; i++)
            {
                float t = i / (float)segments;
                Vector2 point = Vector2.Lerp(start, end, t);
                points.Add(point);
            }
        }

        private Vector2 ParsePoint(float[] parameters, int startIndex, bool isRelative, Vector2 currentPoint)
        {
            float x = parameters[startIndex];
            float y = parameters[startIndex + 1];
            return isRelative ? currentPoint + new Vector2(x, y) : new Vector2(x, y);
        }

        private void ApplyBasicStyles(PowerPoint.Shape shape, XElement element)
        {
            bool isStrokeElement = IsStrokeRequiredElement(element);

            string stroke = element.Attribute("stroke")?.Value;
            if (!string.IsNullOrEmpty(stroke))
            {
                shape.Line.ForeColor.RGB = MyColor.FromSVG(stroke).ToInt();
            }
            else
            {
                shape.Line.Visible = MsoTriState.msoFalse;
            }
            
            string strokeDashArray = element.Attribute("stroke-dasharray")?.Value;
            if (!string.IsNullOrEmpty(strokeDashArray))
            {
                SVGStrokeDashArrayParser.ApplyStrokeDashArray(shape, strokeDashArray);
            }

            if (shape.Fill.ForeColor.Type != MsoColorType.msoColorTypeMixed)
            {
                string fill = element.Attribute("fill")?.Value;
                fill = string.IsNullOrEmpty(fill) ? "#000" : fill;
                string opacity = element.Attribute("opacity")?.Value;

                if (fill != "none")
                {
                    FillColor(shape, fill, opacity);
                }
                else
                {
                    shape.Fill.Visible = MsoTriState.msoFalse;
                }
            }

            string strokeWidth = element.Attribute("stroke-width")?.Value;
            if (!string.IsNullOrEmpty(strokeWidth))
            {
                shape.Line.Weight = float.Parse(strokeWidth);
            }

            string transform = element.Attribute("transform")?.Value;
            if (!string.IsNullOrEmpty(transform))
            {
                SVGTransformParser.ApplyTransform(shape, transform);
            }
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

        private void ApplyGroupStyles(PowerPoint.Shape shape, XElement element)
        {
            string transform = element.Attribute("transform")?.Value;
            if (!string.IsNullOrEmpty(transform))
            {
                SVGTransformParser.ApplyTransform(shape, transform);
            }
        }

        private bool IsStrokeRequiredElement(XElement element)
        {
            string elementName = element.Name.LocalName.ToLower();
            return elementName == "line" || elementName == "path" || elementName == "polyline" || elementName == "polygon";
        }

        private ShapeInfo DrawText(XElement textElement)
        {
            float x = float.Parse(textElement.Attribute("x")?.Value ?? "0");
            float y = float.Parse(textElement.Attribute("y")?.Value ?? "0");

            (float centerX, float centerY, float totalHeight) = CalculateTextBlockCenter(textElement);
            
            string text = ProcessTextContent(textElement);

            ShapeInfo nearestShape = currentQuadTree.FindNearest(x, y);

            // HACK: Currently only allow middle aligned text
            if (nearestShape != null 
                && IsCloseEnough(centerX, centerY, nearestShape) 
                && IsShapeTextEmpty(nearestShape.Shape)
                && textElement.Attribute("text-anchor")?.Value == "middle"
                && false)
            {
                // 도형 내부에 텍스트 추가
                nearestShape.Shape.TextFrame2.TextRange.Text = text;
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

                // Remove margin
                textFrame.MarginBottom = 0;
                textFrame.MarginLeft = 0;
                textFrame.MarginRight = 0;
                textFrame.MarginTop = 0;

                ApplyTextStyles(textFrame, textElement);
                AdjustTextPosition(textBox, x, y, textElement.Attribute("text-anchor")?.Value);
                return new ShapeInfo { Shape = textBox, CenterX = x, CenterY = y };
            }
        }

        private string ProcessTextContent(XElement textElement)
        {
            int paragraphIndex = 1;
            var builder = new StringBuilder();
            foreach (var node in textElement.Nodes())
            {
                if (node is XText textNode)
                {
                    if (paragraphIndex > 1)
                    {
                        builder.Append("\n");
                    }
                    builder.Append(textNode.Value.Replace("&amp;", "&"));
                    paragraphIndex++;
                }
                else if (node is XElement elementNode && elementNode.Name.LocalName.ToLower() == "tspan")
                {
                    if (paragraphIndex > 1)
                    {
                        builder.Append("\n");
                    }
                    builder.Append(elementNode.Value.Replace("&amp;", "&"));
                    paragraphIndex++;
                }
            }
            return builder.ToString();
        }

        private (float centerX, float centerY, float totalHeight) CalculateTextBlockCenter(XElement textElement)
        {
            // TODO: fix it
            float x = float.Parse(textElement.Attribute("x")?.Value ?? "0");
            float y = float.Parse(textElement.Attribute("y")?.Value ?? "0");
            float totalDy = 0;
            float currentY = y;
            int lineCount = 0;

            foreach (var node in textElement.Nodes())
            {
                if (node is XElement tspan)
                {
                    float dy = float.Parse(tspan.Attribute("dy")?.Value ?? "0");
                    currentY += dy;
                    totalDy += currentY;
                    lineCount++;
                }
                else if (node is XText)
                {
                    lineCount++;
                }
            }

            // Estimate line height based on font size if available
            float lineHeight = 20; // Default line height
            string fontSize = textElement.Attribute("font-size")?.Value;
            if (!string.IsNullOrEmpty(fontSize))
            {
                if (fontSize.EndsWith("px"))
                {
                    fontSize = fontSize.Substring(0, fontSize.Length - 2);
                }
                lineHeight = float.Parse(fontSize) * 1.2f; // Assuming line height is 1.2 times font size
            }

            float totalHeight = totalDy + (lineCount - 1) * lineHeight;
            float centerY = totalDy / lineCount;

            return (x, centerY, totalHeight);
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
                textRange.Font.Fill.ForeColor.RGB = MyColor.FromSVG(fill).ToInt();
            }
            else
            {   
                textRange.Font.Fill.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorDark1;
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
            // HACK: 이 값은 대략적인 것으로, 폰트에 따라 조정이 필요할 수 있음
            
            textBox.Top = y - (textBox.Height / 2);
        }

        private bool IsCloseEnough(float x, float y, ShapeInfo shape)
        {
            // 이 임계값은 조정할 수 있습니다.
            float threshold = Math.Min(shape.Shape.Width, shape.Shape.Height) / 8;
            return Math.Abs(x - shape.CenterX) < threshold && Math.Abs(y - shape.CenterY) < threshold;
        }

        private bool IsShapeTextEmpty(PowerPoint.Shape shape)
        {
            try
            {
                // TextFrame이 없거나 텍스트가 비어있으면 true 반환
                return shape.TextFrame.TextRange.Text.Trim() == "";
            }
            catch
            {
                // TextFrame에 접근할 수 없는 경우 (예: 선 객체 등) true 반환
                return true;
            }
        }

        private void FillColor(PowerPoint.Shape shape, string fill, string opacity)
        {
            if (fill.StartsWith("url(#"))
            {
                gradientParser.ApplyGradient(shape, fill, styleParser);
            }
            else
            {
                MyColor fillColor = MyColor.FromSVG(fill, opacity);
                shape.Fill.ForeColor.RGB = fillColor.ToInt();
                shape.Fill.Transparency = fillColor.Transparency;
            }
        }
    }
}