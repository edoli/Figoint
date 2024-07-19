using Microsoft.Office.Core;
using Microsoft.Scripting.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace EdoliAddIn
{
    using ShapeInfoQuadTree = QuadTree<PowerPoint.Shape>;
    using ShapeInfo = Leaf<PowerPoint.Shape>;

    public class SVGtoPPTParser
    {
        private readonly PowerPoint.Slide slide;
        private ShapeInfoQuadTree currentQuadTree;
        private List<ShapeInfo> currentGroup;
        private Dictionary<string, string> currentGroupAttribute;
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
            return ProcessGroup(svg).Data;
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
                    ApplyBasicStyles(shapeInfo.Data, element);
                }
                else if (elementName == "g")
                {
                    ApplyGroupStyles(shapeInfo.Data, element);
                }

                if (textShapeNames.Contains(elementName))
                {
                    currentQuadTree.Insert(shapeInfo);
                }
            }
        }


        private string GetAttribute(XElement element, string attribute, string defaultValue = null)
        {
            return element.Attribute(attribute)?.Value ?? defaultValue;
        }

        private string GetAttributeAlter(XElement element, string attribute, string alternateAttribute, string defaultValue = null)
        {
            return element.Attribute(attribute)?.Value ?? element.Attribute(alternateAttribute)?.Value ?? defaultValue;
        }

        private string GetAttributeInherit(XElement element, string attribute, string defaultValue = null)
        {
            return element.Attribute(attribute)?.Value ?? (currentGroupAttribute.TryGetValue(attribute, out string value) ? value : defaultValue);
        }

        private void AddGroupAttribute(XElement element, string attribute)
        {
            var attrElem = element.Attribute(attribute);
            if (attrElem != null)
            {
                currentGroupAttribute[attribute] = attrElem.Value;
            }
        }

        private readonly string[] inheritAttributes = new string[] {
            "stroke", "stroke-dasharray", "fill", "opacity", "stroke-width", "marker-start", "marker-end",
            "font-family", "font-size", "font-weight", "font-style", "text-decoration", "text-anchor", };

        private ShapeInfo ProcessGroup(XElement groupElement)
        {
            var previousQuadTree = currentQuadTree;
            currentQuadTree = new ShapeInfoQuadTree();

            var previousGroup = currentGroup;
            currentGroup = new List<ShapeInfo>();

            var previousGroupAttribute = currentGroupAttribute;
            currentGroupAttribute = new Dictionary<string, string>();

            PowerPoint.Shape groupedShape = null;
            float centerX = 0f;
            float centerY = 0f;

            styleParser.ParseAndConvertStyle(groupElement);
            
            foreach (var attr in inheritAttributes)
            {
                AddGroupAttribute(groupElement, attr);
            }
            
            foreach (var childElement in groupElement.Elements())
            {
                ProcessElement(childElement);
            }

            if (currentGroup.Count > 1)
            {
                groupedShape = slide.Shapes.Range(currentGroup.Select(s => s.Data.Name).ToArray()).Group();
            }
            else if (currentGroup.Count == 1)
            {
                groupedShape = currentGroup[0].Data;
            }

            currentQuadTree = previousQuadTree;
            currentGroup = previousGroup;
            currentGroupAttribute = previousGroupAttribute;

            if (groupedShape == null)
            {
                return null;
            }
            
            float left = groupedShape.Left;
            float top = groupedShape.Top;
            float right = left + groupedShape.Width;
            float bottom = top + groupedShape.Height;
            centerX = (left + right) / 2;
            centerY = (top + bottom) / 2;

            return new ShapeInfo(groupedShape, centerX, centerY);
        }

        private ShapeInfo DrawRectangle(XElement rect)
        {
            float x = float.Parse(GetAttribute(rect, "x", "0"));
            float y = float.Parse(GetAttribute(rect, "y", "0"));
            float width = float.Parse(GetAttribute(rect, "width"));
            float height = float.Parse(GetAttribute(rect, "height"));
                    
            float rx = float.Parse(GetAttribute(rect, "rx", "0"));
            float ry = float.Parse(GetAttribute(rect, "ry", "0"));
                    
            PowerPoint.Shape shape;
            if (rx > 0 || ry > 0)
            {
                // rx와 ry 중 하나만 지정된 경우 다른 하나도 같은 값으로 설정
                if (rx > 0 && ry == 0) ry = rx;
                if (ry > 0 && rx == 0) rx = ry;

                // rx와 ry의 최대값은 width/2와 height/2
                rx = Math.Min(rx, width / 2);
                ry = Math.Min(ry, height / 2);

                if (rx == ry)
                {
                    shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, x, y, width, height);
                    
                    // PowerPoint에서 모서리 반경은 0에서 0.5 사이의 값으로 표현됨
                    float adjustmentValue = rx / width;
                    shape.Adjustments[1] = adjustmentValue;
                }
                else
                {
                    shape = DrawRoundedRectangle(x, y, width, height, rx, ry);
                }
            }
            else
            {
                shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, width, height);
            }
            return new ShapeInfo(shape, x + width / 2, y + height / 2);
        }

        private PowerPoint.Shape DrawRoundedRectangle(float x, float y, float width, float height, float rx, float ry)
        {
            List<Vector2> points = new List<Vector2>();

            var startPoint = new Vector2(x + rx, y);
            var currentPoint = startPoint;
            var lastControlPoint = startPoint;

            LineTo(ref currentPoint, ref lastControlPoint, new float[] { x + width - rx, y }, false, points);
            EllipticalArc(ref currentPoint, ref lastControlPoint, new float[] { rx, ry, 0, 0, 1, x + width, y + ry }, false, points);
            LineTo(ref currentPoint, ref lastControlPoint, new float[] { x + width, y + height - ry }, false, points);
            EllipticalArc(ref currentPoint, ref lastControlPoint, new float[] { rx, ry, 0, 0, 1, x + width - rx, y + height }, false, points);
            LineTo(ref currentPoint, ref lastControlPoint, new float[] { x + rx, y + height }, false, points);
            EllipticalArc(ref currentPoint, ref lastControlPoint, new float[] { rx, ry, 0, 0, 1, x, y + height - ry }, false, points);
            LineTo(ref currentPoint, ref lastControlPoint, new float[] { x, y + ry }, false, points);
            EllipticalArc(ref currentPoint, ref lastControlPoint, new float[] { rx, ry, 0, 0, 1, x + rx, y }, false, points);

            points.Add(currentPoint);

            // PowerPoint에서 사용할 수 있는 형식으로 점들 변환
            float[,] pointsArray = new float[points.Count, 2];
            for (int i = 0; i < points.Count; i++)
            {
                pointsArray[i, 0] = points[i].X;
                pointsArray[i, 1] = points[i].Y;
            }

            PowerPoint.Shape shape = slide.Shapes.AddCurve(pointsArray);

            return shape;
        }

        private ShapeInfo DrawCircle(XElement circle)
        {
            float cx = float.Parse(GetAttribute(circle, "cx"));
            float cy = float.Parse(GetAttribute(circle, "cy"));
            float r = float.Parse(GetAttribute(circle, "r"));

            var shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, cx - r, cy - r, r * 2, r * 2);
            return new ShapeInfo(shape, cx, cy);
        }

        private ShapeInfo DrawLine(XElement line)
        {
            // TODO: minor improvement by bypass float.Parse when no param
            float x1 = float.Parse(GetAttributeAlter(line, "x1", "x", "0"));
            float y1 = float.Parse(GetAttributeAlter(line, "y1", "y", "0"));
            float x2 = float.Parse(GetAttribute(line, "x2", "0"));
            float y2 = float.Parse(GetAttribute(line, "y2", "0"));

            PowerPoint.Shape shape = slide.Shapes.AddLine(x1, y1, x2, y2);

            ApplyMarkers(shape, line);
            return new ShapeInfo(shape, (x1 + x2) / 2, (y1 + y2) / 2);
        }
        
        private ShapeInfo DrawEllipse(XElement ellipse)
        {
            float cx = float.Parse(GetAttribute(ellipse, "cx", "0"));
            float cy = float.Parse(GetAttribute(ellipse, "cy", "0"));
            float rx = float.Parse(GetAttribute(ellipse, "rx", "0"));
            float ry = float.Parse(GetAttribute(ellipse, "ry", "0"));

            var shape = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeOval, 
                cx - rx, cy - ry, rx * 2, ry * 2
            );

            return new ShapeInfo(shape, cx, cy);
        }

        private ShapeInfo DrawPolygon(XElement polygon)
        {
            string points = GetAttribute(polygon, "points");
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

            return new ShapeInfo(shape, centerX, centerY);
        }

        private ShapeInfo DrawPath(XElement pathElement)
        {
            // TODO: M 커맨드가 중간에 있는 경우 여러개의 path로 나누기
            string d = GetAttribute(pathElement, "d");
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
            PowerPoint.Shape shape = slide.Shapes.AddCurve(pointsArray);

            // 위치와 크기 조정
            shape.Left = minX;
            shape.Top = minY;
            shape.Width = width;
            shape.Height = height;

            ApplyMarkers(shape, pathElement);
            return new ShapeInfo(shape, centerX, centerY);

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
                        MoveTo(ref currentPoint, ref lastCubicControlPoint, parameters, command == 'm', points);
                        break;
                    case 'L':
                    case 'l':
                        LineTo(ref currentPoint, ref lastCubicControlPoint, parameters, command == 'l', points);
                        break;
                    case 'H':
                    case 'h':
                        HorizontalLineTo(ref currentPoint, ref lastCubicControlPoint, parameters, command == 'h', points);
                        break;
                    case 'V':
                    case 'v':
                        VerticalLineTo(ref currentPoint, ref lastCubicControlPoint, parameters, command == 'v', points);
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
                        EllipticalArc(ref currentPoint, ref lastQuadraticControlPoint, parameters, command == 'a', points);
                        break;
                    case 'Z':
                    case 'z':
                        if (points.Count > 0)
                        {
                            var p0 = points[0];
                            LineTo(ref currentPoint, ref lastCubicControlPoint, new float[] { p0.X, p0.Y }, false, points);
                        }
                        break;
                }
            }
            
            points.Add(currentPoint);

            return points;
        }
        private float[] ParseParameters(string cmd)
        {
            string parameterString = cmd.Substring(1);
            string pattern = @"[-+]?[0-9]*\.?[0-9]+(?:e[-+]?[0-9]+)?";
            MatchCollection matches = Regex.Matches(parameterString, pattern);
            return matches.Select(match => float.Parse(((Match) match).Value)).ToArray();
        }
            
        private void MoveTo(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            float x = parameters[0];
            float y = parameters[1];

            Vector2 newPoint = isRelative ? currentPoint + new Vector2(x, y) : new Vector2(x, y);

            currentPoint = newPoint;
            lastControlPoint = newPoint;
        }

        private void LineTo(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            float x = parameters[0];
            float y = parameters[1];

            Vector2 newPoint = isRelative ? currentPoint + new Vector2(x, y) : new Vector2(x, y);

            points.Add(currentPoint);
            points.Add(Vector2.Lerp(currentPoint, newPoint, 1/3f));
            points.Add(Vector2.Lerp(currentPoint, newPoint, 2/3f));

            currentPoint = newPoint;
            lastControlPoint = points[points.Count - 1];
        }

        private void HorizontalLineTo(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            float x = parameters[0];

            Vector2 newPoint = isRelative ? new Vector2(currentPoint.X + x, currentPoint.Y) : new Vector2(x, currentPoint.Y);

            points.Add(currentPoint);
            points.Add(Vector2.Lerp(currentPoint, newPoint, 1/3f));
            points.Add(Vector2.Lerp(currentPoint, newPoint, 2/3f));

            currentPoint = newPoint;
            lastControlPoint = points[points.Count - 1];
        }

        private void VerticalLineTo(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            float y = parameters[0];

            Vector2 newPoint = isRelative ? new Vector2(currentPoint.X, currentPoint.Y + y) : new Vector2(currentPoint.X, y);

            points.Add(currentPoint);
            points.Add(Vector2.Lerp(currentPoint, newPoint, 1/3f));
            points.Add(Vector2.Lerp(currentPoint, newPoint, 2/3f));

            currentPoint = newPoint;
            lastControlPoint = points[points.Count - 1];
        }

        private void CubicBezierCurve(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            Vector2 control1 = ParsePoint(parameters, 0, isRelative, currentPoint);
            Vector2 control2 = ParsePoint(parameters, 2, isRelative, currentPoint);
            Vector2 end = ParsePoint(parameters, 4, isRelative, currentPoint);

            points.Add(currentPoint);
            points.Add(control1);    
            points.Add(control2);    

            currentPoint = end;
            lastControlPoint = control2;
        }

        private void SmoothCubicBezierCurve(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            Vector2 control1 = currentPoint + (currentPoint - lastControlPoint);
            Vector2 control2 = ParsePoint(parameters, 0, isRelative, currentPoint);
            Vector2 end = ParsePoint(parameters, 2, isRelative, currentPoint);

            points.Add(currentPoint);
            points.Add(control1);    
            points.Add(control2);    

            currentPoint = end;
            lastControlPoint = control2;
        }

        private void QuadraticBezierCurve(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            Vector2 control = ParsePoint(parameters, 0, isRelative, currentPoint);
            Vector2 end = ParsePoint(parameters, 2, isRelative, currentPoint);
            
            // HACK: 이거 맞나?
            // 2차 베지어를 3차 베지어로 변환
            Vector2 control1 = currentPoint + 2f/3f * (control - currentPoint);
            Vector2 control2 = end + 2f/3f * (control - end);

            points.Add(currentPoint);
            points.Add(control1);    
            points.Add(control2);    

            currentPoint = end;
            lastControlPoint = control;
        }

        private void SmoothQuadraticBezierCurve(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            Vector2 control = currentPoint + (currentPoint - lastControlPoint);
            Vector2 end = ParsePoint(parameters, 0, isRelative, currentPoint);

            // 2차 베지어를 3차 베지어로 변환
            Vector2 control1 = currentPoint + 2f/3f * (control - currentPoint);
            Vector2 control2 = end + 2f/3f * (control - end);

            points.Add(currentPoint);
            points.Add(control1);    
            points.Add(control2);    

            currentPoint = end;
            lastControlPoint = control;
        }

        private const float Epsilon = 1e-7f;

        private void EllipticalArc(ref Vector2 currentPoint, ref Vector2 lastControlPoint, float[] parameters, bool isRelative, List<Vector2> points)
        {
            float rx = Math.Abs(parameters[0]);
            float ry = Math.Abs(parameters[1]);
            float xAxisRotation = parameters[2] * (float)(Math.PI / 180);
            bool largeArcFlag = parameters[3] != 0;
            bool sweepFlag = parameters[4] != 0;
            Vector2 start = currentPoint;
            Vector2 end = ParsePoint(parameters, 5, isRelative, currentPoint);

            // 반지름이 0이면 직선으로 처리
            if (rx == 0 || ry == 0)
            {
                Vector2 control1 = Vector2.Lerp(start, end, 1/3f);
                Vector2 control2 = Vector2.Lerp(start, end, 2/3f);
                points.Add(control1);
                points.Add(control2);
                lastControlPoint = control2;
                currentPoint = end;
                return;
            }

            // 시작점과 끝점이 같으면 아무것도 하지 않음
            if (Math.Abs(start.X - end.X) < Epsilon && Math.Abs(start.Y - end.Y) < Epsilon)
            {
                return;
            }

            // 각도를 라디안으로 변환
            float cosPhi = (float)Math.Cos(xAxisRotation);
            float sinPhi = (float)Math.Sin(xAxisRotation);

            // 엔드포인트를 타원 좌표계로 변환
            float x1 = cosPhi * (start.X - end.X) / 2 + sinPhi * (start.Y - end.Y) / 2;
            float y1 = -sinPhi * (start.X - end.X) / 2 + cosPhi * (start.Y - end.Y) / 2;

            // 반지름 보정
            float lambda = (x1 * x1) / (rx * rx) + (y1 * y1) / (ry * ry);
            if (lambda > 1)
            {
                float sqrtLambda = (float)Math.Sqrt(lambda);
                rx *= sqrtLambda;
                ry *= sqrtLambda;
            }

            // 중심점 계산
            float sign = (largeArcFlag == sweepFlag) ? -1 : 1;
            float sq = Math.Max(0, (rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1) / (rx * rx * y1 * y1 + ry * ry * x1 * x1));
            Vector2 c = new Vector2(
                sign * (float)Math.Sqrt(sq) * rx * y1 / ry,
                sign * (float)Math.Sqrt(sq) * -ry * x1 / rx
            );

            // 중심점을 원래 좌표계로 변환
            Vector2 center = new Vector2(
                cosPhi * c.X - sinPhi * c.Y + (start.X + end.X) / 2,
                sinPhi * c.X + cosPhi * c.Y + (start.Y + end.Y) / 2
            );

            // 시작각과 각도 범위 계산
            Vector2 startVector = new Vector2((x1 - c.X) / rx, (y1 - c.Y) / ry);
            Vector2 endVector = new Vector2((-x1 - c.X) / rx, (-y1 - c.Y) / ry);

            float startAngle = (float)Math.Atan2(startVector.Y, startVector.X);
            float sweepAngle = (float)Math.Atan2(endVector.Y * startVector.X - endVector.X * startVector.Y,
                                                endVector.X * startVector.X + endVector.Y * startVector.Y);

            if (!sweepFlag && sweepAngle > 0)
            {
                sweepAngle -= 2 * (float)Math.PI;
            }
            else if (sweepFlag && sweepAngle < 0)
            {
                sweepAngle += 2 * (float)Math.PI;
            }

            // 베지어 곡선 생성
            int segments = (int)Math.Ceiling((Math.Abs(sweepAngle) - 0.01f) / (Math.PI / 2));
            float deltaTheta = sweepAngle / segments;
            // float t = (float)(8 / 3 * Math.Sin(deltaTheta / 4) * Math.Sin(deltaTheta / 4) / Math.Sin(deltaTheta / 2));
            float t = 0.5522847498f;

            Vector2 prevPoint = start;

            for (int i = 0; i < segments; i++)
            {
                float theta = startAngle + i * deltaTheta;
                float thetaNext = theta + deltaTheta;

                Vector2 sinCos = new Vector2((float)Math.Sin(theta), (float)Math.Cos(theta));
                Vector2 sinCosNext = new Vector2((float)Math.Sin(thetaNext), (float)Math.Cos(thetaNext));

                Vector2 p1 = new Vector2(
                    cosPhi * rx * sinCos.Y - sinPhi * ry * sinCos.X,
                    sinPhi * rx * sinCos.Y + cosPhi * ry * sinCos.X
                ) + center;

                Vector2 p2 = new Vector2(
                    cosPhi * rx * sinCosNext.Y - sinPhi * ry * sinCosNext.X,
                    sinPhi * rx * sinCosNext.Y + cosPhi * ry * sinCosNext.X
                ) + center;

                // t * 미분값
                Vector2 q1 = p1 + new Vector2(
                    -t * (cosPhi * rx * sinCos.X + sinPhi * ry * sinCos.Y),
                    -t * (sinPhi * rx * sinCos.X - cosPhi * ry * sinCos.Y)
                );

                Vector2 q2 = p2 + new Vector2(
                    t * (cosPhi * rx * sinCosNext.X + sinPhi * ry * sinCosNext.Y),
                    t * (sinPhi * rx * sinCosNext.X - cosPhi * ry * sinCosNext.Y)
                );

                points.Add(prevPoint);
                points.Add(q1);
                points.Add(q2);

                prevPoint = p2;
            }

            lastControlPoint = points[points.Count - 2];
            currentPoint = end;
        }

        private Vector2 ParsePoint(float[] parameters, int startIndex, bool isRelative, Vector2 currentPoint)
        {
            float x = parameters[startIndex];
            float y = parameters[startIndex + 1];
            return isRelative ? currentPoint + new Vector2(x, y) : new Vector2(x, y);
        }

        private void ApplyBasicStyles(PowerPoint.Shape shape, XElement element)
        {
            bool isGroup = element.Name.LocalName.ToLower() == "g";

            string stroke = GetAttributeInherit(element, "stroke");
            if (!string.IsNullOrEmpty(stroke))
            {
                shape.Line.Visible = MsoTriState.msoTrue;
                shape.Line.ForeColor.RGB = MyColor.FromSVG(stroke).ToInt();
            }
            else if (!isGroup)
            {
                shape.Line.Visible = MsoTriState.msoFalse;
            }
            
            string strokeDashArray = GetAttributeInherit(element, "stroke-dasharray");
            if (!string.IsNullOrEmpty(strokeDashArray))
            {
                SVGStrokeDashArrayParser.ApplyStrokeDashArray(shape, strokeDashArray);
            }

            string fill = GetAttributeInherit(element, "fill");
            string opacity = GetAttributeInherit(element, "opacity");
            
            if (!string.IsNullOrEmpty(fill) && fill != "none")
            {
                shape.Fill.Visible = MsoTriState.msoTrue;
                FillColor(shape, fill, opacity);
            }
            else if (!isGroup)
            {
                shape.Fill.Visible = MsoTriState.msoFalse;
            }

            string strokeWidth = GetAttributeInherit(element, "stroke-width");
            if (!string.IsNullOrEmpty(strokeWidth))
            {
                shape.Line.Weight = float.Parse(strokeWidth);
            }

            string transform = GetAttribute(element, "transform");
            if (!string.IsNullOrEmpty(transform))
            {
                SVGTransformParser.ApplyTransform(shape, transform);
            }
        }

        private void ApplyMarkers(PowerPoint.Shape shape, XElement pathElement)
        {
            string markerStart = GetAttributeInherit(pathElement, "marker-start");
            string markerEnd = GetAttributeInherit(pathElement, "marker-end");

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
            string transform = GetAttribute(element, "transform");
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
            float x = float.Parse(GetAttribute(textElement, "x", "0"));
            float y = float.Parse(GetAttribute(textElement, "y", "0"));

            (float centerX, float centerY, float totalHeight) = CalculateTextBlockCenter(textElement);
            
            string text = ProcessTextContent(textElement);

            ShapeInfo nearestShape = currentQuadTree.FindNearest(x, y);

            // HACK: Currently only allow middle aligned text
            if (nearestShape != null 
                && IsCloseEnough(centerX, centerY, totalHeight, nearestShape)
                && IsShapeTextEmpty(nearestShape.Data)
                && GetAttributeInherit(textElement, "text-anchor") == "middle")
            {
                // 도형 내부에 텍스트 추가
                nearestShape.Data.TextFrame2.TextRange.Text = text;
                PowerPoint.TextFrame2 textFrame = nearestShape.Data.TextFrame2;
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
                AdjustTextPosition(textBox, x, y, GetAttributeInherit(textElement, "text-anchor"));
                return new ShapeInfo(textBox, x, y);
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
            float x = float.Parse(GetAttribute(textElement, "x", "0"));
            float y = float.Parse(GetAttribute(textElement, "y", "0"));
            float summedY = 0;
            float currentY = y;
            int lineCount = 0;

            foreach (var node in textElement.Nodes())
            {
                if (node is XElement tspan)
                {
                    float dy = float.Parse(GetAttribute(textElement, "dy", "0"));
                    currentY += dy;
                    summedY += currentY;
                    lineCount++;
                }
                else if (node is XText)
                {
                    summedY += y;
                    lineCount++;
                }
            }

            // Estimate line height based on font size if available
            float lineHeight = 20; // Default line height
            string fontSize = GetAttributeInherit(textElement, "font-size");
            if (!string.IsNullOrEmpty(fontSize))
            {
                if (fontSize.EndsWith("px"))
                {
                    fontSize = fontSize.Substring(0, fontSize.Length - 2);
                }
                lineHeight = float.Parse(fontSize) * 1.2f; // Assuming line height is 1.2 times font size
            }

            float totalHeight = currentY - y + lineHeight;
            float centerY = summedY / lineCount; // 평균값 사용

            return (x, centerY, totalHeight);
        }


        private void ApplyTextStyles(PowerPoint.TextFrame2 textFrame, XElement textElement)
        {
            var textRange = textFrame.TextRange;

            string fontFamily = GetAttributeInherit(textElement, "font-family");
            if (!string.IsNullOrEmpty(fontFamily))
            {
                textRange.Font.Name = fontFamily.Trim('\'', '"');
            }

            string fontSize = GetAttributeInherit(textElement, "font-size");
            if (!string.IsNullOrEmpty(fontSize))
            {
                if (fontSize.EndsWith("px"))
                {
                    fontSize = fontSize.Substring(0, fontSize.Length - 2);
                }
                textRange.Font.Size = float.Parse(fontSize);
            }

            string fill = GetAttributeInherit(textElement, "fill");
            if (!string.IsNullOrEmpty(fill))
            {
                textRange.Font.Fill.ForeColor.RGB = MyColor.FromSVG(fill).ToInt();
            }
            else
            {   
                textRange.Font.Fill.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorDark1;
            }

            string fontWeight = GetAttributeInherit(textElement, "font-weight");
            if (!string.IsNullOrEmpty(fontWeight))
            {
                textRange.Font.Bold = fontWeight == "bold" ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }

            string fontStyle = GetAttributeInherit(textElement, "font-style");
            if (!string.IsNullOrEmpty(fontStyle))
            {
                textRange.Font.Italic = fontStyle == "italic" ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }

            string textDecoration = GetAttributeInherit(textElement, "text-decoration");
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

            string textAnchor = GetAttributeInherit(textElement, "text-anchor");
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
            
            textBox.Top = y - GetApproximateLineHeight(textBox) * 0.7f;
        }

        public float GetApproximateLineHeight(PowerPoint.Shape textBox)
        {
            float totalHeight = textBox.Height;
            string text = textBox.TextFrame.TextRange.Text;
            string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            int lineCount = lines.Length;

            return totalHeight / lineCount;
        }

        private bool IsCloseEnough(float x, float y, float textHeight, ShapeInfo shape)
        {
            // 이 임계값은 조정할 수 있습니다.
            float threshold = textHeight / 2;
            return Math.Abs(x - shape.X) < threshold && Math.Abs(y - shape.Y) < threshold;
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