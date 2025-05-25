using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{

    public static class DimensionTool
    {

        private const float shapeScaleDefault = 28.3465f;

        private static float shapeScale = shapeScaleDefault;

        private static Dictionary<PowerPoint.Shape, PowerPoint.Shape> dimensionTextboxes = new Dictionary<PowerPoint.Shape, PowerPoint.Shape>();

        public static void OnAfterShapeSizeChange()
        {

        }

        private static PowerPoint.Shape AddDimensionTextbox(Slide slide, float x, float y, String text)
        {
            // Add textbox
            var textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, x, y, 0, 0);

            textBox.TextFrame2.MarginLeft = 2;
            textBox.TextFrame2.MarginRight = 2;
            textBox.TextFrame2.MarginTop = 0;
            textBox.TextFrame2.MarginBottom = 0;

            textBox.TextFrame.TextRange.Font.Size = 8;
            textBox.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            textBox.TextFrame.WordWrap = MsoTriState.msoFalse;
            textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

            textBox.TextFrame.TextRange.Text = text;

            // text의 센터가 x, y에 오도록 위치 조정
            textBox.Top -= textBox.Height / 2;

            return textBox;
        }

        public static void DrawAngle()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();
            var slide = Util.CurrentSlide();

            // 두 선이 선택된 경우 두 선 사이의 각도를 표시
            if (shapes.Count == 2 && shapes.All(s => s.Type == MsoShapeType.msoLine))
            {
                DrawAngleBetweenLines(shapes[0], shapes[1], slide);
            }
            // 단일 도형이 선택된 경우 각 꼭지점의 각도를 표시
            else if (shapes.Count == 1)
            {
                var shape = shapes[0];

                if (shape.Nodes != null && shape.Nodes.Count >= 3)
                {
                    DrawAnglesForPolygon(shape, slide);
                }
            }
        }

        private static void DrawAngleBetweenLines(PowerPoint.Shape line1, PowerPoint.Shape line2, PowerPoint.Slide slide)
        {
            var vertices1 = line1.GetVertices();
            var vertices2 = line2.GetVertices();

            Vector2 line1Vector = vertices1[1] - vertices1[0];
            Vector2 line2Vector = vertices2[1] - vertices2[0];

            var intersection = MathExt.CalculateIntersection(
                vertices1[0], line1Vector,
                vertices2[0], line2Vector);

            var intersectionPoint = intersection.Point;

            // If the lines are parallel, intersection will be NaN
            if (float.IsNaN(intersectionPoint.X) || float.IsNaN(intersectionPoint.Y))
            {
                return;
            }

            Vector2 v1 = Vector2.Normalize(line1Vector);
            Vector2 v2 = Vector2.Normalize(line2Vector);
            double angle = MathExt.CalculateAngleBetween(v1, v2);

            if (angle > 180.0)
            {
                angle = 360.0 - angle;
            }

            // 교점에서 각도를 표시 (4위치에 모두 표시)
            DrawAngleArc(intersectionPoint.X, intersectionPoint.Y, 12, v1, v2, angle, slide);
            DrawAngleArc(intersectionPoint.X, intersectionPoint.Y, 12, -v1, -v2, angle, slide);
            DrawAngleArc(intersectionPoint.X, intersectionPoint.Y, 12, -v1, v2, 180 - angle, slide);
            DrawAngleArc(intersectionPoint.X, intersectionPoint.Y, 12, v1, -v2, 180 - angle, slide);
        }

        private static void DrawAnglesForPolygon(PowerPoint.Shape polygon, PowerPoint.Slide slide)
        {
            Vector2[] vertices = polygon.GetVertices();
            int vertexCount = vertices.Length;

            if (vertexCount < 3) return;

            // If the first and last vertices are the same, remove the last vertex and mark as closed
            var isClosed = false;
            if (vertexCount > 1 && vertices[0] == vertices[vertexCount - 1])
            {
                vertexCount--;
                Array.Resize(ref vertices, vertexCount);
                isClosed = true;
            }

            // calculate angles for each vertex
            int startIndex, endIndex;
            if (isClosed)
            {
                startIndex = 0;
                endIndex = vertexCount;
            }
            else
            {
                startIndex = 1;
                endIndex = vertexCount - 1;
            }

            for (int i = startIndex; i < endIndex; i++)
            {
                int prev = (i - 1 + vertexCount) % vertexCount;
                int next = (i + 1) % vertexCount;

                Vector2 v1 = vertices[prev] - vertices[i];
                Vector2 v2 = vertices[next] - vertices[i];

                double angle = MathExt.CalculateAngleBetween(v1, v2);

                if (angle > 180.0)
                {
                    angle = 360.0 - angle;
                }

                DrawAngleArc(vertices[i].X, vertices[i].Y, 12, v1, v2, angle, slide);
            }
        }

        private static void DrawAngleArc(float x, float y, float radius, Vector2 v1, Vector2 v2, double angle, PowerPoint.Slide slide)
        {
            float startAngleRadians = (float)Math.Atan2(v1.Y, v1.X);
            float endAngleRadians = (float)Math.Atan2(v2.Y, v2.X);

            float startAngleDegrees = startAngleRadians * (180f / (float)Math.PI);
            float endAngleDegrees = endAngleRadians * (180f / (float)Math.PI);

            // endAngleDegrees 와 startAngleDegrees 사이각이 180도 이하가 되도록 조정
            while (endAngleDegrees < startAngleDegrees)
            {
                endAngleDegrees += 360f;
            }

            if (endAngleDegrees - startAngleDegrees > 180f)
            {
                float temp = startAngleDegrees;
                startAngleDegrees = endAngleDegrees;
                endAngleDegrees = temp;
            }

            var v1Norm = Vector2.Normalize(v1);
            var v2Norm = Vector2.Normalize(v2);
            float midAngleRadians = (float)Math.Atan2((v1Norm.Y + v2Norm.Y) / 2.0, (v1Norm.X + v2Norm.X) / 2.0);

            // Check if the angle is a right angle
            bool isRightAngle = Math.Abs(angle - 90) < 0.2;

            if (isRightAngle)
            {
                // Draw perpendicular lines
                float squareSize = radius;

                float x1 = x;
                float y1 = y;

                float v1x = (float)Math.Cos(startAngleRadians) * squareSize;
                float v1y = (float)Math.Sin(startAngleRadians) * squareSize;
                x1 += v1x;
                y1 += v1y;

                float x2 = x;
                float y2 = y;
                float v2x = (float)Math.Cos(endAngleRadians) * squareSize;
                float v2y = (float)Math.Sin(endAngleRadians) * squareSize;
                x2 += v2x;
                y2 += v2y;

                // Calculate the corner point of perpendicular lines
                float cornerX = x1 + v2x;
                float cornerY = y1 + v2y;

                // Add perpendicular lines
                slide.Shapes.AddPolyline(new float[,] { { x1, y1 }, { cornerX, cornerY }, { x2, y2 } });
            }
            else
            {
                var shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeArc, x, y - radius, radius, radius);

                shape.Adjustments[1] = startAngleDegrees;
                shape.Adjustments[2] = endAngleDegrees;
            }

            // Add text box
            string angleText = Math.Round(angle, 1).ToString() + "°";

            // Move text to the middle of the arc
            float radiusScale = isRightAngle ? 1.5f : 1.2f;
            var textbox = AddDimensionTextbox(slide, 0, 0, angleText);
            float textX = x + (float)Math.Cos(midAngleRadians) * (radius * radiusScale + textbox.Width / 2);
            float textY = y + (float)Math.Sin(midAngleRadians) * (radius * radiusScale + textbox.Height / 2);
            textbox.Left = textX - textbox.Width / 2;
            textbox.Top = textY - textbox.Height / 2;
        }

        public static void DistanceBetweenPoints()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();
            var slide = Util.CurrentSlide();

            foreach (var shape in shapes)
            {
                Vector2[] vertices = shape.GetVertices();
                int vertexCount = vertices.Length;

                if (vertexCount < 2)
                {
                    continue;
                }

                for (int i = 0; i < vertexCount - 1; i++)
                {
                    int nextIndex = i + 1;

                    Vector2 start = vertices[i];
                    Vector2 end = vertices[nextIndex];

                    float pixelDistance = Vector2.Distance(start, end);

                    // pixels / shapeScale = cm
                    float actualDistance = pixelDistance / shapeScale;

                    string distanceText = Math.Round(actualDistance * 10, 2).ToString() + " mm";

                    Vector2 midPoint = (start + end) / 2;
                    Vector2 lineVector = end - start;

                    float lineAngle = (float)Math.Atan2(lineVector.Y, lineVector.X);
                    float lineAngleDegrees = lineAngle * (180f / (float)Math.PI);

                    Vector2 perpVector = new Vector2(lineVector.Y, -lineVector.X);
                    perpVector = Vector2.Normalize(perpVector) * 8f; // 8 pixels offset

                    Vector2 textPosition = midPoint + perpVector;

                    // Add text box
                    var textbox = AddDimensionTextbox(slide, textPosition.X, textPosition.Y, distanceText);

                    // Rotate text bot to align with line
                    textbox.Rotation = lineAngleDegrees;
                    textbox.Left = textPosition.X - textbox.Width / 2;
                    textbox.Top = textPosition.Y - textbox.Height / 2;

                    // Draw arrows
                    var startOffsetPos = start + perpVector;
                    var endOffsetPos = end + perpVector;

                    var distanceLineA = slide.Shapes.AddLine(startOffsetPos.X, startOffsetPos.Y, 0, 0);
                    var distanceLineB = slide.Shapes.AddLine(endOffsetPos.X, endOffsetPos.Y, 0, 0);
                    distanceLineA.ConnectorFormat.EndConnect(textbox, 2);
                    distanceLineB.ConnectorFormat.EndConnect(textbox, 4);

                    // Arrowhead
                    distanceLineA.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                    distanceLineB.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                }
            }
        }

        public static void ResetDimensionScale()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            if (shapes.Count == 2)
            {
                PowerPoint.Shape lineShape = null;
                PowerPoint.Shape textShape = null;

                foreach (var shape in shapes)
                {
                    if (shape.Type == MsoShapeType.msoLine || shape.Connector == MsoTriState.msoTrue)
                    {
                        lineShape = shape;
                    }
                    else if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        textShape = shape;
                    }
                }

                if (lineShape != null && textShape != null)
                {
                    var vertices = lineShape.GetVertices();
                    float pixelLength = Vector2.Distance(vertices[0], vertices[1]);

                    string text = textShape.TextFrame.TextRange.Text.Trim();

                    // 숫자 부분과 단위 부분 분리
                    float value = ExtractNumberFromText(text, out string unit);

                    if (value <= 0)
                    {
                        System.Windows.Forms.MessageBox.Show("텍스트에서 유효한 치수를 찾을 수 없습니다.", "경고");
                        return;
                    }

                    switch (unit.ToLower())
                    {
                        case "km": value *= 1e+5f; break;
                        case "m": value *= 1e+2f; break;
                        case "mm": value *= 1e-1f; break;
                        case "um": value *= 1e-4f; break;
                        case "nm": value *= 1e-7f; break;
                    }

                    // shapeScale 계산: 픽셀 / cm
                    shapeScale = pixelLength / value;
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("선과 텍스트가 각각 1개씩 선택되어야 합니다.", "경고");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("선과 텍스트가 각각 1개씩 선택되어야 합니다.", "경고");
            }
        }

        private static float ExtractNumberFromText(string text, out string unit)
        {
            unit = "cm"; // 기본 단위는 cm

            System.Text.RegularExpressions.Match match =
                System.Text.RegularExpressions.Regex.Match(text, @"(\d+[.,]?\d*)(?:\s*)([a-zA-Z]*)?");

            if (match.Success)
            {
                // 단위가 있으면 추출
                if (match.Groups[2].Success && !string.IsNullOrWhiteSpace(match.Groups[2].Value))
                {
                    unit = match.Groups[2].Value.ToLower();
                }

                // 숫자 부분 추출 및 변환
                string numberText = match.Groups[1].Value.Replace(',', '.');
                if (float.TryParse(numberText, System.Globalization.NumberStyles.Any,
                                   System.Globalization.CultureInfo.InvariantCulture, out float number))
                {
                    return number;
                }
            }

            return 0;
        }

    }
}
