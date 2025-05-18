using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Linq;
using System.Numerics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{

    public static class DimensionTool
    {

        private static float shapeScale = 28.3465f;
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

            Vector2 intersection = CalculateIntersection(
                vertices1[0], line1Vector,
                vertices2[0], line2Vector);

            // If the lines are parallel, intersection will be NaN
            if (float.IsNaN(intersection.X) || float.IsNaN(intersection.Y))
            {
                return;
            }

            Vector2 v1 = Vector2.Normalize(line1Vector);
            Vector2 v2 = Vector2.Normalize(line2Vector);
            double angle = CalculateAngleBetweenVectors(v1, v2);

            if (angle > 180.0)
            {
                angle = 360.0 - angle;
            }

            // 교점에서 각도를 표시 (4위치에 모두 표시)
            DrawAngleArc(intersection.X, intersection.Y, 12, v1, v2, angle, slide);
            DrawAngleArc(intersection.X, intersection.Y, 12, -v1, -v2, angle, slide);
            DrawAngleArc(intersection.X, intersection.Y, 12, -v1, v2, 180 - angle, slide);
            DrawAngleArc(intersection.X, intersection.Y, 12, v1, -v2, 180 - angle, slide);
        }

        private static Vector2 CalculateIntersection(Vector2 point1, Vector2 direction1, Vector2 point2, Vector2 direction2)
        {
            // 두 직선의 방정식을 이용하여 교점 계산
            // 첫 번째 선: point1 + t * direction1
            // 두 번째 선: point2 + s * direction2

            // Check if the lines are parallel
            float cross = direction1.X * direction2.Y - direction1.Y * direction2.X;
            if (Math.Abs(cross) < 1e-6)
            {
                return new Vector2(float.NaN, float.NaN); // No intersection (parallel lines)
            }

            Vector2 diff = point2 - point1;
            float t = (diff.X * direction2.Y - diff.Y * direction2.X) / cross;

            return point1 + t * direction1;
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

                double angle = CalculateAngleBetweenVectors(v1, v2);

                if (angle > 180.0)
                {
                    angle = 360.0 - angle;
                }

                DrawAngleArc(vertices[i].X, vertices[i].Y, 12, v1, v2, angle, slide);
            }
        }
        private static double CalculateAngleBetweenVectors(Vector2 v1, Vector2 v2)
        {
            if (v1.Length() > 0) v1 = Vector2.Normalize(v1);
            if (v2.Length() > 0) v2 = Vector2.Normalize(v2);

            double dotProduct = Vector2.Dot(v1, v2);
            dotProduct = Math.Min(Math.Max(dotProduct, -1.0), 1.0); // 부동소수점 오류 방지
            double angleRadians = Math.Acos(dotProduct);

            double angleDegrees = angleRadians * (180.0 / Math.PI);

            return angleDegrees;
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
                var shape = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeArc,
                    x,
                    y - radius,
                    radius,
                    radius);

                shape.Adjustments[1] = startAngleDegrees;
                shape.Adjustments[2] = endAngleDegrees;
            }

            // Add text box
            string angleText = Math.Round(angle, 1).ToString() + "°";

            var textShape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 0, 0);

            textShape.TextFrame2.MarginLeft = 0;
            textShape.TextFrame2.MarginRight = 0;
            textShape.TextFrame2.MarginTop = 0;
            textShape.TextFrame2.MarginBottom = 0;

            textShape.TextFrame.WordWrap = MsoTriState.msoFalse;
            textShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            textShape.TextFrame.TextRange.Font.Size = 8;
            textShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            textShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

            textShape.TextFrame.TextRange.Text = angleText;

            // Move text to the middle of the arc
            float textX = x + (float)Math.Cos(midAngleRadians) * (radius * 1.5f);
            float textY = y + (float)Math.Sin(midAngleRadians) * (radius * 1.5f);

            textShape.Left = textX - textShape.Width / 2;
            textShape.Top = textY - textShape.Height / 2;
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

                    // Add textbox
                    var textBox = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        textPosition.X,
                        textPosition.Y,
                        0,
                        0);

                    textBox.TextFrame2.MarginLeft = 2;
                    textBox.TextFrame2.MarginRight = 2;
                    textBox.TextFrame2.MarginTop = 0;
                    textBox.TextFrame2.MarginBottom = 0;

                    textBox.TextFrame.TextRange.Font.Size = 8;
                    textBox.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                    textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                    textBox.TextFrame.WordWrap = MsoTriState.msoFalse;
                    textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

                    textBox.TextFrame.TextRange.Text = distanceText;

                    textBox.Top -= textBox.Height / 2;

                    // Rotate text bot to align with line
                    textBox.Rotation = lineAngleDegrees;

                    // Draw arrows
                    var startOffsetPos = start + perpVector;
                    var endOffsetPos = end + perpVector;

                    var distanceLineA = slide.Shapes.AddLine(startOffsetPos.X, startOffsetPos.Y, 0, 0);
                    var distanceLineB = slide.Shapes.AddLine(endOffsetPos.X, endOffsetPos.Y, 0, 0);
                    distanceLineA.ConnectorFormat.EndConnect(textBox, 2);
                    distanceLineB.ConnectorFormat.EndConnect(textBox, 4);

                    // Arrowhead
                    distanceLineA.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                    distanceLineB.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                }
            }
        }
    }
}
