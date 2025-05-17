using Expressive;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Numerics;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{
    public class PointsShape
    {
        public float[,] points;
        public float width;
        public float height;

        public PointsShape(Vector2[] localPoints, float offsetX, float offsetY)
        {
            int numPoints = localPoints.Length;
            points = new float[numPoints, 2];

            var minV = new Vector2(float.MaxValue, float.MaxValue);
            var maxV = new Vector2(float.MinValue, float.MinValue);
            for (int i = 0; i < numPoints; i++)
            {
                var v = localPoints[i];

                if (v.X < minV.X) { minV.X = v.X; }
                if (v.Y < minV.Y) { minV.Y = v.Y; }
                if (v.X > maxV.X) { maxV.X = v.X; }
                if (v.Y > maxV.Y) { maxV.Y = v.Y; }
            }
            float cx = (minV.X + maxV.X) / 2;
            float cy = (minV.Y + maxV.Y) / 2;

            width = maxV.X - minV.X;
            height = maxV.Y - minV.Y;

            for (int i = 0; i < numPoints; i++)
            {
                var v = localPoints[i];
                points[i, 0] = v.X + offsetX - cx;
                points[i, 1] = -v.Y + offsetY + cy;
            }
        }
    }

    public static class ShapeTool
    {

        private static float shapeScale = 28.3465f;

        // Shape Tag안에 정보들을 저장하기 위한 Tag Name.
        // Shape 상속및 확장이 불가능하기 때문에 Tag를 활용해서 정보를 저장함
        public static string PathTypeTagName = "EquationPath";
        public static string ExpressiveXTagName = "ExpressiveX";
        public static string ExpressiveYTagName = "ExpressiveY";
        public static string ExpressiveStartValueTagName = "ExpressiveRangeFrom";
        public static string ExpressiveEndValueTagName = "ExpressiveRangeTo";
        public static string CurveTag = "EquationCurve";
        public static string PolylineTag = "EquationPolyline";


        public static void ToggleLine()
        {
            var shapes = Util.ListSelectedShapes();

            shapes.ForEach(s =>
            {
                if (s.Line.Visible == MsoTriState.msoFalse)
                {
                    s.Line.Visible = MsoTriState.msoTrue;
                }
                else
                {
                    s.Line.Visible = MsoTriState.msoFalse;
                }
            });
        }

        public static void ChangeLineWeight(float offset)
        {
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                shape.DoRecur(s =>
                {
                    var line = s.Line;
                    if (line.Style > 0)
                    {
                        if (line.Weight > -offset)
                        {
                            line.Weight += offset;
                        }
                    }
                });
            }
        }
        public static void ChangeLineDash(int offset)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                shape.DoRecur(s =>
                {
                    var line = s.Line;
                    if (line.Style > 0)
                    {
                        int style = (int)line.DashStyle;
                        if (style == 2 || style == 3)
                        {
                            style = 1;
                        }
                        if (style > 2)
                        {
                            style -= 2;
                        }

                        int newDashStyle = offset + style;
                        if (newDashStyle > 10)
                        {
                            newDashStyle = 1;
                        }
                        if (newDashStyle < 1)
                        {
                            newDashStyle = 10;
                        }

                        // ignore 2, 3 style
                        if (newDashStyle >= 2)
                        {
                            newDashStyle += 2;
                        }
                        line.DashStyle = (MsoLineDashStyle)newDashStyle;
                    }
                });
            }
        }

        public static void BeginArrowToggle()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                try
                {
                    if (shape.Line.BeginArrowheadStyle == MsoArrowheadStyle.msoArrowheadNone)
                    {
                        shape.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                    }
                    else
                    {
                        shape.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone;
                    }
                }
                catch
                {

                }
            }
        }

        public static void BeginArrowChangeSize(int deltaSize)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                {
                    if (shape.Line.BeginArrowheadStyle != MsoArrowheadStyle.msoArrowheadNone)
                    {
                        var width = (int)shape.Line.BeginArrowheadWidth;
                        var length = (int)shape.Line.BeginArrowheadLength;

                        var newWidth = width + deltaSize;
                        var newLength = length + deltaSize;

                        if (newWidth > 0 && newWidth <= 3)
                        {
                            shape.Line.BeginArrowheadWidth = (MsoArrowheadWidth)newWidth;
                        }

                        if (newLength > 0 && newLength <= 3)
                        {
                            shape.Line.BeginArrowheadLength = (MsoArrowheadLength)newLength;
                        }
                    }
                }
            }
        }

        public static void EndArrowToggle()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                try
                {
                    if (shape.Line.EndArrowheadStyle == MsoArrowheadStyle.msoArrowheadNone)
                    {
                        shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                    }
                    else
                    {
                        shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone;
                    }
                }
                catch
                {

                }
            }
        }

        public static void EndArrowChangeSize(int deltaSize)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                {
                    if (shape.Line.EndArrowheadStyle != MsoArrowheadStyle.msoArrowheadNone)
                    {
                        var width = (int)shape.Line.EndArrowheadWidth;
                        var length = (int)shape.Line.EndArrowheadLength;

                        var newWidth = width + deltaSize;
                        var newLength = length + deltaSize;

                        if (newWidth > 0 && newWidth <= 3)
                        {
                            shape.Line.EndArrowheadWidth = (MsoArrowheadWidth)newWidth;
                        }

                        if (newLength > 0 && newLength <= 3)
                        {
                            shape.Line.EndArrowheadLength = (MsoArrowheadLength)newLength;
                        }
                    }
                }
            }
        }

        public static void ConnectShapesByLine()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            for (int i = 0; i < shapes.Count - 1; i++)
            {
                var shape1 = shapes[i];
                var shape2 = shapes[i + 1];

                var rel = shape1.GetRelativePos(shape2, 0.1f);

                if (rel == ShapeExt.Anchor.None)
                {
                    continue;
                }

                PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                float left1 = shape1.Left;
                float top1 = shape1.Top;
                float right1 = shape1.Right();
                float bottom1 = shape1.Bottom();

                float left2 = shape2.Left;
                float top2 = shape2.Top;
                float right2 = shape2.Right();
                float bottom2 = shape2.Bottom();

                if (rel == ShapeExt.Anchor.TopLeft || rel == ShapeExt.Anchor.BottomRight)
                {
                    slide.Shapes.AddLine(left1, bottom1, left2, bottom2);
                    slide.Shapes.AddLine(right1, top1, right2, top2);
                }

                if (rel == ShapeExt.Anchor.TopRight || rel == ShapeExt.Anchor.BottomLeft)
                {
                    slide.Shapes.AddLine(right1, bottom1, right2, bottom2);
                    slide.Shapes.AddLine(left1, top1, left2, top2);
                }

                if (rel == ShapeExt.Anchor.Left)
                {
                    slide.Shapes.AddLine(left1, top1, right2, top2);
                    slide.Shapes.AddLine(left1, bottom1, right2, bottom2);
                }

                if (rel == ShapeExt.Anchor.Right)
                {
                    slide.Shapes.AddLine(right1, top1, left2, top2);
                    slide.Shapes.AddLine(right1, bottom1, left2, bottom2);
                }

                if (rel == ShapeExt.Anchor.Top)
                {
                    slide.Shapes.AddLine(left1, top1, left2, bottom2);
                    slide.Shapes.AddLine(right1, top1, right2, bottom2);
                }

                if (rel == ShapeExt.Anchor.Bottom)
                {
                    slide.Shapes.AddLine(left1, bottom1, left2, top2);
                    slide.Shapes.AddLine(right1, bottom1, right2, top2);
                }
            }
        }

        public class ConnectorInfo
        {
            public PowerPoint.Shape Shape;
            public int ConnectorSite;

            public ConnectorInfo(PowerPoint.Shape shape, int connectorSite)
            {
                Shape = shape;
                ConnectorSite = connectorSite;
            }
        }

        public static void ConnectAllCandidates()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var connectors = new List<Leaf<ConnectorInfo>>();

            var line = slide.Shapes.AddLine(0, 0, 10, 10);

            var connectorShapes = new List<PowerPoint.Shape>();

            foreach (var shape in shapes)
            {
                if (shape.Connector == MsoTriState.msoTrue)
                {
                    connectorShapes.Add(shape);
                    continue;
                }
                for (int i = 0; i < shape.ConnectionSiteCount; i++)
                {
                    line.ConnectorFormat.EndConnect(shape, i + 1);
                    float x = line.Left + line.Width;
                    float y = line.Top + line.Height;
                    if (x == 0)
                    {
                        x = line.Left;
                    }
                    if (y == 0)
                    {
                        y = line.Top;
                    }
                    connectors.Add(new Leaf<ConnectorInfo>(new ConnectorInfo(shape, i + 1), x, y));
                }
            }

            line.Delete();

            // TODO: KdTree 사용 시 문제 발생
            // var searchTree = new KdTree<ConnectorInfo>(connectors);
            var searchTree = new QuadTree<ConnectorInfo>();
            foreach (var connector in connectors)
            {
                searchTree.Insert(connector);
            }

            // HACK: 선의 정확한 노드 포인트를 가져올 수 없음
            foreach (var shape in connectorShapes)
            {
                var p1 = new Vector2(shape.Left, shape.Top);
                var p2 = new Vector2(shape.Left, shape.Top + shape.Height);
                var p3 = new Vector2(shape.Left + shape.Width, shape.Top);
                var p4 = new Vector2(shape.Left + shape.Width, shape.Top + shape.Height);

                var points = new List<Vector2>();

                points.Add(p1);
                if (!IsCloseEnough(p1, p2))
                {
                    points.Add(p2);
                }
                if (!IsCloseEnough(p1, p3))
                {
                    points.Add(p3);
                }
                if (!IsCloseEnough(p2, p4) && !IsCloseEnough(p3, p4))
                {
                    points.Add(p4);
                }
                var nearests = points.Select(p => searchTree.FindNearest(p.X, p.Y))
                    .Where((n, i) => n != null && IsCloseEnough(points[i], n.X, n.Y)).ToArray();

                if (nearests.Count() == 2)
                {
                    var n1 = nearests[0].Data;
                    var n2 = nearests[1].Data;
                    shape.ConnectorFormat.BeginConnect(n1.Shape, n1.ConnectorSite);
                    shape.ConnectorFormat.EndConnect(n2.Shape, n2.ConnectorSite);
                }
            }
        }

        public static bool IsCloseEnough(Vector2 a, Vector2 b)
        {
            return IsCloseEnough(a, b.X, b.Y);
        }

        public static bool IsCloseEnough(Vector2 a, float x, float y)
        {
            return Math.Abs(a.X - x) + Math.Abs(a.Y - y) < 2f;
        }

        public static bool IsCloseEnough(float a, float b)
        {
            return Math.Abs(a - b) < 2f;
        }

        public static void DrawAngle()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();
            var slide = Util.CurrentSlide();

            // 두 선이 선택된 경우 두 선 사이의 각도를 표시
            if (shapes.Count == 2 && shapes.All(s => s.Type == Microsoft.Office.Core.MsoShapeType.msoLine))
            {
                // 현재로서는 line의 point 위치들을 알아낼 방법이 없음
                //DrawAngleBetweenLines(shapes[0], shapes[1], slide);
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
        private static void DrawAnglesForPolygon(PowerPoint.Shape polygon, PowerPoint.Slide slide)
        {
            int nodeCount = polygon.Nodes.Count;
            if (nodeCount < 3) return;

            // 벡터 배열로 변환
            Vector2[] vertices = polygon.Nodes.GetCornerVertices();
            int vertexCount = vertices.Length;

            // 꼭지점이 3개 미만이면 함수 종료
            if (vertexCount < 3) return;

            // 처음 꼭지점과 마지막 꼭지점이 동일할 경우 제거
            var isClosed = false;
            if (vertexCount > 1 && vertices[0] == vertices[vertexCount - 1])
            {
                vertexCount--;
                Array.Resize(ref vertices, vertexCount);
                isClosed = true;
            }

            // 각 꼭지점에서 각도 계산 및 표시
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

                // 각도 표시
                DrawAngleArc(vertices[i].X, vertices[i].Y, 16, v1, v2, angle, slide);
            }
        }
        private static double CalculateAngleBetweenVectors(Vector2 v1, Vector2 v2)
        {
            // 벡터를 정규화
            if (v1.Length() > 0) v1 = Vector2.Normalize(v1);
            if (v2.Length() > 0) v2 = Vector2.Normalize(v2);

            // 두 벡터 사이의 각도 계산 (라디안)
            double dotProduct = Vector2.Dot(v1, v2);
            dotProduct = Math.Min(Math.Max(dotProduct, -1.0), 1.0); // 부동소수점 오류 방지
            double angleRadians = Math.Acos(dotProduct);

            // 라디안을 도(degree)로 변환
            double angleDegrees = angleRadians * (180.0 / Math.PI);

            return angleDegrees;
        }

        private static void DrawAngleArc(float x, float y, float radius, Vector2 v1, Vector2 v2, double angle, PowerPoint.Slide slide)
        {
            // 두 벡터의 방향 계산
            float startAngleRadians = (float)Math.Atan2(v1.Y, v1.X);
            float endAngleRadians = (float)Math.Atan2(v2.Y, v2.X);

            // 라디안을 도로 변환
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

            // 직각(90±0.5도)인지 확인
            bool isRightAngle = Math.Abs(angle - 90) < 0.5;

            if (isRightAngle)
            {
                // 직각 표시 (정사각형)
                float squareSize = radius;

                // 두 벡터의 방향에 맞게 직각 기호 그리기
                // 시작 지점 계산
                float x1 = x;
                float y1 = y;

                // 벡터 1 방향으로 반지름의 절반만큼 이동
                float v1x = (float)Math.Cos(startAngleRadians) * squareSize;
                float v1y = (float)Math.Sin(startAngleRadians) * squareSize;
                x1 += v1x;
                y1 += v1y;

                // 벡터 2 방향으로 반지름의 절반만큼 이동한 지점 계산
                float x2 = x;
                float y2 = y;
                float v2x = (float)Math.Cos(endAngleRadians) * squareSize;
                float v2y = (float)Math.Sin(endAngleRadians) * squareSize;
                x2 += v2x;
                y2 += v2y;

                // 코너 지점 계산 (직각 표시의 꼭지점)
                float cornerX = x1 + v2x;
                float cornerY = y1 + v2y;

                // 직각 기호 그리기 (두 선분)
                slide.Shapes.AddPolyline(new float[,] { { x1, y1 }, { cornerX, cornerY }, { x2, y2 } });
            }
            else
            {
                // 일반 각도 - 원호 그리기
                var shape = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeArc,
                    x,
                    y - radius,
                    radius,
                    radius);

                // 원호의 시작 및 끝 각도 설정
                shape.Adjustments[1] = startAngleDegrees;
                shape.Adjustments[2] = endAngleDegrees;
            }

            // 텍스트 추가
            string angleText = Math.Round(angle, 1).ToString() + "°";

            float width = angleText.Length * 8;
            float height = 20f;

            float halfWidth = width / 2f;
            float halfHeight = height / 2f;

            // 두 벡터의 중간 방향을 계산
            float textX = x + (float)Math.Cos(midAngleRadians) * (radius * 1.2f + halfWidth) - halfWidth;
            float textY = y + (float)Math.Sin(midAngleRadians) * (radius * 1.2f + halfHeight) - halfHeight;

            var textShape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, textX, textY, width, height);

            textShape.TextFrame.TextRange.Text = angleText;
            textShape.TextFrame.TextRange.Font.Size = 12;
            textShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            textShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            textShape.TextFrame2.MarginLeft = 0;
            textShape.TextFrame2.MarginRight = 0;
            textShape.TextFrame2.MarginTop = 0;
            textShape.TextFrame2.MarginBottom = 0;

            // 텍스트 배경 투명하게
            textShape.Fill.Transparency = 1;
            textShape.Line.Transparency = 1;
        }

        public static void DrawPerpendicularAngleLines()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            if (shapes.Count == 1 && shapes[0].Type == Microsoft.Office.Core.MsoShapeType.msoLine)
            {
                var line = shapes[0];
                var slide = Util.CurrentSlide();

                // 선의 중앙점 계산
                float midX = line.Left + line.Width / 2;
                float midY = line.Top + line.Height / 2;

                // 선의 벡터 계산
                Vector2 lineVector = new Vector2(line.Width, line.Height);
                float lineLength = lineVector.Length();

                // 선이 너무 작은 경우 처리 방지
                if (lineLength < 1) return;

                // 수직 벡터 계산
                Vector2 perpVector = new Vector2(-lineVector.Y, lineVector.X);
                perpVector = Vector2.Normalize(perpVector) * (lineLength / 2); // 원래 선 길이의 절반

                // 수직선 그리기
                var perpLine = slide.Shapes.AddLine(
                    midX,
                    midY,
                    midX + perpVector.X,
                    midY + perpVector.Y);

                // 수직선 스타일 설정
                perpLine.Line.ForeColor.RGB = line.Line.ForeColor.RGB;
                perpLine.Line.Weight = line.Line.Weight;
                perpLine.Line.DashStyle = MsoLineDashStyle.msoLineDashDot;

                // 직각 표시 추가
                float squareSize = 10;
                var rightAngleMark = slide.Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                    midX - squareSize / 2,
                    midY - squareSize / 2,
                    squareSize,
                    squareSize);

                rightAngleMark.Line.ForeColor.RGB = 0; // 검은색
                rightAngleMark.Fill.Transparency = 1; // 투명하게

                // 90도 텍스트 추가
                var textShape = slide.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    midX + perpVector.X / 2 - 15,
                    midY + perpVector.Y / 2 - 10,
                    30,
                    20);

                textShape.TextFrame.TextRange.Text = "90°";
                textShape.TextFrame.TextRange.Font.Size = 12;
                textShape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                textShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                // 텍스트 배경 투명하게
                textShape.Fill.Transparency = 1;
                textShape.Line.Transparency = 1;
            }
        }

        public static void DistanceBetweenPoints()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();
            var slide = Util.CurrentSlide();

            foreach (var shape in shapes)
            {
                Vector2[] vertices = shape.Nodes.GetCornerVertices();
                int vertexCount = vertices.Length;

                if (vertexCount < 2)
                {
                    continue;
                }

                // vertices 사이의 거리를 표시

                for (int i = 0; i < vertexCount - 1; i++)
                {
                    int nextIndex = i + 1;

                    Vector2 start = vertices[i];
                    Vector2 end = vertices[nextIndex];

                    // 두 점 사이의 거리 계산 (PowerPoint에서의 픽셀 단위)
                    float pixelDistance = Vector2.Distance(start, end);

                    // 실제 단위로 변환 (pixels / shapeScale = cm)
                    float actualDistance = pixelDistance / shapeScale;

                    // 거리 표시 (mm 단위로 표시, 소수점 2자리까지)
                    string distanceText = Math.Round(actualDistance * 10, 2).ToString() + " mm";

                    // 두 점 사이의 중간 지점 계산
                    Vector2 midPoint = (start + end) / 2;

                    // 두 점을 연결하는 벡터 계산
                    Vector2 lineVector = end - start;

                    // 선의 각도 계산 (라디안)
                    float lineAngle = (float)Math.Atan2(lineVector.Y, lineVector.X);

                    // 각도를 도(degrees)로 변환
                    float lineAngleDegrees = lineAngle * (180f / (float)Math.PI);

                    // 텍스트가 선과 겹치지 않도록 수직 벡터 계산 (선 위에 위치하도록)
                    Vector2 perpVector = new Vector2(lineVector.Y, -lineVector.X);
                    perpVector = Vector2.Normalize(perpVector) * 8f; // 8 픽셀 offset

                    // 거리 텍스트 추가
                    float textWidth = distanceText.Length * 8; // 텍스트 길이에 따른 적절한 너비
                    float textHeight = 20f;

                    // 텍스트 상자 위치 설정 (선 위에 위치하도록 수직 벡터 사용)
                    Vector2 textPosition = midPoint + perpVector;

                    // 텍스트 상자 추가
                    var textBox = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        textPosition.X, // 중앙 정렬
                        textPosition.Y,
                        textWidth,
                        textHeight);

                    // 텍스트 설정
                    textBox.TextFrame.TextRange.Text = distanceText;
                    textBox.TextFrame.TextRange.Font.Size = 12;
                    textBox.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                    textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                    // 텍스트 마진 설정
                    textBox.TextFrame2.MarginLeft = 0;
                    textBox.TextFrame2.MarginRight = 0;
                    textBox.TextFrame2.MarginTop = 0;
                    textBox.TextFrame2.MarginBottom = 0;

                    textBox.Left -= textBox.Width / 2;
                    textBox.Top -= textBox.Height / 2;

                    // 텍스트 회전 (라인과 같은 각도로)
                    textBox.Rotation = lineAngleDegrees;

                    // 양방향 화살표 선 그리기
                    var startOffsetPos = start + perpVector;
                    var endOffsetPos = end + perpVector;

                    var distanceLineA = slide.Shapes.AddLine(startOffsetPos.X, startOffsetPos.Y, 0, 0);
                    var distanceLineB = slide.Shapes.AddLine(endOffsetPos.X, endOffsetPos.Y, 0, 0);
                    distanceLineA.ConnectorFormat.EndConnect(textBox, 2);
                    distanceLineB.ConnectorFormat.EndConnect(textBox, 4);

                    // 양방향 화살표 설정
                    distanceLineA.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                    distanceLineB.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                }
            }
        }

        public static Expression GenerateExpression(string expression)
        {
            expression = Regex.Replace(expression, @"(?<![a-zA-Z])t(?![a-zA-Z(\[])(?!\s*\()", "[t]");
            return new Expression(expression, ExpressiveOptions.IgnoreCaseForParsing);
        }

        public static void AddPathOfExpression(string expX, string expY, string startValue, string endValue, bool isCurve)
        {
            float startValueEvaluated = Convert.ToSingle(new Expression(startValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());
            float endValueEvaluated = Convert.ToSingle(new Expression(endValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());

            expX = expX == "" ? "t" : expX;
            expY = expY == "" ? "t" : expY;

            var expressiveX = GenerateExpression(expX);
            var expressiveY = GenerateExpression(expY);
            var shape = AddPathOfFunction(t => {
                var dict = new Dictionary<string, object> { ["t"] = t };
                return new Vector2(Convert.ToSingle(expressiveX.Evaluate(dict)),
                                   Convert.ToSingle(expressiveY.Evaluate(dict)));
            }, startValueEvaluated, endValueEvaluated, isCurve);

            if (shape != null)
            {
                shape.Tags.Add(PathTypeTagName, isCurve ? CurveTag : PolylineTag);
                shape.Tags.Add(ExpressiveXTagName, expX);
                shape.Tags.Add(ExpressiveYTagName, expY);
                shape.Tags.Add(ExpressiveStartValueTagName, startValue);
                shape.Tags.Add(ExpressiveEndValueTagName, endValue);
            }
        }

        public static PowerPoint.Shape AddPathOfFunction(Func<float, Vector2> func, float startValue, float endValue, bool isCurve)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var slide = Util.CurrentSlide();
            var currentPresentation = Globals.ThisAddIn.Application.ActivePresentation;

            float slideHeight = currentPresentation.PageSetup.SlideHeight;
            float slideWidth = currentPresentation.PageSetup.SlideWidth;

            try
            {
                var vectors = PathOfFunction(func, startValue, endValue, isCurve);
                var pointsShape = new PointsShape(vectors, slideWidth / 2, slideHeight / 2);
                var points = pointsShape.points;
                var shape = isCurve ? slide.Shapes.AddCurve(points) : slide.Shapes.AddPolyline(points);
                shape.Select();
                
                return shape;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return null;
        }
        public static void UpdatePathOfExpression(string expX, string expY, string startValue, string endValue)
        {
            float startValueEvaluated = Convert.ToSingle(new Expression(startValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());
            float endValueEvaluated = Convert.ToSingle(new Expression(endValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());

            expX = expX == "" ? "t" : expX;
            expY = expY == "" ? "t" : expY;

            var expressiveX = GenerateExpression(expX);
            var expressiveY = GenerateExpression(expY);
            var shape = UpdatePathOfFunction(t => {
                var dict = new Dictionary<string, object> { ["t"] = t };
                return new Vector2(Convert.ToSingle(expressiveX.Evaluate(dict)),
                                   Convert.ToSingle(expressiveY.Evaluate(dict)));
            }, startValueEvaluated, endValueEvaluated);

            if (shape != null)
            {
                shape.Tags.Add(ExpressiveXTagName, expX);
                shape.Tags.Add(ExpressiveYTagName, expY);
                shape.Tags.Add(ExpressiveStartValueTagName, startValue);
                shape.Tags.Add(ExpressiveEndValueTagName, endValue);
            }
        }

        public static PowerPoint.Shape UpdatePathOfFunction(Func<float, Vector2> func, float startValue, float endValue)
        {
            // Globals.ThisAddIn.Application.StartNewUndoEntry();
            var slide = Util.CurrentSlide();
            var currentPresentation = Globals.ThisAddIn.Application.ActivePresentation;
            var selectedShapes = Util.ListSelectedShapes();
            if (selectedShapes.Count == 1)
            {
                var shape = selectedShapes[0];
                var tag = shape.Tags[PathTypeTagName];
                if (tag == CurveTag || tag == PolylineTag)
                {
                    try
                    {
                        bool isCurve = tag == CurveTag;
                        var vectors = PathOfFunction(func, startValue, endValue, isCurve);
                        var pointsShape = new PointsShape(vectors, 0f, 0f);
                        var points = pointsShape.points;

                        // Method 1 (Too slow, Not working)
                        // int nodeCount = shape.Nodes.Count;
                        // for (int i = 0; i < nodeCount; i++)
                        // {
                        //     shape.Nodes.SetPosition(i + 1, points[i, 0], points[i, 1]);
                        //     // var p = shape.Nodes[i + 1].Points;
                        //     // p[1, 1] = points[i, 0];
                        //     // p[1, 2] = points[i, 1];
                        // }

                        // Method 2
                        var newShape = isCurve ? slide.Shapes.AddCurve(points) : slide.Shapes.AddPolyline(points);
                        
                        newShape.Match(shape, true);
                        newShape.Tags.Add(PathTypeTagName, tag);
                        newShape.Select();
                        shape.Delete();

                        return newShape;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            return null;
        }

        public static Vector2[] PathOfFunction(Func<float, Vector2> func, float startValue, float endValue, bool addControlPoints = false)
        {
            float rangeValue = endValue - startValue;
            var numPoints = 100;
            var initVectors = new Vector2[numPoints];
            for (int t = 0; t < numPoints; t++)
            {
                var f = ((float)t) / (numPoints - 1);
                initVectors[t] = func(startValue + f * rangeValue) * shapeScale;
            }
            
            Vector2[] vectors;
            if (addControlPoints)
            {
                // Add control points
                vectors = new Vector2[numPoints];
                for (int t = 0; t < numPoints; t++)
                {
                    if (t % 3 == 0 || t == 1 || t == numPoints - 2)
                    {
                        vectors[t] = initVectors[t];
                        continue;
                    }

                    Vector2 v1 = initVectors[t];
                    Vector2 v2 = new Vector2();
                    Vector2 v0 = new Vector2();
                    if (t % 3 == 1)
                    {
                        v2 = initVectors[t - 2];
                        v0 = initVectors[t - 1];
                    }
                    if (t % 3 == 2)
                    {
                        v2 = initVectors[t + 2];
                        v0 = initVectors[t + 1];
                    }
                    vectors[t] = v0 + (v1 - v2) / 2;
                }
            }
            else
            {
                vectors = initVectors;
            }

            return vectors;
        }
    }
}
