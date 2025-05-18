using Expressive;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
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
