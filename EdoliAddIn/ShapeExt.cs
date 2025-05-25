using System;
using System.Collections.Generic;
using System.Drawing;
using System.Numerics;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{
    public static class ShapeExt
    {
        public enum Anchor
        {
            Top, Bottom, Left, Right,
            TopLeft, TopRight, BottomLeft, BottomRight,
            Center,
            None
        }

        public static Anchor Opposite(this Anchor anchor)
        {
            switch (anchor)
            {
                case Anchor.Top: return Anchor.Bottom;
                case Anchor.Bottom: return Anchor.Top;
                case Anchor.Left: return Anchor.Right;
                case Anchor.Right: return Anchor.Left;
                case Anchor.TopLeft: return Anchor.BottomRight;
                case Anchor.TopRight: return Anchor.BottomLeft;
                case Anchor.BottomLeft: return Anchor.TopRight;
                case Anchor.BottomRight: return Anchor.TopLeft;
                case Anchor.Center: return Anchor.Center;
                case Anchor.None: return Anchor.None;
                default: return Anchor.None;

            }
        }

        /// <summary>
        /// Shape이 group인 경우 안에 item들에 대해서 action을 수행함 
        /// </summary>
        public static void DoRecur(this PowerPoint.Shape shape, Action<PowerPoint.Shape> action)
        {
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape item in shape.GroupItems)
                {
                    item.DoRecur(action);
                }
            }
            else
            {
                action(shape);
            }
        }

        /// <summary>
        /// 현재 <paramref name="shape"/>의 크기, 회전, 위치를 <paramref name="other"/>와 일치시킵니다.
        /// <paramref name="keepAspectRatio"/>가 true이면, 가로 비율에 맞춰 세로 크기를 조정하여 종횡비를 유지합니다.
        /// 위치는 회전된 상태에서의 좌상단 기준으로 맞춥니다.
        /// </summary>
        /// <param name="shape">변경할 대상 도형</param>
        /// <param name="other">기준이 되는 도형</param>
        /// <param name="keepAspectRatio">종횡비 유지 여부</param>
        public static void Match(this PowerPoint.Shape shape, PowerPoint.Shape other, bool keepAspectRatio)
        {
            shape.Rotation = other.Rotation;
            if (keepAspectRatio)
            {
                shape.Height = shape.Height * other.Width / shape.Width;
            }
            else
            {
                shape.Height = other.Height;
            }
            shape.Width = other.Width;

            // Match top left
            float rotationRad = shape.RotationRad();
            Vector2 shapeTopLeft = new Vector2(-shape.Width / 2, -shape.Height / 2);
            Vector2 otherTopLeft = new Vector2(-other.Width / 2, -other.Height / 2);

            Vector2 offset = otherTopLeft.Rotate(rotationRad) - otherTopLeft
                - (shapeTopLeft.Rotate(rotationRad) - shapeTopLeft);

            shape.Left = other.Left + offset.X;
            shape.Top = other.Top + offset.Y;
        }
        public static RectangleF Rect(this PowerPoint.Shape shape)
        {
            return new RectangleF(shape.Left, shape.Top, shape.Width, shape.Height);
        }
        public static RectangleF VisualRect(this PowerPoint.Shape shape)
        {
            return new RectangleF(shape.VisualLeft(), shape.VisualTop(), shape.VisualWidth(), shape.VisualHeight());
        }

        public static float RotationRad(this PowerPoint.Shape shape)
        {
            return shape.Rotation * MathExt.degToRad;
        }

        // Visual size, position. Real visual bound of shapes.
        public static float VisualWidth(this PowerPoint.Shape shape)
        {
            float rotation = shape.RotationRad();
            return (float)(Math.Abs(Math.Cos(rotation)) * shape.Width + Math.Abs(Math.Sin(rotation)) * shape.Height);
        }

        public static float VisualHeight(this PowerPoint.Shape shape)
        {
            float rotation = shape.RotationRad();
            return (float)(Math.Abs(Math.Sin(rotation)) * shape.Width + Math.Abs(Math.Cos(rotation)) * shape.Height);
        }

        public static Vector2 VisualSize(this PowerPoint.Shape shape)
        {
            return new Vector2(VisualWidth(shape), VisualHeight(shape));
        }

        public static void SetVisualSize(this PowerPoint.Shape shape, float width, float height)
        {
            float rotation = shape.RotationRad();
            float cps = (float)(Math.Abs(Math.Cos(rotation)) + Math.Abs(Math.Sin(rotation)));
            float cms = (float)(Math.Abs(Math.Cos(rotation)) - Math.Abs(Math.Sin(rotation)));
            float wph = (width + height) / cps;
            float wmh = (width - height) / cms;
            shape.Width = (wph + wmh) / 2.0f;
            shape.Height = (wph - wmh) / 2.0f;
        }

        public static void SetVisualSize(this PowerPoint.Shape shape, Vector2 size)
        {
            SetVisualSize(shape, size.X, size.Y);
        }

        public static float VisualLeft(this PowerPoint.Shape shape)
        {
            float offset = (VisualWidth(shape) - shape.Width) / 2;
            return shape.Left - offset;
        }

        public static void SetVisualLeft(this PowerPoint.Shape shape, float value)
        {
            float offset = (VisualWidth(shape) - shape.Width) / 2;
            shape.Left = value + offset;
        }

        public static float VisualRight(this PowerPoint.Shape shape)
        {
            float width = VisualWidth(shape);
            float offset = (width - shape.Width) / 2;
            return shape.Left + width - offset;
        }

        public static void SetVisualRight(this PowerPoint.Shape shape, float value)
        {
            float width = VisualWidth(shape);
            float offset = (width - shape.Width) / 2;
            shape.Left = value - width + offset;
        }

        public static float VisualTop(this PowerPoint.Shape shape)
        {
            float offset = (VisualHeight(shape) - shape.Height) / 2;
            return shape.Top - offset;
        }

        public static void SetVisualTop(this PowerPoint.Shape shape, float value)
        {
            float offset = (VisualHeight(shape) - shape.Height) / 2;
            shape.Top = value + offset;
        }

        public static float VisualBottom(this PowerPoint.Shape shape)
        {
            float height = VisualHeight(shape);
            float offset = (height - shape.Height) / 2;
            return shape.Top + height - offset;
        }

        public static void SetVisualBottom(this PowerPoint.Shape shape, float value)
        {
            float height = VisualHeight(shape);
            float offset = (height - shape.Height) / 2;
            shape.Top = value - height + offset;
        }

        public static float DistanceOfShapes(PowerPoint.Shape shapeA, PowerPoint.Shape shapeB)
        {
            var left1 = shapeA.VisualLeft();
            var right1 = shapeA.VisualRight();
            var top1 = shapeA.VisualTop();
            var bottom1 = shapeA.VisualBottom();

            var left2 = shapeB.VisualLeft();
            var right2 = shapeB.VisualRight();
            var top2 = shapeB.VisualTop();
            var bottom2 = shapeB.VisualBottom();

            return Util.RectangleDistance(left1, top1, right1, bottom1, left2, top2, right2, bottom2);
        }

        public static float DistanceOfShapes(PowerPoint.Shape shapeA, PowerPoint.Shape shapeB, Anchor anchor)
        {
            if (anchor == Anchor.None)
            {
                return DistanceOfShapes(shapeA, shapeB);
            } 
            else
            {
                var p1 = shapeA.VisualPosition(anchor);
                var p2 = shapeB.VisualPosition(anchor);
                return Vector2.Distance(p1, p2);
            }
        }
        public static float DistanceOfShapes(PowerPoint.Shape shapeA, PowerPoint.Shape shapeB, Anchor anchorA, Anchor anchorB)
        {
            if (anchorA == Anchor.None && anchorB == Anchor.None)
            {
                return DistanceOfShapes(shapeA, shapeB);
            }
            else if (anchorA == Anchor.None)
            {
                var p2 = shapeB.VisualPosition(anchorB);
                return Util.RectanglePointDistance(shapeA.VisualLeft(), shapeA.VisualTop(), shapeA.VisualRight(), shapeA.VisualBottom(), p2.X, p2.Y);
            }
            else if (anchorB == Anchor.None)
            {
                var p1 = shapeA.VisualPosition(anchorA);
                return Util.RectanglePointDistance(shapeB.VisualLeft(), shapeB.VisualTop(), shapeB.VisualRight(), shapeB.VisualBottom(), p1.X, p1.Y);
            }
            else
            {
                var p1 = shapeA.VisualPosition(anchorA);
                var p2 = shapeB.VisualPosition(anchorB);
                return Vector2.Distance(p1, p2);
            }
        }

        /// <summary>
        /// Get the position of the shape based on the specified anchor point.
        /// based on real visual bound of shapes.
        /// </summary>
        public static Vector2 VisualPosition(this PowerPoint.Shape shape, Anchor anchor)
        {
            switch (anchor)
            {
                case Anchor.Left:
                    return new Vector2(shape.VisualLeft(), shape.VisualTop() + shape.VisualHeight() / 2);
                case Anchor.Right:
                    return new Vector2(shape.VisualLeft() + shape.VisualWidth(), shape.VisualTop() + shape.VisualHeight() / 2);
                case Anchor.Top:
                    return new Vector2(shape.VisualLeft() + shape.VisualWidth() / 2, shape.VisualTop());
                case Anchor.Bottom:
                    return new Vector2(shape.VisualLeft() + shape.VisualWidth() / 2, shape.VisualTop() + shape.VisualHeight());
                case Anchor.TopLeft:
                    return new Vector2(shape.VisualLeft(), shape.VisualTop());
                case Anchor.TopRight:
                    return new Vector2(shape.VisualLeft() + shape.VisualWidth(), shape.VisualTop());
                case Anchor.BottomLeft:
                    return new Vector2(shape.VisualLeft(), shape.VisualTop() + shape.VisualHeight());
                case Anchor.BottomRight:
                    return new Vector2(shape.VisualLeft() + shape.VisualWidth(), shape.VisualTop() + shape.VisualHeight());
                case Anchor.Center:
                    return new Vector2(shape.VisualLeft() + shape.VisualWidth() / 2, shape.VisualTop() + shape.VisualHeight() / 2);
                default:
                    throw new ArgumentException("Invalid anchor point", nameof(anchor));
            }
        }

        /// <summary>
        /// Get the position of the shape based on the specified anchor point.
        /// based on original bound of shapes.
        /// </summary>
        public static Vector2 VertexPosition(this PowerPoint.Shape shape, Anchor anchor)
        {
            Vector2 center = new Vector2(shape.Left + shape.Width / 2, shape.Top + shape.Height / 2);
            Vector2 vector;
            switch (anchor)
            {
                case Anchor.Left:
                    vector = new Vector2(shape.Left, shape.Top + shape.Height / 2);
                    break;
                case Anchor.Right:
                    vector = new Vector2(shape.Left + shape.Width, shape.Top + shape.Height / 2);
                    break;
                case Anchor.Top:
                    vector = new Vector2(shape.Left + shape.Width / 2, shape.Top);
                    break;
                case Anchor.Bottom:
                    vector = new Vector2(shape.Left + shape.Width / 2, shape.Top + shape.Height);
                    break;
                case Anchor.TopLeft:
                    vector = new Vector2(shape.Left, shape.Top);
                    break;
                case Anchor.TopRight:
                    vector = new Vector2(shape.Left + shape.Width, shape.Top);
                    break;
                case Anchor.BottomLeft:
                    vector = new Vector2(shape.Left, shape.Top + shape.Height);
                    break;
                case Anchor.BottomRight:
                    vector = new Vector2(shape.Left + shape.Width, shape.Top + shape.Height);
                    break;
                case Anchor.Center:
                    return center;
                default:
                    throw new ArgumentException("Invalid anchor point", nameof(anchor));
            }
            return center + (vector - center).RotateDeg(shape.Rotation);
        }

        public static void SetPosition(this PowerPoint.Shape shape, Vector2 position, Anchor anchor)
        {
            float newLeft, newTop;

            switch (anchor)
            {
                case Anchor.Left:
                    newLeft = position.X;
                    newTop = position.Y - shape.VisualHeight() / 2;
                    break;
                case Anchor.Right:
                    newLeft = position.X - shape.VisualWidth();
                    newTop = position.Y - shape.VisualHeight() / 2;
                    break;
                case Anchor.Top:
                    newLeft = position.X - shape.VisualWidth() / 2;
                    newTop = position.Y;
                    break;
                case Anchor.Bottom:
                    newLeft = position.X - shape.VisualWidth() / 2;
                    newTop = position.Y - shape.VisualHeight();
                    break;
                case Anchor.TopLeft:
                    newLeft = position.X;
                    newTop = position.Y;
                    break;
                case Anchor.TopRight:
                    newLeft = position.X - shape.VisualWidth();
                    newTop = position.Y;
                    break;
                case Anchor.BottomLeft:
                    newLeft = position.X;
                    newTop = position.Y - shape.VisualHeight();
                    break;
                case Anchor.BottomRight:
                    newLeft = position.X - shape.VisualWidth();
                    newTop = position.Y - shape.VisualHeight();
                    break;
                case Anchor.Center:
                    newLeft = position.X - shape.VisualWidth() / 2;
                    newTop = position.Y - shape.VisualHeight() / 2;
                    break;
                default:
                    throw new ArgumentException("Invalid anchor point", nameof(anchor));
            }

            shape.SetVisualLeft(newLeft);
            shape.SetVisualTop(newTop);
        }

        public static PowerPoint.Shape FindNearestShape(this PowerPoint.Shape shape, List<PowerPoint.Shape> shapes, Anchor anchor)
        {
            return shapes.MinBy(s => DistanceOfShapes(shape, s, anchor));
        }

        public static PowerPoint.Shape FindNearestShape(this PowerPoint.Shape shape, List<PowerPoint.Shape> shapes, Anchor anchorA, Anchor anchorB)
        {
            return shapes.MinBy(s => DistanceOfShapes(shape, s, anchorA, anchorB));
        }

        public static PowerPoint.Shape GetLeftMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }
            return shapes.MinBy(shape => shape.Left);
        }

        public static PowerPoint.Shape GetRightMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }
            return shapes.MaxBy(shape => shape.VisualRight());
        }

        public static PowerPoint.Shape GetTopMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }
            return shapes.MinBy(shape => shape.Top);
        }

        public static PowerPoint.Shape GetBottomMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }
            return shapes.MaxBy(shape => shape.VisualBottom());
        }

        public static Anchor GetRelativePos(this PowerPoint.Shape shape, PowerPoint.Shape other, float epsilon=0)
        {
            float left1 = shape.VisualLeft();
            float top1 = shape.VisualTop();
            float right1 = shape.VisualRight();
            float bottom1 = shape.VisualBottom();

            float left2 = other.VisualLeft();
            float top2 = other.VisualTop();
            float right2 = other.VisualRight();
            float bottom2 = other.VisualBottom();

            if (right2 < left1 - epsilon && bottom2 < top1 - epsilon)
            {
                return Anchor.TopLeft;
            }
            if (left2 > right1 + epsilon && bottom2 < top1 - epsilon)
            {
                return Anchor.TopRight;
            }
            if (right2 < left1 - epsilon && top2 > bottom1 + epsilon)
            {
                return Anchor.BottomLeft;
            }
            if (left2 > right1 + epsilon && top2 > bottom1 + epsilon)
            {
                return Anchor.BottomRight;
            }
            if (right2 < left1)
            {
                return Anchor.Left;
            }
            if (left2 > right1)
            {
                return Anchor.Right;
            }
            if (bottom2 < top1)
            {
                return Anchor.Top;
            }
            if (top2 > bottom1)
            {
                return Anchor.Bottom;
            }
            return Anchor.None;
        }


        public static List<Vector2> ClosePolygon(this List<Vector2> points)
        {
            if (points.Count == 0 || points[0] == points[points.Count - 1])
            {
                return points;
            }
            points.Add(points[0]);
            return points;
        }

        public static Vector2[] ClosePolygon(this Vector2[] points)
        {
            if (points.Length == 0 || points[0] == points[points.Length - 1])
            {
                return points;
            }
            var newArray = new Vector2[points.Length + 1];
            Array.Copy(points, newArray, points.Length);
            newArray[points.Length] = points[0];
            return newArray;
        }

        /// <summary>
        /// Nodes에서 control 포인트들은 제외하고 실제 꼭지점들만 추출
        /// </summary>
        public static Vector2[] GetCornerVertices(this PowerPoint.ShapeNodes nodes)
        {
            int nodeCount = nodes.Count;

            // 실제 꼭지점 노드만 추출
            List<Vector2> vertexList = new List<Vector2>();
            for (int i = 1; i <= nodeCount; i++)
            {
                var node = nodes[i];
                var segmentType = node.SegmentType;
                var editingType = node.EditingType;

                if (i > 1 && segmentType == MsoSegmentType.msoSegmentCurve)
                {
                    // 첫 노드는 무조건 꼭지점
                    // 그 이후로 Curve가 발견되면 2개 뒤의 노드가 꼭지점 (2개는 컨트롤 포인트)
                    var points = nodes[i + 2].Points;
                    vertexList.Add(new Vector2((float)points[1, 1], (float)points[1, 2]));
                    i += 2;
                }
                else
                {
                    var points = node.Points;
                    vertexList.Add(new Vector2((float)points[1, 1], (float)points[1, 2]));
                }
            }

            return vertexList.ToArray();
        }

        /// <summary>
        /// 임시로 line을 만들고 line을 connector에 연결해 가면서 connector의 위치를 추정
        /// </summary>
        public static Vector2[] GetConnectors(this PowerPoint.Shape shape)
        {
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            var connectors = new List<Vector2>();

            var line = slide.Shapes.AddLine(-1000, -1000, 10, 10);

            for (int i = 1; i <= shape.ConnectionSiteCount; i++)
            {
                line.ConnectorFormat.EndConnect(shape, i);
                float x = line.Left + line.Width;
                float y = line.Top + line.Height;

                connectors.Add(new Vector2(x, y));
            }

            line.Delete();

            return connectors.ToArray();
        }

        public static Vector2[] GetConnectors(this PowerPoint.Shape shape, int[] indices = null)
        {
            Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            var connectors = new List<Vector2>();

            var line = slide.Shapes.AddLine(-1000, -1000, 10, 10);

            foreach (int i in indices)
            {
                line.ConnectorFormat.EndConnect(shape, i);
                float x = line.Left + line.Width;
                float y = line.Top + line.Height;

                connectors.Add(new Vector2(x, y));
            }

            line.Delete();

            return connectors.ToArray();
        }

        public static Vector2[] GetLineEndPoints(this PowerPoint.Shape shape)
        {
            // TODO: line에 rotation이 있는 경우가 있음. 이런 경우 제대로 작동하지 않음

            if (shape.Connector != MsoTriState.msoTrue) {
                throw new Exception("Shape is not a connector");
            }

            Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            var line = slide.Shapes.AddLine(-100, -100, -101, -101);

            RectangleF rect = shape.VisualRect();

            PowerPoint.Shape beginConnectedShape = null;
            PowerPoint.Shape endConnectedShape = null;
            int beginConnectionSite = 0;
            int endConnectionSite = 0;
            if (shape.ConnectorFormat.BeginConnected == MsoTriState.msoTrue)
            {
                beginConnectedShape = shape.ConnectorFormat.BeginConnectedShape;
                beginConnectionSite = shape.ConnectorFormat.BeginConnectionSite;
            }
            if (shape.ConnectorFormat.EndConnected == MsoTriState.msoTrue)
            {
                endConnectedShape = shape.ConnectorFormat.EndConnectedShape;
                endConnectionSite = shape.ConnectorFormat.EndConnectionSite;
            }

            shape.ConnectorFormat.BeginConnect(line, 1);

            var right = shape.Left + shape.Width;
            var bottom = shape.Top + shape.Height;

            shape.Left = rect.Left;
            shape.Top = rect.Top;
            shape.Width = rect.Width;
            shape.Height = rect.Height;

            Vector2[] points = new Vector2[2];

            float epsilon = 1e-3f;

            if (right.Approximately(rect.Right, epsilon) && bottom.Approximately(rect.Bottom, epsilon))
            {
                // ↘
                points[0] = new Vector2(rect.Left, rect.Top);
                points[1] = new Vector2(rect.Right, rect.Bottom);
            }
            else if (right.Approximately(rect.Left, epsilon) && bottom.Approximately(rect.Top, epsilon))
            {
                // ↖
                shape.Flip(MsoFlipCmd.msoFlipVertical);
                shape.Flip(MsoFlipCmd.msoFlipHorizontal);
                points[0] = new Vector2(rect.Right, rect.Bottom);
                points[1] = new Vector2(rect.Left, rect.Top);
            }
            else if (right.Approximately(rect.Right, epsilon) && bottom.Approximately(rect.Top, epsilon))
            {
                // ↗
                shape.Flip(MsoFlipCmd.msoFlipVertical);
                points[0] = new Vector2(rect.Left, rect.Bottom);
                points[1] = new Vector2(rect.Right, rect.Top);
            }
            else if (right.Approximately(rect.Left, epsilon) && bottom.Approximately(rect.Bottom, epsilon))
            {
                // ↙
                shape.Flip(MsoFlipCmd.msoFlipHorizontal);
                points[0] = new Vector2(rect.Right, rect.Top);
                points[1] = new Vector2(rect.Left, rect.Bottom);
            }
            else
            {
                throw new Exception("Connector check failed in GetLineEndPoints");
            }

            if (beginConnectedShape != null)
            {
                shape.ConnectorFormat.BeginConnect(beginConnectedShape, beginConnectionSite);
            }
            else
            {
                shape.ConnectorFormat.BeginDisconnect();
            }


            if (endConnectedShape != null)
            {
                shape.ConnectorFormat.EndConnect(endConnectedShape, endConnectionSite);
            }

            line.Delete();

            return points;
        }

        /// <summary>
        /// 일부 shape들에 대해 vertices 위치들을 connector로 부터 추정할 수 있음
        /// </summary>
        public static Vector2[] GetVertices(this PowerPoint.Shape shape)
        {
            if (shape.Type == MsoShapeType.msoFreeform)
            {
                return shape.Nodes.GetCornerVertices();
            }
            else if (shape.Type == MsoShapeType.msoLine)
            {
                return shape.GetConnectors();
            }
            else if (shape.Type == MsoShapeType.msoAutoShape)
            {
                switch (shape.AutoShapeType)
                {
                    case MsoAutoShapeType.msoShapeRectangle:
                        return new Vector2[]
                        {
                        shape.VertexPosition(Anchor.TopLeft),
                        shape.VertexPosition(Anchor.TopRight),
                        shape.VertexPosition(Anchor.BottomRight),
                        shape.VertexPosition(Anchor.BottomLeft),
                        shape.VertexPosition(Anchor.TopLeft)
                        };
                    case MsoAutoShapeType.msoShapeIsoscelesTriangle:
                    case MsoAutoShapeType.msoShapeRightTriangle:
                        return shape.GetConnectors(new int[] { 1, 5, 3, 1 });
                    case MsoAutoShapeType.msoShapeDiamond:
                        return shape.GetConnectors().ClosePolygon();
                    case MsoAutoShapeType.msoShapeRegularPentagon:
                        return shape.GetConnectors(new int[] { 1, 2, 3, 5, 6 });
                    case MsoAutoShapeType.msoShapeHexagon:
                        return shape.GetConnectors().ClosePolygon();
                }

                if (shape.Connector == MsoTriState.msoTrue)
                {
                    return shape.GetLineEndPoints();
                }
            }
            return new Vector2[0];
        }
    }
}
