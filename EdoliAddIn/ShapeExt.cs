using System;
using System.Collections.Generic;
using System.Numerics;
using Microsoft.Office.Core;
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

        public static float RotationRad(this PowerPoint.Shape shape)
        {
            return shape.Rotation * MathExt.degToRad;
        }

        // Visual size, position. Real visual bound of shapes.
        public static float Width(this PowerPoint.Shape shape)
        {
            float rotation = shape.RotationRad();
            return (float)(Math.Abs(Math.Cos(rotation)) * shape.Width + Math.Abs(Math.Sin(rotation)) * shape.Height);
        }

        public static float Height(this PowerPoint.Shape shape)
        {
            float rotation = shape.RotationRad();
            return (float)(Math.Abs(Math.Sin(rotation)) * shape.Width + Math.Abs(Math.Cos(rotation)) * shape.Height);
        }

        public static Vector2 Size(this PowerPoint.Shape shape)
        {
            return new Vector2(Width(shape), Height(shape));
        }

        public static void SetSize(this PowerPoint.Shape shape, float width, float height)
        {
            float rotation = shape.RotationRad();
            float cps = (float)(Math.Abs(Math.Cos(rotation)) + Math.Abs(Math.Sin(rotation)));
            float cms = (float)(Math.Abs(Math.Cos(rotation)) - Math.Abs(Math.Sin(rotation)));
            float wph = (width + height) / cps;
            float wmh = (width - height) / cms;
            shape.Width = (wph + wmh) / 2.0f;
            shape.Height = (wph - wmh) / 2.0f;
        }

        public static void SetSize(this PowerPoint.Shape shape, Vector2 size)
        {
            SetSize(shape, size.X, size.Y);
        }

        public static float Left(this PowerPoint.Shape shape)
        {
            float offset = (Width(shape) - shape.Width) / 2;
            return shape.Left - offset;
        }

        public static void SetLeft(this PowerPoint.Shape shape, float value)
        {
            float offset = (Width(shape) - shape.Width) / 2;
            shape.Left = value + offset;
        }

        public static float Right(this PowerPoint.Shape shape)
        {
            float width = Width(shape);
            float offset = (width - shape.Width) / 2;
            return shape.Left + width - offset;
        }

        public static void SetRight(this PowerPoint.Shape shape, float value)
        {
            float width = Width(shape);
            float offset = (width - shape.Width) / 2;
            shape.Left = value - width + offset;
        }

        public static float Top(this PowerPoint.Shape shape)
        {
            float offset = (Height(shape) - shape.Height) / 2;
            return shape.Top - offset;
        }

        public static void SetTop(this PowerPoint.Shape shape, float value)
        {
            float offset = (Height(shape) - shape.Height) / 2;
            shape.Top = value + offset;
        }

        public static float Bottom(this PowerPoint.Shape shape)
        {
            float height = Height(shape);
            float offset = (height - shape.Height) / 2;
            return shape.Top + height - offset;
        }

        public static void SetBottom(this PowerPoint.Shape shape, float value)
        {
            float height = Height(shape);
            float offset = (height - shape.Height) / 2;
            shape.Top = value - height + offset;
        }

        public static float DistanceOfShapes(PowerPoint.Shape shapeA, PowerPoint.Shape shapeB)
        {
            var left1 = shapeA.Left();
            var right1 = shapeA.Right();
            var top1 = shapeA.Top();
            var bottom1 = shapeA.Bottom();

            var left2 = shapeB.Left();
            var right2 = shapeB.Right();
            var top2 = shapeB.Top();
            var bottom2 = shapeB.Bottom();

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
                var p1 = shapeA.Position(anchor);
                var p2 = shapeB.Position(anchor);
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
                var p2 = shapeB.Position(anchorB);
                return Util.RectanglePointDistance(shapeA.Left(), shapeA.Top(), shapeA.Right(), shapeA.Bottom(), p2.X, p2.Y);
            }
            else if (anchorB == Anchor.None)
            {
                var p1 = shapeA.Position(anchorA);
                return Util.RectanglePointDistance(shapeB.Left(), shapeB.Top(), shapeB.Right(), shapeB.Bottom(), p1.X, p1.Y);
            }
            else
            {
                var p1 = shapeA.Position(anchorA);
                var p2 = shapeB.Position(anchorB);
                return Vector2.Distance(p1, p2);
            }
        }

        /// <summary>
        /// Get the position of the shape based on the specified anchor point.
        /// based on real visual bound of shapes.
        /// </summary>
        public static Vector2 Position(this PowerPoint.Shape shape, Anchor anchor)
        {
            switch (anchor)
            {
                case Anchor.Left:
                    return new Vector2(shape.Left(), shape.Top() + shape.Height() / 2);
                case Anchor.Right:
                    return new Vector2(shape.Left() + shape.Width(), shape.Top() + shape.Height() / 2);
                case Anchor.Top:
                    return new Vector2(shape.Left() + shape.Width() / 2, shape.Top());
                case Anchor.Bottom:
                    return new Vector2(shape.Left() + shape.Width() / 2, shape.Top() + shape.Height());
                case Anchor.TopLeft:
                    return new Vector2(shape.Left(), shape.Top());
                case Anchor.TopRight:
                    return new Vector2(shape.Left() + shape.Width(), shape.Top());
                case Anchor.BottomLeft:
                    return new Vector2(shape.Left(), shape.Top() + shape.Height());
                case Anchor.BottomRight:
                    return new Vector2(shape.Left() + shape.Width(), shape.Top() + shape.Height());
                case Anchor.Center:
                    return new Vector2(shape.Left() + shape.Width() / 2, shape.Top() + shape.Height() / 2);
                default:
                    throw new ArgumentException("Invalid anchor point", nameof(anchor));
            }
        }

        /// <summary>
        /// Get the position of the shape based on the specified anchor point.
        /// based on original bound of shapes.
        /// </summary>
        public static Vector2 PositionRot(this PowerPoint.Shape shape, Anchor anchor)
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
                    newTop = position.Y - shape.Height() / 2;
                    break;
                case Anchor.Right:
                    newLeft = position.X - shape.Width();
                    newTop = position.Y - shape.Height() / 2;
                    break;
                case Anchor.Top:
                    newLeft = position.X - shape.Width() / 2;
                    newTop = position.Y;
                    break;
                case Anchor.Bottom:
                    newLeft = position.X - shape.Width() / 2;
                    newTop = position.Y - shape.Height();
                    break;
                case Anchor.TopLeft:
                    newLeft = position.X;
                    newTop = position.Y;
                    break;
                case Anchor.TopRight:
                    newLeft = position.X - shape.Width();
                    newTop = position.Y;
                    break;
                case Anchor.BottomLeft:
                    newLeft = position.X;
                    newTop = position.Y - shape.Height();
                    break;
                case Anchor.BottomRight:
                    newLeft = position.X - shape.Width();
                    newTop = position.Y - shape.Height();
                    break;
                case Anchor.Center:
                    newLeft = position.X - shape.Width() / 2;
                    newTop = position.Y - shape.Height() / 2;
                    break;
                default:
                    throw new ArgumentException("Invalid anchor point", nameof(anchor));
            }

            shape.SetLeft(newLeft);
            shape.SetTop(newTop);
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
            return shapes.MaxBy(shape => shape.Right());
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
            return shapes.MaxBy(shape => shape.Bottom());
        }

        public static Anchor GetRelativePos(this PowerPoint.Shape shape, PowerPoint.Shape other, float epsilon=0)
        {
            float left1 = shape.Left();
            float top1 = shape.Top();
            float right1 = shape.Right();
            float bottom1 = shape.Bottom();

            float left2 = other.Left();
            float top2 = other.Top();
            float right2 = other.Right();
            float bottom2 = other.Bottom();

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

            for (int i = 0; i < shape.ConnectionSiteCount; i++)
            {
                line.ConnectorFormat.EndConnect(shape, i + 1);
                float x = line.Left + line.Width;
                float y = line.Top + line.Height;

                connectors.Add(new Vector2(x, y));
            }

            line.Delete();

            return connectors.ToArray();
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
                var connectors = shape.GetConnectors();
                return connectors;
            }
            else if (shape.Type == MsoShapeType.msoAutoShape)
            {
            }
            return new Vector2[0];
        }
    }
}
