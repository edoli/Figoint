using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Core = Microsoft.Office.Core;
using static EdoliAddIn.ShapeExt;
using System.Numerics;

namespace EdoliAddIn
{
    public class AlignTool
    {
        public enum Align
        {
            Top, Bottom, Left, Right
        }

        public static void AlignLeft()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tLeft = lastShape.VisualLeft();

                foreach (var shape in shapes)
                {
                    shape.SetVisualLeft(tLeft);
                }
            }
        }
        public static void AlignRight()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tRight = lastShape.VisualRight();

                foreach (var shape in shapes)
                {
                    shape.SetVisualRight(tRight);
                }
            }
        }
        public static void AlignTop()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tTop = lastShape.VisualTop();

                foreach (var shape in shapes)
                {
                    shape.SetVisualTop(tTop);
                }
            }
        }
        public static void AlignBottom()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tBottom = lastShape.VisualBottom();

                foreach (var shape in shapes)
                {
                    shape.SetVisualBottom(tBottom);
                }
            }
        }

        public static void AlignLeftOf()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tLeft = lastShape.VisualLeft();

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.SetVisualRight(tLeft);
                }
            }
        }
        public static void AlignRightOf()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tRight = lastShape.VisualRight();

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.SetVisualLeft(tRight);
                }
            }
        }
        public static void AlignTopOf()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tTop = lastShape.VisualTop();

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.SetVisualBottom(tTop);
                }
            }
        }
        public static void AlignBottomOf()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tBottom = lastShape.VisualBottom();

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.SetVisualTop(tBottom);
                }
            }
        }

        public static void AlignCenter()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float hCenter = lastShape.Left + lastShape.Width / 2;
                float vCenter = lastShape.Top + lastShape.Height / 2;

                foreach (var shape in shapes)
                {
                    shape.Left = hCenter - shape.Width / 2;
                    shape.Top = vCenter - shape.Height / 2;
                }
            }
        }
        public static void AlignCenterHorizontal()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float hCenter = lastShape.Left + lastShape.Width / 2;

                foreach (var shape in shapes)
                {
                    shape.Left = hCenter - shape.Width / 2;
                }
            }
        }
        public static void AlignCenterVertical()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float vCenter = lastShape.Top + lastShape.Height / 2;

                foreach (var shape in shapes)
                {
                    shape.Top = vCenter - shape.Height / 2;
                }
            }
        }

        public static void AlignInRow()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                shapes.Sort((shapeA, shapeB) => Math.Sign(shapeA.Left - shapeB.Left));

                var leftMostShape = ShapeExt.GetLeftMostShape(shapes);
                var rightMostShape = ShapeExt.GetRightMostShape(shapes);

                var left = leftMostShape.Left;
                var right = rightMostShape.Left + rightMostShape.Width;
                var top = leftMostShape.Top;

                var height = leftMostShape.Height;
                foreach (var shape in shapes)
                {
                    shape.Width = shape.Width * height / shape.Height;
                    shape.Height = height;
                }

                float sumWidth = 0;
                foreach (var shape in shapes)
                {
                    sumWidth += shape.Width;
                }
                float interval = (right - left - sumWidth) / (shapes.Count - 1);
                float culLeft = left;
                for (int i = 0; i < shapes.Count; i++)
                {
                    var shape = shapes[i];
                    shape.Left = culLeft;
                    shape.Top = top;

                    culLeft += shape.Width + interval;
                }
            }
        }
        public static void AlignLabels(Anchor anchor)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            var images = new List<PowerPoint.Shape>();
            var textboxes = new List<PowerPoint.Shape>();

            foreach (var shape in shapes)
            {
                if (shape.HasTextFrame == Core.MsoTriState.msoFalse
                    || shape.AutoShapeType == Core.MsoAutoShapeType.msoShapeMixed
                    || shape.TextFrame.TextRange.Text.Equals(""))
                {
                    images.Add(shape);
                }
                else
                {
                    textboxes.Add(shape);
                    shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                }
            }

            foreach (var textbox in textboxes)
            {
                var textAnchor = anchor.Opposite();
                var nearestImage = textbox.FindNearestShape(images, textAnchor, anchor);

                switch (anchor)
                {
                    case Anchor.Top:
                        textbox.SetVisualLeft(nearestImage.VisualLeft() + nearestImage.VisualWidth() / 2 - textbox.VisualWidth() / 2);
                        textbox.SetVisualBottom(nearestImage.VisualTop());
                        break;
                    case Anchor.Bottom:
                        textbox.SetVisualLeft(nearestImage.VisualLeft() + nearestImage.VisualWidth() / 2 - textbox.VisualWidth() / 2);
                        textbox.SetVisualTop(nearestImage.VisualBottom());
                        break;
                    case Anchor.Left:
                        textbox.SetVisualRight(nearestImage.VisualLeft());
                        textbox.SetVisualTop(nearestImage.VisualTop() + nearestImage.VisualHeight() / 2 - textbox.VisualHeight() / 2);
                        break;
                    case Anchor.Right:
                        textbox.SetVisualLeft(nearestImage.VisualRight());
                        textbox.SetVisualTop(nearestImage.VisualTop() + nearestImage.VisualHeight() / 2 - textbox.VisualHeight() / 2);
                        break;
                }
            }
        }

        public static void DistributeHorizontal()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                shapes.Sort((shapeA, shapeB) => Math.Sign(shapeA.Left - shapeB.Left));

                var leftMostShape = ShapeExt.GetLeftMostShape(shapes);
                var rightMostShape = ShapeExt.GetRightMostShape(shapes);

                var left = leftMostShape.VisualLeft() + leftMostShape.VisualWidth() / 2;
                var right = rightMostShape.VisualLeft() + rightMostShape.VisualWidth() / 2;

                float interval = (right - left) / (shapes.Count - 1);
                float culLeft = left + interval;
                for (int i = 1; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.SetVisualLeft(culLeft - shape.VisualWidth() / 2);

                    culLeft += interval;
                }

            }
        }

        public static void DistributeVertical()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                shapes.Sort((shapeA, shapeB) => Math.Sign(shapeA.Top - shapeB.Top));

                var topMostShape = ShapeExt.GetTopMostShape(shapes);
                var bottomMostShape = ShapeExt.GetBottomMostShape(shapes);

                var top = topMostShape.VisualTop() + topMostShape.VisualHeight() / 2;
                var bottom = bottomMostShape.VisualTop() + bottomMostShape.VisualHeight() / 2;

                float interval = (bottom - top) / (shapes.Count - 1);
                float culTop = top + interval;
                for (int i = 1; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.SetVisualTop(culTop - shape.VisualHeight() / 2);

                    culTop += interval;
                }

            }
        }

        public static void GroupLabels()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            var images = new List<PowerPoint.Shape>();
            var textboxes = new List<PowerPoint.Shape>();

            foreach (var shape in shapes)
            {
                if (shape.HasTextFrame == Core.MsoTriState.msoFalse
                    || shape.AutoShapeType == Core.MsoAutoShapeType.msoShapeMixed
                    || shape.TextFrame.TextRange.Text.Equals(""))
                {
                    images.Add(shape);
                }
                else
                {
                    textboxes.Add(shape);
                    shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                }
            }

            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            foreach (var textbox in textboxes)
            {
                try
                {
                    var nearestImage = textbox.FindNearestShape(images, Anchor.None);
                    slide.Shapes.Range(new string[] { textbox.Name, nearestImage.Name }).Group();
                }
                catch
                {

                }
            }
        }

        public static void Transpose()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            var minLeft = shapes.Min(shape => shape.VisualLeft());
            var maxLeft = shapes.Max(shape => shape.VisualLeft());

            var minTop = shapes.Min(shape => shape.VisualTop());
            var maxTop = shapes.Max(shape => shape.VisualTop());

            var diag = new Vector2(maxLeft - minLeft, maxTop - minTop);

            foreach (var shape in shapes)
            {
                float x = shape.VisualLeft() - minLeft;
                float y = shape.VisualTop() - minTop;
                float newX = y * diag.X / diag.Y;
                float newY = x * diag.Y / diag.X;
                shape.SetVisualLeft(newX + minLeft);
                shape.SetVisualTop(newY + minTop);
            }
        }

        public static void AlignGrid()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            float padding = 0;
            int numColumn = 0;
            try
            {
                padding = float.Parse(Globals.Ribbons.EdoliRibbon.gridPadding.Text);
                numColumn = int.Parse(Globals.Ribbons.EdoliRibbon.gridNumColumn.Text);
            }
            catch
            {
                return;
            }

            if (numColumn < 1)
            {
                numColumn = int.MaxValue;
            }

            float left = 0;
            float top = shapes[0].VisualTop();
            float maxHeight = 0;

            for (int i = 0; i < shapes.Count(); i++)
            {
                int col = i % numColumn;
                int row = i / numColumn;
                var shape = shapes[i];

                if (col == 0)
                {
                    left = shapes[0].VisualLeft();
                    if (row >= 1)
                    {
                        top += maxHeight + padding;
                    }
                    maxHeight = 0;
                }
                shape.SetVisualLeft(left);
                shape.SetVisualTop(top);

                left += shapes[col].VisualWidth() + padding;
                var height = shape.VisualHeight();
                if (height > maxHeight)
                {
                    maxHeight = height;
                }
            }
        }

        public static void AlignWithSiblingSlide(int indexOffset)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var slides = Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides;

            int index = slide.SlideIndex;
            int siblingSlideIndex = index + indexOffset;

            if (siblingSlideIndex < 1 || siblingSlideIndex > slides.Count)
            {
                return;
            }

            var prevSlide = slides[siblingSlideIndex];

            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            IEnumerable<PowerPoint.Shape> shapes;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionNone
                || selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                shapes = Util.ListSlideShapes();
            }
            else
            {
                shapes = Util.ListSelectedShapes();
            }

            var prevShapes = Util.ListSlideShapes(prevSlide);

            foreach (var shape in shapes)
            {
                var matchedShape = shape.FindNearestShape(prevShapes, Anchor.Center);
                shape.SetVisualLeft(matchedShape.VisualLeft());
                shape.SetVisualTop(matchedShape.VisualTop());

                shape.SetVisualSize(matchedShape.VisualWidth(), shape.VisualHeight() * matchedShape.VisualWidth() / shape.VisualWidth());
            }
        }

        public static void SwapCycle()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count < 2)
            {
                return;
            }

            float firstLeft = shapes[0].VisualLeft();
            float firstTop = shapes[0].VisualTop();

            for (int i = 0; i < shapes.Count - 1; i++)
            {
                shapes[i].SetVisualLeft(shapes[i + 1].VisualLeft());
                shapes[i].SetVisualTop(shapes[i + 1].VisualTop());
            }

            shapes.Last().SetVisualLeft(firstLeft);
            shapes.Last().SetVisualTop(firstTop);
        }

        public static void SwapCycleReverse()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count < 2)
            {
                return;
            }

            float lastLeft = shapes.Last().VisualLeft();
            float lastTop = shapes.Last().VisualTop();

            for (int i = shapes.Count - 1; i > 0; i--)
            {
                shapes[i].SetVisualLeft(shapes[i - 1].VisualLeft());
                shapes[i].SetVisualTop(shapes[i - 1].VisualTop());
            }
            shapes[0].SetVisualLeft(lastLeft);
            shapes[0].SetVisualTop(lastTop);
        }

        public static void SnapDownRight()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            var firstShape = shapes[0];
            var left = firstShape.VisualRight();
            var top = firstShape.VisualBottom();

            for (int i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                shape.SetVisualLeft(left);
                shape.SetVisualTop(top);

                top += shape.VisualHeight();
                left += shape.VisualWidth();
            }

        }

        public static void SnapUpRight()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            var firstShape = shapes[0];
            var left = firstShape.VisualRight();
            var top = firstShape.VisualTop();

            for (int i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                top -= shape.VisualHeight();
                shape.SetVisualLeft(left);
                shape.SetVisualTop(top);
                left += shape.VisualWidth();
            }
        }

        public static void MatchWidth()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float width = lastShape.Width;

                foreach (var shape in shapes)
                {
                    shape.Width = width;
                }
            }
        }

        public static void MatchHeight()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float height = lastShape.Height;

                foreach (var shape in shapes)
                {
                    shape.Height = height;
                }
            }
        }

        public static void AlignHorizontalVertical()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            if (shapes.Count > 0)
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                float minLeft = shapes.Min(s => s.Left);
                float minTop = shapes.Min(s => s.Top);

                float maxLeft = shapes.Max(s => s.Left);
                float maxTop = shapes.Max(s => s.Top);

                float meanWidth = shapes.Average(s => s.Width);
                float meanHeight = shapes.Average(s => s.Height);

                var lefts = shapes.Select(s => s.Left);
                var tops = shapes.Select(s => s.Top);

                int numHorizontalCluster = Util.NumCluster(lefts, meanWidth);
                int numVerticalCluster = Util.NumCluster(tops, meanHeight);

                float hInterval = 0;
                float vInterval = 0;

                if (numHorizontalCluster > 1)
                {
                    hInterval = (maxLeft - minLeft) / (numHorizontalCluster - 1);
                }

                if (numVerticalCluster > 1)
                {
                    vInterval = (maxTop - minTop) / (numVerticalCluster - 1);
                }

                foreach (var shape in shapes)
                {
                    int row = (int)((shape.Top - minTop + vInterval / 2) / vInterval);
                    shape.Top = minTop + row * vInterval;

                    int col = (int)((shape.Left - minLeft + hInterval / 2) / hInterval);
                    shape.Left = minLeft + col * hInterval;
                }
            }
        }

        public static void BringToFront()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            shapes.ForEach(shape =>
            {
                shape.ZOrder(Core.MsoZOrderCmd.msoBringToFront);
            });
        }

        public static void SendToBack()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            shapes.ForEach(shape =>
            {
                shape.ZOrder(Core.MsoZOrderCmd.msoSendToBack);
            });
        }

        public static void BringForward()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            shapes.ForEach(shape =>
            {
                shape.ZOrder(Core.MsoZOrderCmd.msoBringForward);
            });
        }

        public static void SendBackward()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            shapes.ForEach(shape =>
            {
                shape.ZOrder(Core.MsoZOrderCmd.msoSendBackward);
            });
        }
    }
}
