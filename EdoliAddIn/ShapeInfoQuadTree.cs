
using System;
using System.Collections.Generic;
using System.Linq;

namespace EdoliAddIn
{
    public class ShapeInfoQuadTree
    {
        private const int MAX_OBJECTS = 10;
        private const int MAX_LEVELS = 5;

        private int level;
        private List<ShapeInfo> shapes;
        private Rect bounds;
        private ShapeInfoQuadTree[] nodes;

        public ShapeInfoQuadTree() : this(new Rect(float.NegativeInfinity, float.NegativeInfinity, float.PositiveInfinity, float.PositiveInfinity), 0) {}

        public ShapeInfoQuadTree(Rect bounds, int level)
        {
            this.bounds = bounds;
            this.level = level;
            shapes = new List<ShapeInfo>();
            nodes = new ShapeInfoQuadTree[4];
        }

        public void Insert(ShapeInfo shape)
        {
            if (nodes[0] != null)
            {
                int index = GetIndex(shape);
                if (index != -1)
                {
                    nodes[index].Insert(shape);
                    return;
                }
            }

            shapes.Add(shape);

            if (shapes.Count > MAX_OBJECTS && level < MAX_LEVELS)
            {
                if (nodes[0] == null)
                {
                    Split();
                }

                int i = 0;
                while (i < shapes.Count)
                {
                    int index = GetIndex(shapes[i]);
                    if (index != -1)
                    {
                        nodes[index].Insert(shapes[i]);
                        shapes.RemoveAt(i);
                    }
                    else
                    {
                        i++;
                    }
                }
            }
        }

        public ShapeInfo FindNearest(float x, float y)
        {
            List<ShapeInfo> possibleNearest = new List<ShapeInfo>();
            RetrievePossibleNearest(x, y, possibleNearest);

            return possibleNearest
                .OrderBy(s => Math.Sqrt(Math.Pow(s.CenterX - x, 2) + Math.Pow(s.CenterY - y, 2)))
                .FirstOrDefault();
        }

        private void RetrievePossibleNearest(float x, float y, List<ShapeInfo> possibleNearest)
        {
            int index = GetIndex(x, y);
            if (index != -1 && nodes[0] != null)
            {
                nodes[index].RetrievePossibleNearest(x, y, possibleNearest);
            }

            possibleNearest.AddRange(shapes);
        }

        private void Split()
        {
            float subWidth = bounds.Width / 2;
            float subHeight = bounds.Height / 2;
            float x = bounds.X;
            float y = bounds.Y;

            nodes[0] = new ShapeInfoQuadTree(new Rect(x + subWidth, y, subWidth, subHeight), level + 1);
            nodes[1] = new ShapeInfoQuadTree(new Rect(x, y, subWidth, subHeight), level + 1);
            nodes[2] = new ShapeInfoQuadTree(new Rect(x, y + subHeight, subWidth, subHeight), level + 1);
            nodes[3] = new ShapeInfoQuadTree(new Rect(x + subWidth, y + subHeight, subWidth, subHeight), level + 1);
        }

        private int GetIndex(ShapeInfo shape)
        {
            return GetIndex(shape.CenterX, shape.CenterY);
        }

        private int GetIndex(float x, float y)
        {
            int index = -1;
            double verticalMidpoint = bounds.X + (bounds.Width / 2);
            double horizontalMidpoint = bounds.Y + (bounds.Height / 2);

            bool topQuadrant = (y < horizontalMidpoint && y > bounds.Y);
            bool bottomQuadrant = (y >= horizontalMidpoint && y < bounds.Y + bounds.Height);

            if (x < verticalMidpoint && x > bounds.X)
            {
                if (topQuadrant)
                {
                    index = 1;
                }
                else if (bottomQuadrant)
                {
                    index = 2;
                }
            }
            else if (x >= verticalMidpoint && x < bounds.X + bounds.Width)
            {
                if (topQuadrant)
                {
                    index = 0;
                }
                else if (bottomQuadrant)
                {
                    index = 3;
                }
            }

            return index;
        }
    }

    public struct Rect
    {
        public float X, Y, Width, Height;

        public Rect(float x, float y, float width, float height)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }
    }

    public class ShapeInfo
    {
        public Microsoft.Office.Interop.PowerPoint.Shape Shape { get; set; }
        public float CenterX { get; set; }
        public float CenterY { get; set; }
    }
}