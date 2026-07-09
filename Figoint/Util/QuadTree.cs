
using System.Collections.Generic;
using System.Linq;

namespace Figoint
{
    public class QuadTree<T> : SearchTree<T>
    {
        private const int MAX_OBJECTS = 10;
        private const int MAX_LEVELS = 5;

        private int level;
        private List<Leaf<T>> leaves;
        private Rect bounds;
        private QuadTree<T>[] nodes;

        public QuadTree() : this(new Rect(float.NegativeInfinity, float.NegativeInfinity, float.PositiveInfinity, float.PositiveInfinity), 0) {}

        public QuadTree(Rect bounds, int level)
        {
            this.bounds = bounds;
            this.level = level;
            leaves = new List<Leaf<T>>();
            nodes = new QuadTree<T>[4];
        }

        public void Insert(Leaf<T> leaf)
        {
            if (nodes[0] != null)
            {
                int index = GetIndex(leaf);
                if (index != -1)
                {
                    nodes[index].Insert(leaf);
                    return;
                }
            }

            leaves.Add(leaf);

            if (leaves.Count > MAX_OBJECTS && level < MAX_LEVELS)
            {
                if (nodes[0] == null)
                {
                    Split();
                }

                int i = 0;
                while (i < leaves.Count)
                {
                    int index = GetIndex(leaves[i]);
                    if (index != -1)
                    {
                        nodes[index].Insert(leaves[i]);
                        leaves.RemoveAt(i);
                    }
                    else
                    {
                        i++;
                    }
                }
            }
        }

        public override Leaf<T> FindNearest(float x, float y)
        {
            List<Leaf<T>> possibleNearest = new List<Leaf<T>>();
            RetrievePossibleNearest(x, y, possibleNearest);

            return possibleNearest
                .OrderBy(s => Distance(s.X, s.Y, x, y))
                .FirstOrDefault();
        }

        private void RetrievePossibleNearest(float x, float y, List<Leaf<T>> possibleNearest)
        {
            int index = GetIndex(x, y);
            if (index != -1 && nodes[0] != null)
            {
                nodes[index].RetrievePossibleNearest(x, y, possibleNearest);
            }
            else
            {
                // HACK: is it right?
                possibleNearest.AddRange(leaves);                
            }
        }

        private void Split()
        {
            float subWidth = bounds.Width / 2;
            float subHeight = bounds.Height / 2;
            float x = bounds.X;
            float y = bounds.Y;

            nodes[0] = new QuadTree<T>(new Rect(x + subWidth, y, subWidth, subHeight), level + 1);
            nodes[1] = new QuadTree<T>(new Rect(x, y, subWidth, subHeight), level + 1);
            nodes[2] = new QuadTree<T>(new Rect(x, y + subHeight, subWidth, subHeight), level + 1);
            nodes[3] = new QuadTree<T>(new Rect(x + subWidth, y + subHeight, subWidth, subHeight), level + 1);
        }

        private int GetIndex(Leaf<T> leaf)
        {
            return GetIndex(leaf.X, leaf.Y);
        }

        private int GetIndex(float x, float y)
        {
            int index = -1;
            float verticalMidpoint = bounds.X + (bounds.Width / 2);
            float horizontalMidpoint = bounds.Y + (bounds.Height / 2);

            bool topQuadrant = y < horizontalMidpoint && y > bounds.Y;
            bool bottomQuadrant = y >= horizontalMidpoint && y < bounds.Y + bounds.Height;

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

        private float Distance(float x1, float y1, float x2, float y2)
        {
            float dx = x1 - x2;
            float dy = y1 - y2;
            return dx * dx + dy * dy;
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
}