
using System;
using System.Collections.Generic;
using System.Linq;

namespace EdoliAddIn
{
    public class KdTree<T> : SearchTree<T>
    {
        private readonly Node root;

        public KdTree(List<Leaf<T>> leaves)
        {
            root = BuildTree(leaves, 0);
        }

        private Node BuildTree(List<Leaf<T>> leaves, int depth)
        {
            if (!leaves.Any())
            {
                return null;
            }

            int axis = depth % 2;
            leaves.Sort((a, b) => axis == 0 ? a.X.CompareTo(b.X) : a.Y.CompareTo(b.Y));

            int medianIndex = leaves.Count / 2;
            Leaf<T> leaf = leaves[medianIndex];
            Node node = new Node(leaf)
            {
                Left = BuildTree(leaves.Take(medianIndex).ToList(), depth + 1),
                Right = BuildTree(leaves.Skip(medianIndex + 1).ToList(), depth + 1)
            };

            return node;
        }

        public override Leaf<T> FindNearest(float x, float y)
        {
            float bestDistance = float.MaxValue;
            var bestNode = FindNearestNeighbor(root, x, y, 0, null, ref bestDistance);
            if (bestNode == null)
            {
                return null;
            }
            return new Leaf<T>(bestNode.Data, bestNode.X, bestNode.Y);
        }

        private Node FindNearestNeighbor(Node node, float x, float y, int depth, Node best, ref float bestDistance)
        {
            if (node == null)
            {
                return best;
            }

            float distance = Distance(node, x, y);
            if (distance < bestDistance)
            {
                best = node;
                bestDistance = distance;
            }

            int axis = depth % 2;
            Node nextBranch;
            Node oppositeBranch;

            if ((axis == 0 && x < node.X) || (axis == 1 && y < node.Y))
            {
                nextBranch = node.Left;
                oppositeBranch = node.Right;
            }
            else
            {
                nextBranch = node.Right;
                oppositeBranch = node.Left;
            }

            best = FindNearestNeighbor(nextBranch, x, y, depth + 1, best, ref bestDistance);

            float axisDistance = axis == 0 ? node.Y - y : node.X - x;
            if (axisDistance * axisDistance < bestDistance)
            {
                best = FindNearestNeighbor(oppositeBranch, x, y, depth + 1, best, ref bestDistance);
            }

            return best;
        }

        private float Distance(Node a, float x, float y)
        {
            float dx = a.X - x;
            float dy = a.Y - y;
            return dx * dx + dy * dy;
        }

        public class Node
        {
            public T Data;
            public float X;
            public float Y;
            public Node Left;
            public Node Right;

            public Node(Leaf<T> leaf)
            {
                Data = leaf.Data;
                X = leaf.X;
                Y = leaf.Y;
            }
        }
    }
}