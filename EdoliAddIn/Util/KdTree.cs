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
            if (leaves == null || leaves.Count == 0)
            {
                return null;
            }

            int axis = depth % 2;
            leaves.Sort((a, b) => axis == 0 ? a.X.CompareTo(b.X) : a.Y.CompareTo(b.Y));

            int medianIndex = leaves.Count / 2;
            Leaf<T> leaf = leaves[medianIndex];
            Node node = new Node(leaf)
            {
                Left = BuildTree(leaves.GetRange(0, medianIndex), depth + 1),
                Right = BuildTree(leaves.Count > medianIndex + 1 ?
                    leaves.GetRange(medianIndex + 1, leaves.Count - medianIndex - 1) :
                    new List<Leaf<T>>(), depth + 1)
            };

            return node;
        }

        public override Leaf<T> FindNearest(float x, float y)
        {
            float bestDistanceSq = float.MaxValue;
            var bestNode = FindNearestNeighbor(root, x, y, 0, null, ref bestDistanceSq);
            if (bestNode == null)
            {
                return null;
            }
            return new Leaf<T>(bestNode.Data, bestNode.X, bestNode.Y);
        }

        private Node FindNearestNeighbor(Node node, float x, float y, int depth, Node best, ref float bestDistanceSq)
        {
            if (node == null)
            {
                return best;
            }

            float distanceSq = DistanceSquared(node, x, y);
            if (distanceSq < bestDistanceSq)
            {
                best = node;
                bestDistanceSq = distanceSq;
            }

            int axis = depth % 2;
            float axisValue = axis == 0 ? x : y;
            float nodeAxisValue = axis == 0 ? node.X : node.Y;

            Node firstBranch, secondBranch;
            if (axisValue < nodeAxisValue)
            {
                firstBranch = node.Left;
                secondBranch = node.Right;
            }
            else
            {
                firstBranch = node.Right;
                secondBranch = node.Left;
            }

            best = FindNearestNeighbor(firstBranch, x, y, depth + 1, best, ref bestDistanceSq);

            // 초평면(hyperplane)과의 거리를 계산하여 다른 브랜치 탐색 여부 결정
            float axisDistanceSq = (nodeAxisValue - axisValue) * (nodeAxisValue - axisValue);
            if (axisDistanceSq < bestDistanceSq)
            {
                best = FindNearestNeighbor(secondBranch, x, y, depth + 1, best, ref bestDistanceSq);
            }

            return best;
        }

        private float DistanceSquared(Node a, float x, float y)
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
