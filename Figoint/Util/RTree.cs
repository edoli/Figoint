
using System;
using System.Collections.Generic;
using System.Linq;

namespace Figoint
{
    public class RTree<T> : SearchTree<T>
    {
        private const int MaxEntries = 4;
        private const int MinEntries = 2;
        private Node root;

        public RTree()
        {
            root = new Node(0);
        }

        public void Insert(Leaf<T> leaf)
        {
            root.Insert(new Entry { MBR = new Rectangle(leaf.X, leaf.Y), Leaf = leaf }, root.Level);
            if (root.Entries.Count > MaxEntries)
            {
                Node newRoot = new Node(root.Level + 1);
                newRoot.Entries.Add(new Entry { MBR = root.GetMBR(), Child = root });
                root = newRoot;
                root.Split();
            }
        }

        public override Leaf<T> FindNearest(float x, float y)
        {
            var bestEntry = root.FindNearest(x, y, float.MaxValue);
            if (bestEntry == null)
            {
                return null;
            }
            return bestEntry.Leaf;
        }

        public class Node
        {
            public int Level;
            public List<Entry> Entries;

            public Node(int level)
            {
                Level = level;
                Entries = new List<Entry>();
            }

            public void Insert(Entry entry, int level)
            {
                if (Level == level)
                {
                    Entries.Add(entry);
                    return;
                }

                Entry bestEntry = Entries
                    .OrderBy(e => e.MBR.EnlargedArea(entry.MBR) - e.MBR.Area)
                    .First();

                bestEntry.Child.Insert(entry, level);
                bestEntry.MBR = bestEntry.MBR.Enlarge(entry.MBR);

                if (bestEntry.Child.Entries.Count > MaxEntries)
                {
                    bestEntry.Child.Split();
                }
            }

            public void Split()
            {
                var (group1, group2) = QuadraticSplit(Entries);
                Entries.Clear();
                Entries.Add(new Entry { Child = new Node(Level - 1) { Entries = group1 }, MBR = GetMBR(group1) });
                Entries.Add(new Entry { Child = new Node(Level - 1) { Entries = group2 }, MBR = GetMBR(group2) });
            }

            public Rectangle GetMBR()
            {
                return GetMBR(Entries);
            }

            private static Rectangle GetMBR(List<Entry> entries)
            {
                return entries.Aggregate(entries[0].MBR, (current, entry) => current.Enlarge(entry.MBR));
            }

            private static (List<Entry>, List<Entry>) QuadraticSplit(List<Entry> entries)
            {
                var pair = entries
                    .SelectMany((e1, i) => entries.Skip(i + 1).Select(e2 => (e1, e2)))
                    .OrderByDescending(p => p.e1.MBR.Enlarge(p.e2.MBR).Area - p.e1.MBR.Area - p.e2.MBR.Area)
                    .First();

                var group1 = new List<Entry> { pair.e1 };
                var group2 = new List<Entry> { pair.e2 };
                
                foreach (var entry in entries.Except(new[] { pair.e1, pair.e2 }))
                {
                    if (group1.Count >= MinEntries && group2.Count < MinEntries)
                    {
                        group2.Add(entry);
                    }
                    else if (group2.Count >= MinEntries && group1.Count < MinEntries)
                    {
                        group1.Add(entry);
                    }
                    else
                    {
                        var mbr1 = GetMBR(group1);
                        var mbr2 = GetMBR(group2);
                        if (mbr1.EnlargedArea(entry.MBR) < mbr2.EnlargedArea(entry.MBR))
                            group1.Add(entry);
                        else
                            group2.Add(entry);
                    }
                }

                return (group1, group2);
            }

            public Entry FindNearest(float x, float y, float bestDistanceSq)
            {
                Entry nearest = null;
                foreach (var entry in Entries.OrderBy(e => e.MBR.MinDistanceSquared(x, y)))
                {
                    if (entry.MBR.MinDistanceSquared(x, y) > bestDistanceSq)
                        break;

                    if (entry.Leaf != null)
                    {
                        float distSq = entry.Distance(x, y);
                        if (distSq < bestDistanceSq)
                        {
                            bestDistanceSq = distSq;
                            nearest = entry;
                        }
                    }
                    else
                    {
                        var childNearest = entry.Child.FindNearest(x, y, bestDistanceSq);
                        if (childNearest != null)
                        {
                            float distSq = childNearest.Distance(x, y);
                            if (distSq < bestDistanceSq)
                            {
                                bestDistanceSq = distSq;
                                nearest = childNearest;
                            }
                        }
                    }
                }
                return nearest;
            }
        }

        public class Entry
        {
            public Rectangle MBR { get; set; }
            public Node Child { get; set; }
            public Leaf<T> Leaf;

            public float Distance(float x, float y)
            {
                float dx = Leaf.X - x;
                float dy = Leaf.Y - y;
                return dx * dx + dy * dy;
            }
        }
                
        public struct Rectangle
        {
            public float MinX { get; set; }
            public float MinY { get; set; }
            public float MaxX { get; set; }
            public float MaxY { get; set; }

            public Rectangle(float x, float y) : this(x, y, x, y) { }

            public Rectangle(float minX, float minY, float maxX, float maxY)
            {
                MinX = minX;
                MinY = minY;
                MaxX = maxX;
                MaxY = maxY;
            }

            public float Area => (MaxX - MinX) * (MaxY - MinY);

            public Rectangle Enlarge(Rectangle other)
            {
                return new Rectangle(
                    Math.Min(MinX, other.MinX),
                    Math.Min(MinY, other.MinY),
                    Math.Max(MaxX, other.MaxX),
                    Math.Max(MaxY, other.MaxY)
                );
            }

            public float EnlargedArea(Rectangle other)
            {
                float minX = Math.Min(MinX, other.MinX);
                float minY = Math.Min(MinY, other.MinY);
                float maxX = Math.Max(MaxX, other.MaxX);
                float maxY = Math.Max(MaxY, other.MaxY);
                return (maxX - minX) * (maxY - minY);
            }

            public float MinDistanceSquared(float x, float y)
            {
                float dx = Math.Max(0, Math.Max(MinX - x, x - MaxX));
                float dy = Math.Max(0, Math.Max(MinY - y, y - MaxY));
                return dx * dx + dy * dy;
            }
        }
    }
}