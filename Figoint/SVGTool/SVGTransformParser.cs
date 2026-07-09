using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Figoint
{
    public static class SVGTransformParser
    {
        public static void ApplyTransform(PowerPoint.Shape shape, string transform)
        {
            var transformations = ParseTransformString(transform);
            foreach (var transformation in transformations)
            {
                ApplySingleTransform(shape, transformation);
            }
        }

        private static List<(string Type, float[] Values)> ParseTransformString(string transform)
        {
            var transformList = new List<(string, float[])>();
            var regex = new Regex(@"(\w+)\s*\(([-\d\s,\.]+)\)");
            var matches = regex.Matches(transform);

            foreach (Match match in matches)
            {
                string type = match.Groups[1].Value;
                float[] values = match.Groups[2].Value.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                                     .Select(float.Parse)
                                                     .ToArray();
                transformList.Add((type, values));
            }

            return transformList;
        }

        private static void ApplySingleTransform(PowerPoint.Shape shape, (string Type, float[] Values) transformation)
        {
            switch (transformation.Type.ToLower())
            {
                case "translate":
                    ApplyTranslate(shape, transformation.Values);
                    break;
                case "rotate":
                    ApplyRotate(shape, transformation.Values);
                    break;
                case "scale":
                    ApplyScale(shape, transformation.Values);
                    break;
                case "skew":
                case "skewx":
                case "skewy":
                    ApplySkew(shape, transformation.Type, transformation.Values);
                    break;
                case "matrix":
                    ApplyMatrix(shape, transformation.Values);
                    break;
            }
        }

        private static void ApplyTranslate(PowerPoint.Shape shape, float[] values)
        {
            if (values.Length >= 2)
            {
                shape.Left += values[0];
                shape.Top += values[1];
            }
            else if (values.Length == 1)
            {
                shape.Left += values[0];
            }
        }

        private static void ApplyRotate(PowerPoint.Shape shape, float[] values)
        {
            if (values.Length >= 1)
            {
                float angle = values[0];
                Vector2 pivot;

                // 기본 피벗 포인트는 도형의 왼쪽 아래 모서리
                pivot = new Vector2(shape.Left, shape.Top + shape.Height);

                // SVG에서 회전 중심이 지정된 경우 (cx, cy가 제공된 경우)
                if (values.Length >= 3)
                {
                    float cx = values[1];
                    float cy = values[2];
                    pivot = new Vector2(cx, cy);
                }

                // 회전 적용
                RotateShape(shape, angle, pivot);
            }
        }

        private static void RotateShape(PowerPoint.Shape shape, float angle, Vector2 pivot)
        {
            Vector2 center = shape.VisualPosition(ShapeExt.Anchor.Center);
            Vector2 pivotDir = pivot - center;
            Vector2 pivotDirRotated = pivotDir.RotateDeg(angle);
            Vector2 moved = pivotDirRotated - pivotDir;

            shape.Left -= moved.X;
            shape.Top -= moved.Y;

            shape.Rotation = angle;
        }

        private static void ApplyScale(PowerPoint.Shape shape, float[] values)
        {
            if (values.Length >= 1)
            {
                float scaleX = values[0];
                float scaleY = values.Length >= 2 ? values[1] : scaleX;

                shape.Width *= scaleX;
                shape.Height *= scaleY;
            }
        }

        private static void ApplySkew(PowerPoint.Shape shape, string type, float[] values)
        {
            // PowerPoint doesn't have a direct skew transformation
            // We can approximate it by adjusting the shape's adjustments if it's a basic shape
            // For complex shapes, skew might not be possible to apply accurately
            if (shape.AutoShapeType == MsoAutoShapeType.msoShapeRectangle)
            {
                if (type == "skewx" && values.Length >= 1)
                {
                    shape.Adjustments[1] = (float)Math.Tan(values[0] * Math.PI / 180) * 100;
                }
                else if (type == "skewy" && values.Length >= 1)
                {
                    shape.Adjustments[2] = (float)Math.Tan(values[0] * Math.PI / 180) * 100;
                }
            }
        }

        private static void ApplyMatrix(PowerPoint.Shape shape, float[] values)
        {
            if (values.Length >= 6)
            {
                // Matrix transformation: [a b c d e f]
                // We can approximate this by decomposing into scale, rotate, and translate
                float a = values[0], b = values[1], c = values[2], d = values[3], e = values[4], f = values[5];

                // Extract scale
                float scaleX = (float)Math.Sqrt(a * a + b * b);
                float scaleY = (float)Math.Sqrt(c * c + d * d);

                // Extract rotation
                float angle = (float)(Math.Atan2(b, a) * 180 / Math.PI);

                // Apply transformations
                shape.Width *= scaleX;
                shape.Height *= scaleY;
                shape.Rotation = angle;
                shape.Left += e;
                shape.Top += f;
            }
        }
    }
}