using System;
using System.Numerics;

namespace Figoint
{
    public static class MathExt
    {
        public const float degToRad = ((float)Math.PI) / 180.0f;

        public static Vector2 Rotate(this Vector2 pointToRotate, float angleInRadians)
        {
            float cosTheta = (float)Math.Cos(angleInRadians);
            float sinTheta = (float)Math.Sin(angleInRadians);

            return new Vector2
            {
                X = cosTheta * pointToRotate.X - sinTheta * pointToRotate.Y,
                Y = sinTheta * pointToRotate.X + cosTheta * pointToRotate.Y
            };
        }

        public static Vector2 RotateDeg(this Vector2 pointToRotate, float angleInDegrees)
        {
            float angleInRadians = angleInDegrees * degToRad;
            return Rotate(pointToRotate, angleInRadians);
        }

        public static Vector2 RotateWithPivot(this Vector2 pointToRotate, Vector2 pivot, float angleInRadians)
        {
            float offsetX = pointToRotate.X - pivot.X;
            float offsetY = pointToRotate.Y - pivot.Y;

            float cosTheta = (float)Math.Cos(angleInRadians);
            float sinTheta = (float)Math.Sin(angleInRadians);

            float rotatedX = cosTheta * offsetX - sinTheta * offsetY;
            float rotatedY = sinTheta * offsetX + cosTheta * offsetY;

            return new Vector2(
                rotatedX + pivot.X,
                rotatedY + pivot.Y
            );
        }
        public static Vector2 RotateWithPivotDeg(this Vector2 pointToRotate, Vector2 pivot, float angleInDegrees)
        {
            float angleInRadians = angleInDegrees * degToRad;
            return RotateWithPivot(pointToRotate, pivot, angleInRadians);
        }

        public class Intersection
        {
            public Vector2 Point;
            public bool IsIntersecting;
            public float t1;
            public float t2;
        }

        public static Intersection CalculateIntersection(Vector2 point1, Vector2 direction1, Vector2 point2, Vector2 direction2)
        {
            // 두 직선의 방정식을 이용하여 교점 계산
            // 첫 번째 선: point1 + t1 * direction1
            // 두 번째 선: point2 + t2 * direction2

            Intersection result = new Intersection();

            // Check if the lines are parallel
            float cross = direction1.X * direction2.Y - direction1.Y * direction2.X;
            if (Math.Abs(cross) < 1e-6)
            {
                result.Point = new Vector2(float.NaN, float.NaN); // No intersection (parallel lines)
                result.IsIntersecting = false;
                return result;
            }

            Vector2 diff = point2 - point1;
            float t1 = (diff.X * direction2.Y - diff.Y * direction2.X) / cross;
            float t2 = (diff.X * direction1.Y - diff.Y * direction1.X) / cross;

            result.Point = point1 + t1 * direction1;

            result.IsIntersecting = (t1 >= 0 && t1 <= 1 && t2 >= 0 && t2 <= 1);

            result.t1 = t1;
            result.t2 = t2;

            return result;
        }

        public static double CalculateAngleBetween(this Vector2 v1, Vector2 v2)
        {
            if (v1.Length() > 0) v1 = Vector2.Normalize(v1);
            if (v2.Length() > 0) v2 = Vector2.Normalize(v2);

            double dotProduct = Vector2.Dot(v1, v2);
            dotProduct = Math.Min(Math.Max(dotProduct, -1.0), 1.0); // 부동소수점 오류 방지
            double angleRadians = Math.Acos(dotProduct);

            double angleDegrees = angleRadians * (180.0 / Math.PI);

            return angleDegrees;
        }

        public static bool Approximately(this float a, float b, float epsilon = 1e-4f)
        {
            return Math.Abs(a - b) < epsilon;
        }

        public static bool Approximately(this double a, double b, double epsilon = 1e-4)
        {
            return Math.Abs(a - b) < epsilon;
        }

        public static bool Approximately(this decimal a, decimal b, decimal epsilon = 1e-4m)
        {
            return Math.Abs(a - b) < epsilon;
        }
    }
}