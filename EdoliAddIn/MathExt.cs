using System;
using System.Numerics;

namespace EdoliAddIn
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

    }
}