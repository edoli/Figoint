using System.Collections.Generic;
using System.Drawing;

namespace EdoliAddIn
{
    public class MyColor
    {
        private static readonly Dictionary<string, Color> ColorKeywords = new Dictionary<string, Color>
        {
            {"none", Color.Transparent},
        };

        public static MyColor FromSVG(string colorString, string opacityString = null)
        {   
            Color color;
            if (ColorKeywords.TryGetValue(colorString, out Color knownColor))
            {
                color = knownColor;
            }
            else
            {
                color = ColorTranslator.FromHtml(colorString);
            }
            
            if (string.IsNullOrEmpty(opacityString))
            {
                return new MyColor(color);
            }
            
            float opacity = float.Parse(opacityString);

            // Apply opacity to color
            int alpha = (int)(opacity * 255);
            return new MyColor(Color.FromArgb(alpha, color), opacity);
        }

        public Color Color;
        public float Transparency = 0.0f;

        public MyColor(Color color)
        {
            this.Color = color;
        }
        public MyColor(Color color, float opacity)
        {
            Color = color;
            Transparency = 1 - opacity;
        }

        public int ToInt()
        {
            return (Color.B << 16) | (Color.G << 8) | Color.R;
        }

    }
}