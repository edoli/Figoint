using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;

namespace EdoliAddIn
{
    public class MyColor
    {
        public static MyColor FromSVG(string colorString, string opacityString = null)
        {
            colorString = colorString.Trim();
            Color color;

            // RGB 형식 확인
            Match rgbMatch = Regex.Match(colorString, @"^rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$");
            if (rgbMatch.Success)
            {
                color = Color.FromArgb(
                    int.Parse(rgbMatch.Groups[1].Value),
                    int.Parse(rgbMatch.Groups[2].Value),
                    int.Parse(rgbMatch.Groups[3].Value)
                );
            } else if (colorString == "none") {
                return new MyColor(Color.FromArgb(0, 0, 0, 0), 0);
            } else {
                // RGB 형식 확인
                Match rgbaMatch = Regex.Match(colorString, @"^rgba\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([\d.]+)\s*\)$");
                if (rgbaMatch.Success)
                {
                    color = Color.FromArgb(
                        (int)(float.Parse(rgbaMatch.Groups[4].Value) * 255),
                        int.Parse(rgbaMatch.Groups[1].Value),
                        int.Parse(rgbaMatch.Groups[2].Value),
                        int.Parse(rgbaMatch.Groups[3].Value)
                    );
                }
                else
                {
                    color = ColorTranslator.FromHtml(colorString);
                }
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