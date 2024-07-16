using System.Drawing;

namespace EdoliAddIn
{
    public static class ColorExt
    {

        public static int ToRGB(this Color color)
        {
            return (color.B << 16) | (color.G << 8) | color.R;
        }
    }
}