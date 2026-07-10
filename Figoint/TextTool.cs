using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using Expressive;
using System.Linq;

namespace Figoint
{
    public static class TextTool
    {
        public static void IncreaseNumber()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            ReplaceSelectedText(s => (float.Parse(s) + 1).ToString());
        }

        public static void DecreaseNumber()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            ReplaceSelectedText(s => (float.Parse(s) - 1).ToString());
        }

        public static void EvaluateExpression()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            ReplaceSelectedText(s => {
                if (string.IsNullOrEmpty(s)) return s;

                // Split while keeping separators (CRLF, CR, LF)
                var parts = System.Text.RegularExpressions.Regex.Split(s, "(\\r\\n|\\r|\\n)");
                var sb = new System.Text.StringBuilder();

                for (int i = 0; i < parts.Length; i += 2)
                {
                    var line = parts[i];
                    var separator = (i + 1 < parts.Length) ? parts[i + 1] : string.Empty;

                    var trimmedLine = line.Trim();
                    string processed;

                    if (trimmedLine.EndsWith("="))
                    {
                        try
                        {
                            // if line ends with '=', add result of expression after '='
                            // e.g. "1+2=" => "1+2=3"
                            var expr = trimmedLine.Substring(0, trimmedLine.Length - 1);
                            var value = new Expression(expr, ExpressiveOptions.IgnoreCaseForParsing).Evaluate();
                            processed = line + (value?.ToString() ?? string.Empty);
                        }
                        catch (Exception ex)
                        {
                            AppLog.Warn("Failed to evaluate expression line.", ex, "Evaluate expression");
                            processed = line; // if evaluation fails, return original line
                        }
                    }
                    else
                    {
                        try
                        {
                            // if line does not end with '=', replace line with result
                            // e.g. "1+2" => "3"
                            processed = new Expression(trimmedLine, ExpressiveOptions.IgnoreCaseForParsing).Evaluate().ToString();
                        }
                        catch (Exception ex)
                        {
                            AppLog.Warn("Failed to evaluate expression line.", ex, "Evaluate expression");
                            processed = line; // if evaluation fails, return original line
                        }
                    }
                    sb.Append(processed);
                    if (!string.IsNullOrEmpty(separator))
                    {
                        sb.Append(separator);
                    }
                }
                if (sb[sb.Length - 1] == '\n' || sb[sb.Length - 1] == '\r')
                {
                    sb.Remove(sb.Length - 1, 1);
                }

                return sb.ToString();
            });
        }

        public static String GetSelectedText()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection != null && selection.TextRange != null)
            {
                return selection.TextRange.Text;
            }
            return "";
        }

        public static void ReplaceSelectedText(Func<string, string> replacer)
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            var isShapeSelection = (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes);

            if (isShapeSelection)
            {
                var shapeRange = selection.ShapeRange;
                foreach (PowerPoint.Shape shape in shapeRange)
                {
                    var text = shape.TextFrame.TextRange.Text;
                    shape.TextFrame.TextRange.Text = replacer(text);
                }

            }
            else
            {
                if (selection != null && selection.TextRange != null)
                {
                    try
                    {
                        if (selection.TextRange.Text == "")
                        {
                            var shapes = Util.ListSelectedShapes();
                            var shape = shapes[0];
                            var text = shape.TextFrame.TextRange.Text;
                            shape.TextFrame.TextRange.Text = replacer(text);
                        } else
                        {
                            var text = selection.TextRange.Text;
                            selection.TextRange.Text = replacer(text);
                        }
                    }
                    catch (Exception ex)
                    {
                        AppLog.Warn("Failed to replace selected text.", ex, "Replace selected text");
                    }
                }
            }
        }
    }
}
