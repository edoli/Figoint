

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace EdoliAddIn
{
    public class SVGStyleParser
    {
        public static readonly char[] styleSeparators = new char[] { ';' };
        public static readonly char[] pairSeparators = new char[] { ':' };

        private Dictionary<string, Dictionary<string, string>> classStyles;
        private Dictionary<string, Dictionary<string, string>> elementStyles;

        public SVGStyleParser(XElement svgRoot)
        {
            classStyles = new Dictionary<string, Dictionary<string, string>>();
            elementStyles = new Dictionary<string, Dictionary<string, string>>();
            ParseStylesFromSVG(svgRoot);
        }

        private void ParseStylesFromSVG(XElement svgRoot)
        {
            var styleElements = svgRoot.Descendants()
                .Where(e => e.Name.LocalName == "style");
            foreach (var styleElement in styleElements)
            {
                ParseStyleElement(styleElement.Value);
            }
        }

        private void ParseStyleElement(string styleContent)
        {
            var styleRules = Regex.Split(styleContent, @"}\s*")
                                  .Where(rule => !string.IsNullOrWhiteSpace(rule));

            foreach (var rule in styleRules)
            {
                var parts = rule.Split(new[] { '{' }, 2);
                if (parts.Length == 2)
                {
                    var selectors = parts[0].Trim().Split(',');
                    var styles = ParseStyleProperties(parts[1]);

                    foreach (var selector in selectors)
                    {
                        var trimmedSelector = selector.Trim();
                        if (trimmedSelector.StartsWith("."))
                        {
                            classStyles[trimmedSelector.Substring(1)] = styles;
                        }
                        else
                        {
                            elementStyles[trimmedSelector] = styles;
                        }
                    }
                }
            }
        }

        private Dictionary<string, string> ParseStyleProperties(string styleString)
        {
            return styleString.Split(styleSeparators, StringSplitOptions.RemoveEmptyEntries)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s.Split(pairSeparators, StringSplitOptions.RemoveEmptyEntries))
                .Where(pair => pair.Length == 2)
                .ToDictionary(pair => pair[0].Trim(), pair => pair[1].Trim());
        }

        public Dictionary<string, string> GetStylesForElement(XElement element)
        {
            var styles = new Dictionary<string, string>();

            // Apply element styles
            if (elementStyles.TryGetValue(element.Name.LocalName, out var elementStyle))
            {
                foreach (var style in elementStyle)
                {
                    styles[style.Key] = style.Value;
                }
            }

            // Apply class styles
            var classAttribute = element.Attribute("class");
            if (classAttribute != null)
            {
                var classes = classAttribute.Value.Split(' ');
                foreach (var className in classes)
                {
                    if (classStyles.TryGetValue(className, out var classStyle))
                    {
                        foreach (var style in classStyle)
                        {
                            styles[style.Key] = style.Value;
                        }
                    }
                }
            }

            // Apply inline styles (these take precedence)
            var styleAttribute = element.Attribute("style");
            if (styleAttribute != null)
            {
                var inlineStyles = ParseStyleProperties(styleAttribute.Value);
                foreach (var style in inlineStyles)
                {
                    styles[style.Key] = style.Value;
                }
            }

            return styles;
        }
        
        public XElement ParseAndConvertStyle(XElement element)
        {
            var styleDict = GetStylesForElement(element);

            foreach (var kvp in styleDict)
            {
                element.SetAttributeValue(kvp.Key, kvp.Value);
            }

            element.Attribute("style")?.Remove();
            return element;
        }
    }
}