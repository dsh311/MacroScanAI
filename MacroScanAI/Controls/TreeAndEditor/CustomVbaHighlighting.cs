/*
 * Copyright (C) 2025 David S. Shelley <davidsmithshelley@gmail.com>
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License 
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */

using ICSharpCode.AvalonEdit.Highlighting;
using System.Text.RegularExpressions;
using System.Windows.Media;
using System.Windows;

namespace MacroScanAI.Controls.TreeAndEditor
{
    public class CustomVbaHighlighting : IHighlightingDefinition
    {
        private readonly HighlightingRuleSet mainRuleSet;

        SimpleHighlightingBrush dkThemeGreen = new SimpleHighlightingBrush(Color.FromRgb(87, 166, 74));
        SimpleHighlightingBrush ltThemeGreen = new SimpleHighlightingBrush(Colors.Green);

        SimpleHighlightingBrush dkThemeBlue = new SimpleHighlightingBrush(Color.FromRgb(86, 152, 189));
        SimpleHighlightingBrush ltThemeBlue = new SimpleHighlightingBrush(Colors.Blue);

        SimpleHighlightingBrush dkThemeBrown = new SimpleHighlightingBrush(Color.FromRgb(209, 157, 111));
        SimpleHighlightingBrush ltThemeBrown = new SimpleHighlightingBrush(Colors.Brown);

        SimpleHighlightingBrush dkThemeSeaGreen = new SimpleHighlightingBrush(Color.FromRgb(181, 206, 168));
        SimpleHighlightingBrush ltThemeSeaGreen = new SimpleHighlightingBrush(Colors.DarkSeaGreen);

        SimpleHighlightingBrush dkThemeYellow = new SimpleHighlightingBrush(Color.FromRgb(218, 214, 131));
        SimpleHighlightingBrush ltThemeYellow = new SimpleHighlightingBrush(Color.FromRgb(218, 214, 131));

        public CustomVbaHighlighting(bool useDarkTheme = true)
        {
            mainRuleSet = new HighlightingRuleSet();

            SimpleHighlightingBrush theGreen = useDarkTheme ? dkThemeGreen : ltThemeGreen;
            SimpleHighlightingBrush theBlue = useDarkTheme ? dkThemeBlue : ltThemeBlue;
            SimpleHighlightingBrush theBrown = useDarkTheme ? dkThemeBrown : ltThemeBrown;
            SimpleHighlightingBrush theSeaGreen = useDarkTheme ? dkThemeSeaGreen : ltThemeSeaGreen;
            SimpleHighlightingBrush theYellow = useDarkTheme ? dkThemeYellow : ltThemeYellow;

            // Comments (green) - highest priority
            mainRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"'.*$", RegexOptions.Compiled),
                Color = new HighlightingColor
                {
                    Foreground = theGreen
                }
            });

            var keywords = new[]
            {
                "Alias", "And", "As", "Attribute", "Base", "Beep", "Binary", "Boolean", "ByRef", "Byte",
                "ByVal", "Call", "Case", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt",
                "CLng", "CLngLng", "CLngPtr", "Compare", "Const", "CSng", "CStr", "Currency", "CVar", "Date",
                "Decimal", "Declare", "DefBool", "DefByte", "DefCur", "DefDate", "DefDbl", "DefDec", "DefInt", "DefLng",
                "DefLngLng", "DefLngPtr", "DefObj", "DefSng", "DefStr", "DefVar", "Debug.Assert", "Debug.Print", "Dim", "Do",
                "DoEvents", "Double", "Each", "Else", "ElseIf", "End", "EndIf", "Enum", "Eqv", "Erase", "Error",
                "Event", "Exit", "Explicit", "False", "For", "Friend", "Function", "Get", "Global", "GoSub",
                "GoTo", "If", "Imp", "Implements", "In", "Input", "Integer", "Is", "Let", "Lib",
                "Like", "Line", "Lock", "Long", "LongLong", "LongPtr", "Loop", "LSet", "Me", "Mid",
                "Mod", "Module", "New", "Next", "Not", "Nothing", "Null", "Object", "On",
                "Option", "Optional", "Or", "ParamArray", "Preserve", "Private", "Property", "Public", "RaiseEvent", "ReDim",
                "Rem", "Resume", "Return", "RSet", "Select", "Set", "Single", "Static",
                "Step", "Stop", "String", "Sub", "Then", "To", "True", "Type", "TypeOf", "Until",
                "Variant", "Wend", "While", "With", "WithEvents", "Xor", "Application", "Workbook", "Worksheets", "Worksheet", "Cells",
                "Rows", "Columns", "ActiveSheet", "ActiveWorkbook", "Chart", "Charts", "PivotTable", "PivotTables",
                "Shapes", "UsedRange", "Document", "Documents", "Paragraphs", "Paragraph", "Tables", "Table", "Bookmarks", "Bookmark",
                "InlineShapes", "ContentControls", "Presentation", "Presentations", "Slide", "Slides", "TextFrame", "TextRange", "SlideShowWindow", "SlideShowWindows",

                // Common VBA methods
                "Activate", "ActiveChart", "ActiveCell", "Add", "Asc", "Array", "Beep", "Calculate", "CDate", "CDbl", "CInt", "CLng",
                "Close", "Clear", "ClearContents", "ClearFormats", "Copy", "Cut", "Date", "Debug.Assert", "Debug.Print", "Delete",
                "Dir", "EOF", "Erase", "Find", "Format", "Get", "InputBox", "Insert", "IsArray", "IsDate",
                "InStr", "IsEmpty", "IsNull", "IsNumeric", "Len", "LBound", "Left", "Mid", "Move", "Next", "Now",
                "OptionBase", "Print", "Range", "Randomize", "Refresh", "Repaint", "Replace", "RGB", "Right", "Sheets", "Select",
                "Selection", "SetFocus", "Show", "Space", "Split", "Str", "String", "Time", "Timer", "Trim", "TypeName", "UBound",
                "Val", "Variant", "MsgBox"
            };

            foreach (var kw in keywords)
            {
                mainRuleSet.Rules.Add(new HighlightingRule
                {
                    Regex = new Regex(@"\b" + kw + @"\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                    Color = new HighlightingColor
                    {
                        Foreground = theBlue,
                        FontWeight = FontWeights.Bold
                    }
                });
            }


            // Strings
            mainRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex("\"(?:[^\"]|\"\")*\"", RegexOptions.Compiled),
                Color = new HighlightingColor
                {
                    Foreground = theBrown,
                    FontWeight = FontWeights.Bold
                }
            });


            // Highlight name follows sub
            mainRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"(?<=\b(?:Public|Private|Friend|Declare)?\s*Sub\s+)\w+", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                Color = new HighlightingColor
                {
                    Foreground = theYellow,
                    FontWeight = FontWeights.Bold
                }
            });

            // Highlight name that follows Function
            mainRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"(?<=\b(?:Public|Private|Friend|Declare)?\s*Function\s+)\w+", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                Color = new HighlightingColor
                {
                    Foreground = theYellow,
                    FontWeight = FontWeights.Bold
                }
            });



            // Numeric constants (light green)
            mainRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"\b(&H[0-9A-Fa-f]+|&O[0-7]+|\d+(\.\d+)?)\b", RegexOptions.Compiled),
                Color = new HighlightingColor
                {
                    Foreground = theSeaGreen,
                    FontWeight = FontWeights.Normal
                }
            });
        }

        public string Name => "Custom VBA";

        public HighlightingRuleSet MainRuleSet => mainRuleSet;

        public IEnumerable<HighlightingColor> NamedHighlightingColors => new List<HighlightingColor>();

        public IDictionary<string, string> Properties => new Dictionary<string, string>();

        public HighlightingColor GetNamedColor(string name) => null;

        public HighlightingRuleSet GetNamedRuleSet(string name) => null;
    }
}
