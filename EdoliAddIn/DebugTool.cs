using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Numerics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{
    public static class DebugTool
    {
        public static RibbonGroup ribbonGroup;

        public static RibbonEditBox numConnector;
        public static RibbonEditBox isConnector;
        public static RibbonEditBox typeField;
        public static RibbonEditBox autoTypeField;
        public static RibbonEditBox nodeCount;
        public static RibbonEditBox node1;
        public static RibbonEditBox node2;
        public static RibbonEditBox rotation;
        public static RibbonEditBox bound;
        public static RibbonEditBox length;
        public static RibbonEditBox lengthMM;
        public static void OnStart()
        {
            AddButton("ShowConnectors", ShowConnectors);
            AddButton("ConnectAll", ShapeTool.ConnectAllCandidates);
            numConnector = AddField("NumConnector");
            isConnector = AddField("isConnector");
            typeField = AddField("Type");
            autoTypeField = AddField("AutoType");
            nodeCount = AddField("NodeCount");
            node1 = AddField("Node1");
            node2 = AddField("Node2");
            rotation = AddField("Rotation");
            bound = AddField("Bound");
            length = AddField("Length_pt");
            lengthMM = AddField("Length_mm");
        }

        public static void OnAfterShapeSizeChange(PowerPoint.Shape shape)
        {
            Update();
        }

        public static void OnAfterDragDropOnSlide(PowerPoint.Slide slide, float x, float y)
        {
            Update();
        }

        public static void OnWindowSelectionChange()
        {
            Update();
        }

        public static void Update()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 0)
            {
                var shape = shapes[0];
                isConnector.Text = shape.Connector.ToString();
                numConnector.Text = shape.ConnectionSiteCount.ToString();
                typeField.Text = shape.Type.ToString();
                autoTypeField.Text = shape.AutoShapeType.ToString();
                var nodes = shape.Nodes;
                nodeCount.Text = nodes.Count.ToString();
                node1.Text = nodes.Count > 0 ? NodeToString(nodes[1]) : "";
                node2.Text = nodes.Count > 1 ? NodeToString(nodes[nodes.Count]) : "";
                rotation.Text = shape.Rotation.ToString();
                bound.Text = shape.Rect().ToString();
                var vertices = shape.GetVertices();
                if (vertices.Length > 2)
                {
                    var distance = Vector2.Distance(vertices[0], vertices[1]);
                    length.Text = distance.ToString();
                    lengthMM.Text = (distance * ShapeTool.PtToMm).ToString();
                }
            }
        }

        public static string NodeToString(PowerPoint.ShapeNode node)
        {
            try
            {
                float x = node.Points[1, 1];
                float y = node.Points[1, 2];
                return $"{x}, {y}";
            }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x800A01A8))
            {
                return "Unknown";
            }
        }

        public static void CheckDebugGroup()
        {
            if (ribbonGroup == null)
            {
                var ribbon = Globals.Ribbons.EdoliRibbon;
                ribbonGroup = ribbon.Factory.CreateRibbonGroup();
                Globals.Ribbons.EdoliRibbon.tab1.Groups.Add(ribbonGroup);
            }
        }

        public static RibbonButton AddButton(string name, Action action)
        {
            CheckDebugGroup();
            var ribbon = Globals.Ribbons.EdoliRibbon;
            var button = ribbon.Factory.CreateRibbonButton();

            name = name.Replace(" ", "");

            button.Label = name;
            button.Name = name;
            button.ScreenTip = name;
            button.Click += new RibbonControlEventHandler((object sender, RibbonControlEventArgs e) => {
                action();
            });

            ribbonGroup.Items.Add(button);
            return button;
        }

        public static RibbonEditBox AddField(string name, Action<string> action = null)
        {
            CheckDebugGroup();
            var ribbon = Globals.Ribbons.EdoliRibbon;
            var field = ribbon.Factory.CreateRibbonEditBox();

            name = name.Replace(" ", "");

            field.Label = name;
            field.Name = name;
            field.ScreenTip = name;
            if (action != null) 
            {
                field.TextChanged += new RibbonControlEventHandler((object sender, RibbonControlEventArgs e) => {
                    action(field.Text);
                });
            }

            ribbonGroup.Items.Add(field);
            return field;
        }

        public static Dictionary<MsoShapeType, Vector2> ConnectMap = new Dictionary<MsoShapeType, Vector2>();

        public static void ShowConnectors()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            
            var line = slide.Shapes.AddLine(-100, -100, 10, 10);

            foreach (var shape in shapes)
            {
                for (int i = 0; i < shape.ConnectionSiteCount; i++)
                {
                    line.ConnectorFormat.EndConnect(shape, i + 1);
                    float x = line.Left + line.Width;
                    float y = line.Top + line.Height;

                    var textBox = AddDebugTextbox(slide, x, y, (i + 1).ToString());
                }
            }

            line.Delete();
        }

        private static PowerPoint.Shape AddDebugTextbox(Slide slide, float x, float y, String text)
        {
            // Add textbox
            var textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, x, y, 0, 0);

            textBox.TextFrame2.MarginLeft = 0;
            textBox.TextFrame2.MarginRight = 0;
            textBox.TextFrame2.MarginTop = 0;
            textBox.TextFrame2.MarginBottom = 0;

            textBox.TextFrame.TextRange.Font.Size = 12;
            textBox.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            textBox.TextFrame.TextRange.Font.Color.RGB = (int)ColorTranslator.ToOle(Color.Red);
            textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            textBox.TextFrame.WordWrap = MsoTriState.msoFalse;
            textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

            textBox.TextFrame2.TextRange.Font.Line.Visible = MsoTriState.msoTrue;
            textBox.TextFrame2.TextRange.Font.Line.ForeColor.RGB = (int)ColorTranslator.ToOle(Color.Black);
            textBox.TextFrame2.TextRange.Font.Line.Weight = 0.25f;

            textBox.TextFrame.TextRange.Text = text;

            textBox.Top -= textBox.Height / 2;

            return textBox;
        }
    }
}
