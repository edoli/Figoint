using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
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
        public static RibbonEditBox node1;
        public static RibbonEditBox node2;
        public static void OnStart()
        {
            AddButton("ShowConnectors", ShowConnectors);
            AddButton("ConnectAll", ShapeTool.ConnectAllCandidates);
            AddButton("TestShape", TestShape);
            numConnector = AddField("NumConnector");
            isConnector = AddField("isConnector");
            typeField = AddField("Type");
            autoTypeField = AddField("AutoType");
            node1 = AddField("Node1");
            node2 = AddField("Node2");
        }

        public static void OnAfterShapeSizeChange()
        {
            // try {
            //     Globals.Ribbons.EdoliRibbon.animationName.Text = shape.ConnectorFormat.BeginConnectionSite.ToString();
            //     Globals.Ribbons.EdoliRibbon.animationName.Text = shape.ConnectorFormat.EndConnectionSite.ToString();
            // } catch {

            // }
        }

        public static void OnWindowSelectionChange()
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
                node1.Text = nodes.Count > 0 ? NodeToString(nodes[1]) : "";
                node2.Text = nodes.Count > 1 ? NodeToString(nodes[nodes.Count]) : "";
            }
        }

        public static void TestShape()
        {
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            PowerPoint.FreeformBuilder freeform = slide.Shapes.BuildFreeform(MsoEditingType.msoEditingCorner, 0, 0);

            freeform.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, 
                100, 100, 200, 200);

            freeform.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, 
                300, 100, 400, 200);

            PowerPoint.Shape shape = freeform.ConvertToShape();
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

                    var textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, x, y, 1, 1);

                    var textFrame = textBox.TextFrame2;
                    textBox.TextFrame2.TextRange.Text = (i + 1).ToString();
                    textFrame.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
                    textFrame.WordWrap = MsoTriState.msoFalse;

                    textBox.Left -= textBox.Width / 2;
                    textBox.Top -= textBox.Height / 2;
                }
            }

            line.Delete();
        }
    }
}
