using Microsoft.Office.Tools.Ribbon;

namespace Figoint
{
    public partial class EdoliRibbon
    {
        private void EdoliRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Execute(string operation, System.Action action, bool notifyUserOnError = true)
        {
            AppCommand.Run(operation, action, notifyUserOnError);
        }

        private void grid_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Align grid", () => AlignTool.AlignGrid());
        }

        private void labelBottom_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Align label bottom", () => AlignTool.AlignLabels(ShapeExt.Anchor.Bottom));
        }

        private void labelTop_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Align label top", () => AlignTool.AlignLabels(ShapeExt.Anchor.Top));
        }

        private void labelLeft_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Align label left", () => AlignTool.AlignLabels(ShapeExt.Anchor.Left));
        }

        private void labelRight_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Align label right", () => AlignTool.AlignLabels(ShapeExt.Anchor.Right));
        }

        private void transpose_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Transpose", () => AlignTool.Transpose());
        }

        private void groupLabel_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Group labels", () => AlignTool.GroupLabels());
        }

        private void alignPrevSlide_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Align with previous slide", () => AlignTool.AlignWithSiblingSlide(-1));
        }

        private void alignNextSlide_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Align with next slide", () => AlignTool.AlignWithSiblingSlide(1));
        }

        private void curve_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Execute("Update curve expression", () =>
            {
                var equationX = Globals.Ribbons.EdoliRibbon.curveOfEquationX.Text;
                var equationY = Globals.Ribbons.EdoliRibbon.curveOfEquationY.Text;
                var startValue = Globals.Ribbons.EdoliRibbon.curveStart.Text;
                var endValue = Globals.Ribbons.EdoliRibbon.curveEnd.Text;
                ShapeTool.UpdatePathOfExpression(equationX, equationY, startValue, endValue);
            }, false);
        }

        private void animation_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Execute("Set animation name", () => AnimationTool.SetNameOfActive(this.animationName.Text), false);
        }

        private void swapCycle_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Swap cycle", () => AlignTool.SwapCycle());
        }

        private void swapCycleReverse_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Swap cycle reverse", () => AlignTool.SwapCycleReverse());
        }

        private void snapDownRight_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Snap down right", () => AlignTool.SnapDownRight());
        }

        private void snapUpRight_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Snap up right", () => AlignTool.SnapUpRight());
        }

        private void alignGrid_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Align horizontal vertical", () => AlignTool.AlignHorizontalVertical());
        }

        private void beginArrowToggle_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Toggle begin arrow", () => ShapeTool.BeginArrowToggle());
        }

        private void beginArrowChangeSize_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Increase begin arrow size", () => ShapeTool.BeginArrowChangeSize(1));
        }

        private void beginArrowSizeDown_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Decrease begin arrow size", () => ShapeTool.BeginArrowChangeSize(-1));
        }

        private void endArrowToggle_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Toggle end arrow", () => ShapeTool.EndArrowToggle());
        }

        private void endArrowSizeUp_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Increase end arrow size", () => ShapeTool.EndArrowChangeSize(1));
        }

        private void endArrowSizeDown_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Decrease end arrow size", () => ShapeTool.EndArrowChangeSize(-1));
        }

        private void connectShapeByLine_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Connect shapes by line", () => ShapeTool.ConnectShapesByLine());
        }
        private void drawAngle_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Draw angle", () => DimensionTool.DrawAngle());
        }
        private void drawDimension_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Draw dimension", () => DimensionTool.DistanceBetweenPoints());
        }
        private void resetDimension_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Reset dimension", () => DimensionTool.ResetDimensionScale());
        }

        private void curveOfEquation_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Add curve of equation", () =>
            {
                var equationX = Globals.Ribbons.EdoliRibbon.curveOfEquationX.Text;
                var equationY = Globals.Ribbons.EdoliRibbon.curveOfEquationY.Text;
                var startValue = Globals.Ribbons.EdoliRibbon.curveStart.Text;
                var endValue = Globals.Ribbons.EdoliRibbon.curveEnd.Text;
                ShapeTool.AddPathOfExpression(equationX, equationY, startValue, endValue, true);
            });
        }

        private void polylineOfEquation_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Add polyline of equation", () =>
            {
                var equationX = Globals.Ribbons.EdoliRibbon.curveOfEquationX.Text;
                var equationY = Globals.Ribbons.EdoliRibbon.curveOfEquationY.Text;
                var startValue = Globals.Ribbons.EdoliRibbon.curveStart.Text;
                var endValue = Globals.Ribbons.EdoliRibbon.curveEnd.Text;
                ShapeTool.AddPathOfExpression(equationX, equationY, startValue, endValue, false);
            });
        }

        private void trimImage_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Trim image", () => ImageTool.TrimImage());
        }

        private void resizeWidth_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Match width", () => AlignTool.MatchWidth());
        }

        private void resizeHeight_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Match height", () => AlignTool.MatchHeight());
        }

        private void followAnimation_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Follow animation", () => AnimationTool.FollowAnimation());
        }

        private void SVGTool_Click(object sender, RibbonControlEventArgs e)
        {
            Execute("Add SVG from clipboard", () => SVGtoPPTParser.AddSVGFigureFromClipboard());
        }
    }
}
