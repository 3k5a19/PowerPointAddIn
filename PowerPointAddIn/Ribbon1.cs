using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            ToggleGlowEffectOnSelection();
        }

        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            ToggleTextBoxFormatSettings();
        }


        public void ToggleGlowEffectOnSelection()
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app.ActiveWindow.Selection;

                // 光彩設定
                float desiredRadius = 10f;
                float desiredTransparency = 0f;
                int desiredColorRGB = Color.White.ToArgb();

                switch (selection.Type)
                {
                    case PowerPoint.PpSelectionType.ppSelectionText:
                        {
                            var textRange = selection.TextRange2;
                            var glow = textRange.Font.Glow;

                            if (glow.Radius > 0)
                            {
                                glow.Radius = 0;
                                //MessageBox.Show("選択されたテキストの光彩をオフにしました。");
                            }
                            else
                            {
                                glow.Radius = desiredRadius;
                                glow.Transparency = desiredTransparency;
                                glow.Color.RGB = desiredColorRGB;
                                //MessageBox.Show("選択されたテキストに光彩を設定しました。");
                            }
                            break;
                        }

                    case PowerPoint.PpSelectionType.ppSelectionShapes:
                        {
                            var shapes = selection.ShapeRange;
                            bool allHaveGlow = true;

                            // 全てに光彩があるかチェック
                            for (int i = 1; i <= shapes.Count; i++)
                            {
                                var tf2 = shapes[i].TextFrame2;
                                if (tf2.HasText == Office.MsoTriState.msoTrue)
                                {
                                    var glow = tf2.TextRange.Font.Glow;
                                    if (glow.Radius == 0)
                                    {
                                        allHaveGlow = false;
                                        break;
                                    }
                                }
                            }

                            for (int i = 1; i <= shapes.Count; i++)
                            {
                                var tf2 = shapes[i].TextFrame2;
                                if (tf2.HasText == Office.MsoTriState.msoTrue)
                                {
                                    var glow = tf2.TextRange.Font.Glow;

                                    if (allHaveGlow)
                                    {
                                        glow.Radius = 0;
                                    }
                                    else
                                    {
                                        glow.Radius = desiredRadius;
                                        glow.Transparency = desiredTransparency;
                                        glow.Color.RGB = desiredColorRGB;
                                    }
                                }
                            }

                            //MessageBox.Show(allHaveGlow ? "全図形の光彩をオフにしました。" : "全図形に光彩を設定しました。");
                            break;
                        }

                    default:
                        //MessageBox.Show("図形またはテキストを選択してください。");
                        break;
                }
            }

        public void ToggleTextBoxFormatSettings()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                //MessageBox.Show("図形を選択してください。");
                return;
            }

            var shapes = selection.ShapeRange;
            bool needsAutoSizeOn = false;

            // AutoSize 状態をチェック（すべてが msoAutoSizeShapeToFitText でなければ ON に統一）
            for (int i = 1; i <= shapes.Count; i++)
            {
                var tf = shapes[i].TextFrame2;
                if (tf.HasText == Office.MsoTriState.msoTrue)
                {
                    if (tf.AutoSize != Office.MsoAutoSize.msoAutoSizeShapeToFitText)
                    {
                        needsAutoSizeOn = true;
                        break;
                    }
                }
            }

            // 各図形に対して設定を適用
            for (int i = 1; i <= shapes.Count; i++)
            {
                var shape = shapes[i];
                var tf = shape.TextFrame2;

                if (tf.HasText != Office.MsoTriState.msoTrue)
                    continue;

                var textRange = tf.TextRange;
                string text = textRange.Text;

                // 1. 改段落 ↔ 改行
                if (text.Contains("\r"))
                {
                    textRange.Text = text.Replace("\r", "\n");
                }
                else if (text.Contains("\n"))
                {
                    textRange.Text = text.Replace("\n", "\r");
                }

                // 2. AutoSize: サイズに合わせて図形を調整
                tf.AutoSize = needsAutoSizeOn
                    ? Office.MsoAutoSize.msoAutoSizeShapeToFitText
                    : Office.MsoAutoSize.msoAutoSizeNone;

                // 3. 折り返しのトグル
                tf.WordWrap = (tf.WordWrap == Office.MsoTriState.msoTrue)
                    ? Office.MsoTriState.msoFalse
                    : Office.MsoTriState.msoTrue;
            }

            //MessageBox.Show("テキストボックスの設定を更新しました。");
        }

    }

}
