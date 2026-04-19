using Microsoft.Office.Tools.Ribbon;
using SlideSorterOnPlay.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SlideSorterOnPlay
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            toggleButton1.Checked = Settings.Default.Enabled;
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            if (Settings.Default.Enabled != toggleButton1.Checked)
            {
                Settings.Default.Enabled = toggleButton1.Checked;
                Settings.Default.Save();
                if (Settings.Default.Enabled)
                {
                    Globals.ThisAddIn.InstallEvents();
                }
                else
                {
                    Globals.ThisAddIn.UninstallEvents();
                }
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(@"
**PowerPoint 浏览式播放插件说明**

本插件改变了 PowerPoint 的放映方式：

1. **进入放映**
   * 当你点击「放映」时，主窗口会保持在 **幻灯片浏览视图（Tiled/Slide Sorter）**，可以直接看到所有幻灯片。

2. **切换页面**
   * 在主窗口中点击任意幻灯片，即可立即切换到该页进行播放。
   * 按 ESC 退出播放。使用 → 和 ← 键或鼠标点击切换页面。

点击上方""浏览式播放""按钮,可以启用或禁用此功能。", "帮助");
        }
    }
}
