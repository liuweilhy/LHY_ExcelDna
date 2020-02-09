using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Hook;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace LHY_ExcelDna
{
    /// <summary>
    /// Load custom Excel Fluent/Ribbon
    /// </summary>
    [ComVisible(true)]
    public class RibbonUI : ExcelRibbon, IExcelAddIn
    {
        #region Member

        // 记录IRibbonUI对象
        private static IRibbonUI customRibbon;

        // App
        private Application xlapp = (Application)ExcelDnaUtil.Application;

        // Workbook
        private Workbook workbook = null;

        // Worksheet
        private Worksheet worksheet = null;

        // 处理键盘钩子
        private KeyboardHook keyboardHook = null;

        // 其它成员变量
        private bool isAbs = false;

        private bool isR1c1 = false;
        private string searchDir = string.Empty;
        private List<string> extensions = new List<string>();
        private bool isOnlyFile = true;
        private bool isIncluedSubDir = true;
        private bool isSearchAll = false;
        private bool isOpenFile = true;

        #endregion Member

        #region RibbonUI

        //https://msdn.microsoft.com/en-us/library/aa722523(v=office.12).aspx         //Ribbon函数回调定义
        //https://msdn.microsoft.com/zh-cn/library/office/ee691833(v=office.14).aspx  //Office 2010 Backstage 视图介绍

        public RibbonUI()
        {
            keyboardHook = new KeyboardHook("Test");
            keyboardHook.KeyDownEvent += ShortCut;
        }

        ~RibbonUI()
        {
            keyboardHook.KeyDownEvent -= ShortCut;
        }

        public void AutoOpen()
        {
            xlapp.WorkbookActivate += OnWorkbookActivate;
            xlapp.WorkbookBeforeClose += OnWorkbookBeforeClose;
        }

        private void OnWorkbookActivate(Workbook wb)
        {
            Invalidate();
        }

        private void OnWorkbookBeforeClose(Workbook wb, ref Boolean cancel)
        {
            if (!cancel)
                Invalidate();
        }

        public void AutoClose()
        {
            try
            {
                xlapp.WorkbookActivate -= OnWorkbookActivate;
                xlapp.WorkbookBeforeClose -= OnWorkbookBeforeClose;
            }
            catch { return; }
        }

        /// <summary>
        /// 定义热键
        /// </summary>
        /// <param name="e"></param>
        private void ShortCut(KeyboardHookEventArgs e)
        {
            if (xlapp == null || customRibbon == null)
            {
                return;
            }

            // Ctrl+Alt+F
            if (e.isCtrlPressed && e.isAltPressed && e.ToString().Equals("F"))
            {
                e.isShield = true;
                buttonSearch_onAction();
            }
            else if (e.isCtrlPressed && e.isAltPressed && e.ToString().Equals("W"))
            {
                e.isShield = true;
                isAbs = !isAbs;
                customRibbon.InvalidateControl("buttonAbs");
                buttonAbs_onAction(null, isAbs);
            }
            else if (e.isCtrlPressed && e.isAltPressed && e.ToString().Equals("E"))
            {
                e.isShield = true;
                isR1c1 = !isR1c1;
                customRibbon.InvalidateControl("buttonR1C1");
                buttonR1C1_onAction(null, isR1c1);
            }
            else
            {
                e.isShield = false;
            }
        }

        /// <summary>
        /// ribbon callback, get IRibbonUI object.
        /// </summary>
        public void OnLoad(IRibbonUI ribbon)
        {
            customRibbon = ribbon;
            customRibbon.Invalidate();
            Debug.WriteLine("IRibbonUI加载成功！");
        }

        public void Invalidate()
        {
            customRibbon.Invalidate();
        }

        /// <summary>
        /// read CustomUI.xml, xml file must be UTF-8 encode and Embedded resources.
        /// </summary>
        public override string GetCustomUI(string uiName)
        {
            string ribbonxml = string.Empty;
            try
            {
                if (ExcelDnaUtil.ExcelVersion > 12)
                {
                    ribbonxml = ResourceHelper.GetResourceText("CustomUI14.xml");
                }
                else
                {
                    throw new Exception("Do not support this Office Version.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return ribbonxml;
        }

        /// <summary>
        /// Ribbon callback，load image in XML element
        /// </summary>
        public override object LoadImage(string imageId)
        {
            return ResourceHelper.GetResourceBitmap(imageId);
        }

        /// <summary>
        /// control_getEnabled
        /// </summary>
        public bool control_getEnabled(IRibbonControl control)
        {
            // 如果没有打开工作表，则全部控件不可用
            if (xlapp.ActiveSheet == null)
                return false;
            // 对于comboBoxExName控件，Enabled取决于isOnlyFile
            if (control.Id == "comboBoxExName")
                return isOnlyFile;
            // 其余情况下控件可用
            return true;
        }

        /// <summary>
        /// toggleButton_getPressed
        /// </summary>
        public bool toggleButton_getPressed(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "buttonAbs":
                    return isAbs;

                case "buttonR1C1":
                    return isR1c1;
            }
            return true;
        }

        /// <summary>
        /// buttonAbs_onAction
        /// </summary>
        public void buttonAbs_onAction(IRibbonControl control, bool pressed)
        {
            isAbs = pressed;
            object referenceType;
            if (pressed)
            {
                referenceType = XlReferenceType.xlAbsolute;
            }
            else
            {
                referenceType = XlReferenceType.xlRelative;
            }

            try
            {
                worksheet = xlapp.ActiveSheet;
                Range formulaRange = xlapp.Selection;
                if (formulaRange == null || !formulaRange.HasFormula)
                {
                    return;
                }

                formulaRange = xlapp.Intersect(formulaRange,
                    formulaRange.SpecialCells(XlCellType.xlCellTypeFormulas));

                foreach (Range cell in formulaRange)
                {
                    cell.Formula = xlapp.ConvertFormula(cell.Formula,
                        XlReferenceStyle.xlA1, XlReferenceStyle.xlA1,
                        referenceType);
                }
            }
            catch { return; }
        }

        /// <summary>
        /// buttonR1C1_onAction
        /// </summary>
        public void buttonR1C1_onAction(IRibbonControl control, bool pressed)
        {
            isR1c1 = pressed;

            // 使用.Net API
            if (pressed)
            {
                xlapp.ReferenceStyle = XlReferenceStyle.xlR1C1;
            }
            else
            {
                xlapp.ReferenceStyle = XlReferenceStyle.xlA1;
            }
        }

        /// <summary>
        /// buttonCrack_onAction
        /// </summary>
        public void buttonCrack_onAction(IRibbonControl control)
        {
            worksheet = xlapp.ActiveSheet;
            if (worksheet == null)
            {
                return;
            }

            if (worksheet.ProtectContents == false)
            {
                MessageBox.Show("当前工作表无保护密码！", "无密码",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            /*
            char[] word = new char[] { '3','9' };
            string str = new string(word);
            try
            {
                worksheet = xlapp.ActiveSheet;
                worksheet.Unprotect(str);
            }
            catch (Exception ex)
            {
            }
            if (worksheet.ProtectContents == false)
            {
                MessageBox.Show("破解成功！等效密码：" + str, "破解成功",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            return;
            */

            // 工作表的写保护密码，被替换为12位字符
            // 前11位为“A”或“B”，第12位为char(32)～char(126)
            char[] passCharArray = new char[12];
            long count = 0;
            for (int i = 0; i <= 0b11111111111; i++)
            {
                for (int j = 0; j < 11; j++)
                {
                    if ((1 << j & i) != 0)
                    {
                        passCharArray[j] = 'B';
                    }
                    else
                    {
                        passCharArray[j] = 'A';
                    }
                }
                for (int c = 32; c <= 126; c++)
                {
                    count++;
                    passCharArray[11] = (char)c;
                    string password = new string(passCharArray);
                    System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
                    try
                    {
                        stopwatch.Start();
                        worksheet.Unprotect(password);
                    }
                    catch (Exception ex)
                    {
                        stopwatch.Stop();
                        TimeSpan timeSpan = stopwatch.Elapsed;
                        continue;
                    }
                    if (worksheet.ProtectContents == false)
                    {
                        MessageBox.Show("破解成功！等效密码：" + password, "破解成功",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// editBoxSearchDir_onChange
        /// </summary>
        public void editBoxSearchDir_onChange(IRibbonControl control, string text)
        {
            searchDir = text;
        }

        /// <summary>
        /// checkBox_getPressed
        /// </summary>
        public bool checkBox_getPressed(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "checkBoxOnlyFile":
                    return isOnlyFile;

                case "checkBoxIncludeSubDir":
                    return isIncluedSubDir;

                case "checkBoxSearchAll":
                    return isSearchAll;

                case "checkBoxOpenFile":
                    return isOpenFile;
            }
            return true;
        }

        /// <summary>
        /// checkBox_onAction
        /// </summary>
        public void checkBox_onAction(IRibbonControl control, bool pressed)
        {
            switch (control.Id)
            {
                case "checkBoxOnlyFile":
                    isOnlyFile = pressed;
                    break;

                case "checkBoxIncludeSubDir":
                    isIncluedSubDir = pressed;
                    break;

                case "checkBoxSearchAll":
                    isSearchAll = pressed;
                    break;

                case "checkBoxOpenFile":
                    isOpenFile = pressed;
                    break;

                default:
                    break;
            }
            customRibbon.Invalidate();
        }

        /// <summary>
        /// comboBoxExName_onChange
        /// </summary>
        public void comboBoxExName_onChange(IRibbonControl control, string text)
        {
            extensions.Clear();
            text = text.Replace(" ", "").Replace("*", "").Replace(".", "");
            foreach (string str in text.Split(';'))
            {
                if (!string.IsNullOrWhiteSpace(str) && !extensions.Contains(str))
                {
                    extensions.Add(str);
                }
            }
        }

        /// <summary>
        /// buttonSearch_onAction
        /// </summary>
        [ExcelCommand(Name = "buttonSearch_onAction")]
        public void buttonSearch_onAction(IRibbonControl control = null)
        {
            if (string.IsNullOrWhiteSpace(searchDir.Trim()))
            {
                MessageBox.Show("未指定目录！", "目录错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!searchDir.EndsWith(@"\"))
            {
                searchDir += @"\";
            }

            DirectoryInfo dir = new DirectoryInfo(searchDir);
            if (!dir.Exists)
            {
                MessageBox.Show("目录不存在！", "目录错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string keyword = null;
                if (xlapp.ActiveCell != null)
                {
                    keyword = xlapp.ActiveCell.Text;
                }
                if (string.IsNullOrWhiteSpace(keyword))
                {
                    MessageBox.Show("未指定要查找的文件", "查找错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                FileInfo[] files = null;
                DirectoryInfo[] directories = null;
                int found = 0;

                if (isOnlyFile)
                {
                    List<string> patterns = new List<string>();
                    foreach (string extension in extensions)
                    {
                        patterns.Add("*" + keyword + "*." + extension);
                    }
                    if (patterns.Count == 0)
                    {
                        patterns.Add("*" + keyword + "*");
                    }

                    foreach (string pattern in patterns)
                    {
                        if (isIncluedSubDir)
                        {
                            files = dir.GetFiles(pattern, SearchOption.AllDirectories);
                        }
                        else
                        {
                            files = dir.GetFiles(pattern, SearchOption.TopDirectoryOnly);
                        }

                        found = files.Length;
                        if (found == 0)
                        {
                            MessageBox.Show("在" + searchDir + "下未找到文件" + pattern, "未找到文件",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        int opened = 0;
                        foreach (FileInfo file in files)
                        {
                            if (isOpenFile)
                            {
                                System.Diagnostics.Process.Start("explorer.exe", file.FullName);
                            }
                            else
                            {
                                System.Diagnostics.Process.Start("explorer.exe", @" /select, " + file.FullName);
                            }
                            opened++;
                            if (!isSearchAll || opened >= 10)
                            {
                                return;
                            }
                        }
                    }
                }
                else
                {
                    string pattern = "*" + keyword + "*";
                    if (isIncluedSubDir)
                    {
                        files = dir.GetFiles(pattern, SearchOption.AllDirectories);
                        directories = dir.GetDirectories(pattern, SearchOption.AllDirectories);
                    }
                    else
                    {
                        files = dir.GetFiles(pattern, SearchOption.TopDirectoryOnly);
                        directories = dir.GetDirectories(pattern, SearchOption.TopDirectoryOnly);
                    }

                    found = files.Length + directories.Length;
                    if (found == 0)
                    {
                        MessageBox.Show("在" + searchDir + "下未找到文件（夹）" + pattern, "未找到文件（夹）",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    int opened = 0;
                    foreach (DirectoryInfo directory in directories)
                    {
                        if (isOpenFile)
                        {
                            System.Diagnostics.Process.Start("explorer.exe", directory.FullName);
                        }
                        else
                        {
                            System.Diagnostics.Process.Start("explorer.exe", @" /select, " + directory.FullName);
                        }
                        opened++;
                        if (!isSearchAll || opened >= 10)
                        {
                            return;
                        }
                    }

                    foreach (FileInfo file in files)
                    {
                        if (isOpenFile)
                        {
                            System.Diagnostics.Process.Start("explorer.exe", file.FullName);
                        }
                        else
                        {
                            System.Diagnostics.Process.Start("explorer.exe", @" /select, " + file.FullName);
                        }
                        opened++;
                        if (!isSearchAll || opened >= 10)
                        {
                            return;
                        }
                    }
                }
            }
            catch { return; }
        }

        /// <summary>
        /// buttonAboutShortcut_onAction
        /// </summary>
        public void buttonAboutShortcut_onAction(IRibbonControl control)
        {
            MessageBox.Show("Ctrl+Alt+W:  绝对引用和相对引用转换\n" +
                "Ctrl+Alt+E:   A1模式和R1C1模式转换\n" +
                "Ctrl+Alt+F:   查找文件\n" +
                "Ctrl+Q:         合并居中单元格",
                "快捷键说明", MessageBoxButtons.OK);
        }

        #endregion RibbonUI

        #region ExcelCommand
        
        /// <summary>
        /// 合并单元格
        /// </summary>
        /// 注意宏命令的快捷键只支持Ctrl+字母或数字，ShortCut里的字母要小写
        [ExcelCommand(Name = "Merge_UnMerge", ShortCut = "^q")]
        public static void Merge_UnMerge()
        {
            try
            {
                Application app = (Application)ExcelDnaUtil.Application;
                Range selection = app.Selection;
                if (true.Equals(selection.MergeCells))
                {
                    selection.UnMerge();
                    selection.HorizontalAlignment = XlHAlign.xlHAlignGeneral;
                    selection.VerticalAlignment = XlVAlign.xlVAlignCenter;
                }
                else
                {
                    selection.Merge();
                    selection.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    selection.VerticalAlignment = XlVAlign.xlVAlignCenter;
                }
            }
            catch { return; }
        }

        /*
        [ExcelCommand(MenuName = "功能示例", MenuText = "求和",
            //ShortCut = "^2",
            Name = "SumSelectRange")]
        public static void SumSelectRange()
        {
            ExcelReference selection = null;
            try
            {
                selection = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);
            }
            catch { return; }

            object sum = XlCall.Excel(XlCall.xlfSum, selection);
            ExcelReference target = new ExcelReference(0, 0);
            target.SetValue(sum);
        }
        */

        #endregion ExcelCommand

        #region ExcelFunction

        /*
        [ExcelFunction(Category = "LHY_ExcelDna插件",
            Description = "测试",
            IsHidden = false,
            IsMacroType = true,
            IsThreadSafe = false,
            Name = "NewAdd")]
        public static string NewAdd(int a, int b)
        {
            int c = a + b;
            return c.ToString();
        }
        */

        #endregion ExcelFunction
    }
}