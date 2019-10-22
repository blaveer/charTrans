#region using
using DotNetSpeech;
using Microsoft.International.Converters.PinYinConverter;
using Microsoft.International.Converters.TraditionalChineseToSimplifiedConverter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using WinForm=System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.IO;
using System.Text.RegularExpressions;
using Fibonacci;
using System.Runtime.InteropServices;
using MsWord = Microsoft.Office.Interop.Word;
using MsExcel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Reflection;
using System.Diagnostics;
#endregion
namespace charTrans
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        #region 初始化代码
        private string SSelectPath;
        private string[] SFileNames;
        private string SSavePath;
        [DllImport("AADTIME.dll", EntryPoint = "add", CallingConvention = CallingConvention.StdCall)]
        extern static double add(double a, double b);
        [DllImport("AADTIME.dll", EntryPoint = "time", CallingConvention = CallingConvention.StdCall)]
        extern static Double time(double a, double b);
        public MainWindow()
        {
            InitializeComponent();
            SFileNames = null;
            SSelectPath = "";
            SSavePath = "";
            GetFile();
            this.Z.Visibility = Visibility.Hidden;
        }
        #endregion

        #region 已完成
        #region 拼音转换
        private void Button_Click(object sender, RoutedEventArgs e)     //获取拼音
        {
            string text = this.TB.Text.Trim();
            if (text.Length == 0)
            {
                return;
            }
            try
            {
                //for(int i = 0; i < text.Length; i++)
                //{

                //}
                char one_char = text.ToCharArray()[0];
                int ch_int = (int)one_char;
                string str_char_int = string.Format("{0}", ch_int);
                if (ch_int > 127)
                {
                    ChineseChar chineseChar = new ChineseChar(one_char);
                    IReadOnlyCollection<string> pinyin = chineseChar.Pinyins;
                    string pin_str = "\n  ";
                    foreach (string pin in pinyin)
                    {
                        pin_str += pin + "\r\n  ";
                    }
                    this.OUT.Text = pin_str;
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show("出现错误" + e1.ToString());
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)     //获得简体
        {
            string text = this.TB.Text.Trim();
            if (text.Length == 0)
            {
                return;
            }
            try
            {
                this.OUT.Text = "\n  " + ChineseConverter.Convert(text, ChineseConversionDirection.TraditionalToSimplified);
            }
            catch (Exception e1)
            {
                MessageBox.Show("出现错误" + e1.ToString());
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)      //获得繁体
        {
            string text = this.TB.Text.Trim();
            if (text.Length == 0)
            {
                return;
            }
            try
            {
                this.OUT.Text = "\n  " + ChineseConverter.Convert(text, ChineseConversionDirection.SimplifiedToTraditional);
            }
            catch (Exception e1)
            {
                MessageBox.Show("出现错误" + e1.ToString());
            }
        }


        private void Button_Click_3(object sender, RoutedEventArgs e)    //获得发音
        {
            string text = this.TB.Text.Trim();
            if (text.Length == 0)
            {
                return;
            }
            try
            {
                SpeechVoiceSpeakFlags spFlags = SpeechVoiceSpeakFlags.SVSFlagsAsync;
                SpVoice voice = new SpVoice();
                voice.Speak(text, spFlags);
            }
            catch (Exception e1)
            {
                MessageBox.Show("发生错误" + e1.ToString());
            }
        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void TB_TextChanged(object sender, TextChangedEventArgs e)
        {
            //string text = this.TB.Text;
            //if (text.Length < 2)
            //{
            //    return;
            //}
            //char one_char = text.ToCharArray()[0];
            //if(one_char.Equals(" "))
            //{
            //    return;
            //}
            //else
            //{
            //    this.TB.Text = " " + this.TB.Text; 
            //}
        }
        #endregion

        #region 文件合并
        private void Button_Click_4(object sender, RoutedEventArgs e)     //打开文件夹的按钮
        {
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.Title = "选择数据源文件";
            //openFileDialog.Filter = "txt文件|*.txt|所有文件|*.*";
            //openFileDialog.FileName = string.Empty;
            //openFileDialog.FilterIndex = 1;
            //openFileDialog.Multiselect = true;  
            //openFileDialog.RestoreDirectory = true;
            //openFileDialog.DefaultExt = "txt";
            //if (openFileDialog.ShowDialog() == false)
            //{
            //    return;
            //}
            //string []txtFile = openFileDialog.FileNames;
            WinForm.FolderBrowserDialog dialog = new WinForm.FolderBrowserDialog();
            WinForm.DialogResult result = dialog.ShowDialog();
            if (result == WinForm.DialogResult.Cancel)
            {
                return;
            }
            SSelectPath = dialog.SelectedPath;
            SFileNames = null;
            GetFile(SSelectPath);
        }
        private void GetFile()    //测试用函数
        {
            SSelectPath = @"E:\Course\windowsProgramDesign\H\TXT";
            string[] strNames = Directory.GetFiles(SSelectPath);
            SFileNames = Directory.GetFiles(SSelectPath);
            foreach (string s in strNames)
            {
                this.FilePath.Items.Add(s);
            }
        }

        private void GetFile(string dir)
        {
            string[] strNames = Directory.GetFiles(dir);
            SFileNames = Directory.GetFiles(SSelectPath);
            foreach (string s in strNames)
            {
                this.FilePath.Items.Add(s);
            }
        }

        private void FilePath_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Button_Click_5(object sender, RoutedEventArgs e)  //查找文件
        {
            string search = this.STB.Text.Trim();
            if (search.Equals(""))
            {
                return;
            }
            this.FilePath.Items.Clear();
            foreach (string s in SFileNames)
            {
                if (Regex.IsMatch(s, search))
                {
                    this.FilePath.Items.Add(s);
                }
            }

        }

        private void Button_Click_6(object sender, RoutedEventArgs e)     //添加到目标集中
        {
            foreach (string s in this.FilePath.Items)
            {
                this.FileFinal.Items.Add(s);
            }
        }

        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            this.FileFinal.Items.Clear();
        }

        private void Button_Click_8(object sender, RoutedEventArgs e)      //打开文件
        {
            string path = this.FileFinal.SelectedItem.ToString();
            //string path = SelectedItems"; //测试一个word文档
            System.Diagnostics.Process.Start(path);
        }

        private void Button_Click_9(object sender, RoutedEventArgs e)    //上移
        {
            int i = this.FileFinal.Items.Count;
            if (i == 0)
            {
                return;
            }
            if (this.FileFinal.SelectedItem == null)
            {
                return;
            }
            if (this.FileFinal.Items.GetItemAt(0).ToString().Equals(this.FileFinal.SelectedItem.ToString()))
            {
                return;
            }
            string temp = this.FileFinal.SelectedItem.ToString();
            List<string> tempStringS = new List<string>();
            int counter = 0;
            for (; counter < i; counter++)
            {
                string s = this.FileFinal.Items.GetItemAt(counter).ToString();
                if (s.Equals(temp))
                {
                    break;
                }
            }
            if (counter == i)
            {
                return;
            }
            for (int counter1 = 0; counter1 < i; counter1++)
            {
                if ((counter1 + 1) == counter)
                {
                    tempStringS.Add(this.FileFinal.Items.GetItemAt(counter).ToString());
                    tempStringS.Add(this.FileFinal.Items.GetItemAt(counter1).ToString());
                    counter1++;
                }
                else
                {
                    tempStringS.Add(this.FileFinal.Items.GetItemAt(counter1).ToString());
                }
            }
            this.FileFinal.Items.Clear();
            foreach (string s in tempStringS)
            {
                this.FileFinal.Items.Add(s);
            }

        }

        private void Button_Click_10(object sender, RoutedEventArgs e)    //下移
        {
            int i = this.FileFinal.Items.Count;
            if (i == 0)
            {
                return;
            }
            if (this.FileFinal.SelectedItem == null)
            {
                return;
            }
            if (this.FileFinal.Items.GetItemAt(i - 1).ToString().Equals(this.FileFinal.SelectedItem.ToString()))
            {
                return;
            }

            string temp = this.FileFinal.SelectedItem.ToString();
            List<string> tempStringS = new List<string>();
            int counter = 0;
            for (; counter < i; counter++)
            {
                string s = this.FileFinal.Items.GetItemAt(counter).ToString();
                if (s.Equals(temp))
                {
                    break;
                }
            }
            if (counter == i)
            {
                return;
            }
            for (int counter1 = 0; counter1 < i; counter1++)
            {
                if (counter1 == counter)
                {
                    tempStringS.Add(this.FileFinal.Items.GetItemAt(counter + 1).ToString());
                    tempStringS.Add(this.FileFinal.Items.GetItemAt(counter1).ToString());
                    counter1++;
                }
                else
                {
                    tempStringS.Add(this.FileFinal.Items.GetItemAt(counter1).ToString());
                }
            }
            this.FileFinal.Items.Clear();
            foreach (string s in tempStringS)
            {
                this.FileFinal.Items.Add(s);
            }
        }

        private void Button_Click_11(object sender, RoutedEventArgs e)    //设置合并后的文件的位置
        {
            WinForm.FolderBrowserDialog dialog = new WinForm.FolderBrowserDialog();
            WinForm.DialogResult result = dialog.ShowDialog();
            if (result == WinForm.DialogResult.Cancel)
            {
                return;
            }
            SSavePath = dialog.SelectedPath + @"\save.txt";
        }

        private void Button_Click_12(object sender, RoutedEventArgs e)    //合并文件
        {
            if (SSavePath.Equals(""))
            {
                MessageBox.Show("请选择输出目录");
                return;
            }
            //= this.ChangeLine.IsChecked;
            List<string> txt = new List<string>();
            foreach (string s in this.FileFinal.Items)
            {
                if (this.AddName.IsChecked == true)
                {
                    txt.Add(s);
                }
                using (StreamReader sr = new StreamReader(s, Encoding.Default))
                {
                    int lineCount = 0;
                    while (sr.Peek() > 0)
                    {
                        lineCount++;
                        string temp = sr.ReadLine();
                        txt.Add(temp);
                    }
                }
                if (this.ChangeLine.IsChecked == true)
                {
                    txt.Add("\n");
                }
            }
            using (FileStream fs = new FileStream(SSavePath, FileMode.Create))
            {
                using (StreamWriter sw = new StreamWriter(fs, Encoding.Default))
                {
                    for (int i = 0; i < txt.Count; i++)
                    {
                        sw.WriteLine(txt[i]);
                    }
                }
            }
            if (this.OpenMergeFile.IsChecked == true)
            {
                System.Diagnostics.Process.Start(SSavePath);
            }
        }

        #endregion

        #region DLL使用
        private void Button_Click_13(object sender, RoutedEventArgs e)   //求阶乘
        {
            string s = this.NumGetForF.Text.ToString().Trim();
            if (s.Equals(""))
            {
                return;
            }
            else
            {
                try
                {
                    int f = int.Parse(s);
                    int result = Fibonacci.Fibonacci.Factorial(f);
                    this.RESULT.Text = result.ToString();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("请输入整数");
                }
            }
        }

        private void Button_Click_14(object sender, RoutedEventArgs e)//求斐波那契
        {
            string s = this.NumGetForF.Text.ToString().Trim();
            if (s.Equals(""))
            {
                return;
            }
            else
            {
                try
                {
                    int f = int.Parse(s);
                    int result = Fibonacci.Fibonacci.SICFibonacci(f);
                    this.RESULT.Text = result.ToString();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("请输入整数");
                }
            }
        }

        private void Button_Click_15(object sender, RoutedEventArgs e)   //相加
        {
            // TODO
            string num1 = this.NUMONE.Text.ToString().Trim();
            string num2 = this.NUMTWO.Text.ToString().Trim();
            if (num1.Equals(""))
            {
                MessageBox.Show("请输入第一个运算数");
                return;
            }
            else if (num2.Equals(""))
            {
                MessageBox.Show("请输入第二个运算数");
                return;
            }
            try
            {
                double d1 = Double.Parse(num1);
                double d2 = Double.Parse(num2);

                double result = add(d1, d2);
                this.RESULT.Text = result.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("请输入自然数");
            }

        }

        private void Button_Click_16(object sender, RoutedEventArgs e)  //相乘
        {
            string num1 = this.NUMONE.Text.ToString().Trim();
            string num2 = this.NUMTWO.Text.ToString().Trim();
            if (num1.Equals(""))
            {
                MessageBox.Show("请输入第一个运算数");
                return;
            }
            else if (num2.Equals(""))
            {
                MessageBox.Show("请输入第二个运算数");
                return;
            }
            try
            {
                double d1 = Double.Parse(num1);
                double d2 = Double.Parse(num2);
                double result = time(d1, d2);
                this.RESULT.Text = result.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("请输入自然数");
            }

        }

        private void Button_Click_19(object sender, RoutedEventArgs e)  //反射机制
        {
            this.DLLMethod.Items.Clear();
            Type t = typeof(Fibonacci.Fibonacci);
            foreach (MethodInfo m in t.GetMethods())
            {
                this.DLLMethod.Items.Add(m.ToString());
            }
        }
        #endregion

        #region 自定义COM的使用
        #endregion

        #region 自定义COM的使用
        private void Button_Click_17(object sender, RoutedEventArgs e)
        {
            string s1 = this.ComNumOne.Text.ToString().Trim();
            string s2 = this.ComNumTwo.Text.ToString().Trim();
            if (s1.Equals(""))
            {
                MessageBox.Show("请输入第一个数");
                return;
            }
            else if (s2.Equals(""))
            {
                MessageBox.Show("请输入第二个数");
                return;
            }
            try
            {
                int i1 = int.Parse(s1);
                int i2 = int.Parse(s2);
                MyCOMTest.IADD cAdd = new MyCOMTest.CADD();
                int res = cAdd.add(i1, i2);
                this.ComNumRes.Text = res.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("请输入整数");
                return;
            }
        }

        #endregion
        #endregion

        #region COM组件的使用
        private void Button_Click_18(object sender, RoutedEventArgs e)
        {
            //this.Z.Visibility = Visibility.Visible;
            //Thread t2 = new Thread(useWord2);
            //t2.Start();
            //t2.Join();
            useWord2();
            MessageBox.Show("已完成");
            //while (SV.WorkingWORD)
            //{
            //    //Thread.Sleep(100);
            //}
            ////this.Z.Visibility = Visibility.Collapsed;
            //SV.WorkingWORD = true;
            //MessageBox.Show("sbdjas");

        }
        private void Button_Click_20(object sender, RoutedEventArgs e)
        {
            this.useExcel();
        }
        private void useWord()
        {
            MsWord.Application oWordApplic;
            MsWord.Document oDoc;
            try
            {
                ////Console.WriteLine("开始了");
                string doc_file_name = SV.outUrl + @"\content.doc";
                if (File.Exists(doc_file_name))
                {
                    File.Delete(doc_file_name);
                }
                oWordApplic = new MsWord.Application();
                object missing = System.Reflection.Missing.Value;

                //创建小节
                MsWord.Range curRange;
                object curTxt;
                int curSectionNum = 1;
                oDoc = oWordApplic.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                oDoc.Activate();
                ////Console.WriteLine("正在生成小节文档");
                object section_nextPage = MsWord.WdBreakType.wdSectionBreakNextPage;
                object page_break = MsWord.WdBreakType.wdPageBreak;
                for (int i = 0; i < 4; i++)
                {
                    oDoc.Paragraphs[1].Range.InsertParagraphAfter();
                    oDoc.Paragraphs[1].Range.InsertBreak(ref section_nextPage);
                }

                ////Console.WriteLine("正在插入摘要内容");
                #region
                curSectionNum = 1;
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                curRange.Select();
                string one_str, key_word;
                StreamReader le_abstract = new StreamReader(SV.basicUrl + @"\abstract.txt");
                oWordApplic.Options.Overtype = false;
                MsWord.Selection currentSelection = oWordApplic.Selection;
                if (currentSelection.Type == MsWord.WdSelectionType.wdSelectionNormal)
                {
                    one_str = le_abstract.ReadLine();
                    currentSelection.TypeText(one_str);
                    currentSelection.TypeParagraph();
                    currentSelection.TypeText("摘要");
                    currentSelection.TypeParagraph();
                    key_word = le_abstract.ReadLine();
                    one_str = le_abstract.ReadLine();
                    while (one_str != null)
                    {
                        currentSelection.TypeText(one_str);
                        currentSelection.TypeParagraph();
                        one_str = le_abstract.ReadLine();
                    }
                    currentSelection.TypeText("关键字:");
                    currentSelection.TypeText(key_word);
                    currentSelection.TypeParagraph();

                }
                le_abstract.Close();

                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                curTxt = curRange.Paragraphs[1].Range.Text;
                curRange.Font.Name = "宋体";
                curRange.Font.Size = 22;
                curRange.Paragraphs[1].Alignment = MsWord.WdParagraphAlignment.wdAlignParagraphCenter;
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[2].Range;
                curRange.Select();
                curRange.Paragraphs[1].Alignment = MsWord.WdParagraphAlignment.wdAlignParagraphCenter;
                curRange.Font.Name = "黑体";
                curRange.Font.Size = 16;
                //摘要正文
                oDoc.Sections[curSectionNum].Range.Paragraphs[1].Alignment = MsWord.WdParagraphAlignment.wdAlignParagraphCenter;
                for (int i = 3; i < oDoc.Sections[curSectionNum].Range.Paragraphs.Count; i++)
                {
                    curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[i].Range;
                    curTxt = curRange.Paragraphs[1].Range.Text;
                    curRange.Select();
                    curRange.Font.Name = "宋体";
                    curRange.Font.Size = 12;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].LineSpacingRule = MsWord.WdLineSpacing.wdLineSpaceMultiple;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].LineSpacing = 15f;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].IndentFirstLineCharWidth(2);

                }
                curRange = curRange.Paragraphs[curRange.Paragraphs.Count].Range;
                curTxt = curRange.Paragraphs[1].Range.Text;
                object range_start, range_end;
                range_start = curRange.Start;
                range_end = curRange.Start + 4;
                curRange = oDoc.Range(ref range_start, ref range_end);
                curTxt = curRange.Text;
                curRange.Select();
                curRange.Font.Bold = 1;
                #endregion


                //oDoc.Fields[1].Update();
                #region 
                ////Console.WriteLine("正在保存文档");
                object file_name;
                file_name = doc_file_name;
                oDoc.SaveAs2(ref file_name);
                oDoc.Close();
                ////Console.WriteLine("正在释放COM资源");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
                oDoc = null;
                oWordApplic.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWordApplic);
                oWordApplic = null;
                #endregion

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ////Console.WriteLine("正在结束word进程");
                System.Diagnostics.Process[] AllProces = System.Diagnostics.Process.GetProcesses();
                for (int i = 0; i < AllProces.Length; i++)
                {
                    string processName = AllProces[i].ProcessName;
                    if (String.Compare(processName, "WINWORD") == 0)
                    {
                        if (AllProces[i].Responding && !AllProces[i].HasExited)
                        {
                            AllProces[i].Kill();
                        }
                    }

                }
                MessageBox.Show("成功了");
                SV.WorkingWORD = false;
            }
            ////Console.WriteLine("结束了");
            ////Console.ReadLine();
        }

        private void useWord2()
        {
            MsWord.Application oWordApplic;
            MsWord.Document oDoc;
            string doc_file_name = SV.outUrl + @"\content.doc";
            try
            {
                if (File.Exists(doc_file_name))
                {
                    File.Delete(doc_file_name);
                }
                oWordApplic = new MsWord.Application();
                object missing = System.Reflection.Missing.Value;

                MsWord.Range curRange;
                object curTxt;
                int curSectionNum = 1;
                oDoc = oWordApplic.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                oDoc.Activate();
                ////Console.WriteLine(" 正在生成文档小节");


                object section_nextPage = MsWord.WdBreakType.wdSectionBreakNextPage;
                object page_break = MsWord.WdBreakType.wdPageBreak;
                for (int si = 0; si < 4; si++)
                {
                    oDoc.Paragraphs[1].Range.InsertParagraphAfter();
                    oDoc.Paragraphs[1].Range.InsertBreak(ref section_nextPage);
                }

                ////Console.WriteLine(" 正在插入摘要内容");

                #region 摘要部分
                curSectionNum = 1;
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                curRange.Select();
                string one_str, key_word;
                StreamReader file_abstract = new StreamReader(SV.basicUrl + @"\abstract.txt");
                oWordApplic.Options.Overtype = false;//overtype 改写模式
                MsWord.Selection currentSelection = oWordApplic.Selection;
                if (currentSelection.Type == MsWord.WdSelectionType.wdSelectionNormal)
                {
                    one_str = file_abstract.ReadLine();//读入题目
                    currentSelection.TypeText(one_str);
                    currentSelection.TypeParagraph(); //添加段落标记
                    currentSelection.TypeText(" 摘要");//写入" 摘要" 二字
                    currentSelection.TypeParagraph(); //添加段落标记
                    key_word = file_abstract.ReadLine();//读入题目
                    one_str = file_abstract.ReadLine();//读入段落文本
                    while (one_str != null)
                    {
                        currentSelection.TypeText(one_str);
                        currentSelection.TypeParagraph(); //添加段落标记
                        one_str = file_abstract.ReadLine();
                    }
                    currentSelection.TypeText(" 关键字：");
                    currentSelection.TypeText(key_word);
                    currentSelection.TypeParagraph(); //添加段落标记
                }
                file_abstract.Close();
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                curTxt = curRange.Paragraphs[1].Range.Text;
                curRange.Font.Name = " 宋体";
                curRange.Font.Size = 22;
                curRange.Paragraphs[1].Alignment =
                MsWord.WdParagraphAlignment.wdAlignParagraphCenter;
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[2].Range;
                curRange.Select();
                curRange.Paragraphs[1].Alignment =
                MsWord.WdParagraphAlignment.wdAlignParagraphCenter;
                curRange.Font.Name = " 黑体";
                curRange.Font.Size = 16;
                oDoc.Sections[curSectionNum].Range.Paragraphs[1].Alignment =
                MsWord.WdParagraphAlignment.wdAlignParagraphCenter;
                for (int i = 3; i < oDoc.Sections[curSectionNum].Range.Paragraphs.Count; i++)
                {
                    curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[i].Range;
                    curTxt = curRange.Paragraphs[1].Range.Text;
                    curRange.Select();
                    curRange.Font.Name = " 宋体";
                    curRange.Font.Size = 12;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].LineSpacingRule =
                    MsWord.WdLineSpacing.wdLineSpaceMultiple;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].LineSpacing = 15f;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].IndentFirstLineCharWidth(2);
                }
                curRange = curRange.Paragraphs[curRange.Paragraphs.Count].Range;
                curTxt = curRange.Paragraphs[1].Range.Text;
                object range_start, range_end;
                range_start = curRange.Start;
                range_end = curRange.Start + 4;
                curRange = oDoc.Range(ref range_start, ref range_end);
                curTxt = curRange.Text;
                curRange.Select();
                curRange.Font.Bold = 1;
                #endregion 摘要部分




                ////Console.WriteLine(" 正在插入目录");

                #region 目录
                curSectionNum = 2;
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                curRange.Select();
                object useheading_styles = true;//使用内置的目录标题样式
                object upperheading_level = 1;//最高的标题级别
                object lowerheading_level = 3;//最低标题级别
                object usefields = 1;//true 表示创建的是目录
                object tableid = 1;
                object RightAlignPageNumbers = true;//右边距对齐的页码
                object IncludePageNumbers = true;//目录中包含页码
                currentSelection = oWordApplic.Selection;
                currentSelection.TypeText(" 目录");
                currentSelection.TypeParagraph();
                currentSelection.Select();
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[2].Range;
                curRange.Collapse();
                oDoc.TablesOfContents.Add(curRange, ref useheading_styles, ref upperheading_level,
                ref lowerheading_level, ref usefields, ref tableid, ref RightAlignPageNumbers,
                ref IncludePageNumbers, ref missing, ref missing, ref missing, ref missing);
                oDoc.Sections[curSectionNum].Range.Paragraphs[1].Alignment =
                MsWord.WdParagraphAlignment.wdAlignParagraphCenter;
                oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range.Font.Bold = 1;
                oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range.Font.Name = " 黑体";
                oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range.Font.Size = 16;
                #endregion 目录


                #region 第一章

                curSectionNum = 3;
                oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range.Select();
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                ////Console.WriteLine(" 正在设置标题样式");


                object wdFontSizeIndex;
                wdFontSizeIndex = 14;
                oWordApplic.ActiveDocument.Styles.get_Item(ref wdFontSizeIndex).ParagraphFormat.Alignment =
                MsWord.WdParagraphAlignment.wdAlignParagraphCenter;
                oWordApplic.ActiveDocument.Styles.get_Item(ref wdFontSizeIndex).Font.Name = "黑体";
                oWordApplic.ActiveDocument.Styles.get_Item(ref wdFontSizeIndex).Font.Size = 16;//三号
                wdFontSizeIndex = 15;
                oWordApplic.ActiveDocument.Styles.get_Item(ref wdFontSizeIndex).Font.Name = "黑体";
                oWordApplic.ActiveDocument.Styles.get_Item(ref wdFontSizeIndex).Font.Size = 15;//小三
                object Style1 = MsWord.WdBuiltinStyle.wdStyleHeading1;
                object Style2 = MsWord.WdBuiltinStyle.wdStyleHeading2;
                oDoc.Sections[curSectionNum].Range.Select();
                currentSelection = oWordApplic.Selection;
                StreamReader file_content = new StreamReader(SV.basicUrl + @"\content.txt");
                one_str = file_content.ReadLine();//一级标题
                currentSelection.TypeText(one_str);
                currentSelection.TypeParagraph(); //添加段落标记
                one_str = file_content.ReadLine();//二级标题
                currentSelection.TypeText(one_str);
                currentSelection.TypeParagraph(); //添加段落标记
                one_str = file_content.ReadLine();//正文
                while (one_str != null)
                {
                    currentSelection.TypeText(one_str);
                    currentSelection.TypeParagraph(); //添加段落标记
                    one_str = file_content.ReadLine();//正文
                }
                file_content.Close();
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                curRange.set_Style(ref Style1);
                oDoc.Sections[curSectionNum].Range.Paragraphs[1].Alignment =
                MsWord.WdParagraphAlignment.wdAlignParagraphCenter;
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[2].Range;
                curRange.set_Style(ref Style2);
                for (int i = 3; i < oDoc.Sections[curSectionNum].Range.Paragraphs.Count; i++)
                {
                    curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[i].Range;
                    curRange.Select();
                    curRange.Font.Name = "宋体";
                    curRange.Font.Size = 12;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].LineSpacingRule =
                    MsWord.WdLineSpacing.wdLineSpaceMultiple;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].LineSpacing = 15f;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].IndentFirstLineCharWidth(2);
                }
                #endregion 第一章

                ////Console.WriteLine(" 正在插入第二章内容");

                #region 第二章表格
                curSectionNum = 4;
                oDoc.Sections[curSectionNum].Range.Select();
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                currentSelection = oWordApplic.Selection;
                currentSelection.TypeText("2 表格");
                currentSelection.TypeParagraph();
                currentSelection.TypeText(" 表格示例");
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[3].Range;

                oDoc.Sections[curSectionNum].Range.Paragraphs[3].Range.Select();
                currentSelection = oWordApplic.Selection;
                MsWord.Table oTable;
                oTable = curRange.Tables.Add(curRange, 5, 3, ref missing, ref missing);
                oTable.Range.ParagraphFormat.Alignment =
                MsWord.WdParagraphAlignment.wdAlignParagraphCenter;
                oTable.Range.Font.Name = " 宋体";
                oTable.Range.Font.Size = 16;
                oTable.Range.Cells.VerticalAlignment =
                MsWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oTable.Range.Rows.Alignment = MsWord.WdRowAlignment.wdAlignRowCenter;
                oTable.Columns[1].Width = 80;
                oTable.Columns[2].Width = 180;
                oTable.Columns[3].Width = 80;
                oTable.Cell(1, 1).Range.Text = " 字段";
                oTable.Cell(1, 2).Range.Text = " 描述";
                oTable.Cell(1, 3).Range.Text = " 数据类型";
                oTable.Cell(2, 1).Range.Text = "ProductID";
                oTable.Cell(2, 2).Range.Text = " 产品标识";
                oTable.Cell(2, 3).Range.Text = " 字符串";
                oTable.Borders.InsideLineStyle = MsWord.WdLineStyle.wdLineStyleSingle;
                oTable.Borders.OutsideLineStyle = MsWord.WdLineStyle.wdLineStyleSingle;
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                curRange.set_Style(ref Style1);
                curRange.ParagraphFormat.Alignment =
                MsWord.WdParagraphAlignment.wdAlignParagraphCenter;



                #endregion 第二章

                ////Console.WriteLine(" 正在插入第三章内容");

                #region 第三章图片
                curSectionNum = 5;
                oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range.Select();
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                currentSelection = oWordApplic.Selection;
                currentSelection.TypeText("3 图片");
                currentSelection.TypeParagraph();
                currentSelection.TypeText(" 图片示例");
                currentSelection.TypeParagraph();

                currentSelection.InlineShapes.AddPicture(SV.basicUrl + @"\whu.png",
                ref missing, ref missing, ref missing);
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                curRange.set_Style(ref Style1);
                curRange.ParagraphFormat.Alignment =
                MsWord.WdParagraphAlignment.wdAlignParagraphCenter;


                #endregion 第三章


                ////Console.WriteLine(" 正在设置第一节摘要页眉内容");

                curSectionNum = 1;
                oDoc.Sections[curSectionNum].Range.Select();
                oWordApplic.ActiveWindow.ActivePane.View.SeekView =
                MsWord.WdSeekView.wdSeekCurrentPageFooter;
                oDoc.Sections[curSectionNum].
                Headers[MsWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].
                Range.Borders[MsWord.WdBorderType.wdBorderBottom].LineStyle =
                MsWord.WdLineStyle.wdLineStyleNone;
                oWordApplic.Selection.HeaderFooter.PageNumbers.RestartNumberingAtSection = true;
                oWordApplic.Selection.HeaderFooter.PageNumbers.NumberStyle
                = MsWord.WdPageNumberStyle.wdPageNumberStyleUppercaseRoman;
                oWordApplic.Selection.HeaderFooter.PageNumbers.StartingNumber = 1;
                oWordApplic.ActiveWindow.ActivePane.View.SeekView =
                MsWord.WdSeekView.wdSeekMainDocument;
                ////Console.WriteLine(" 正在设置第二节目录页眉内容");

                curSectionNum = 2;
                oDoc.Sections[curSectionNum].Range.Select();
                oWordApplic.ActiveWindow.ActivePane.View.SeekView =
                MsWord.WdSeekView.wdSeekCurrentPageFooter;
                oDoc.Sections[curSectionNum].
                Headers[MsWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].
                Range.Borders[MsWord.WdBorderType.wdBorderBottom].LineStyle =
                MsWord.WdLineStyle.wdLineStyleNone;
                oWordApplic.Selection.HeaderFooter.PageNumbers.RestartNumberingAtSection = false;
                oWordApplic.Selection.HeaderFooter.PageNumbers.NumberStyle
                = MsWord.WdPageNumberStyle.wdPageNumberStyleUppercaseRoman;
                oWordApplic.ActiveWindow.ActivePane.View.SeekView =
                MsWord.WdSeekView.wdSeekMainDocument;
                curSectionNum = 3;
                oDoc.Sections[curSectionNum].Range.Select();
                oWordApplic.ActiveWindow.ActivePane.View.SeekView =
                MsWord.WdSeekView.wdSeekCurrentPageFooter;
                currentSelection = oWordApplic.Selection;
                curRange = currentSelection.Range;
                oWordApplic.Selection.HeaderFooter.PageNumbers.RestartNumberingAtSection = true;
                oWordApplic.Selection.HeaderFooter.PageNumbers.NumberStyle
                = MsWord.WdPageNumberStyle.wdPageNumberStyleArabic;
                oWordApplic.Selection.HeaderFooter.PageNumbers.StartingNumber = 1;
                object fieldpage = MsWord.WdFieldType.wdFieldPage;
                oWordApplic.Selection.Fields.Add(oWordApplic.Selection.Range,
                ref fieldpage, ref missing, ref missing);
                oWordApplic.Selection.ParagraphFormat.Alignment =
                MsWord.WdParagraphAlignment.wdAlignParagraphCenter;
                oDoc.Sections[curSectionNum].Headers[Microsoft.Office.Interop.
                Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;

                oWordApplic.ActiveWindow.ActivePane.View.SeekView = MsWord.WdSeekView.wdSeekMainDocument;
                ////Console.WriteLine(" 正在更新目录");
                ;
                oDoc.Fields[1].Update();
                #region 保存文档

                ////Console.WriteLine(" 正在保存 Word 文档");

                object fileName;
                fileName = doc_file_name;
                oDoc.SaveAs2(ref fileName);
                oDoc.Close();
                ////Console.WriteLine("正在释放 COM 资源");




                //释放 COM 资源
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
                oDoc = null;
                oWordApplic.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWordApplic);
                oWordApplic = null;

                System.GC.Collect();
                #endregion 保存文档 
            }
            catch (Exception e2)
            {
                MessageBox.Show("生成错误" + e2.Message);
            }
            finally
            {

                //Console.WriteLine(" 正在结束 Word 进程");

                //关闭 word 进程
                Process[] AllProces = Process.GetProcesses();
                for (int j = 0; j < AllProces.Length; j++)
                {
                    string theProcName = AllProces[j].ProcessName;
                    if (String.Compare(theProcName, "WINWORD") == 0)
                    {
                        if (AllProces[j].Responding && !AllProces[j].HasExited)
                        {
                            AllProces[j].Kill();
                        }
                    }
                }



            }
        }

        private void useExcel()
        {
            string src_file_name = SV.basicUrl+@"\list.csv";
            string dest_file_name = SV.outUrl+@"\list.xlsx";
            MsExcel.Application oExcApp;
            MsExcel.Workbook oExcBook;
            try
            {
                if (File.Exists(dest_file_name))
                {
                    File.Delete(dest_file_name);
                }
                oExcApp = new MsExcel.Application();
                object missing = System.Reflection.Missing.Value;
                oExcBook = oExcApp.Workbooks.Add(true);
                MsExcel.Worksheet worksheet1 = (MsExcel.Worksheet)oExcBook.Worksheets["sheet1"];
                worksheet1.Activate();
                oExcApp.Visible = false;
                oExcApp.DisplayAlerts = false;
                MsExcel.Range range1 = worksheet1.get_Range("B1", "H2");
                range1.Columns.ColumnWidth = 8;
                range1.Columns.RowHeight = 20;
                range1.Merge(false);
                range1.VerticalAlignment = MsExcel.XlVAlign.xlVAlignCenter;
                range1.HorizontalAlignment = MsExcel.XlHAlign.xlHAlignCenter;
                range1.Font.Size = 20;
                range1.Font.Bold = true;

                worksheet1.Cells[1, 2] = "学生成绩单";
                worksheet1.Cells[3, 1] = "学号";
                worksheet1.Cells[3, 2] = "姓名";
                worksheet1.Columns[1].ColumnWidth = 12;
                StreamReader sw = new StreamReader(src_file_name);
                string a_str;
                string[] str_list;
                int i = 4;
                a_str = sw.ReadLine();
                while (a_str != null)
                {
                    str_list = a_str.Split(",".ToCharArray());
                    worksheet1.Cells[i, 1] = str_list[0];
                    worksheet1.Cells[i, 2] = str_list[1];
                    i++;
                    a_str = sw.ReadLine();
                }
                sw.Close();
                for (int i1 = 0; i1 < 5; i1++)
                {
                    for (int j = 0; j < 8; j++)
                    {
                        worksheet1.Cells[i1 + 18, j + 3].Value2 = new Random(Guid.NewGuid().GetHashCode()).Next(0, 100);
                        worksheet1.Cells[i1 + 4, j + 3].Value2 = worksheet1.Cells[i1 + 18, j + 3].Value;
                    }
                }

                //添加图表
                MsExcel.Shape theShape = worksheet1.Shapes.AddChart(MsExcel.XlChartType.xl3DColumn, 120, 130, 380, 250);


                worksheet1.Cells[3, 3].Value2 = "美术";
                worksheet1.Cells[3, 4].Value2 = "物理";
                worksheet1.Cells[3, 5].Value2 = "政治";
                worksheet1.Cells[3, 6].Value2 = "化学";
                worksheet1.Cells[3, 7].Value2 = "体育";
                worksheet1.Cells[3, 8].Value2 = "英语";
                worksheet1.Cells[3, 9].Value2 = "数学";
                worksheet1.Cells[3, 10].Value2 = "历史";
                //设定图表的数据区域
                MsExcel.Range range = worksheet1.get_Range("b3:j8");
                theShape.Chart.SetSourceData(range, Type.Missing);

                //设置图标题文本
                theShape.Chart.HasTitle = true;
                theShape.Chart.ChartTitle.Text = "学生成绩";
                theShape.Chart.ChartTitle.Caption = "学生成绩";

                //设置单元格边框线型
                range1 = worksheet1.get_Range("a3", "j8");
                range1.Borders.LineStyle = MsExcel.XlLineStyle.xlContinuous;

                oExcBook.RefreshAll();
                worksheet1 = null;
                object file_name = dest_file_name;
                oExcBook.Close(true, file_name, null);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcBook);
                oExcBook = null;

                oExcApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcApp);
                oExcApp = null;
                System.GC.Collect();

            }
            catch (Exception e2)
            {
                //Console.WriteLine(e2.Message);
                MessageBox.Show("生成出错" + e2.Message);
            }
            finally
            {
                //Console.WriteLine(" 正在结束 excel 进程");

                //关闭 excel 进程
                Process[] AllProces = Process.GetProcesses();
                for (int j = 0; j < AllProces.Length; j++)
                {
                    string theProcName = AllProces[j].ProcessName;
                    if (String.Compare(theProcName, "EXCEL") == 0)
                    {
                        if (AllProces[j].Responding && !AllProces[j].HasExited)
                        {
                            AllProces[j].Kill();
                        }
                    }
                }

                MessageBox.Show("生成成功");


            }
        }


        #endregion

        #region 线程使用

        #endregion


    }

    #region 自定义常量类
    class SV
    {     //放一些静态常量
        public static string basicUrl = @"E:\Course\windowsProgramDesign\ProjectForHomeWork\Test\ConsoleCom\03_COM_material";
        public static string outUrl = @"E:\Course\windowsProgramDesign\ProjectForHomeWork\Test\ConsoleCom\out";
        public static bool WorkingWORD = true;
    }
    #endregion
}
