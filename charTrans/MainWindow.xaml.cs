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
using System.Threading;

namespace charTrans
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private string SSelectPath;
        private string[] SFileNames;
        private string SSavePath;
        [DllImport("AT.dll", EntryPoint = "add")]
        extern static Double add(double a, double b);
        [DllImport("ADDTIME.dll", EntryPoint = "time")]
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

                double result = d1 + d2;
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
                double result = d1 * d2;
                this.RESULT.Text = result.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("请输入自然数");
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

        #region wordCOM组件的使用
        private void Button_Click_18(object sender, RoutedEventArgs e)
        {
            //this.Z.Visibility = Visibility.Visible;
            changeZ();
            //while (SV.WorkingWORD)
            //{
            //    //Thread.Sleep(100);
            //}
            ////this.Z.Visibility = Visibility.Collapsed;
            //SV.WorkingWORD = true;
            //MessageBox.Show("sbdjas");

        }
        private void changeZ()
        {
            //Thread.Sleep(1000);
            Thread t2 = new Thread(useWord);
            t2.Start();
            //t2.Join();
            //this.Z.Visibility = Visibility.Collapsed;
        }
        private void useWord()
        {
            MsWord.Application oWordApplic;
            MsWord.Document oDoc;
            try
            {
                //Console.WriteLine("开始了");
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
                //Console.WriteLine("正在生成小节文档");
                object section_nextPage = MsWord.WdBreakType.wdSectionBreakNextPage;
                object page_break = MsWord.WdBreakType.wdPageBreak;
                for (int i = 0; i < 4; i++)
                {
                    oDoc.Paragraphs[1].Range.InsertParagraphAfter();
                    oDoc.Paragraphs[1].Range.InsertBreak(ref section_nextPage);
                }

                //Console.WriteLine("正在插入摘要内容");
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
                //Console.WriteLine("正在保存文档");
                object file_name;
                file_name = doc_file_name;
                oDoc.SaveAs2(ref file_name);
                oDoc.Close();
                //Console.WriteLine("正在释放COM资源");
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
                //Console.WriteLine("正在结束word进程");
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
            //Console.WriteLine("结束了");
            //Console.ReadLine();
        }
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
