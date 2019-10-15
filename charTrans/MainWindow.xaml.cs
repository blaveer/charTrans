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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace charTrans
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

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
                    foreach(string pin in pinyin)
                    {
                        pin_str += pin + "\r\n  ";
                    }
                    this.OUT.Text = pin_str;
                }
            }catch(Exception e1)
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
            catch(Exception e1)
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
            }catch(Exception e1)
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
    }
}
