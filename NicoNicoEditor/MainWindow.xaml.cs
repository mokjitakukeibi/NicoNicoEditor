using System;
using System.IO;
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
using System.Windows.Forms;
using System.Text.RegularExpressions;
using MarkupConverter;
using HtmlAgilityPack;
using System.Globalization;

namespace WpfApplication4
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Color_Click(object sender, RoutedEventArgs e)
        {
            ColorDialog cd = new ColorDialog();
            if (cd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                System.Windows.Media.Color color = new System.Windows.Media.Color();
                color.A = cd.Color.A;
                color.R = cd.Color.R;
                color.G = cd.Color.G;
                color.B = cd.Color.B;
                if (rtb != null)
                {
                    SolidColorBrush brush = new SolidColorBrush(color);
                    rtb.Selection.ApplyPropertyValue(Run.ForegroundProperty, brush);
                }
            }
        }

        private void Bold_Click(object sender, RoutedEventArgs e)
        {
            if (rtb != null)
            {
                if (rtb.Selection.GetPropertyValue(Run.FontWeightProperty) is FontWeight && ((FontWeight)rtb.Selection.GetPropertyValue(Run.FontWeightProperty)) == FontWeights.Normal)
                    rtb.Selection.ApplyPropertyValue(Run.FontWeightProperty, FontWeights.Bold);
                else
                    rtb.Selection.ApplyPropertyValue(Run.FontWeightProperty, FontWeights.Normal);
            }
        }

        private void Italic_Click(object sender, RoutedEventArgs e)
        {
            if (rtb != null)
            {
                if (rtb.Selection.GetPropertyValue(Run.FontStyleProperty) is FontStyle && ((FontStyle)rtb.Selection.GetPropertyValue(Run.FontStyleProperty)) == FontStyles.Normal)
                    rtb.Selection.ApplyPropertyValue(Run.FontStyleProperty, FontStyles.Italic);
                else
                    rtb.Selection.ApplyPropertyValue(Run.FontStyleProperty, FontStyles.Normal);
            }
        }

        private void Underline_Click(object sender, RoutedEventArgs e)
        {
            if (rtb != null)
            {
                TextDecorationCollection tdc;
                try
                {
                    tdc = (TextDecorationCollection)rtb.Selection.GetPropertyValue(Run.TextDecorationsProperty);
                }
                catch (InvalidCastException)
                {
                    tdc = null;
                }
                if (tdc == null || !tdc.SequenceEqual(TextDecorations.Underline))
                    rtb.Selection.ApplyPropertyValue(Run.TextDecorationsProperty, TextDecorations.Underline);
                else
                    rtb.Selection.ApplyPropertyValue(Run.TextDecorationsProperty, null);
            }
        }

        private void Strikethrough_Click(object sender, RoutedEventArgs e)
        {
            if (rtb != null)
            {
                TextDecorationCollection tdc;
                try
                {
                    tdc = (TextDecorationCollection)rtb.Selection.GetPropertyValue(Run.TextDecorationsProperty);
                }
                catch (InvalidCastException)
                {
                    tdc = null;
                }
                if (tdc == null || !tdc.SequenceEqual(TextDecorations.Strikethrough))
                    rtb.Selection.ApplyPropertyValue(Run.TextDecorationsProperty, TextDecorations.Strikethrough);
                else
                    rtb.Selection.ApplyPropertyValue(Run.TextDecorationsProperty, null);
            }
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (rtb != null)
            {
                double fontSize = 12;
                switch ((Sizes.SelectedItem as ComboBoxItem).Tag.ToString())
                {
                    case "1":
                        fontSize = 10;
                        break;
                    case "2":
                        fontSize = 11;
                        break;
                    case "3":
                        fontSize = 18;
                        break;
                    case "4":
                        fontSize = 24;
                        break;
                    case "5":
                        fontSize = 30;
                        break;
                    default:
                        fontSize = 12;
                        break;
                }
                rtb.Selection.ApplyPropertyValue(Run.FontSizeProperty, fontSize);
            }
        }

        private void ToHtml_Click(object sender, RoutedEventArgs e)
        {
            TextRange range;
            FileStream fStream;
            IMarkupConverter mc = new MarkupConverter.MarkupConverter();
            range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            fStream = new FileStream("toHtml.txt", FileMode.Create);
            range.Save(fStream, System.Windows.DataFormats.Rtf);
            fStream.Close();
            fStream = new FileStream("toHtml.txt", FileMode.Open);
            StreamReader sr = new StreamReader(fStream);
            string str = mc.ConvertRtfToHtml(sr.ReadToEnd());
            //delete inner <SPAN> and </SPAN> tag
            str = Regex.Replace(str, "<SPAN (.*?)><SPAN>", "<SPAN $1>");
            str = str.Replace("<SPAN><SPAN>", "<SPAN>");
            str = str.Replace("</SPAN></SPAN>", "</SPAN>");
            str = HtmlToNiconamaCode(str);
            htmlText.Text = str;
            sr.Close();
            fStream.Close();
            StringInfo si = new StringInfo(htmlText.Text);
            Count.Content = "コードの文字数：" + si.LengthInTextElements.ToString();
        }

        private string HtmlToNiconamaCode(string html)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);

            //span
            if (doc.DocumentNode.SelectNodes("//span") != null)
            {
                foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//span"))
                {
                    int fontSize = 6;
                    bool isBold = false;
                    bool isUnderline = false;
                    bool isThrough = false;
                    string colorCode = "";
                    if (node.HasAttributes)
                    {
                        string attribute = node.GetAttributeValue("style", "");
                        if (attribute.IndexOf("font-size:10") >= 0)
                        {
                            fontSize = 1;
                        }
                        if (attribute.IndexOf("font-size:11") >= 0)
                        {
                            fontSize = 2;
                        }
                        if (attribute.IndexOf("font-size:18") >= 0)
                        {
                            fontSize = 3;
                        }
                        if (attribute.IndexOf("font-size:24") >= 0)
                        {
                            fontSize = 4;
                        }
                        if (attribute.IndexOf("font-size:30") >= 0)
                        {
                            fontSize = 5;
                        }
                        int index = 0;
                        if ((index = attribute.IndexOf("color:")) >= 0)
                        {
                            colorCode = attribute.Substring(index + 6, 7);
                        }
                        if (attribute.IndexOf("font-weight:bold") >= 0)
                        {
                            isBold = true;
                        }
                        if (attribute.IndexOf("text-decoration:underline") >= 0)
                        {
                            isUnderline = true;
                        }
                        if (attribute.IndexOf("text-decoration:line-through") >= 0)
                        {
                            isThrough = true;
                        }
                    }
                    if (fontSize == 6)
                    {
                        if (colorCode == "")
                        {
                            node.ParentNode.ReplaceChild(doc.CreateTextNode((isBold ? "<b>" : "") + (isUnderline ? "<u>" : "") + (isThrough ? "<s>" : "") + node.InnerHtml + (isThrough ? "</s>" : "") + (isUnderline ? "</u>" : "") + (isBold ? "</b>" : "")), node);
                        }
                        else
                        {
                            node.ParentNode.ReplaceChild(doc.CreateTextNode("<font color=\"" + colorCode + "\">" + (isBold ? "<b>" : "") + (isUnderline ? "<u>" : "") + (isThrough ? "<s>" : "") + node.InnerHtml + (isThrough ? "</s>" : "") + (isUnderline ? "</u>" : "") + (isBold ? "</b>" : "") + "</font>"), node);
                        }
                    }
                    else
                    {
                        if (colorCode == "")
                        {
                            string fontTag = string.Format("<font size=\"{0}\">", fontSize);
                            node.ParentNode.ReplaceChild(doc.CreateTextNode(fontTag + (isBold ? "<b>" : "") + (isUnderline ? "<u>" : "") + (isThrough ? "<s>" : "") + node.InnerHtml + (isThrough ? "</s>" : "") + (isUnderline ? "</u>" : "") + (isBold ? "</b>" : "") + "</font>"), node);
                        }
                        else
                        {
                            string fontTag = string.Format("<font size=\"{0}\" ", fontSize);
                            node.ParentNode.ReplaceChild(doc.CreateTextNode(fontTag + "color=\"" + colorCode + "\">" + (isBold ? "<b>" : "") + (isUnderline ? "<u>" : "") + (isThrough ? "<s>" : "") + node.InnerHtml + (isThrough ? "</s>" : "") + (isUnderline ? "</u>" : "") + (isBold ? "</b>" : "") + "</font>"), node);
                        }

                    }
                }
            }
            //p
            if (doc.DocumentNode.SelectNodes("//p") != null)
            {
                foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//p"))
                {
                    int fontSize = 6;
                    if (node.HasAttributes)
                    {
                        string attribute = node.GetAttributeValue("style", "");
                        if (attribute.IndexOf("font-size:10") >= 0)
                        {
                            fontSize = 1;
                        }
                        if (attribute.IndexOf("font-size:11") >= 0)
                        {
                            fontSize = 2;
                        }
                        if (attribute.IndexOf("font-size:18") >= 0)
                        {
                            fontSize = 3;
                        }
                        if (attribute.IndexOf("font-size:24") >= 0)
                        {
                            fontSize = 4;
                        }
                        if (attribute.IndexOf("font-size:30") >= 0)
                        {
                            fontSize = 5;
                        }
                    }
                    if (fontSize == 6)
                    {
                        node.ParentNode.ReplaceChild(doc.CreateTextNode(node.InnerHtml + "\r\n"), node);
                    }
                    else
                    {
                        string fontTag = string.Format("<font size=\"{0}\">", fontSize);
                        node.ParentNode.ReplaceChild(doc.CreateTextNode(fontTag + node.InnerHtml + "</font>" + "\r\n"), node);
                    }
                }
            }
            //div
            if (doc.DocumentNode.SelectNodes("//div") != null)
            {
                foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//div"))
                {
                    node.ParentNode.ReplaceChild(doc.CreateTextNode(node.InnerHtml), node);
                }
            }
            
            string codeText = doc.DocumentNode.OuterHtml;
            return codeText;
        }

        private void New_Click(object sender, RoutedEventArgs e)
        {
            TextRange txt = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            txt.Text = "";
            ToHtml_Click(sender, e);
        }

        private string GetRTF()
        {
            TextRange range = new TextRange(rtb.Document.ContentStart,
                rtb.Document.ContentEnd);

            // Exception abfangen für StreamReader und MemoryStream
            try
            {
                using (MemoryStream rtfMemoryStream = new MemoryStream())
                {
                    using (StreamWriter rtfStreamWriter =
                                             new StreamWriter(rtfMemoryStream))
                    {
                        range.Save(rtfMemoryStream, System.Windows.DataFormats.Rtf);

                        rtfMemoryStream.Flush();
                        rtfMemoryStream.Position = 0;
                        StreamReader sr = new StreamReader(rtfMemoryStream);
                        return sr.ReadToEnd();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void SetRTF(string rtf)
        {
            TextRange range = new TextRange(rtb.Document.ContentStart,
                rtb.Document.ContentEnd);

            try
            {
                using (MemoryStream rtfMemoryStream = new MemoryStream())
                {
                    using (StreamWriter rtfStreamWriter =
                                                   new StreamWriter(rtfMemoryStream))
                    {
                        rtfStreamWriter.Write(rtf);
                        rtfStreamWriter.Flush();
                        rtfMemoryStream.Seek(0, SeekOrigin.Begin);

                        range.Load(rtfMemoryStream, System.Windows.DataFormats.Rtf);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void Open_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
            ofd.Multiselect = false;
            ofd.Filter = "Saved Files|*.rtf|All Files|*.*";

            Nullable<bool> result = ofd.ShowDialog();
            if (result == true)
            {
                SetRTF(File.ReadAllText(ofd.FileName));
                ToHtml_Click(sender, e);            
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog sfd = new Microsoft.Win32.SaveFileDialog();
            sfd.DefaultExt = ".rtf";
            sfd.Filter = "Saved Files|*.rtf|All Files|*.*";

            Nullable<bool> result = sfd.ShowDialog();
            if (result == true)
            {
                string rtf = GetRTF();
                using (Stream fs = sfd.OpenFile())
                {
                    System.Text.UTF8Encoding enc = new System.Text.UTF8Encoding();
                    byte[] buffer = enc.GetBytes(rtf);
                    fs.Write(buffer, 0, buffer.Length);
                    fs.Close();
                }

            }
        }

        private void ShowBR_Checked(object sender, RoutedEventArgs e)
        {
            htmlText.Text = htmlText.Text.Replace("\r\n", "<br>\r\n");
        }

        private void ShowBR_Unchecked(object sender, RoutedEventArgs e)
        {
            htmlText.Text = htmlText.Text.Replace("<br>\r\n", "\r\n");
        } 
    }
}
