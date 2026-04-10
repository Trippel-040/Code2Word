using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Interop;
using System.Windows.Media.Animation;

namespace Code2Word
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        private Brush selectedBorderBrush = Brushes.CornflowerBlue;

        public MainWindow()
        {
            InitializeComponent();

            SourceInitialized += (s, e) =>
            {
                ApplyTheme(chkDarkMode.IsChecked == true);
            };
        }

        private void ColorSelected_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn)
            {
                selectedBorderBrush = btn.Background;
                BtnConvert_Click(null, null); // 重新生成以应用新颜色
            }
        }

        private void CustomColor_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var color = (Color)ColorConverter.ConvertFromString(txtCustomColor.Text);
                selectedBorderBrush = new SolidColorBrush(color);
                BtnConvert_Click(null, null); // 重新生成以应用新颜色
            }
            catch (FormatException)
            {
                MessageBox.Show("请输入有效的16进制颜色代码。例如：#FF0000", "无效的颜色代码", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ChkDarkMode_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded) return;
            ApplyTheme(chkDarkMode.IsChecked == true);
        }

        private void ApplyTheme(bool isDark)
        {
            var bgDark = (SolidColorBrush)new BrushConverter().ConvertFrom("#1E1E1E");
            var txtBgDark = (SolidColorBrush)new BrushConverter().ConvertFrom("#232122");
            var txtFgDark = (SolidColorBrush)new BrushConverter().ConvertFrom("#D4D4D4");

            Background = isDark ? bgDark : Brushes.White;
            Foreground = isDark ? Brushes.White : Brushes.Black;

            chkDarkMode.Foreground = isDark ? Brushes.White : Brushes.Black;
            chkShowLineNumbers.Foreground = isDark ? Brushes.White : Brushes.Black;
            chkRemoveEmptyLines.Foreground = isDark ? Brushes.White : Brushes.Black;
            chkAddShading.Foreground = isDark ? Brushes.White : Brushes.Black;
            txtBorderColor.Foreground = isDark ? Brushes.White : Brushes.Black;

            txtInput.Background = isDark ? txtBgDark : Brushes.White;
            txtInput.Foreground = isDark ? txtFgDark : Brushes.Black;

            rtbOutput.Background = isDark ? txtBgDark : Brushes.White;
            rtbOutput.Foreground = isDark ? txtFgDark : Brushes.Black;

            var interop = new WindowInteropHelper(this);
            if (interop.Handle != IntPtr.Zero)
            {
                int darkVal = isDark ? 1 : 0;
                DwmSetWindowAttribute(interop.Handle, 20, ref darkVal, sizeof(int));
                DwmSetWindowAttribute(interop.Handle, 19, ref darkVal, sizeof(int));
            }
        }

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            string input = txtInput.Text;
            bool removeEmpty = chkRemoveEmptyLines.IsChecked == true;
            bool showLineNums = chkShowLineNumbers.IsChecked == true;
            bool addShading = chkAddShading.IsChecked == true;

            string[] lines = input.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            if (removeEmpty)
            {
                lines = lines.Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();
            }

            FlowDocument doc = new FlowDocument();
            doc.PagePadding = new Thickness(72);
            doc.FontFamily = new FontFamily("Consolas");


            List rootList = new List()
            {
                MarkerStyle = TextMarkerStyle.None,
                Margin = new Thickness(0),
                Padding = new Thickness(0)
            };
            ListItem dummyItem = new ListItem()
            {
                Margin = new Thickness(0),
                Padding = new Thickness(0)
            };
            rootList.ListItems.Add(dummyItem);

            List list = new List();
            // 设置多级序号（内嵌于 rootList 形成多级结构）
            list.MarkerStyle = showLineNums ? TextMarkerStyle.Decimal : TextMarkerStyle.None;
            list.StartIndex = 1;
            dummyItem.Blocks.Add(list);

            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];

                string cleanLine = line.Replace("\t", "    ");

                // 仅将行首的普通空格替换为不间断空格 (\u00A0)，防止复制进 Word 时前导空格(缩进)被自动修剪。
                // 保留代码中间的正常空格，从而不影响 Word 的自动换行功能。
                int leadingSpaceCount = cleanLine.TakeWhile(c => c == ' ').Count();
                string text = new string('\u00A0', leadingSpaceCount) + cleanLine.Substring(leadingSpaceCount);

                if (string.IsNullOrEmpty(text)) 
                {
                    text = "\u00A0";
                }

                // 将竖向彩条添加到段落边框上
                Paragraph p = new Paragraph()
                {
                    // 左右Margin原本为NaN。为了解决Word中色条包围行号（导致行号跑到色条右侧）的问题，
                    // 强制指定 LeftMargin = 0，从而将段落边框锚定在行号和文本之间。
                    Margin = new Thickness(0, 0, 0, 0),
                    Padding = new Thickness(10, 0, 0, 0),
                    BorderThickness = new Thickness(3, 0, 0, 0),
                    BorderBrush = selectedBorderBrush // 自定义左侧竖线条颜色
                };

                if (addShading && i % 2 == 1)
                {
                    p.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f4f4f4"));
                }

                ParseAndColorizeCode(text, p);

                ListItem li = new ListItem(p)
                {
                    Foreground = Brushes.Gray // 让前面的编号变成灰色，内容在里面覆盖层
                };

                p.Foreground = Brushes.Black;

                list.ListItems.Add(li);
            }

            doc.Blocks.Add(rootList);
            rtbOutput.Document = doc;
        }

        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            rtbOutput.SelectAll();
            rtbOutput.Copy();

            DoubleAnimation animation = new DoubleAnimation
            {
                From = 1.0,
                To = 0.0,
                Duration = new Duration(TimeSpan.FromSeconds(2)),
                BeginTime = TimeSpan.FromSeconds(0.2) // Stay visible for 1 second before fading
            };
            toastNotification.BeginAnimation(UIElement.OpacityProperty, animation);
        }

        private void ParseAndColorizeCode(string text, Paragraph p)
        {
            // 匹配注释、字符串、关键字、类名(包括大写字母开头)、方法名
            string keywords = @"(?<!@)\b(abstract|as|base|bool|break|byte|case|catch|char|checked|class|const|continue|decimal|default|delegate|do|double|else|enum|event|explicit|extern|false|finally|fixed|float|for|foreach|goto|if|implicit|in|int|interface|internal|is|lock|long|namespace|new|null|object|operator|out|override|params|private|protected|public|readonly|ref|return|sbyte|sealed|short|sizeof|stackalloc|static|string|struct|switch|this|throw|true|try|typeof|uint|ulong|unchecked|unsafe|ushort|using|virtual|void|volatile|while|add|alias|ascending|async|await|by|descending|dynamic|equals|from|get|global|group|into|join|let|nameof|on|orderby|partial|remove|select|set|unmanaged|value|var|when|where|yield)\b";
            string pattern = $@"(?<Comment>//.*?$)|(?<String>@?""(\\""|""""|[^""])*"")|(?<Keyword>{keywords})|(?<Method>(?<!@)\b[a-zA-Z_]\w*(?=\s*\())|(?<Class>(?<!@)\b[A-Z][A-Za-z0-9_]*\b)";

            int lastIndex = 0;
            Regex regex = new Regex(pattern);
            foreach (Match match in regex.Matches(text))
            {
                if (match.Index > lastIndex)
                {
                    p.Inlines.Add(new Run(text.Substring(lastIndex, match.Index - lastIndex)) { Foreground = Brushes.Black });
                }

                Run run = new Run(match.Value);
                if (match.Groups["Comment"].Success)
                {
                    run.Foreground = Brushes.Green;
                }
                else if (match.Groups["String"].Success)
                {
                    run.Foreground = Brushes.Brown; // Rust equivalent
                }
                else if (match.Groups["Keyword"].Success)
                {
                    run.Foreground = Brushes.Blue;
                }
                else if (match.Groups["Method"].Success)
                {
                    run.Foreground = Brushes.DarkGoldenrod;
                }
                else if (match.Groups["Class"].Success)
                {
                    run.Foreground = Brushes.DarkCyan;
                }

                p.Inlines.Add(run);
                lastIndex = match.Index + match.Length;
            }

            if (lastIndex < text.Length)
            {
                p.Inlines.Add(new Run(text.Substring(lastIndex)) { Foreground = Brushes.Black });
            }
        }
    }
}