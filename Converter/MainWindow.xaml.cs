using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace Converter
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string[] fileNames;
        string font = "PT Sans";
        string[] oldSymbols = { "8", "*", "щ", "Щ", "ъ", "Ъ", "=", "+", "5", "%", "й", "Й", "ю", "Ю", "ё", "Ё", "7", "?", "6", ":", "3", "№", "э", "Э", "я", "Я", "0", ")", "1", "2", "4", "9", };
        string[] newSymbols = { "ԥ", "Ԥ", "ҳ", "Ҳ", "ә", "Ә", "ҿ", "Ҿ", "џ", "Џ", "ҟ", "Ҟ", "ҩ", "Ҩ", "ӡ", "Ӡ", "ҵ", "Ҵ", "қ", "Қ", "ҷ", "Ҷ", "ҽ", "Ҽ", "ӷ", "Ӷ", "ҭ", "Ҭ", "?", "№", ":", ")" };
        string[] oldSymbolsArialABS = { "ё", "Ё", "1", "2", "3", "№", "4", ";", "5", "%", "6", ":", "7", "?", "8", "*", "9", "(", "й", "Й", "ц", "Ц", "у", "У", "к", "К", "е", "Е", "г", "Г", "ш", "Ш", "щ", "Щ", "з", "З", "х", "Х", "ъ", "Ъ", "ф", "Ф", "ы", "Ы", "в", "В", "а", "А", "п", "П", "р", "Р", "о", "О", "л", "Л", "д", "Д", "ж", "Ж", "э", "Э", "я", "Я", "ч", "Ч", "м", "М", "ю", "Ю" };
        string[] newSymbolsArialABS = { "џ", "Џ", "?", "џ", "ж", "Ж", "ҵ", "Ҵ", "ҟ", "Ҟ", "ҳ", "Ҳ", "ҩ", "Ҩ", "в", "В", "ф", "Ф", "ӡ", "Ӡ", "ӷ", "Ӷ", "ц", "Ц", "қ", "Қ", "х", "Х", "з", "З", "л", "Л", "о", "О", "ԥ", "Ԥ", "ҿ", "Ҿ", "ҷ", "Ҷ", "г", "Г", "ш", "Ш", "к", "К", "ы", "Ы", "р", "Р", "а", "А", "ә", "Ә", "у", "У", "п", "П", "ҽ", "Ҽ", "ч", "Ч", "д", "Д", "ҭ", "Ҭ", "е", "Е", "м", "М" };

        public MainWindow()
        {
            InitializeComponent();
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // ... Get the ComboBox.
            var comboBox = sender as ComboBox;

            // ... Set SelectedItem as Window Title.
            string value = (comboBox.SelectedItem as ComboBoxItem).Content.ToString();

            switch (value)
            {
                case "PT Sans":
                    font = "PT Sans";
                    break;
                case "PT Serif":
                    font = "PT Serif";
                    break;
            }

        }
        private async void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            //dlg.DefaultExt = ".doc";
            dlg.Multiselect = true;
            dlg.Filter = "Word documents |*.doc*";

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                progressBar.IsIndeterminate = true;

                if (dlg.FileNames.Length > 0)
                {
                    BrowseButton.IsEnabled = false;
                    SelectedFileTextBox.Text = "";
                    textBlock.Text = "Идет конвертация файлов";

                    string path = System.IO.Path.GetDirectoryName(dlg.FileName) + "\\UnicodeFiles";
                    if (Directory.Exists(path))
                    {
                        Console.WriteLine("That path exists already.");

                    }
                    else Directory.CreateDirectory(path);


                    fileNames = dlg.FileNames;


                    var WordApp = new Word.Application();
                    WordApp.Visible = false;

                    for (int j = 0; j <= fileNames.Length - 1; j++)
                    {

                        var WordDoc = WordApp.Documents.Open(fileNames[j]);
                        object fileName = path + "\\(Unicode)" + WordDoc.Name;
                        Object missing = Type.Missing;
                        WordApp.ActiveDocument.SaveAs(ref fileName,
        ref missing, ref missing, ref missing, ref missing, ref missing,
        ref missing, ref missing, ref missing, ref missing, ref missing,
        ref missing, ref missing, ref missing, ref missing, ref missing);
                        WordDoc = WordApp.Documents.Open(fileName);
                        WordApp.Visible = false;


                        var range = WordDoc.Range();
                        // var range1 = WordDoc.Range();
                        //var range2 = WordDoc.Range();

                        //range1.Find.Execute(FindText: "ҧ", Replace: WdReplace.wdReplaceAll, ReplaceWith: "ԥ");                   
                        //range2.Find.Execute(FindText: "ҕ", Replace: WdReplace.wdReplaceAll, ReplaceWith: "ӷ");

                        var rangeTimes = WordDoc.Range();
                        var rangeArialABS = WordDoc.Range();
                        var footnotes = WordDoc.Footnotes;
                        range.Find.Font.Name = "Arial Abkh";
                        rangeTimes.Find.Font.Name = "Times New Roman Abkh";
                        rangeArialABS.Find.Font.Name = "ArialABS";

                        if (footnotes.Count > 0)
                        {
                            foreach (Word.Footnote fn in footnotes)
                            {
                                var rangeFootnote = fn.Range;
                                rangeFootnote.Find.Font.Name = "Arial Abkh";
                                var rangeFootnoteTimes = fn.Range;
                                rangeFootnoteTimes.Find.Font.Name = "Times New Roman Abkh";
                                for (int i = 0; i <= oldSymbols.Length - 1; i++)
                                {
                                    await Task.Run(() =>
                                    {

                                        rangeFootnote.Find.Execute(FindText: oldSymbols[i], Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: newSymbols[i]);
                                        rangeFootnoteTimes.Find.Execute(FindText: oldSymbols[i], Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: newSymbols[i]);

                                    });
                                }
                            }
                        }


                        for (int i = 0; i <= oldSymbols.Length - 1; i++)
                        {
                            await Task.Run(() =>
                            {

                                range.Find.Execute(FindText: oldSymbols[i], Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: newSymbols[i]);
                                rangeTimes.Find.Execute(FindText: oldSymbols[i], Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: newSymbols[i]);

                            });
                        }


                        for (int i = 0; i <= oldSymbolsArialABS.Length - 1; i++)
                        {
                            await Task.Run(() =>
                            {

                                rangeArialABS.Find.Execute(FindText: oldSymbolsArialABS[i], Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: newSymbolsArialABS[i]);

                            });
                        }
                        range.Font.Name = font;
                        rangeTimes.Font.Name = font;


                        WordApp.ActiveDocument.Save();
                        WordApp.ActiveDocument.Close();
                        SelectedFileTextBox.Text += fileName + "\n";
                    }

                    WordApp.Quit();
                    BrowseButton.IsEnabled = true;
                    textBlock.Text = "Выберите один или более файлов .doc или .docx\nУ вас должен быть установлен Microsoft Word";
                    progressBar.IsIndeterminate = false;
                    System.Diagnostics.Process.Start("explorer.exe", path);
                }
            }
        }
    }
}
