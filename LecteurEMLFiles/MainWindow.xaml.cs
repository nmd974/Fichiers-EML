using System;
using System.Linq;
using System.Windows;
using System.IO;

namespace LecteurEMLFiles
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string MyDirectory = @"__YOUR_DIRECTORY__\sources\";
        public class DataObject
        {
            public string A { get; set; }
            public string B { get; set; }
        }

        public void InitializeContent(CDO.Message msg)
        {
            if (msg == null)
            {
                var list = new System.Collections.ObjectModel.ObservableCollection<DataObject>();
                list.Add(new DataObject() { A = "Headers", B = "Content" });
                list.Add(new DataObject() { A = "From", B = "" });
                list.Add(new DataObject() { A = "To", B = "" });
                list.Add(new DataObject() { A = "Date", B = "" });
                list.Add(new DataObject() { A = "Subject", B = "" });
                list.Add(new DataObject() { A = "Message", B = "" });
                this.Headers.ItemsSource = list;
            }
            else
            {
                var list = new System.Collections.ObjectModel.ObservableCollection<DataObject>();
                list.Add(new DataObject() { A = "Headers", B = "Content" });
                list.Add(new DataObject() { A = "From", B = msg.From });
                list.Add(new DataObject() { A = "To", B = msg.To });
                list.Add(new DataObject() { A = "Date", B = msg.ReceivedTime.ToString() });
                list.Add(new DataObject() { A = "Subject", B = msg.Subject });
                list.Add(new DataObject() { A = "Message", B = msg.TextBody });
                this.Headers.ItemsSource = list;
            }
        }
        public MainWindow()
        {
            InitializeComponent();
            CDO.Message msg = null;
            InitializeContent(msg);

            //Taille type du message et pièces jointes

            if (Directory.Exists(this.MyDirectory))
            {
                string[] files = Directory.GetFiles(this.MyDirectory, "*.eml").Select(Path.GetFileName).ToArray();

                foreach (string file in files)
                {
                    this.File_list.Items.Add(file.Split(".")[0]);
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string path = this.MyDirectory + this.File_list.SelectedItem + ".eml";
            InitializeContent(LoadEmlFromFile(path));
        }

        public CDO.Message LoadEmlFromFile(String emlFileName)
        {
            CDO.Message msg = new CDO.Message();
            ADODB.Stream stream = new ADODB.Stream();

            stream.Open(Type.Missing, ADODB.ConnectModeEnum.adModeUnknown, ADODB.StreamOpenOptionsEnum.adOpenStreamUnspecified, String.Empty, String.Empty);
            stream.LoadFromFile(emlFileName);
            stream.Flush();
            msg.DataSource.OpenObject(stream, "_Stream");
            msg.DataSource.Save();

            stream.Close();
            return msg;
        }
    }
}
