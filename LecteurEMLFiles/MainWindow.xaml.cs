using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.IO;
using System.Diagnostics;

using ADODB;
using System.Net.Mail;
using System.Reflection;

namespace LecteurEMLFiles
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
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
            string path = @"__YOUR_DIRECTORY__\sources";
            if (Directory.Exists(path))
            {
                string[] files = Directory.GetFiles(path, "*.eml")
                            .Select(Path.GetFileName)
                            .ToArray();

                foreach (string file in files)
                {
                    this.File_list.Items.Add(file.Split(".")[0]);
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string path = @"J:\Projets\C#\APPS\LecteurEMLFiles\sources\" + this.File_list.SelectedItem + ".eml";

            //if (File.Exists(path))
            //{

            //    Debug.WriteLine("Click"); 
            //    using (FileStream fs = File.OpenRead(path))
            //    {
            //        byte[] b = new byte[1024];
            //        UTF8Encoding temp = new UTF8Encoding(true);
            //        while (fs.Read(b, 0, b.Length) > 0)
            //        {
            //            Debug.WriteLine(temp.GetString(b));
            //        }
            //    }
            //}
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
