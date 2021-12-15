using System;
using System.Linq;
using System.Windows;
using System.IO;
using System.Diagnostics;

namespace LecteurEMLFiles
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string MyDirectory = @"J:\Projets\C#\APPS\LecteurEMLFiles\sources\";
        public class DataObject
        {
            public string A { get; set; }
            public string B { get; set; }
            public string ColorA { get; set; }
            public string ColorB { get; set; }
            public string ForegroundA { get; set; }
            public string ForegroundB { get; set; }
        }

        public void InitializeContent(CDO.Message msg)
        {
            string colorHead = "white";
            string backgroundHead = "gray";
            string colorData = "black";
            string backgroundData = "white";

            if (msg == null)
            {
                var list = new System.Collections.ObjectModel.ObservableCollection<DataObject>();
                list.Add(new DataObject() { A = "Headers", B = "Content", ColorA = backgroundHead, ColorB = backgroundHead, ForegroundA = colorHead, ForegroundB = colorHead});
                list.Add(new DataObject() { A = "From", B = "", ColorA = backgroundData, ColorB = backgroundData, ForegroundA = colorData, ForegroundB = colorData });
                list.Add(new DataObject() { A = "To", B = "", ColorA = backgroundData, ColorB = backgroundData, ForegroundA = colorData, ForegroundB = colorData });
                list.Add(new DataObject() { A = "Date", B = "" , ColorA = backgroundData, ColorB = backgroundData, ForegroundA = colorData, ForegroundB = colorData });
                list.Add(new DataObject() { A = "Subject", B = "" , ColorA = backgroundData, ColorB = backgroundData, ForegroundA = colorData, ForegroundB = colorData });
                list.Add(new DataObject() { A = "Message", B = "", ColorA = backgroundData, ColorB = backgroundData, ForegroundA = colorData, ForegroundB = colorData });
                this.Headers.ItemsSource = list;
            }
            else
            {
                var list = new System.Collections.ObjectModel.ObservableCollection<DataObject>();
                list.Add(new DataObject() { A = "Headers", B = "Content", ColorA = backgroundHead, ColorB = backgroundHead, ForegroundA = colorHead, ForegroundB = colorHead });
                list.Add(new DataObject() { A = "From", B = msg.From, ColorA = backgroundData, ColorB = backgroundData, ForegroundA = colorData, ForegroundB = colorData });
                list.Add(new DataObject() { A = "To", B = msg.To, ColorA = backgroundData, ColorB = backgroundData, ForegroundA = colorData, ForegroundB = colorData });
                list.Add(new DataObject() { A = "Date", B = msg.ReceivedTime.ToString(), ColorA = backgroundData, ColorB = backgroundData, ForegroundA = colorData, ForegroundB = colorData });
                list.Add(new DataObject() { A = "Subject", B = msg.Subject, ColorA = backgroundData, ColorB = backgroundData, ForegroundA = colorData, ForegroundB = colorData });
                list.Add(new DataObject() { A = "Message", B = msg.TextBody, ColorA = backgroundData, ColorB = backgroundData, ForegroundA = colorData, ForegroundB = colorData });
                
                if(msg.Attachments.Count > 1)
                {
                    for (int i = 1; i < msg.Attachments.Count + 1; i++)
                    {
                        Debug.WriteLine(msg.Attachments[i].ContentMediaType);
                        Debug.WriteLine(msg.Attachments[i].GetStream().Size);
                        Debug.WriteLine(msg.Attachments[i].GetDecodedContentStream().Size);
                        list.Add(new DataObject() { A = "Attachment " + i, B = msg.Attachments[i].FileName + " (" + msg.Attachments[i].ContentMediaType + ")" + " - " + (msg.Attachments[i].GetDecodedContentStream().Size / 1000) + "Ko", ColorA = backgroundData, ColorB = backgroundData, ForegroundA = colorData, ForegroundB = colorData });
                    }
                }
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
            if (File.Exists(path))
            {
                InitializeContent(LoadEmlFromFile(path));
            }
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
