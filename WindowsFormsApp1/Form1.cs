using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json.Linq;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Text;
using System.Linq;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1() => InitializeComponent();
        static string elastic(string dosya)
        {
            var httpWebRequest = (HttpWebRequest)WebRequest.Create("http://localhost:9200/my_index6/_analyze");
            httpWebRequest.ContentType = "application/json; charset=UTF-8";
            httpWebRequest.Method = "POST";

            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                string json = "{\"analyzer\": \"my_analyzer\",\"text\": \"" + dosya + "\"}";
                streamWriter.Write(json);
            }

            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                return streamReader.ReadToEnd();
            }
        }
        public string wordAc(string uzantı)
        {
            string word = "";
            try
            {
                Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                Document document;
                document = application.Documents.Open(uzantı);
                word = document.Range().Text;
                word = HttpUtility.JavaScriptStringEncode(word);
                document.Close();
                //Console.WriteLine(word);
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Console.WriteLine(e.Message);
                MessageBox.Show("DOSYA OKUNAMADI! \n" + uzantı + "\nHATA! : " + e.Message);
            }
            return word;
        }
        public string pdfAc(string uzantı)
        {
            string js = "";
            try
            {
                using (PdfReader reader = new PdfReader(uzantı))
                {
                    StringBuilder textBuilder = new StringBuilder();

                    for (int j = 1; j <= reader.NumberOfPages; j++)
                    {
                        textBuilder.Append(PdfTextExtractor.GetTextFromPage(reader, j));
                    }
                    js = HttpUtility.JavaScriptStringEncode(textBuilder.ToString());
                }
            }
            catch (iTextSharp.text.exceptions.InvalidPdfException e)
            {
                Console.WriteLine(e.Message);
                MessageBox.Show("DOSYA OKUNAMADI! \n" + uzantı + "\nHATA! : " + e.Message);
            }
            return js;
        }
        public string sonucKontrol(string txt, string dosya, string uzantı)
        {
            var jsondata = JObject.Parse(elastic(dosya));
            List<OzelVeri> verilist = new List<OzelVeri>();
            foreach (var data in jsondata["tokens"])
            {
                OzelVeri ov = new OzelVeri();
                ov.Token = data["token"].ToString();
                ov.Start_offset = (int)data["start_offset"];
                ov.End_offset = (int)data["end_offset"];
                ov.Position = (int)data["position"];
                verilist.Add(ov);
            }
            Regex email = new Regex(@"[a-z0-9][a-z0-9._]*@[a-z0-9][a-z0-9]*\.[a-z0-9]+\.?[a-z0-9]*");
            Regex telefon = new Regex(@"(0?5[0-9]{9})|(0 ?[0-9]{3} [0-9]{3} [0-9]{4})");
            Regex tckn = new Regex(@"[1-9][0-9]{10}");
            Regex tarih = new Regex(@"([0-2][0-9]|3[0-1])[/.](0[0-9]|1[0-2])[/.]([0-9]{4})");
            txt = "\n" + uzantı + "\n";
            txt += verilist.Count + " tane özel veri bulundu\n\n";
            foreach (OzelVeri veri in verilist)
            {
                if (email.IsMatch(veri.Token))
                    txt += veri.Token + " email olabilir\n";
                if (telefon.IsMatch(veri.Token))
                    txt += veri.Token + " telefon olabilir\n";
                if (tckn.IsMatch(veri.Token))
                    txt += veri.Token + " tckn olabilir\n";
                if (tarih.IsMatch(veri.Token))
                    txt += veri.Token + " tarih olabilir\n";
            }
            txt += "----------------------------------------\n";
            return txt;
        }
        public void raporYaz(string txt)
        {
            //string dosya_yolu = @"D:\\siber_güvenlik_kursu\\özelveriler.txt";
            //FileStream fs = new FileStream(dosya_yolu, FileMode.OpenOrCreate, FileAccess.Write);
            //StreamWriter sw = new StreamWriter(fs);
            //sw.WriteLine(txt);
            //sw.Close();
            //fs.Close();
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "Metin Dosyası|*.txt";
            if (save.ShowDialog() == DialogResult.OK)
            {
                StreamWriter sw = new StreamWriter(save.FileName);
                sw.WriteLine(txt);
                sw.Close();
            }
        }
        private void AçToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ac = new OpenFileDialog();
            ac.Filter = "Word Documents|*.docx|Word Documents|*.doc|PDF Documents|*.pdf|All Files|*.*";
            ac.Multiselect = true;
            string txt = "", dosya = "";
            if (ac.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            for (int i = 0; i < ac.FileNames.Length; i++)
            {
                if (ac.SafeFileNames[i].EndsWith(".doc") | ac.SafeFileNames[i].EndsWith(".docx"))
                {
                    dosya = wordAc(ac.FileNames[i]);
                }
                if (ac.SafeFileNames[i].EndsWith(".pdf"))
                {
                    dosya = pdfAc(ac.FileNames[i]);
                }
                txt += sonucKontrol(txt, dosya, ac.FileNames[i]);
            }
            raporYaz(txt);
        }

        List<String> dosyaList = new List<string>();
        public List<string> WalkDirectoryTree(DirectoryInfo root)
        {
            FileInfo[] files = null;
            DirectoryInfo[] subDirs = null;
            try
            {
                files = root.GetFiles("*.*")
                    .Where(s => s.ToString().EndsWith(".doc") || s.ToString().EndsWith(".docx")|| s.ToString().EndsWith(".pdf"))
                    .ToArray();
                //files = root.GetFiles("*.pdf");
            }
            catch (UnauthorizedAccessException e)
            {
                Console.WriteLine(e.Message);
            }
            catch (DirectoryNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }
            if (files != null)
            {
                foreach (FileInfo fi in files)
                {
                    dosyaList.Add(fi.FullName);
                }
                subDirs = root.GetDirectories();
                foreach (DirectoryInfo dirInfo in subDirs)
                {
                    WalkDirectoryTree(dirInfo);
                }
            }
            return dosyaList;
        }
        private void Aç2ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {

            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            DirectoryInfo info = new DirectoryInfo(fbd.SelectedPath);
            string[] dosyalar = WalkDirectoryTree(info).ToArray();
            dosyaList.Clear();
            //BURADA DOSYA TARANMASI BİTTİ YAZDIR
            //MessageBox.Show("Toplam "+dosyalar.Length+" Tane dosya bulundu. Dosyalar taranıyor");
            //label2.Text = "Toplam " + dosyalar.Length + " Tane dosya bulundu. Dosyalar taranıyor";
            //string dosya = "", txt = "Toplam " + dosyalar.Length + " Tane dosya bulundu.\n";
            string dosya = "", txt = "";
            for (int i = 0; i < dosyalar.Length; i++)
            {
                //MessageBox.Show(i + "/" + dosyalar.Length + "Tane dosya tarandı\n"+dosyalar[i]+" taranıyor");
               // label1.Text = i + "/" + dosyalar.Length + " Tane dosya tarandı";
               // label2.Text = dosyalar[i] + " taranıyor";
                if (dosyalar[i].EndsWith(".doc") | dosyalar[i].EndsWith(".docx"))
                {
                    dosya = wordAc(dosyalar[i]);
                }
                if (dosyalar[i].EndsWith(".pdf"))
                {
                    dosya = pdfAc(dosyalar[i]);
                }
                txt += sonucKontrol(txt, dosya, dosyalar[i]);
                //BURADA i. DOSYA TARANDI YAZDIR
                //progressBar1.Value += 100 / dosyalar.Length;
            }
            //progressBar1.Value = 100;
            //label1.Text = dosyalar.Length + "/" + dosyalar.Length + " Tane dosya tarandı";
            //label2.Text = "Tarama işlemi bitti";
            raporYaz(txt);
        }
    }
}