using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using Spire.Pdf.Lists;
using Spire.Pdf.Exporting.XPS.Schema;
using iTextSharp.text;
using iTextSharp.text.pdf;
using PdfDocument = iTextSharp.text.pdf.PdfDocument;
//using Aspose.Pdf;
using System.Runtime.InteropServices;
using Color = System.Drawing.Color;
///using Bytescout.PDFExtractor;
using SautinSoft;
using Rectangle = iTextSharp.text.Rectangle;
using Path = System.IO.Path;

namespace elfatek_3
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        static extern bool HideCaret(IntPtr hWnd);
        public Form1()
        {
            InitializeComponent();

            textBox1.GotFocus += (s1, s2) => { HideCaret(textBox1.Handle); };

        }
      
        System.Diagnostics.Process process = new System.Diagnostics.Process();
        string bat = "key.bat";
        string officekey,disk;
        double disk3, disk4, disk5, disk6, disk7;
        
        #region panelin harekti için gerekli kodlar
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr one, int two, int three, int four);
        #endregion
        private void Form1_Load(object sender, EventArgs e)
        {
            textBox2.Visible = false;
            textBox3.Visible = false;
            textBox4.Visible = false;
            label3.Text = "";
            panel1.BackColor = Color.FromArgb(0, 127, 127, 127);            
            panel1.BackColor = Color.FromArgb(0, 127, 127, 127);

            //ekran boyutu

            int ourScreenWidth = Screen.FromControl(this).WorkingArea.Width;
            int ourScreenHeight = Screen.FromControl(this).WorkingArea.Height;
            float scaleFactorWidth = (float)ourScreenWidth / 1600f;
            float scaleFactorHeigth = (float)ourScreenHeight / 900f;
            SizeF scaleFactor = new SizeF(scaleFactorWidth, scaleFactorHeigth);
            Scale(scaleFactor);
            CenterToScreen();

            //mevcut diskler çekildi
            foreach (var drive in System.IO.DriveInfo.GetDrives())
            {
                comboBox1.Items.Add(drive);
            }
        }
        private void execute()
        {
            process.StartInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + bat;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = true;
            process.Start();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        { //form kapandığında bat boşsa 1 yadırıyor???
            string readText = File.ReadAllText(bat);
            if (String.IsNullOrEmpty(readText))
            {
                string txt = "1";
                File.WriteAllText(bat, txt);

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (cBirim.SelectedItem!=null && txtAd.Text != "")
            {

                if (comboBox1.SelectedItem != null)
                {

                    execute();
                    StringBuilder sb = new StringBuilder();
                    progressBar1.Value = 20;
                    string fold = "c:";
                    sb.AppendLine(fold);
                    File.WriteAllText(bat, sb.ToString());

                    //komutlar gönderildi
                    string ramManufacturer = "wmic MemoryChip get manufacturer & wmic baseboard get product,Manufacturer  & wmic csproduct get identifyingnumber & wmic memorychip get devicelocator & wmic MEMORYCHIP get Capacity & wmic diskdrive get Model & wmic diskdrive get size & wmic diskdrive get size & wmic csproduct get name & wmic cpu get name & wmic path win32_VideoController get name & wmic path softwarelicensingservice get OA3xOriginalProductKey";
                    sb.AppendLine(ramManufacturer);
                    File.WriteAllText(bat, sb.ToString());
                    StreamReader r = process.StandardOutput;
                    string output1 = r.ReadToEnd();
                    string fix = Regex.Replace(output1, @"^\s*$\n", string.Empty, RegexOptions.Multiline);

                    int s = fix.LastIndexOf("get OA3xOriginalProductKey");
                    string res = fix.Remove(0, s + "get OA3xOriginalProductKey".Length);

                    //çıktı split edidi
                    string[] arr = res.Split(new string[] { "Manufacturer" }, StringSplitOptions.RemoveEmptyEntries);
                    string[] arr2 = res.Split(new string[] { "IdentifyingNumber" }, StringSplitOptions.RemoveEmptyEntries);
                    string[] arr3 = res.Split(new string[] { "DeviceLocator" }, StringSplitOptions.RemoveEmptyEntries);
                    string[] arr4 = res.Split(new string[] { "Capacity" }, StringSplitOptions.RemoveEmptyEntries);
                    string[] arr5 = res.Split(new string[] { "Model" }, StringSplitOptions.RemoveEmptyEntries);
                    string[] arr6 = res.Split(new string[] { "Size" }, StringSplitOptions.RemoveEmptyEntries);
                    string[] arr7 = res.Split(new string[] { "Name" }, StringSplitOptions.RemoveEmptyEntries);
                    string[] arr9 = res.Split(new string[] { "OA3xOriginalProductKey" }, StringSplitOptions.RemoveEmptyEntries);

                    string ramManu = arr[1];
                    textBox1.Text = "RAM MARKASI" + "\n" + ramManu;

                    int b1 = arr2[0].IndexOf("Product") + "Product".Length;
                    string mbModel = arr2[0].Substring(b1);
                    textBox1.Text += "ANAKART MODELİ" + "\n" + mbModel;


                    int b2 = arr3[0].IndexOf("IdentifyingNumber") + "IdentifyingNumber".Length;
                    string serial = arr3[0].Substring(b2);
                    textBox1.Text += "ANAKART SERİ NUMARASI" + "\n" + serial;


                    int b3 = arr4[0].IndexOf("DeviceLocator") + "DeviceLocator".Length;
                    string ram = arr4[0].Substring(b3);
                    textBox1.Text += "RAM TÜRÜ" + "\n" + ram;


                    int b4 = arr5[0].IndexOf("Capacity") + "Capacity".Length;
                    string ramCapacity = arr5[0].Substring(b4);
                    string fix3 = ramCapacity.Replace("\r\n", string.Empty);
                    textBox2.Text = fix3;

                    //ram byttan gba dönüştrüldü.
                    if (textBox2.Lines.Length == 3)
                    {

                        double ram3 = double.Parse(textBox2.Lines[1]);
                        ram3 = (ram3 / (1024 * 1024 * 1024));
                        textBox1.Text += "RAM KAPASİTESİ" + Environment.NewLine + ram3 + " GB" + Environment.NewLine;
                    }
                    else if (textBox2.Lines.Length == 4)
                    {

                        double ram3 = double.Parse(textBox2.Lines[1]);
                        ram3 = (ram3 / (1024 * 1024 * 1024));
                        double ram4 = double.Parse(textBox2.Lines[2]);
                        ram4 = (ram4 / (1024 * 1024 * 1024));
                        string total = (ram3 + ram4).ToString();
                        textBox1.Text += "RAM KAPASİTESİ" + Environment.NewLine + total + " GB" + Environment.NewLine;
                    }

                    string[] arr8 = textBox1.Text.Split(new string[] { "GB" }, StringSplitOptions.RemoveEmptyEntries);

                    int b10 = arr8[0].IndexOf("RAM KAPASİTESİ") + "RAM KAPASITESI".Length;
                    string rams = textBox1.Text.Substring(b10);


                    string system = arr7[1];
                    textBox1.Text += "SİSTEM MODELİ" + "\n" + system;
                    string cpu = arr7[2];
                    textBox1.Text += "İŞLEMCİ" + "\n" + cpu;

                    int last = arr7[3].LastIndexOf("OA3xOriginalProductKey");
                    string display = arr7[3].Substring(0, last);
                    textBox1.Text += "EKRAN KARTI" + "\n" + display;


                    string diskSize = arr6[1];
                    string fix4 = diskSize.Replace("\r\n", string.Empty);
                    textBox3.Text = fix4;

                    //dik kapasitesi dönüştürüldü 
                    if (textBox3.Lines.Length == 3)
                    {
                        disk3 = double.Parse(textBox3.Lines[1]);
                        disk3 = Math.Round(disk3 / (1024 * 1024 * 1024), 2);
                    }
                    else if (textBox3.Lines.Length == 4)
                    {

                        disk3 = double.Parse(textBox3.Lines[1]);
                        disk3 = Math.Round(disk3 / (1024 * 1024 * 1024), 2);
                        disk4 = double.Parse(textBox3.Lines[2]);
                        disk4 = Math.Round(disk4 / (1024 * 1024 * 1024), 2);
                        string total = (disk3 + disk4).ToString();
                    }
                    else if (textBox3.Lines.Length == 5)
                    {
                        disk3 = double.Parse(textBox3.Lines[1]);
                        disk3 = Math.Round(disk3 / (1024 * 1024 * 1024), 2);
                        disk4 = double.Parse(textBox3.Lines[2]);
                        disk4 = Math.Round(disk4 / (1024 * 1024 * 1024), 2);
                        disk5 = double.Parse(textBox3.Lines[3]);
                        disk5 = Math.Round(disk5 / (1024 * 1024 * 1024), 2);
                        string total = (disk3 + disk4 + disk5).ToString();
                    }
                    else if (textBox3.Lines.Length == 6)
                    {
                        disk3 = double.Parse(textBox3.Lines[1]);
                        disk3 = Math.Round(disk3 / (1024 * 1024 * 1024), 2);
                        disk4 = double.Parse(textBox3.Lines[2]);
                        disk4 = Math.Round(disk4 / (1024 * 1024 * 1024), 2);
                        disk5 = double.Parse(textBox3.Lines[3]);
                        disk5 = Math.Round(disk5 / (1024 * 1024 * 1024), 2);
                        disk6 = double.Parse(textBox3.Lines[4]);
                        disk6 = Math.Round(disk5 / (1024 * 1024 * 1024), 2);
                        string total = (disk3 + disk4 + disk5 + disk6).ToString();
                    }
                    else if (textBox3.Lines.Length > 6)
                    {
                        disk3 = double.Parse(textBox3.Lines[1]);
                        disk3 = Math.Round(disk3 / (1024 * 1024 * 1024), 2);
                        disk4 = double.Parse(textBox3.Lines[2]);
                        disk4 = Math.Round(disk4 / (1024 * 1024 * 1024), 2);
                        disk5 = double.Parse(textBox3.Lines[3]);
                        disk5 = Math.Round(disk5 / (1024 * 1024 * 1024), 2);
                        disk6 = double.Parse(textBox3.Lines[4]);
                        disk6 = Math.Round(disk6 / (1024 * 1024 * 1024), 2);
                        disk7 = double.Parse(textBox3.Lines[5]);
                        disk7 = Math.Round(disk7 / (1024 * 1024 * 1024), 2);
                        string total = (disk3 + disk4 + disk5 + disk6 + disk7).ToString();
                    }

                    int b5 = arr6[0].IndexOf("Model") + "Model".Length;
                    string diskModel = arr6[0].Substring(b5);
                    string fix5 = diskModel.Replace("\r\n", string.Empty);
                    textBox4.Text = fix5;
                    //her disk modelinin yanına disk kapasitesi yazdırıldı

                    if (textBox4.Lines.Length == 3)
                    {

                        disk = textBox4.Lines[1] + " " + disk3 + " GB";
                        textBox1.Text += "DİSK MODELİ" + Environment.NewLine + disk + Environment.NewLine;


                    }
                    else if (textBox4.Lines.Length == 4)
                    {
                        disk = textBox4.Lines[1] + " " + disk3 + " GB" + Environment.NewLine + textBox4.Lines[2] + " " + disk4 + " GB";
                        textBox1.Text += "DİSK MODELİ" + Environment.NewLine + disk + Environment.NewLine;
                    }
                    else if (textBox4.Lines.Length == 5)
                    {
                        disk = textBox4.Lines[1] + " " + disk3 + " GB" + Environment.NewLine + textBox4.Lines[2] + " " + disk4 + " GB" + Environment.NewLine + textBox4.Lines[3] + " " + disk5 + " GB";
                        textBox1.Text += "DİSK MODELİ" + Environment.NewLine + disk + Environment.NewLine;
                    }
                    else if (textBox4.Lines.Length == 6)
                    {
                        disk = textBox4.Lines[1] + " " + disk3 + " GB" + Environment.NewLine + textBox4.Lines[2] + " " + disk4 + " GB" + Environment.NewLine + textBox4.Lines[3] + " " + disk5 + " GB" + Environment.NewLine + textBox4.Lines[4] + " " + disk6 + " GB";
                        textBox1.Text += "DİSK MODELİ" + Environment.NewLine + disk + Environment.NewLine;
                    }
                    else if (textBox4.Lines.Length > 6)
                    {
                        disk = textBox4.Lines[1] + " " + disk3 + " GB" + Environment.NewLine + textBox4.Lines[2] + " " + disk4 + " GB" + Environment.NewLine + textBox4.Lines[3] + " " + disk5 + " GB" + Environment.NewLine + textBox4.Lines[4] + " " + disk6 + " GB" + Environment.NewLine + textBox4.Lines[5] + " " + disk7 + " GB";
                        textBox1.Text += "DİSK MODELİ" + Environment.NewLine + disk + Environment.NewLine;
                    }

                    string win = arr9[1];
                    textBox1.Text += "WINDOWS ÜRÜN ANAHTARI" + " \n" + win;
                     
                    //office sürümlerine göre office ürün anahtarını bulan kod gönderildi
                    progressBar1.Value = 40;
                    StringBuilder sb2 = new StringBuilder();
                    string office19_32 = "cd " + comboBox1.Text + "Program Files (x86)/Microsoft Office/Office16 & cscript ospp.vbs /dstatus";
                    File.WriteAllText(bat, office19_32);
                    StreamReader reader = process.StandardOutput;
                    string output2 = reader.ReadToEnd();
                    if (output2.Contains("PRODUCT ID:"))
                    {
                        execute();
                        sb2.AppendLine(office19_32);
                        File.WriteAllText(bat, sb2.ToString());
                        StreamReader reader2 = process.StandardOutput;
                        output2 = reader2.ReadToEnd();
                        int baslangic = output2.IndexOf("key:") + "key:".Length;
                        string office = output2.Substring(baslangic, 6);
                        textBox1.Text += "OFFİCE ÜRÜN ANAHTARI" + "\n" + office;

                    }
                    else
                    {
                        execute();
                        string office15_64 = "cd " + comboBox1.Text + "Program Files/Microsoft Office/Office15 & cscript ospp.vbs /dstatus ";
                        sb2.Clear();
                        sb2.AppendLine(office15_64);
                        File.WriteAllText(bat, sb2.ToString());
                        StreamReader reader3 = process.StandardOutput;
                        output2 = reader3.ReadToEnd();

                        if (output2.Contains("PRODUCT ID:"))
                        {
                            execute();
                            File.WriteAllText(bat, sb2.ToString());
                            StreamReader reader4 = process.StandardOutput;
                            output2 = reader4.ReadToEnd();
                            int baslangic = output2.IndexOf("key:") + "key:".Length;
                            string office = output2.Substring(baslangic, 6);
                            textBox1.Text += "OFFİCE ÜRÜN ANAHTARI" + Environment.NewLine + office;


                        }
                        else
                        {
                            execute();
                            string office15_32 = "cd " + comboBox1.Text + "Program Files (x86)/Microsoft Office/Office15 & cscript ospp.vbs /dstatus";
                            sb2.Clear();
                            sb2.AppendLine(office15_32);
                            File.WriteAllText(bat, sb2.ToString());
                            StreamReader reader5 = process.StandardOutput;
                            output2 = reader5.ReadToEnd();

                            if (output2.Contains("PRODUCT ID:"))
                            {
                                execute();
                                File.WriteAllText(bat, sb2.ToString());
                                StreamReader reader8 = process.StandardOutput;
                                output2 = reader8.ReadToEnd();
                                int baslangic = output2.IndexOf("key:") + "key:".Length;
                                string office = output2.Substring(baslangic, 6);
                                textBox1.Text += "OFFİCE ÜRÜN ANAHTARI" + Environment.NewLine + office;
                            }
                            else
                            {
                                execute();
                                string office19_64 = "cd " + comboBox1.Text + "Program Files/Microsoft Office/Office16 & cscript ospp.vbs /dstatus";
                                sb2.Clear();
                                sb2.AppendLine(office19_64);
                                File.WriteAllText(bat, sb2.ToString());
                                StreamReader reader7 = process.StandardOutput;
                                output2 = reader7.ReadToEnd();

                                if (output2.Contains("Processing"))
                                {
                                    execute();
                                    File.WriteAllText(bat, sb2.ToString());
                                    StreamReader reader6 = process.StandardOutput;
                                    output2 = reader6.ReadToEnd();
                                    int baslangic = output2.IndexOf("key:") + "key:".Length;
                                    string office = output2.Substring(baslangic, 6);
                                    textBox1.Text += "OFFİCE ÜRÜN ANAHTARI" + Environment.NewLine + office;
                                }
                                else
                                {
                                    execute();
                                    string office14_64 = "cd " + comboBox1.Text + " Program Files / Microsoft Office / Office14 & cscript ospp.vbs / dstatus";
                                    sb2.Clear();
                                    sb.AppendLine(office14_64);
                                    File.WriteAllText(bat, sb.ToString());
                                    StreamReader reader8 = process.StandardOutput;
                                    output2 = reader8.ReadToEnd();
                                    if (output2.Contains("Processing"))
                                    {
                                        MessageBox.Show("14 varsa");
                                        execute();
                                        File.WriteAllText(bat, sb.ToString());
                                        StreamReader reader9 = process.StandardOutput;
                                        output2 = reader9.ReadToEnd();
                                        int baslangic = output2.IndexOf("key:") + "key:".Length;
                                        string office = output2.Substring(baslangic, 6);
                                        textBox1.Text += "OFFİCE ÜRÜN ANAHTARI" + Environment.NewLine + office;

                                    }
                                    else
                                    {
                                        execute();
                                        string office14_32 = "cd " + comboBox1.Text + "Program Files (x86)/Microsoft Office/Office14 & cscript ospp.vbs /dstatus";
                                        sb.AppendLine(office14_32);
                                        sb.Replace(office14_64, office14_32);
                                        File.WriteAllText(bat, sb.ToString());
                                        StreamReader reader10 = process.StandardOutput;
                                        output2 = reader10.ReadToEnd();
                                        if (output2.Contains("Processing"))
                                        {
                                            execute();
                                            File.WriteAllText(bat, sb.ToString());
                                            StreamReader reader11 = process.StandardOutput;
                                            output2 = reader11.ReadToEnd();
                                            int baslangic = output2.IndexOf("key:") + "key:".Length;
                                            string office = output2.Substring(baslangic, 6);
                                            textBox1.Text += "OFFİCE ÜRÜN ANAHTARI" + Environment.NewLine + office;

                                        }
                                        else
                                        {
                                            MessageBox.Show("Bilgisayaranızda Office Bulunmamaktadır.");

                                        }
                                    }
                                }
                            }
                        }
                    }
                    progressBar1.Value = 50;

                    if (textBox1.Text.Contains("OFFİCE ÜRÜN ANAHTARI"))
                    {
                        int office2 = textBox1.Text.IndexOf("OFFİCE ÜRÜN ANAHTARI") + "OFFİCE ÜRÜN ANAHTARI".Length;
                        officekey = textBox1.Text.Substring(office2);
                    }
                    else
                    {
                        officekey = "";
                    }

                    string date = DateTime.Now.ToShortDateString();
                    string time = DateTime.Now.ToShortTimeString();

                    string fileName = cBirim.SelectedItem.ToString();
                    Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + fileName);

                    //veriler pdf'e aktarıldı
                    string file =txtAd.Text + ".pdf";
                    FileStream fs = new FileStream(Path.Combine(AppDomain.CurrentDomain.BaseDirectory + fileName, file), FileMode.Create);

                    iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4);

                    PdfWriter write = PdfWriter.GetInstance(doc, fs);
                    BaseFont bF = BaseFont.CreateFont("C:\\windows\\fonts\\arial.ttf", "windows-1254", true);
                    iTextSharp.text.Font font = new iTextSharp.text.Font(bF, 12f, iTextSharp.text.Font.NORMAL);
                    iTextSharp.text.Font font2 = new iTextSharp.text.Font(bF, 11f, iTextSharp.text.Font.BOLD);
                    doc.Open();

                    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(AppDomain.CurrentDomain.BaseDirectory + "elfatek.png");
                    doc.Add(img);
                    img.SetAbsolutePosition(1, 1);

                    PdfPTable tbl = new PdfPTable(1);
                    tbl.DefaultCell.BorderWidth = 0;
                    tbl.WidthPercentage = 100;
                    tbl.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    Chunk tittle = new Chunk("BİLGiSAYARIN TEKNİK ÖZELLİKLERİ", FontFactory.GetFont("Arial"));
                    tittle.Font = font;
                    tittle.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);
                    tittle.Font.SetStyle(3);
                    tittle.Font.Size = 14;
                    Phrase p1 = new Phrase();
                    p1.Add(tittle);
                    tbl.AddCell(p1);
                    doc.Add(tbl);


                    PdfPTable tbl4 = new PdfPTable(1);
                    tbl4.DefaultCell.BorderWidth = 0;
                    tbl4.WidthPercentage = 100;
                    PdfPCell cell16 = new PdfPCell(new Phrase(date + "  " + time));
                    cell16.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell16.VerticalAlignment = Element.ALIGN_TOP;
                    cell16.MinimumHeight = 20;
                    cell16.Border = 0;
                    tbl4.AddCell(cell16);
                    doc.Add(tbl4);

                    PdfPTable tbl1 = new PdfPTable(2);
                    tbl1.DefaultCell.Padding = 5;
                    tbl1.WidthPercentage = 100;
                    tbl1.DefaultCell.BorderWidth = 0.5f;
                    tbl1.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;                 
                    PdfPCell cell = new PdfPCell(new Phrase("ÜRÜNLER", font2));
                    cell.BackgroundColor = new iTextSharp.text.BaseColor(22, 85, 160);
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    tbl1.AddCell(cell);                
                    PdfPCell cell2 = new PdfPCell(new Phrase("TEKNİK ÖZELLİKLER", font2));          
                    cell2.BackgroundColor = new iTextSharp.text.BaseColor(22, 85, 160);
                    cell2.HorizontalAlignment = Element.ALIGN_CENTER;
                    tbl1.AddCell(cell2);
                    tbl1.AddCell(new Phrase("WINDOWS ÜRÜN ANAHTARI",font2));
                    PdfPCell cell3 = new PdfPCell(new Phrase(win.Trim()));
                    cell3.VerticalAlignment = Element.ALIGN_CENTER;
                    tbl1.AddCell(cell3);
                    tbl1.AddCell(new Phrase("OFFICE ÜRÜN ANAHTARI",font2));
                    PdfPCell cell4 = new PdfPCell(new Phrase(officekey.Trim()));
                    tbl1.AddCell(cell4);
                    tbl1.AddCell(new Phrase("ANAKART MODELİ",font2));
                    PdfPCell cell5 = new PdfPCell(new Phrase(mbModel.Trim()));
                    tbl1.AddCell(cell5);
                    tbl1.AddCell(new Phrase("SİSTEM MODELİ",font2));
                    PdfPCell cell6 = new PdfPCell(new Phrase(system.Trim()));
                    tbl1.AddCell(cell6);
                    tbl1.AddCell(new Phrase("İŞLEMCİ ADI",font2));
                    PdfPCell cell7 = new PdfPCell(new Phrase(cpu.Trim()));
                    tbl1.AddCell(cell7);
                    tbl1.AddCell(new Phrase("RAM TÜRÜ",font2));
                    PdfPCell cell8 = new PdfPCell(new Phrase(ram.Trim()));
                    tbl1.AddCell(cell8);
                    tbl1.AddCell(new Phrase("RAM KAPASİTESİ",font2));
                    PdfPCell cell9 = new PdfPCell(new Phrase(rams.Replace("\r", string.Empty).Replace("\n", string.Empty)));
                    tbl1.AddCell(cell9);
                    tbl1.AddCell(new Phrase("RAM MARKASI",font2));
                    PdfPCell cell10 = new PdfPCell(new Phrase(ramManu.Trim()));
                    tbl1.AddCell(cell10);
                    tbl1.AddCell(new Phrase("DİSK",font2));
                    PdfPCell cell11 = new PdfPCell(new Phrase(disk.Trim()));
                    tbl1.AddCell(cell11);                
                    tbl1.AddCell(new Phrase("EKRAN KARTI",font2));
                    PdfPCell cell12 = new PdfPCell(new Phrase(display.Trim()));
                    tbl1.AddCell(cell12);
                    tbl1.AddCell(new Phrase("ANAKART SERİ NUMARASI",font2));
                    PdfPCell cell13 = new PdfPCell(new Phrase(serial.Trim()));
                    cell13.VerticalAlignment = Element.ALIGN_CENTER;                  
                    doc.Add(tbl1);

                    iTextSharp.text.Paragraph p = new iTextSharp.text.Paragraph(" ");
                    doc.Add(p);
                    doc.Add(p);

                    PdfPTable tbl9 = new PdfPTable(2);
                    tbl9.DefaultCell.Padding = 5;
                    tbl9.WidthPercentage = 50;
                    tbl9.DefaultCell.BorderWidth = 0.5f;
                    tbl9.HorizontalAlignment = Element.ALIGN_RIGHT;                
                    PdfPCell cell22 = new PdfPCell(new Phrase("AD SOYAD", font2));
                    cell22.BackgroundColor = new iTextSharp.text.BaseColor(22, 85, 160);
                    cell22.HorizontalAlignment = Element.ALIGN_CENTER;
                    tbl9.AddCell(cell22);                   
                    PdfPCell cell23 = new PdfPCell(new Phrase("BİRİM", font2));
                    cell23.BackgroundColor = new iTextSharp.text.BaseColor(22, 85, 160);
                    cell23.HorizontalAlignment = Element.ALIGN_CENTER;
                    tbl9.AddCell(cell23);
                    tbl9.AddCell(new Phrase(txtAd.Text,font2));
                    tbl9.AddCell(new Phrase(cBirim.SelectedItem.ToString(),font2));
                    doc.Add(tbl9);
              
                   progressBar1.Value = 75;
                    doc.Close();
                    fs.Close();

                    //pdf excel'e dönüştürüldü
                    string pdfFile = fileName+"/" + file;
                    string excelFile = System.IO.Path.ChangeExtension(pdfFile, ".xls");

                    PdfFocus f = new PdfFocus();
                    f.ExcelOptions.ConvertNonTabularDataToSpreadsheet = true;

                    f.ExcelOptions.PreservePageLayout = true;

                    f.OpenPdf(pdfFile);

                    if (f.PageCount > 0)
                    {
                        f.ToExcel(excelFile);
                    }
                    progressBar1.Value = 100;
                    label3.Text = "Pdf ve Excel Formatında Kaydedildi.";
                }
                else
                {
                    MessageBox.Show("Office'in Kayıtlı olduğu Depolama Alanını Seçiniz!");
                }
            }
            else
            {
                MessageBox.Show("AD SOYAD VE BİRİM BİLGİLERİNİ DOLDURUNUZ");
            }
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }


        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(Handle, 0x112, 0xf012, 0);
        }

     
        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.Width += 5;
            button1.Height += 5;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.Width -= 5;
            button1.Height -= 5;
        }

        private void panel3_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(Handle, 0x112, 0xf012, 0);
        }
    }
    
}
