
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using OfficeOpenXml;
using System.Text;
using System.IO;
using LinqToExcel;
using Application = Microsoft.Office.Interop.Excel.Application;
using OpenQA.Selenium.Support.Events;
using DataTable = System.Data.DataTable;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        IWebDriver driver;
        public void OkulAdiBrowser()
        {

            driver = new FirefoxDriver();
            driver.Navigate().GoToUrl("http://mtegm.meb.gov.tr/kurumlar/?ara");
            driver.Manage().Window.Maximize();

        }
        public void MYKBrowser()
        {
            driver = new FirefoxDriver();
            driver.Navigate().GoToUrl("https://portal.myk.gov.tr/index.php?option=com_yeterlilik&view=arama&belge_zorunlu=1");
            driver.Manage().Window.Maximize();
        }

        public void UniversiteBrowser()
        {
            driver = new FirefoxDriver();
            driver.Navigate().GoToUrl("https://yokatlas.yok.gov.tr/tercih-sihirbazi-t4-tablo.php?p=say");
            driver.Manage().Window.Maximize();

            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OkulAdiBrowser();
            driver.FindElement(By.XPath("/html/body/section[1]/div/div[2]/div/div[3]/table/tbody/tr[4]/td[2]/form/span/span[1]/span/span[2]")).Click();
            System.Threading.Thread.Sleep(5000);
            driver.FindElement(By.XPath("/html/body/span/span/span[1]/input")).SendKeys("MTAL");
            System.Threading.Thread.Sleep(5000);
            driver.FindElement(By.XPath("/html/body/span/span/span[1]/input")).SendKeys(OpenQA.Selenium.Keys.Enter);
            System.Threading.Thread.Sleep(10000);

            var kayitGorunumu = driver.FindElement(By.XPath("/html/body/section[1]/div/div[2]/div/div[3]/div/div/div/div[1]/div[1]/div/label/select"));
            var tumu = new SelectElement(kayitGorunumu);
            tumu.SelectByValue("50");

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Sıra No");
            dt.Columns.Add("İl");
            dt.Columns.Add("İlçe");
            dt.Columns.Add("Okul Adı");
            dt.Columns.Add("Kurum Kodu");

            for (int i = 1; i <= 50; i++)
            {
                var tablo = driver.FindElement(By.XPath("/html/body/section[1]/div/div[2]/div/div[3]/div/div/div/div[2]/div/table/tbody"));
                foreach (var row in tablo.FindElements(By.TagName("tr")))
                {
                    //Almak istediğim sütun sayısını belirler.
                    int cellIndex = 0;

                    string[] okul = new string[5];
                    foreach (var cell in row.FindElements(By.TagName("td")))
                    {
                        if (cellIndex <= 4)
                        {
                            okul[cellIndex] = cell.Text;

                        }
                        else
                            continue;

                        cellIndex++;
                    }
                    dt.Rows.Add(okul);
                }
                dataGridView1.DataSource = dt;
                driver.FindElement(By.XPath("/html/body/section[1]/div/div[2]/div/div[3]/div/div/div/div[3]/div[2]/div/ul/li[9]")).Click();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(@"C:\Users\Lenovo\Desktop\Nota\WindowsFormsApp1\TürkiyeMTALiseleri.xlsx");
            Worksheet ws = wb.Worksheets[1];
            string[] kurumKoduListesi = new string[2487];
            for (int i = 2; i < 2487; i++)
            {
                if (ws.Cells[i, 5].Value != null)
                {
                    var kurumKodu = ws.Cells[i, 5].Value.ToString();
                    kurumKoduListesi[i] = kurumKodu;
                }
            }

            driver = new FirefoxDriver();

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Kurum Kodu");
            dt.Columns.Add("Okul Adı");
            dt.Columns.Add("Program Türü");
            dt.Columns.Add("Alan Adı");
            dt.Columns.Add("Dal Adı");
            dt.Columns.Add("Yabancı Dil");
            dt.Columns.Add("Öğretim Şekli");

            for (int i = 2; i < 2487; i++)
            {
                driver.Navigate().GoToUrl($"http://mtegm.meb.gov.tr/kurumlar/?s=kurumDetay&d=kurum.alandal&kk={kurumKoduListesi[i]}");

                System.Threading.Thread.Sleep(5000);

                var okulAdiTag = driver.FindElement(By.XPath("/html/body/section[1]/div/div[2]/div/div[3]/div/div[1]/table/tbody/tr[1]/td[2]"));
                var okulAdi = "";
                foreach (var item in okulAdiTag.FindElements(By.TagName("strong")))
                {

                    okulAdi = item.Text;
                }
                System.Threading.Thread.Sleep(2000);



                var tablo = driver.FindElement(By.XPath("/html/body/section[1]/div/div[2]/div/div[3]/div/div[2]/table/tbody"));
                foreach (var row in tablo.FindElements(By.TagName("tr")))
                {
                    //Almak istediğim sütun sayısını belirler.
                    int cellIndex = 0;

                    string[] okul = new string[5];
                    foreach (var cell in row.FindElements(By.TagName("td")))
                    {
                        if (cellIndex <= 4)
                        {
                            okul[cellIndex] = cell.Text;

                        }
                        else
                            continue;

                        cellIndex++;
                    }


                    dt.Rows.Add(kurumKoduListesi[i], okulAdi, okul[0], okul[1], okul[2], okul[3], okul[4]);

                }
                dataGridView1.DataSource = dt;

                System.Threading.Thread.Sleep(2000);
            }

        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {

                Application xcelApp = new Application();
                xcelApp.Application.Workbooks.Add(Type.Missing);

                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    xcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        xcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
                xcelApp.Columns.AutoFit();
                xcelApp.Visible = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MYKBrowser();

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Sıra No");
            dt.Columns.Add("Yeterlilik Kodu");
            dt.Columns.Add("Yeterlilik Adı");
            dt.Columns.Add("Seviye");
            dt.Columns.Add("Sektör");


            var tablo = driver.FindElement(By.XPath("/html/body/div[8]/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div/div[5]/table/tbody"));

            foreach (var row in tablo.FindElements(By.TagName("tr")))
            {

                //Almak istediğim sütun sayısını belirler.
                int cellIndex = 0;


                string[] temelBilgiler = new string[5];
                foreach (var cell in row.FindElements(By.TagName("td")))
                {
                    if (cellIndex <= 4)
                    {
                        temelBilgiler[cellIndex] = cell.Text;
                    }
                    else
                        continue;

                    cellIndex++;

                }
                if (temelBilgiler[0].Length < 4)
                {
                    dt.Rows.Add(temelBilgiler);
                }
                else continue;
            }
            dataGridView1.DataSource = dt;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            UniversiteBrowser();
            System.Threading.Thread.Sleep(5000);




            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("YOP Kodu");
            dt.Columns.Add("Universite Adı");
            dt.Columns.Add("Fakülte Adı");
            dt.Columns.Add("Bölüm Adı");
            dt.Columns.Add("Bölüm Tipi");
            dt.Columns.Add("Şehir");
            dt.Columns.Add("Üniversite Tür");
            dt.Columns.Add("Ücret/Burs");
            dt.Columns.Add("Öğrenim Türü");




            System.Threading.Thread.Sleep(7000);

            string[] YOPkodu = new string[1];
            string[] UniAdi = new string[2];
            string[] UniFakulteAdi = new string[2];
            string[] Sehir = new string[1];
            string[] UniTuru = new string[1];
            string[] UniTuruBursUcret = new string[1];
            string[] OgrenimTuru = new string[1];

            for (int i = 0; i <= 97; i++)
            {

                for (int j = 1; j <= 50; j++)
                {

                    
                    var yopKodu = driver.FindElement(By.XPath($"/html/body/div/div[2]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[{j}]/td[2]/a")).Text;
                    YOPkodu[0] = yopKodu;

                    var uniAdi = driver.FindElement(By.XPath($"/html/body/div/div[2]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[{j}]/td[3]/strong")).Text;
                    UniAdi[0] = uniAdi;

                    var uniBolumAdi = driver.FindElement(By.XPath($"/html/body/div/div[2]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[{j}]/td[4]/strong")).Text;
                    UniAdi[1] = uniBolumAdi;

                    var uniFakulteAdi = driver.FindElement(By.XPath($"/html/body/div/div[2]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[{j}]/td[3]/font")).Text;
                    UniFakulteAdi[0] = uniFakulteAdi;

                    var uniFakulteCesidiAdi = driver.FindElement(By.XPath($"/html/body/div/div[2]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[{j}]/td[4]/font")).Text;
                    UniFakulteAdi[1] = uniFakulteCesidiAdi;

                    var sehirAdi = driver.FindElement(By.XPath($"/html/body/div/div[2]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[{j}]/td[5]")).Text;
                    Sehir[0] = sehirAdi;

                    var uniTuru = driver.FindElement(By.XPath($"/html/body/div/div[2]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[{j}]/td[6]")).Text;
                    UniTuru[0] = uniTuru;

                    var uniTuruBursUcret = driver.FindElement(By.XPath($"/html/body/div/div[2]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[{j}]/td[7]")).Text;
                    UniTuruBursUcret[0] = uniTuruBursUcret;

                    var ogrenimTuru = driver.FindElement(By.XPath($"/html/body/div/div[2]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[{j}]/td[8]")).Text;
                    OgrenimTuru[0] = ogrenimTuru;

                    dt.Rows.Add(YOPkodu[0], UniAdi[0], UniFakulteAdi[0], UniAdi[1], UniFakulteAdi[1], Sehir[0], UniTuru[0], UniTuruBursUcret[0], OgrenimTuru[0]);
                }

                dataGridView1.DataSource = dt;

                System.Threading.Thread.Sleep(3000);

                //Scroll aşağıya doğru hareket ettirmek için.
                IJavaScriptExecutor js1 = driver as IJavaScriptExecutor;
                js1.ExecuteScript("window.scrollBy(0,4000);");

                System.Threading.Thread.Sleep(3000);
                driver.FindElement(By.XPath("/html/body/div/div[2]/div[2]/div[2]/div/div/div[3]/div[2]/div/ul/li[9]")).Click();

                System.Threading.Thread.Sleep(8000);

            }
            

        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Sıra No");
            dt.Columns.Add("Ulusal Yeterlilik Kodu");
            dt.Columns.Add("ISCO 08");
            dt.Columns.Add("Millî Eğitim Bakanlığına Bağlı Mesleki ve Teknik Eğitim Kurumlarınca Verilen Diplomalar (Alan/Dal/Bölüm-FOET Kodu)");
           

            _Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = excel.Workbooks.Open(@"C:\Users\Lenovo\Desktop\Nota\MYK\Muafiyet_Tablosu_2020-Rev.04.xlsx");
            Worksheet ws = wb.Worksheets[1];

            
                for (int j = 2; j < 105; j++)
                {
                 var  SiraNo =ws.Cells[j,1].Value;

                var UYK = ws.Cells[j, 2].Value;
                var ISCOkodu = ws.Cells[j, 3].Value;

                var MEB = (string)ws.Cells[j, 4].Value;

                string[] ayirici = new string[] { ")," };


                var MeslekFoetListe = MEB.Split(ayirici, StringSplitOptions.None);

                foreach (var item in MeslekFoetListe)
                {
                    var MeslekFoet = item;
                    dt.Rows.Add(SiraNo, UYK, ISCOkodu, MeslekFoet);
                }



               

                }
            
            dataGridView1.DataSource = dt;
                
            
        }

    }
}











