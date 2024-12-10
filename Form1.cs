using System.Collections;
using System.Data.SqlClient;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelVTEntegrasyonProjesi
{
    public partial class Form1 : Form
    {
        [Obsolete]
        SqlConnection baglanti = new SqlConnection(@"Server=LAPTOP-SFNHHOMM\SQLEXPRESS;Database=ProjelerVT;Trusted_Connection=True;");
        private string okunanHucre;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnVTdenOku_Click(object sender, EventArgs e)
        {
            Excel.Application exeUygulama = new Excel.Application();
            exeUygulama.Visible = true;
            Excel.Workbook workbook = exeUygulama.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet sayfa1 = workbook.Sheets[1];

            string[] basliklar = { "Personel no", "Ad", "Soyad", "Semt", "Sehir" };
            Excel.Range range;
            for (int i = 0; i < basliklar.Length; i++)
            {
                range = sayfa1.Cells[1 + i];
                range.Value2 = basliklar[i];



            }
            try
            {
                if (baglanti != null && baglanti.State == System.Data.ConnectionState.Closed)
                {
                    baglanti.Open();
                }

                string sqlCumlesi = "SELECT PersonelNo, Ad, Soyad, Semt, Sehir FROM Personel";
                SqlCommand sqlKomut = new SqlCommand(sqlCumlesi, baglanti);
                SqlDataReader sdr = sqlKomut.ExecuteReader();

                if (sdr != null && sdr.HasRows)
                {
                    int satir = 2; // ilk satırda başlıklar bulunur o yüzden 2
                    while (sdr.Read())
                    {
                        string pno = sdr[0]?.ToString() ?? "NULL";
                        string ad = sdr[1]?.ToString() ?? "NULL";
                        string soyad = sdr[2]?.ToString() ?? "NULL";
                        string semt = sdr[3]?.ToString() ?? "NULL";
                        string sehir = sdr[4]?.ToString() ?? "NULL";

                        richTextBox1.Text += $"{pno}  {ad}  {sehir}\n";

                        range = sayfa1.Cells[satir, 1];
                        range.Value2 = pno;
                        range = sayfa1.Cells[satir, 2];
                        range.Value2 = ad;
                        range = sayfa1.Cells[satir, 3];
                        range.Value2 = soyad;
                        range = sayfa1.Cells[satir, 4];
                        range.Value2 = semt;
                        range = sayfa1.Cells[satir, 5];
                        range.Value2 = sehir;
                        satir++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("SQL Query sırasında bir hata oluştu: " + ex.Message);
            }
            finally
            {
                if (baglanti != null && baglanti.State == System.Data.ConnectionState.Open)
                {
                    baglanti.Close();
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnExceldenOku_Click(object sender, EventArgs e)
        {
            Excel.Application exlApp = new Excel.Application();
            Excel.Workbook exlWorkbook = exlApp.Workbooks.Open(@"C:\Masaüstü\test.xlsx");
            Excel.Worksheet exlWorkSheet = (Excel.Worksheet)exlWorkbook.Worksheets[1];
            Excel.Range range = exlWorkSheet.UsedRange;

            int rCnt = 0;
            int cCnt = 0;
            //  exlApp= new Excel.Application();
            // Excel.Workbook exlWorkbook; ; = exlApp.WExcel.Workbook exlWorkbook;.Open("\"C:\\Masaüstü\\test.xlsx\"");
            // exlWorkSheet = (Excel.Worksheet)exlWorkook.WorkSheets.get_Item(1);

            //İlk olarak rich.TextBox2 içeriğini temizleyelim.
            richTextBox2.Clear();

            //ilk satır başlıkları içerdiği için rowcount u 2 den başlatmamız gerekiyor. 
            //eğer ilk satır veriler başlamışş olsaydı 1'den başlatmamız gerekirdi.

            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {
                List<string> list = new List<string>();

                // ArrayList list = new ArrayList();
                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {
                    string okunanHucre = (range.Cells[rCnt, cCnt] as Excel.Range)?.Value2?.ToString() ?? "NULL";

                    //string okunanHucre =Convert.ToString( range.Cells[rCnt,cCnt] as Excel.Range).Value2);
                    richTextBox2.Text += okunanHucre + "  ";
                    list.Add(okunanHucre);
                }
                richTextBox2.Text += "\n";

                try
                {
                    baglanti.Open();
                    SqlCommand sqlCommend = new SqlCommand("INSERT INTO Personel (PersonelNo, Ad, Soyad, Semt, Sehir)"
                                                            + "VALUES (@P1, @P2, @P3, @P4, @P5)", baglanti);
                    sqlCommend.Parameters.AddWithValue("@P1", list[0]);
                    sqlCommend.Parameters.AddWithValue("@P2", list[1]);
                    sqlCommend.Parameters.AddWithValue("@P3", list[2]);
                    sqlCommend.Parameters.AddWithValue("@P4", list[3]);
                    sqlCommend.Parameters.AddWithValue("@P5", list[4]);
                    sqlCommend.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Veritabanına yazarken hata oluştu! Hata kodu: SQLWRİTE02");
                }
                finally
                {
                    if (baglanti != null)
                        baglanti.Close();
                }

            }
        
        }
    }
}

 




