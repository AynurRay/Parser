using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using HtmlAgilityPack;
using ScrapySharp.Extensions;
using System.Text.RegularExpressions;

using System.Data.SqlClient;
using System.Data;

namespace ParserGP
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        SqlConnection conn;
        SqlConnectionStringBuilder connStringBuilder;

        public void ConnectOpen()
        {
            connStringBuilder = new SqlConnectionStringBuilder();
            connStringBuilder.DataSource = @"(localdb)\v11.0";
            connStringBuilder.InitialCatalog = "ProductDB";
            connStringBuilder.IntegratedSecurity = true;

            conn = new SqlConnection(connStringBuilder.ToString());
            conn.Open();
        }

        public void ConnectClose()
        {
            conn.Close();
        }

        async void Pasring()
        {

            dataGridView1.Rows.Clear();
            start.Enabled = false;

            int counts = Int32.Parse(finishBox.Text);
            int index = 0;
            string ComURL = ComUrlBox.Text;

            /*Постраничный перебор*/
            for (int i = Int32.Parse(startBox.Text); i <= counts; i++)
            {               
                /*Селекторы*/
                var catalog_selector = blockBox.Text;
                var name_selector = nameBox.Text;
                var url_selector = urlBox.Text;
                var price_selector = priceBox.Text;
                var img_selector = imgBox.Text;

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); //Кодировка

                var webGet = new HtmlWeb();

                var FullURL = string.Format(UrlCatalog.Text, i);//получение данных с сервера

                //var FullURL = string.Format(@"C:\Users\aynur\Desktop\shop.rt.ru\pars_file(shop.rt.ru)\example\main.html");//локальный файл


                if (webGet.Load(FullURL) is HtmlAgilityPack.HtmlDocument doc)
                {
                    var items = doc.DocumentNode.CssSelect(catalog_selector).ToList();

                    foreach (var item in items)
                    { 
                        string filter_for_name = filter_nameBox.Text;//фильтр для наименования
                        string filter_for_price = filter_priceBox.Text;//фильтр для цены

                        string brend = null;
                        string color = null;
                        string model = null;
                        

                        /*Наименование*/
                        string name = item.CssSelect(name_selector).Single().InnerText.Trim().ToUpper();//получаем наименование                
                        name = System.Text.RegularExpressions.Regex.Replace(name, @"\s+", " ");//убираем лишние пробелы из наименования
                        name = Regex.Replace(name, filter_for_name.ToUpper(), String.Empty);//применяем фильтр


                        string brendToFind = brendBox.Text;                      
                        MatchCollection matches = Regex.Matches(name, @"\b" + brendToFind + @"\b", RegexOptions.IgnoreCase);

                        if (matches.Count > 0)
                        {
                            brend = matches[0].Groups[0].Value;
                        }

                        else brend = null;



                        string colorToFind = colorBox.Text;
                        MatchCollection matches2 = Regex.Matches(name, @"\b" + colorToFind + @"\b", RegexOptions.IgnoreCase);

                        if (matches2.Count > 0)
                        {
                            color = matches2[0].Groups[0].Value;
                        }

                        else color = null;



                        string model_filter = brend + "|" + color;//фильр
                        model = Regex.Replace(name, model_filter.ToUpper(), String.Empty);//применяем фильтр



                        /*Ссылка*/
                        string url = ComURL + item.CssSelect(url_selector).Single().Attributes["href"].Value;

                        /*Цена*/
                        string price = item.CssSelect(price_selector).Single().InnerText.Trim().ToUpper();
                        price = System.Text.RegularExpressions.Regex.Replace(price, @"\s+", " ");
                        price = Regex.Replace(price, filter_for_price.ToUpper(), String.Empty);

                        /*Время и дата*/
                        string date = DateTime.Now.ToShortDateString();
                        string time = DateTime.Now.ToLongTimeString();

                        string img = ComURL + item.CssSelect(img_selector).Single().Attributes["src"].Value;

                        index++;

                        dataGridView1.Rows.Add(
                            index,
                            name,
                            brend,
                            model,
                            color,
                            url,  
                            price,
                            date,
                            time,
                            img
                            );

                        await Task.Delay(TimeSpan.FromSeconds(short.Parse(this.time.Text)));

                    }                   

                }           
            }

            MessageBox.Show("Было найдено " + index.ToString() + " наименований.");
            start.Enabled = true;
        }



        private void button1_Click(object sender, EventArgs e)
        {
                Pasring();
        }


        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            filter_nameBox.Clear();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            filter_priceBox.Clear();
        }


        private void button3_Click(object sender, EventArgs e)
        {
            var openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filter_nameBox.Text = File.ReadAllText(openFileDialog1.FileName, Encoding.Default);          
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filter_priceBox.Text = File.ReadAllText(openFileDialog1.FileName, Encoding.Default);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                brendBox.Text = File.ReadAllText(openFileDialog1.FileName, Encoding.Default);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            var openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                colorBox.Text = File.ReadAllText(openFileDialog1.FileName, Encoding.Default);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            brendBox.Clear();
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            colorBox.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
                int row = dataGridView1.RowCount;
                string FullName=null;
                string Brend=null;
                string Model=null;
                string Color=null;
                string Linq=null;
                int Price;
                string Date;
                string Time;
                string Img;

                int CompanyID = Int32.Parse(textBox1.Text);

            try
            {

                for (int n = 0; n < (row - 1); n++)
                {

                    //считываем данные с таблицы
                    FullName = Convert.ToString(dataGridView1[1, n].Value.ToString());
                    Brend = Convert.ToString(dataGridView1[2, n].Value.ToString());
                    Model = Convert.ToString(dataGridView1[3, n].Value.ToString());
                    Color = Convert.ToString(dataGridView1[4, n].Value.ToString());          
                    Linq = Convert.ToString(dataGridView1[5, n].Value.ToString());
                    Price = Convert.ToInt32(dataGridView1[6, n].Value.ToString());
                    Date = Convert.ToString(dataGridView1[7, n].Value.ToString());
                    Time = Convert.ToString(dataGridView1[8, n].Value.ToString());
                    Img = Convert.ToString(dataGridView1[9, n].Value.ToString());


                //поиск записи товара в БД
                var commandTextAll = "SELECT COUNT(*) FROM Products WHERE Linq='" + Linq + "' AND CompanyID=" + CompanyID + "";
                SqlCommand cmd = new SqlCommand(commandTextAll, conn);

                int count = (int)cmd.ExecuteScalar();

                 //если запись существует, то добавляем цену и дату  
                if (count !=0)
                {
                    //ищем добавленную ID товара
                    var commandText = "SELECT ID FROM Products WHERE Linq='" + Linq + "' AND CompanyID=" + CompanyID + "";//поиск ID товара
                    SqlCommand cmd2 = new SqlCommand(commandText, conn);

                    int count_id = (int)cmd2.ExecuteScalar();

                    //ищем последнюю ID
                    var commandText_ = "SELECT MAX(ID) FROM HistoryPrice";
                    SqlCommand cmd3 = new SqlCommand(commandText_, conn);
                    int count_id_ = (int)cmd3.ExecuteScalar();

                    count_id_++;

                    //добавляем цену
                    var cmdText = "INSERT INTO HistoryPrice (ID, Price, Date, Time, ProductID) VALUES (@ID, @Price, @Date, @Time, @ProductID)";
                    SqlCommand command = new SqlCommand(cmdText, conn);
                    command.Parameters.AddWithValue("@ID", count_id_);
                    command.Parameters.AddWithValue("@Price", Price);
                    command.Parameters.AddWithValue("@Date", Date);
                    command.Parameters.AddWithValue("@Time", Time);
                    command.Parameters.AddWithValue("@ProductID", count_id);
                    command.ExecuteNonQuery();

                    //обновляем цену
                    var commandTextUP = "UPDATE Products SET CurrentPrice=" + Price + " WHERE ID="+ count_id +"";
                    SqlCommand cmdUP = new SqlCommand(commandTextUP, conn);
                    cmdUP.ExecuteNonQuery();
                    }
                    
                //если НЕ существует, то создаем новую запись и потом добавляем цену и дату
                if (count == 0)
                {
                    //ищем последнюю ID
                    var commandText_ = "SELECT MAX(ID) FROM Products";
                    SqlCommand cmd3 = new SqlCommand(commandText_, conn);
                    int count_id_ = (int)cmd3.ExecuteScalar();

                    count_id_++;

                   //добавляем товар
                    var cmdText1 = "INSERT INTO Products (ID, FullName, Brend, Model, Color, Linq, Img, CurrentPrice, CompanyID) VALUES (@ID, @FullName, @Brend, @Model, @Color, @Linq, @Img, @CurrentPrice, @CompanyID)";
                    SqlCommand command1 = new SqlCommand(cmdText1, conn);
                    command1.Parameters.AddWithValue("@ID", count_id_);
                    command1.Parameters.AddWithValue("@FullName", FullName);
                    command1.Parameters.AddWithValue("@Brend", Brend);
                    command1.Parameters.AddWithValue("@Model", Model);
                    command1.Parameters.AddWithValue("@Color", Color);
                    command1.Parameters.AddWithValue("@Linq", Linq);
                    command1.Parameters.AddWithValue("@Img", Img);
                    command1.Parameters.AddWithValue("@CurrentPrice", Price);
                    command1.Parameters.AddWithValue("@CompanyID", CompanyID);

                    command1.ExecuteNonQuery();

                    //ищем добавленную ID товара
                    var commandText = "SELECT MAX(ID) FROM Products";
                    SqlCommand cmd2 = new SqlCommand(commandText, conn);
                    int count_id = (int)cmd2.ExecuteScalar();



                    //ищем последнюю ID
                    var commandText_4 = "SELECT MAX(ID) FROM HistoryPrice";
                    SqlCommand cmd4 = new SqlCommand(commandText_4, conn);
                    int count_id_4 = (int)cmd4.ExecuteScalar();

                    count_id_4++;

                    //добавляем цену и дату
                    var cmdText = "INSERT INTO HistoryPrice (ID, Price, Date, Time, ProductID) VALUES (@ID, @Price, @Date, @Time, @ProductID)";
                    SqlCommand command = new SqlCommand(cmdText, conn);
                    command.Parameters.AddWithValue("@ID", count_id_4);
                    command.Parameters.AddWithValue("@Price", Price);
                    command.Parameters.AddWithValue("@Date", Date);
                    command.Parameters.AddWithValue("@Time", Time);
                    command.Parameters.AddWithValue("@ProductID", count_id);

                    command.ExecuteNonQuery();

                }

            }     

                MessageBox.Show("Успешно!");
            }


            catch { MessageBox.Show("Ошибка!"); }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        { 
            ConnectOpen(); label15.Text = "Connection: Open";
        }

        private void button11_Click(object sender, EventArgs e)
        {
            ConnectClose(); label15.Text = "Connection: Close";
        }
    }
}
