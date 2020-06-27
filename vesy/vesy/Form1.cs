using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO.Ports;
using System.Data.OleDb;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        SerialPort SerialPort1 = new SerialPort("COM2");
        SerialPort SerialPort2 = new SerialPort("COM5");
       OleDbConnection myOleDbConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ves.mdb");
      
        static int rzr;
        static int rzr0 = 1;
        public Form1()
        {
          
            InitializeComponent();
            SerialPort1.BaudRate = 2400;
            SerialPort1.Parity = Parity.None;
            SerialPort1.StopBits = StopBits.One;
            SerialPort1.DataBits = 8;
            SerialPort2.BaudRate = 4800;
            SerialPort2.Parity = Parity.None;
            SerialPort2.StopBits = StopBits.One;
            SerialPort2.DataBits = 8;
            SerialPort1.Open();
            SerialPort2.Open();
            timer1.Enabled = true;
            timer2.Enabled = false;
            Greed();
            
        }
       
        private void timer1_Tick(object sender, EventArgs e)
        {
          
            int ff = SerialPort1.BytesToRead;
            int ff1 = SerialPort2.BytesToRead;
            SerialPort1.DiscardInBuffer();
            System.Threading.Thread.Sleep(300); //
            string dd = System.Convert.ToString(ff);
            string input = SerialPort1.ReadExisting();

           if (input.IndexOf("a") >= 0)
           {
            
               timer1.Enabled = false;
               timer2.Enabled = true;
               SerialPort2.DiscardInBuffer();
               string gg = SerialPort2.ReadExisting(); 
               rzr = 1;
               pictureBox1.BackColor = Color.Red;
           }
        }
        
        private void timer2_Tick(object sender, EventArgs e)
        {
            label7.Text = "";
            label8.Text = "";
            label9.Text = "";
            SerialPort2.DiscardInBuffer(); // очистка буфера приема
           System.Threading.Thread.Sleep(300); // задержка на 300 мс
           int d = SerialPort2.BytesToRead; 
           int idx;
            string dd = System.Convert.ToString(d);
            double output = -1;
            string input1 = SerialPort2.ReadExisting();
            label7.Text = input1;
            label1.Text = input1;
            if (myOleDbConnection.State != System.Data.ConnectionState.Open) myOleDbConnection.Open(); 
              if (input1.IndexOf("kg") >= 0)
            {
                  //парсер
                idx = input1.IndexOf("WW");

                if (idx > -1)
                {
                    input1 = input1.Substring(idx + 2, 9);
                    label8.Text = input1;
                }
                idx = input1.IndexOf("kg");
                if (idx == 7)
                {
                 
                    input1 = input1.Substring(0, 7);
                    label9.Text = input1;
                    string ves = input1;
             
                }
                
                idx = input1.IndexOf('.');
                if (idx == 5)
                {
                    string input_ = input1.Replace('.', ',');
                    output = Convert.ToDouble(input_);
                
                }
                if (output > 0 & output < 301 & rzr == 1 & rzr0 == 1)
               
                { // если вес больше 0 и меньше 300 и запись в БД разрешена 
                    // зажечь светодиод
                  /* --узнать текущее дату и время и преобразовать его в символьный вид ---------------*/
                 //   label3.Text = output.ToString() + " кг";
                   DateTime dk = DateTime.Now; //
                   string dat = dk.ToString();
                   string  dd1 = dat.Substring(0, 2);
                   string  mm = dat.Substring(3, 2);
                   string  yyyy = dat.Substring(6, 4);
                   string  chmm = dat.Substring(11, 7);
                   string  tek_data = "#" + yyyy + "/" + mm + "/" + dd1 + " " + chmm + "#";
                    /* --сформировать SQL запрос вставки строки в БД -----------------*/
                   string textCommand1 = "INSERT INTO ves(data,ves) values(" + tek_data + "," + input1 + ")";
                    OleDbCommand command = new OleDbCommand(textCommand1, myOleDbConnection);
                    SerialPort1.WriteLine("+"); // Записать в SerialPort1 символ "+"
                    command.ExecuteNonQuery(); // записать данные в БД
                    myOleDbConnection.Close();
                    
                    DateTime df = DateTime.Today;
                    string tek_data1 = df.ToString();
                    string dd2 = tek_data1.Substring(0, 2);
                    string mm2 = tek_data1.Substring(3, 2);
                    string yyyy2 = tek_data1.Substring(6, 4);
                    tek_data = "#" + yyyy2 + "/" + mm2 + "/" + dd2 + "#";
                    Greed();
                    
                    rzr = 0; // сбросить флаг разрешения записи в БД
                    rzr0 = 0;

                }
                // АНАЛИЗ НА НУЛЕВОЙ ВЕС
                if (output == 0) //ЕСЛИ ВЕС равен НУЛю  отправка "-" в порт1
                { SerialPort1.WriteLine("-"); // потушить светодиод
             //   SerialPort1.DiscardInBuffer();
                timer1.Enabled = true;
                timer2.Enabled = false;
                rzr0 = 1;
                pictureBox1.BackColor = Color.Lime;
                }
                    

            }
        }

        private void Greed()
        {
            DateTime dk = DateTime.Now; //
            string dat = dk.ToString();
            string dd1 = dat.Substring(0, 2);
            string mm = dat.Substring(3, 2);
            string yyyy = dat.Substring(6, 4);
            string chmm = dat.Substring(11, 7);
         //   string tek_data = "#" + yyyy + "/" + mm + "/" + dd1 + " " + chmm + "#";
            string tek_data = "#" + yyyy + "/" + mm + "/" + dd1+"#";
            myOleDbConnection.Open();
            string Zapros1 = "Select  data, ves  from ves where  data >= " + tek_data + " AND data >= " + tek_data + " order by data desc"; // ZP.StartData + " and data <= " + ZP.EndData; ORDER BY data  DESC
            string textCommand = Zapros1;
            OleDbCommand myOleDbCommand = new OleDbCommand(textCommand, myOleDbConnection);
            //   label2.Text = ZP.Zapros1;
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(Zapros1, myOleDbConnection);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds, "ves");
            dataGridView1.DataSource = ds.Tables["ves"];
            int cont = dataGridView1.RowCount;
            if (cont > 0)
            {
                double sum = 0;
                label6.Text = cont.ToString();
                //    dataGridView1.DataSource = ds.Tables["ves"];
                dataGridView1.Columns[1].HeaderText = "Вес";
                dataGridView1.Columns[0].HeaderText = "Дата";
                for (int i = 0; i < cont; i++)
                {
                    sum = sum + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                }
                label3.Text = Convert.ToString(sum) + " кг";
            }
            myOleDbConnection.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
         

        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            timer2.Enabled = true;


            rzr = 1;
            pictureBox1.BackColor = Color.Red;
        }

     

        }

     

    
    }


