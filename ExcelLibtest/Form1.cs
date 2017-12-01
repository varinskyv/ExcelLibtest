using System;
using System.Windows.Forms;
using ExcelLib;
using System.Collections.Generic;
using System.Threading;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using System.IO.Compression;
using ExcelDataReader;
using System.Linq;
using DetailedOperatorServicesCore;
using Rostelecom;

namespace ExcelLibtest
{
    public partial class Form1 : Form
    {
        private Excel excel;

        private DateTime startDate;
        private object[,] data;
        private List<List<object>> sheets = new List<List<object>>();
        private LocalBase lbase;

        public Form1()
        {
            InitializeComponent();

            lbase = LocalBase.getInstance();
            lbase.Init("subscribers");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                excel = new Excel();

                startDate = DateTime.Now;

                excel.GetData(openFileDialog1.FileName, ExcelCallBack);

            }
        }

        private void ExcelCallBack(bool result)
        {
            MessageBox.Show(Convert.ToString(startDate - DateTime.Now));

            startDate = DateTime.Now;
            Rostelecom.Rostelecom rostelecom = new Rostelecom.Rostelecom();
            rostelecom.Parse(excel.Sheets, RostelecomCallBack);
        }

        private void RostelecomCallBack(CallBackResult result)
        {
            MessageBox.Show(Convert.ToString(startDate - DateTime.Now));
            MessageBox.Show(Convert.ToString(result.StartPeriodDate));
            MessageBox.Show(lbase.SubscriberList.Count.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Thread convert = new Thread(convertData);
            convert.Start();
        }

        private void convertData()
        {
            startDate = DateTime.Now;

            BinaryFormatter formatter = new BinaryFormatter();
            // получаем поток, куда будем записывать сериализованный объект
            using (FileStream fs = new FileStream("base.dat", FileMode.Create))
            {
                formatter.Serialize(fs, excel.Sheets);

                MessageBox.Show("Объект сериализован за " + Convert.ToString(startDate - DateTime.Now));
            }

            startDate = DateTime.Now;

            using (FileStream fs = new FileStream("base.ds", FileMode.Create))
            {
                // поток архивации
                using (DeflateStream compressionStream = new DeflateStream(fs, CompressionMode.Compress))
                {
                    formatter.Serialize(compressionStream, excel.Sheets);

                    MessageBox.Show("Объект сжат в DeflateStream за " + Convert.ToString(startDate - DateTime.Now));
                }
            }

            data = null;
            excel.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            startDate = DateTime.Now;

            BinaryFormatter formatter = new BinaryFormatter();

            using (FileStream fs = new FileStream("base.dat", FileMode.OpenOrCreate))
            {
                data = (object[,])formatter.Deserialize(fs);

                MessageBox.Show("Объект десериализован за " + Convert.ToString(startDate - DateTime.Now));

                MessageBox.Show(Convert.ToString(data[0, 0]));
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            List<Connection> con = new List<Connection>();

            Connection c = new Connection();
            c.Id = 1;
            c.Type = ConnectionType.OutgoingCall;
            c.Date = DateTime.Now;
            c.IOTarget = "79507441700";
            c.Cost = 0.00m;
            c.Value = 53;

            con.Add(c);

            c.Id = 2;
            c.Type = ConnectionType.OutgoingCall;
            c.Date = DateTime.Now;
            c.IOTarget = "79507441700";
            c.Cost = 0.00m;
            c.Value = 125;

            con.Add(c);

            startDate = DateTime.Now;

            BinaryFormatter formatter = new BinaryFormatter();
            using (FileStream fs = new FileStream("connections.db", FileMode.Create))
            {
                // поток архивации
                using (DeflateStream compressionStream = new DeflateStream(fs, CompressionMode.Compress))
                {
                    formatter.Serialize(compressionStream, con);

                    MessageBox.Show("Объект сжат в DeflateStream за " + Convert.ToString(startDate - DateTime.Now));
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            startDate = DateTime.Now;

            BinaryFormatter formatter = new BinaryFormatter();

            using (FileStream fs = new FileStream("connections.db", FileMode.Open))
            {
                using (DeflateStream decompressionStream = new DeflateStream(fs, CompressionMode.Decompress))
                {
                    List<Connection> con = (List<Connection>)formatter.Deserialize(decompressionStream);

                    MessageBox.Show("Объект десериализован за " + Convert.ToString(startDate - DateTime.Now));

                    MessageBox.Show(con[0].IOTarget);
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            

            MessageBox.Show(lbase.SubscriberList.Count.ToString());

            lbase.Commit();
        }
    }
}
