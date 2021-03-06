﻿using System;
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

            excel.Dispose();
            excel = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Convert.ToString(lbase.SubscriberList[0].ConnectionsList.Count));

            int i = 0;
            while(i < lbase.SubscriberList.Count)
            {
                int j = 0;
                string msg = lbase.SubscriberList[i].Number + "\n";
                while(j < lbase.SubscriberList[i].ConnectionsList.Count)
                {
                    msg += lbase.SubscriberList[i].ConnectionsList[j].IOTarget + "\n";

                    j++;
                }

                MessageBox.Show(Convert.ToString(lbase.SubscriberList[i].ConnectionsList.Count));

                i++;
            }
        }
    }
}
