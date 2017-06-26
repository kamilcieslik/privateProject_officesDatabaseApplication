using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RogalApplication.Properties;

namespace BazaKlientów
{
    public partial class TradePredefinitions : Form
    {
        public ComboBox comboBoxTrade;
        public TradePredefinitions(ComboBox comboBoxTrade)
        {
            this.comboBoxTrade = comboBoxTrade;

            InitializeComponent();

            for (int i = 0; i < comboBoxTrade.Items.Count; i++)
            {
                listBoxTradePredefinitions.Items.Add(comboBoxTrade.Items[i].ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBoxTradePredefinitions.SelectedIndex != -1)
            {
                listBoxTradePredefinitions.Items.Remove(listBoxTradePredefinitions.SelectedItem.ToString());
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            if (textBoxPredefinition.Text != "")
            {
                listBoxTradePredefinitions.Items.Add(textBoxPredefinition.Text);
                textBoxPredefinition.Text="";
            }
        }

  
        private void TradePredefinitions_FormClosing(object sender, FormClosingEventArgs e)
        {
            Settings.Default.ListTradePredefinitions.Clear();
            for (int i = 0; i < comboBoxTrade.Items.Count; i++)
            {
                comboBoxTrade.Items.RemoveAt(i);
            }
            for (int i = 0; i < listBoxTradePredefinitions.Items.Count; i++)
            {
                comboBoxTrade.Items.Add(listBoxTradePredefinitions.Items[i].ToString());
                Settings.Default.ListTradePredefinitions.Add(listBoxTradePredefinitions.Items[i].ToString());
            }
            Settings.Default.Save();
        }
    }
}
