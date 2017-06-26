using System;
using System.Windows.Forms;
using RogalApplication.Properties;

namespace RogalApplication
{
    public partial class TradePredefinitions : Form
    {
        public ComboBox ComboBoxTrade;
        public TradePredefinitions(ComboBox comboBoxTrade)
        {
            ComboBoxTrade = comboBoxTrade;

            InitializeComponent();

            foreach (var t in comboBoxTrade.Items)
            {
                listBoxTradePredefinitions.Items.Add(t.ToString());
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
            if (textBoxPredefinition.Text == "") return;
            listBoxTradePredefinitions.Items.Add(textBoxPredefinition.Text);
            textBoxPredefinition.Text="";
        }

  
        private void TradePredefinitions_FormClosing(object sender, FormClosingEventArgs e)
        {
            Settings.Default.ListTradePredefinitions.Clear();
            for (var i = 0; i < ComboBoxTrade.Items.Count; i++)
            {
                ComboBoxTrade.Items.RemoveAt(i);
            }
            foreach (var t in listBoxTradePredefinitions.Items)
            {
                ComboBoxTrade.Items.Add(t.ToString());
                Settings.Default.ListTradePredefinitions.Add(t.ToString());
            }
            Settings.Default.Save();
        }
    }
}
