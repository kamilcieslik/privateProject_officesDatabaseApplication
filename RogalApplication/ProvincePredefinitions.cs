using System;
using System.Windows.Forms;
using RogalApplication.Properties;

namespace RogalApplication
{
    public partial class ProvincePredefinitions : Form
    {
        public ComboBox ComboBoxProvince;
        public ProvincePredefinitions(ComboBox comboBoxProvince)
        {
            ComboBoxProvince = comboBoxProvince;

            InitializeComponent();

            foreach (var t in comboBoxProvince.Items)
            {
                listBoxProvincePredefinitions.Items.Add(t.ToString());
            }
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (listBoxProvincePredefinitions.SelectedIndex != -1)
            {
                listBoxProvincePredefinitions.Items.Remove(listBoxProvincePredefinitions.SelectedItem.ToString());
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            if (textBoxPredefinition.Text == "") return;
            listBoxProvincePredefinitions.Items.Add(textBoxPredefinition.Text);
            textBoxPredefinition.Text = "";
        }

        private void ProvincePredefinitions_FormClosing(object sender, FormClosingEventArgs e)
        {
            Settings.Default.ListProvincePredefinitions.Clear();
            
            for (var i = 0; i < ComboBoxProvince.Items.Count; i++)
            {
                ComboBoxProvince.Items.RemoveAt(i);
            }
            foreach (var t in listBoxProvincePredefinitions.Items)
            {
                ComboBoxProvince.Items.Add(t.ToString());
                Settings.Default.ListProvincePredefinitions.Add(t.ToString());
            }
            Settings.Default.Save();
        }
    }
}
