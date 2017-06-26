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

namespace RogalApplication
{
    public partial class ProvincePredefinitions : Form
    {
        public ComboBox comboBoxProvince;
        public ProvincePredefinitions(ComboBox comboBoxProvince)
        {
            this.comboBoxProvince = comboBoxProvince;

            InitializeComponent();

            for (int i = 0; i < comboBoxProvince.Items.Count; i++)
            {
                listBoxProvincePredefinitions.Items.Add(comboBoxProvince.Items[i].ToString());
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
            if (textBoxPredefinition.Text != "")
            {
                listBoxProvincePredefinitions.Items.Add(textBoxPredefinition.Text);
                textBoxPredefinition.Text = "";
            }
        }

        private void ProvincePredefinitions_FormClosing(object sender, FormClosingEventArgs e)
        {
            Settings.Default.ListProvincePredefinitions.Clear();
            
            for (int i = 0; i < comboBoxProvince.Items.Count; i++)
            {
                comboBoxProvince.Items.RemoveAt(i);
            }
            for (int i = 0; i < listBoxProvincePredefinitions.Items.Count; i++)
            {
                comboBoxProvince.Items.Add(listBoxProvincePredefinitions.Items[i].ToString());
                Settings.Default.ListProvincePredefinitions.Add(listBoxProvincePredefinitions.Items[i].ToString());
            }
            Settings.Default.Save();
        }
    }
}
