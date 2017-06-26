using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity.Core;
using System.Data.Entity.Validation;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BazaKlientów.Model;
using BazaKlientów.Repository.Commands;
using BazaKlientów.Repository.Queries;
using Microsoft.Vbe.Interop;
using RogalApplication;
using RogalApplication.Properties;

namespace BazaKlientów
{
    public partial class MainFrame : Form
    {
        private CustomerDatabaseContext context;
        private readonly ReadRepository<Customer> readCustomerRepository;
        private readonly WriteRepository<Customer> writeCustomerRepository;
        private Boolean searchingOn = false;
        private Boolean editingOn = false;

        public MainFrame()
        {
            context = new CustomerDatabaseContext();
            readCustomerRepository = new ReadRepository<Customer>(context);
            writeCustomerRepository = new WriteRepository<Customer>(context);
            InitializeComponent();
            ShowCustomers();
        }

        public void ClearTextBoxes()
        {
            textBoxName.Text = "";
            textBoxPhoneNumber.Text = "";
            textBoxEmail.Text = "";
            comboBoxTrade.Text = "";
            comboBoxProvince.Text = "";
            textBoxCity.Text = "";
            textBoxPostalCode.Text = "";
            textBoxStreet.Text = "";
            textBoxDescription.Text = "";
            textBoxComments.Text = "";
            radioButtonEditionOff.Checked = true;
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {

            Customer customer = new Customer()
            {
                Name = textBoxName.Text,
                PhoneNumber = textBoxPhoneNumber.Text,
                Email = textBoxEmail.Text,
                Trade = comboBoxTrade.Text,
                Province = comboBoxProvince.Text,
                City = textBoxCity.Text,
                PostalCode = textBoxPostalCode.Text,
                Street = textBoxStreet.Text,
                Description = textBoxDescription.Text,
                Comments = textBoxComments.Text

            };
            try
            {
                writeCustomerRepository.Create(customer);
                ShowCustomers();
                radioButtonEditionOff.Checked = true;
                ClearTextBoxes();
                dataGridViewCustomers.ClearSelection();
            }
            catch (DbEntityValidationException dbEx)
            {
                string errors = "";
                errors = "Błędy zapisu: " + Environment.NewLine;
                foreach (var validationErrors in dbEx.EntityValidationErrors)
                {
                    foreach (var validationError in validationErrors.ValidationErrors)
                    {

                        errors = errors + "Dotyczy: " + validationError.PropertyName + Environment.NewLine + "Rodzaj: " +
                                 validationError.ErrorMessage + Environment.NewLine + Environment.NewLine;
                    }
                }
                MessageBox.Show(errors);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd zapisu: " + ex.Message);
            }

        }

        public void ShowCustomers()
        {
            dataGridViewCustomers.DataSource = null;
            dataGridViewCustomers.DataSource = readCustomerRepository
                .GetAll()
                .Select(x => new
                {
                    ID = x.ID,
                    Nazwa = x.Name,
                    Telefon = x.PhoneNumber,
                    Email = x.Email,
                    Branża = x.Trade,
                    Województwo = x.Province,
                    Miasto = x.City,
                    Kod_pocztowy = x.PostalCode,
                    Ulica = x.Street
                }).ToList();
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (dataGridViewCustomers.SelectedRows.Count > 0)
            {
                var customerId = dataGridViewCustomers.SelectedRows[0].Cells["ID"].Value.ToString();
                writeCustomerRepository.Delete(readCustomerRepository.GetById(int.Parse(customerId)));
                ShowCustomers();
                radioButtonEditionOff.Checked = true;
                ClearTextBoxes();
            }
            else
            {
                MessageBox.Show("Aby usunąc obiekt w pierwszej kolejności musisz jakiś wybrać.");
            }
        }

        private void buttonEdit_Click(object sender, EventArgs e)
        {
            if (dataGridViewCustomers.SelectedRows.Count > 0)
            {
                if (searchingOn == false)
                {
                    string customerId = dataGridViewCustomers.SelectedRows[0].Cells["ID"].Value.ToString();
                    try
                    {
                        Customer customer = new Customer()
                        {

                            ID = int.Parse(customerId),
                            Name = textBoxName.Text,
                            PhoneNumber = textBoxPhoneNumber.Text,
                            Email = textBoxEmail.Text,
                            Trade = comboBoxTrade.Text,
                            Province = comboBoxProvince.Text,
                            City = textBoxCity.Text,
                            PostalCode = textBoxPostalCode.Text,
                            Street = textBoxStreet.Text,
                            Description = textBoxDescription.Text,
                            Comments = textBoxComments.Text


                        };
                        writeCustomerRepository.Edit(readCustomerRepository.GetById(int.Parse(customerId)), customer);
                        ShowCustomers();
                        radioButtonEditionOff.Checked = true;
                        dataGridViewCustomers.ClearSelection();
                    }
                    catch (DbEntityValidationException dbEx)
                    {
                        string errors = "";
                        errors = "Błędy edycji: " + Environment.NewLine;
                        foreach (var validationErrors in dbEx.EntityValidationErrors)
                        {
                            foreach (var validationError in validationErrors.ValidationErrors)
                            {

                                errors = errors + "Dotyczy: " + validationError.PropertyName + Environment.NewLine +
                                         "Rodzaj: " +
                                         validationError.ErrorMessage + Environment.NewLine + Environment.NewLine;
                            }
                        }
                        MessageBox.Show(errors);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Błąd edycji: " + ex.Message);
                    }
                }
                else
                {
                    string customerId = dataGridViewCustomers.SelectedRows[0].Cells["ID"].Value.ToString();
                    string customerName = dataGridViewCustomers.SelectedRows[0].Cells["Nazwa"].Value.ToString();
                    string customerPhoneNumber = dataGridViewCustomers.SelectedRows[0].Cells["Telefon"].Value.ToString();
                    string customerEmail = dataGridViewCustomers.SelectedRows[0].Cells["Email"].Value.ToString();
                    string customerTrade = dataGridViewCustomers.SelectedRows[0].Cells["Branża"].Value.ToString();
                    string customerProvince =
                        dataGridViewCustomers.SelectedRows[0].Cells["Województwo"].Value.ToString();
                    string customerCity = dataGridViewCustomers.SelectedRows[0].Cells["Miasto"].Value.ToString();
                    string customerPostalCode =
                        dataGridViewCustomers.SelectedRows[0].Cells["Kod_pocztowy"].Value.ToString();
                    string customerStreet = dataGridViewCustomers.SelectedRows[0].Cells["Ulica"].Value.ToString();
                    try
                    {
                        Customer customer = new Customer()
                        {

                            ID = int.Parse(customerId),
                            Name = customerName,
                            PhoneNumber = customerPhoneNumber,
                            Email = customerEmail,
                            Trade = customerTrade,
                            Province = customerProvince,
                            City = customerCity,
                            PostalCode = customerPostalCode,
                            Street = customerStreet,
                            Description = textBoxDescription.Text,
                            Comments = textBoxComments.Text

                        };
                        writeCustomerRepository.Edit(readCustomerRepository.GetById(int.Parse(customerId)), customer);
                        radioButtonEditionOff.Checked = true;
                        dataGridViewCustomers.ClearSelection();
                    }
                    catch (DbEntityValidationException dbEx)
                    {
                        string errors = "";
                        errors = "Błędy edycji: " + Environment.NewLine;
                        foreach (var validationErrors in dbEx.EntityValidationErrors)
                        {
                            foreach (var validationError in validationErrors.ValidationErrors)
                            {

                                errors = errors + "Dotyczy: " + validationError.PropertyName + Environment.NewLine +
                                         "Rodzaj: " +
                                         validationError.ErrorMessage + Environment.NewLine + Environment.NewLine;
                            }
                        }
                        MessageBox.Show(errors);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Błąd edycji: " + ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show("Aby edytować obiekt w pierwszej kolejności musisz jakiś wybrać.");
            }
        }

        private void radioButtonEditionOn_CheckedChanged(object sender, EventArgs e)
        {
            editingOn = true;
            buttonEdit.Enabled = true;
            buttonAdd.Enabled = true;
            buttonDelete.Enabled = true;
            buttonClear.Enabled = true;
            textBoxDescription.BorderStyle = BorderStyle.Fixed3D;
            textBoxComments.BorderStyle = BorderStyle.Fixed3D;
            textBoxName.ReadOnly = false;
            textBoxPhoneNumber.ReadOnly = false;
            textBoxEmail.ReadOnly = false;
            comboBoxTrade.Enabled = true;
            comboBoxProvince.Enabled = true;
            textBoxCity.ReadOnly = false;
            textBoxPostalCode.ReadOnly = false;
            textBoxStreet.ReadOnly = false;
            textBoxDescription.ReadOnly = false;
            textBoxComments.ReadOnly = false;
        }

        private void radioButtonEditionOff_CheckedChanged(object sender, EventArgs e)
        {
            if (searchingOn == true)
            {
                buttonEdit.Enabled = false;
                buttonAdd.Enabled = false;
                buttonDelete.Enabled = false;
                textBoxDescription.BorderStyle = BorderStyle.None;
                textBoxComments.BorderStyle = BorderStyle.None;
                textBoxDescription.ReadOnly = true;
                textBoxComments.ReadOnly = true;
            }
            else
            {
                editingOn = false;
                buttonEdit.Enabled = false;
                buttonAdd.Enabled = false;
                buttonDelete.Enabled = false;
                buttonClear.Enabled = false;
                textBoxDescription.BorderStyle = BorderStyle.None;
                textBoxComments.BorderStyle = BorderStyle.None;
                textBoxName.ReadOnly = true;
                textBoxPhoneNumber.ReadOnly = true;
                textBoxEmail.ReadOnly = true;
                comboBoxTrade.Enabled = false;
                comboBoxProvince.Enabled = false;
                textBoxCity.ReadOnly = true;
                textBoxPostalCode.ReadOnly = true;
                textBoxStreet.ReadOnly = true;
                textBoxDescription.ReadOnly = true;
                textBoxComments.ReadOnly = true;
            }

        }


        private void dataGridViewCustomers_Click(object sender, EventArgs e)
        {
            try
            {
                if (searchingOn == false)
                {

                    int id = (int)dataGridViewCustomers.SelectedRows[0].Cells["ID"].Value;
                    if (readCustomerRepository.GetById(id).Description != null)
                    {
                        textBoxDescription.Text = readCustomerRepository.GetById(id).Description;
                    }
                    else
                    {
                        textBoxDescription.Text = "";
                    }
                    if (readCustomerRepository.GetById(id).Comments != null)
                    {
                        textBoxComments.Text = readCustomerRepository.GetById(id).Comments;
                    }
                    else
                    {
                        textBoxComments.Text = "";
                    }
                    if (dataGridViewCustomers.SelectedRows[0].Cells["Nazwa"].Value != null)
                    {
                        textBoxName.Text = dataGridViewCustomers.SelectedRows[0].Cells["Nazwa"].Value.ToString();
                    }
                    else
                    {
                        textBoxName.Text = "";
                    }
                    if (dataGridViewCustomers.SelectedRows[0].Cells["Telefon"].Value != null)
                    {
                        textBoxPhoneNumber.Text =
                            dataGridViewCustomers.SelectedRows[0].Cells["Telefon"].Value.ToString();
                    }
                    else
                    {
                        textBoxPhoneNumber.Text = "";
                    }
                    if (dataGridViewCustomers.SelectedRows[0].Cells["Email"].Value != null)
                    {
                        textBoxEmail.Text = dataGridViewCustomers.SelectedRows[0].Cells["Email"].Value.ToString();
                    }
                    else
                    {
                        textBoxEmail.Text = "";
                    }
                    if (dataGridViewCustomers.SelectedRows[0].Cells["Branża"].Value != null)
                    {
                        comboBoxTrade.Text = dataGridViewCustomers.SelectedRows[0].Cells["Branża"].Value.ToString();
                    }
                    else
                    {
                        comboBoxTrade.Text = "";
                    }
                    if (dataGridViewCustomers.SelectedRows[0].Cells["Miasto"].Value != null)
                    {
                        textBoxCity.Text = dataGridViewCustomers.SelectedRows[0].Cells["Miasto"].Value.ToString();
                    }
                    else
                    {
                        textBoxCity.Text = "";
                    }
                    if (dataGridViewCustomers.SelectedRows[0].Cells["Kod_pocztowy"].Value != null)
                    {
                        textBoxPostalCode.Text =
                            dataGridViewCustomers.SelectedRows[0].Cells["Kod_pocztowy"].Value.ToString();
                    }
                    else
                    {
                        textBoxPostalCode.Text = "";
                    }
                    if (dataGridViewCustomers.SelectedRows[0].Cells["Ulica"].Value != null)
                    {
                        textBoxStreet.Text = dataGridViewCustomers.SelectedRows[0].Cells["Ulica"].Value.ToString();
                    }
                    else
                    {
                        textBoxStreet.Text = "";
                    }
                    if (dataGridViewCustomers.SelectedRows[0].Cells["Województwo"].Value != null)
                    {
                        comboBoxProvince.Text =
                            dataGridViewCustomers.SelectedRows[0].Cells["Województwo"].Value.ToString();
                    }
                    else
                    {
                        comboBoxProvince.Text = "";
                    }

                }
                else
                {
                    int id = (int)dataGridViewCustomers.SelectedRows[0].Cells["ID"].Value;
                    if (readCustomerRepository.GetById(id).Description != null)
                    {
                        textBoxDescription.Text = readCustomerRepository.GetById(id).Description;
                    }
                    else
                    {
                        textBoxDescription.Text = "";
                    }
                    if (readCustomerRepository.GetById(id).Comments != null)
                    {
                        textBoxComments.Text = readCustomerRepository.GetById(id).Comments;
                    }
                    else
                    {
                        textBoxComments.Text = "";
                    }
                }
            }
            catch (Exception)
            {
                ClearTextBoxes();
            }
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            textBoxName.Text = "";
            textBoxPhoneNumber.Text = "";
            textBoxEmail.Text = "";
            comboBoxTrade.Text = "";
            comboBoxProvince.Text = "";
            textBoxCity.Text = "";
            textBoxPostalCode.Text = "";
            textBoxStreet.Text = "";
            textBoxDescription.Text = "";
            textBoxComments.Text = "";
        }

        private void radioButtonSearchingOff_CheckedChanged(object sender, EventArgs e)
        {
            radioButtonEditionOn.Enabled = true;
            radioButtonEditionOff.Enabled = true;

            textBoxDescription.ReadOnly = true;
            textBoxComments.ReadOnly = true;
            buttonEdit.Enabled = false;

            searchingOn = false;
            buttonClear.Enabled = false;
            textBoxName.ReadOnly = true;
            textBoxPhoneNumber.ReadOnly = true;
            textBoxEmail.ReadOnly = true;
            textBoxPostalCode.ReadOnly = true;
            comboBoxTrade.Enabled = false;
            comboBoxProvince.Enabled = false;
            textBoxCity.ReadOnly = true;
            textBoxStreet.ReadOnly = true;
            textBoxName.Text = "";
            textBoxPhoneNumber.Text = "";
            textBoxEmail.Text = "";
            textBoxPostalCode.Text = "";
            comboBoxTrade.Text = "";
            comboBoxProvince.Text = "";
            textBoxCity.Text = "";
            textBoxStreet.Text = "";
            textBoxDescription.Text = "";
            textBoxComments.Text = "";
        }

        private void radioButtonSearchingOn_CheckedChanged(object sender, EventArgs e)
        {
            radioButtonEditionOn.Enabled = false;
            radioButtonEditionOff.Enabled = false;
            radioButtonEditionOff.Checked = true;

            textBoxDescription.ReadOnly = false;
            textBoxComments.ReadOnly = false;
            buttonEdit.Enabled = true;

            searchingOn = true;
            buttonClear.Enabled = true;
            textBoxName.ReadOnly = false;
            textBoxPhoneNumber.ReadOnly = false;
            textBoxEmail.ReadOnly = false;
            textBoxPostalCode.ReadOnly = false;
            comboBoxTrade.Enabled = true;
            comboBoxProvince.Enabled = true;
            textBoxCity.ReadOnly = false;
            textBoxStreet.ReadOnly = false;
            textBoxName.Text = "";
            textBoxPhoneNumber.Text = "";
            textBoxEmail.Text = "";
            textBoxPostalCode.Text = "";
            comboBoxTrade.Text = "";
            comboBoxProvince.Text = "";
            textBoxCity.Text = "";
            textBoxStreet.Text = "";
            textBoxDescription.Text = "";
            textBoxComments.Text = "";
        }


        private void textBoxes_TextChangedForDynamicSearching(object sender, EventArgs e)
        {
            if (searchingOn == true)
            {
                string buttonName;
                string option;
                if (sender.GetType().ToString() == "System.Windows.Forms.TextBox")
                {
                    TextBox textBox = (TextBox)sender;
                    buttonName = textBox.Name;
                    option = textBox.Text;
                }
                else
                {
                    ComboBox comboBox = (ComboBox)sender;
                    buttonName = comboBox.Name;
                    option = comboBox.Text;
                }

                switch (buttonName)
                {
                    case "textBoxName":
                        dataGridViewCustomers.DataSource = null;
                        dataGridViewCustomers.DataSource =
                            readCustomerRepository.GetAll().Where(x => x.Name.Contains(option)).Select(x => new
                            {
                                ID = x.ID,
                                Nazwa = x.Name,
                                Telefon = x.PhoneNumber,
                                Email = x.Email,
                                Branża = x.Trade,
                                Województwo = x.Province,
                                Miasto = x.City,
                                Kod_pocztowy = x.PostalCode,
                                Ulica = x.Street
                            }).ToList();
                        break;
                    case "textBoxPhoneNumber":
                        dataGridViewCustomers.DataSource = null;
                        dataGridViewCustomers.DataSource =
                            readCustomerRepository.GetAll().Where(x => x.PhoneNumber.Contains(option)).Select(x => new
                            {
                                ID = x.ID,
                                Nazwa = x.Name,
                                Telefon = x.PhoneNumber,
                                Email = x.Email,
                                Branża = x.Trade,
                                Województwo = x.Province,
                                Miasto = x.City,
                                Kod_pocztowy = x.PostalCode,
                                Ulica = x.Street
                            }).ToList();
                        break;
                    case "textBoxEmail":
                        dataGridViewCustomers.DataSource = null;
                        dataGridViewCustomers.DataSource =
                            readCustomerRepository.GetAll().Where(x => x.Email.Contains(option)).Select(x => new
                            {
                                ID = x.ID,
                                Nazwa = x.Name,
                                Telefon = x.PhoneNumber,
                                Email = x.Email,
                                Branża = x.Trade,
                                Województwo = x.Province,
                                Miasto = x.City,
                                Kod_pocztowy = x.PostalCode,
                                Ulica = x.Street
                            }).ToList();
                        break;
                    case "comboBoxTrade":
                        dataGridViewCustomers.DataSource = null;
                        dataGridViewCustomers.DataSource =
                            readCustomerRepository.GetAll().Where(x => x.Trade.Contains(option)).Select(x => new
                            {
                                ID = x.ID,
                                Nazwa = x.Name,
                                Telefon = x.PhoneNumber,
                                Email = x.Email,
                                Branża = x.Trade,
                                Województwo = x.Province,
                                Miasto = x.City,
                                Kod_pocztowy = x.PostalCode,
                                Ulica = x.Street
                            }).ToList();
                        break;
                    case "comboBoxProvince":
                        dataGridViewCustomers.DataSource = null;
                        dataGridViewCustomers.DataSource =
                            readCustomerRepository.GetAll().Where(x => x.Province.Contains(option)).Select(x => new
                            {
                                ID = x.ID,
                                Nazwa = x.Name,
                                Telefon = x.PhoneNumber,
                                Email = x.Email,
                                Branża = x.Trade,
                                Województwo = x.Province,
                                Miasto = x.City,
                                Kod_pocztowy = x.PostalCode,
                                Ulica = x.Street
                            }).ToList();
                        break;
                    case "textBoxCity":
                        dataGridViewCustomers.DataSource = null;
                        dataGridViewCustomers.DataSource =
                            readCustomerRepository.GetAll().Where(x => x.City.Contains(option)).Select(x => new
                            {
                                ID = x.ID,
                                Nazwa = x.Name,
                                Telefon = x.PhoneNumber,
                                Email = x.Email,
                                Branża = x.Trade,
                                Województwo = x.Province,
                                Miasto = x.City,
                                Kod_pocztowy = x.PostalCode,
                                Ulica = x.Street
                            }).ToList();
                        break;
                    case "textBoxPostalCode":
                        dataGridViewCustomers.DataSource = null;
                        dataGridViewCustomers.DataSource =
                            readCustomerRepository.GetAll().Where(x => x.PostalCode.Contains(option)).Select(x => new
                            {
                                ID = x.ID,
                                Nazwa = x.Name,
                                Telefon = x.PhoneNumber,
                                Email = x.Email,
                                Branża = x.Trade,
                                Województwo = x.Province,
                                Miasto = x.City,
                                Kod_pocztowy = x.PostalCode,
                                Ulica = x.Street
                            }).ToList();
                        break;
                    case "textBoxStreet":
                        dataGridViewCustomers.DataSource = null;
                        dataGridViewCustomers.DataSource =
                            readCustomerRepository.GetAll().Where(x => x.Street.Contains(option)).Select(x => new
                            {
                                ID = x.ID,
                                Nazwa = x.Name,
                                Telefon = x.PhoneNumber,
                                Email = x.Email,
                                Branża = x.Trade,
                                Województwo = x.Province,
                                Miasto = x.City,
                                Kod_pocztowy = x.PostalCode,
                                Ulica = x.Street
                            }).ToList();
                        break;
                }
            }
        }

        private void predefinicjeBranżToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            TradePredefinitions tradePredefinitions = new TradePredefinitions(comboBoxTrade);
            this.Visible = false;
            tradePredefinitions.ShowDialog();
            if (tradePredefinitions.IsAccessible == false) Visible = true;

            comboBoxTrade.Items.Clear();
            foreach (var item in Settings.Default.ListTradePredefinitions)
            {
                comboBoxTrade.Items.Add(item);
            }
        }

        private void predefinicjeWojewództwToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ProvincePredefinitions provincePredefinitions = new ProvincePredefinitions(comboBoxProvince);
            this.Visible = false;
            provincePredefinitions.ShowDialog();
            if (provincePredefinitions.IsAccessible == false) Visible = true;

            comboBoxProvince.Items.Clear();
            foreach (var item in Settings.Default.ListProvincePredefinitions)
            {
                comboBoxProvince.Items.Add(item);
            }
        }

        private void copyAlltoClipboard()
        {
            dataGridViewAllDataExportToExcel.DataSource = null;
            dataGridViewAllDataExportToExcel.DataSource = readCustomerRepository
                .GetAll()
                .Select(x => new
                {
                    ID = x.ID,
                    Nazwa = x.Name,
                    Telefon = x.PhoneNumber,
                    Email = x.Email,
                    Branża = x.Trade,
                    Województwo = x.Province,
                    Miasto = x.City,
                    Kod_pocztowy = x.PostalCode,
                    Ulica = x.Street,
                    Opis = x.Description,
                    Uwagi = x.Comments
                }).ToList();

            dataGridViewAllDataExportToExcel.SelectAll();
            DataObject dataObj = dataGridViewAllDataExportToExcel.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void copyAlltoClipboardWithoutCommentsAndDescriptions()
        {
            dataGridViewAllDataExportToExcel.DataSource = null;
            dataGridViewAllDataExportToExcel.DataSource = readCustomerRepository
                .GetAll()
                .Select(x => new
                {
                    ID = x.ID,
                    Nazwa = x.Name,
                    Telefon = x.PhoneNumber,
                    Email = x.Email,
                    Branża = x.Trade,
                    Województwo = x.Province,
                    Miasto = x.City,
                    Kod_pocztowy = x.PostalCode,
                    Ulica = x.Street
                }).ToList();

            dataGridViewAllDataExportToExcel.SelectAll();
            DataObject dataObj = dataGridViewAllDataExportToExcel.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void eksportujDoExcelaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult =
                MessageBox.Show("Czy obiekty w tabeli mają posiadać kolumny 'Opis' oraz 'Uwagi'?",
                    "Wybierz wersję tabeli.", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
            {
                copyAlltoClipboard();
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Microsoft.Office.Interop.Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
            else if (dialogResult == DialogResult.No)
            {
                copyAlltoClipboardWithoutCommentsAndDescriptions();
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Microsoft.Office.Interop.Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
        }

        private void zamknijToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void dataGridViewCustomers_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string option = dataGridViewCustomers.Columns[e.ColumnIndex].Name;

            switch (option)
            {
                case "ID":
                    dataGridViewCustomers.DataSource =
                        readCustomerRepository.GetAll().OrderBy(x => x.ID).Select(x => new
                        {
                            ID = x.ID,
                            Nazwa = x.Name,
                            Telefon = x.PhoneNumber,
                            Email = x.Email,
                            Branża = x.Trade,
                            Województwo = x.Province,
                            Miasto = x.City,
                            Kod_pocztowy = x.PostalCode,
                            Ulica = x.Street
                        }).ToList();
                    break;
                case "Nazwa":
                    dataGridViewCustomers.DataSource =
                        readCustomerRepository.GetAll().OrderBy(x => x.Name).Select(x => new
                        {
                            ID = x.ID,
                            Nazwa = x.Name,
                            Telefon = x.PhoneNumber,
                            Email = x.Email,
                            Branża = x.Trade,
                            Województwo = x.Province,
                            Miasto = x.City,
                            Kod_pocztowy = x.PostalCode,
                            Ulica = x.Street
                        }).ToList();
                    break;
                case "Telefon":
                    dataGridViewCustomers.DataSource =
                        readCustomerRepository.GetAll().OrderBy(x => x.PhoneNumber).Select(x => new
                        {
                            ID = x.ID,
                            Nazwa = x.Name,
                            Telefon = x.PhoneNumber,
                            Email = x.Email,
                            Branża = x.Trade,
                            Województwo = x.Province,
                            Miasto = x.City,
                            Kod_pocztowy = x.PostalCode,
                            Ulica = x.Street
                        }).ToList();
                    break;
                case "Email":
                    dataGridViewCustomers.DataSource =
                        readCustomerRepository.GetAll().OrderBy(x => x.Email).Select(x => new
                        {
                            ID = x.ID,
                            Nazwa = x.Name,
                            Telefon = x.PhoneNumber,
                            Email = x.Email,
                            Branża = x.Trade,
                            Województwo = x.Province,
                            Miasto = x.City,
                            Kod_pocztowy = x.PostalCode,
                            Ulica = x.Street
                        }).ToList();
                    break;
                case "Branża":
                    dataGridViewCustomers.DataSource =
                        readCustomerRepository.GetAll().OrderBy(x => x.Trade).Select(x => new
                        {
                            ID = x.ID,
                            Nazwa = x.Name,
                            Telefon = x.PhoneNumber,
                            Email = x.Email,
                            Branża = x.Trade,
                            Województwo = x.Province,
                            Miasto = x.City,
                            Kod_pocztowy = x.PostalCode,
                            Ulica = x.Street
                        }).ToList();
                    break;
                case "Województwo":
                    dataGridViewCustomers.DataSource =
                        readCustomerRepository.GetAll().OrderBy(x => x.Province).Select(x => new
                        {
                            ID = x.ID,
                            Nazwa = x.Name,
                            Telefon = x.PhoneNumber,
                            Email = x.Email,
                            Branża = x.Trade,
                            Województwo = x.Province,
                            Miasto = x.City,
                            Kod_pocztowy = x.PostalCode,
                            Ulica = x.Street
                        }).ToList();
                    break;
                case "Miasto":
                    dataGridViewCustomers.DataSource =
                        readCustomerRepository.GetAll().OrderBy(x => x.City).Select(x => new
                        {
                            ID = x.ID,
                            Nazwa = x.Name,
                            Telefon = x.PhoneNumber,
                            Email = x.Email,
                            Branża = x.Trade,
                            Województwo = x.Province,
                            Miasto = x.City,
                            Kod_pocztowy = x.PostalCode,
                            Ulica = x.Street
                        }).ToList();
                    break;
                case "Kod_pocztowy":
                    dataGridViewCustomers.DataSource =
                        readCustomerRepository.GetAll().OrderBy(x => x.PostalCode).Select(x => new
                        {
                            ID = x.ID,
                            Nazwa = x.Name,
                            Telefon = x.PhoneNumber,
                            Email = x.Email,
                            Branża = x.Trade,
                            Województwo = x.Province,
                            Miasto = x.City,
                            Kod_pocztowy = x.PostalCode,
                            Ulica = x.Street
                        }).ToList();
                    break;
                case "Ulica":
                    dataGridViewCustomers.DataSource =
                        readCustomerRepository.GetAll().OrderBy(x => x.Street).Select(x => new
                        {
                            ID = x.ID,
                            Nazwa = x.Name,
                            Telefon = x.PhoneNumber,
                            Email = x.Email,
                            Branża = x.Trade,
                            Województwo = x.Province,
                            Miasto = x.City,
                            Kod_pocztowy = x.PostalCode,
                            Ulica = x.Street
                        }).ToList();
                    break;
            }
        }

        private void MainFrame_FormClosing(object sender, FormClosingEventArgs e)
        {
            Settings.Default.Column1Width = dataGridViewCustomers.Columns[0].Width;
            Settings.Default.Column2Width = dataGridViewCustomers.Columns[1].Width;
            Settings.Default.Column3Width = dataGridViewCustomers.Columns[2].Width;
            Settings.Default.Column4Width = dataGridViewCustomers.Columns[3].Width;
            Settings.Default.Column5Width = dataGridViewCustomers.Columns[4].Width;
            Settings.Default.Column6Width = dataGridViewCustomers.Columns[5].Width;
            Settings.Default.Column7Width = dataGridViewCustomers.Columns[6].Width;
            Settings.Default.Column8Width = dataGridViewCustomers.Columns[7].Width;
            Settings.Default.Column9Width = dataGridViewCustomers.Columns[8].Width;
            Settings.Default.FormSize = this.Size;
            Settings.Default.Save();
        }
        private void MainFrame_Load(object sender, EventArgs e)
        {
            try
            {
                foreach (var item in Settings.Default.ListTradePredefinitions)
                {
                    comboBoxTrade.Items.Add(item);
                }
            }
            catch (Exception)
            {
                List<String> tmpList = new List<string>();
                Settings.Default.ListTradePredefinitions = tmpList;
                Settings.Default.Save();
            }
            try
            {
                foreach (var item in Settings.Default.ListProvincePredefinitions)
                {
                    comboBoxProvince.Items.Add(item);
                }
            }
            catch (Exception)
            {
                List<String> tmpList = new List<string>();
                Settings.Default.ListProvincePredefinitions = tmpList;
                Settings.Default.Save();
            }
            if ((Settings.Default.FormSize.Width != 0) && (Settings.Default.FormSize.Height != 0))
            {
                this.Size = Settings.Default.FormSize;
            }
            if (Settings.Default.Column1Width != 0)
            {
                dataGridViewCustomers.Columns[0].Width = Settings.Default.Column1Width;
            }
            if (Settings.Default.Column2Width != 0)
            {
                dataGridViewCustomers.Columns[1].Width = Settings.Default.Column2Width;
            }
            if (Settings.Default.Column3Width != 0)
            {
                dataGridViewCustomers.Columns[2].Width = Settings.Default.Column3Width;
            }
            if (Settings.Default.Column4Width != 0)
            {
                dataGridViewCustomers.Columns[3].Width = Settings.Default.Column4Width;
            }
            if (Settings.Default.Column5Width != 0)
            {
                dataGridViewCustomers.Columns[4].Width = Settings.Default.Column5Width;
            }
            if (Settings.Default.Column6Width != 0)
            {
                dataGridViewCustomers.Columns[5].Width = Settings.Default.Column6Width;
            }
            if (Settings.Default.Column7Width != 0)
            {
                dataGridViewCustomers.Columns[6].Width = Settings.Default.Column7Width;
            }
            if (Settings.Default.Column8Width != 0)
            {
                dataGridViewCustomers.Columns[7].Width = Settings.Default.Column8Width;
            }
            if (Settings.Default.Column9Width != 0)
            {
                dataGridViewCustomers.Columns[8].Width = Settings.Default.Column9Width;
            }
        }
    }
}
