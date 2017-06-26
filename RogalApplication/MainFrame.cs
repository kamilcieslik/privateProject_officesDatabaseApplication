using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.Linq;
using System.Windows.Forms;
using BazaKlientów;
using RogalApplication.Model;
using RogalApplication.Properties;
using RogalApplication.Repository.Commands;
using RogalApplication.Repository.Queries;
using TextBox = System.Windows.Forms.TextBox;

namespace RogalApplication
{
    public partial class MainFrame : Form
    {
        private readonly ReadRepository<Customer> _readCustomerRepository;
        private readonly WriteRepository<Customer> _writeCustomerRepository;
        private bool _searchingOn = false;
        private bool _editingOn = false;

        public MainFrame()
        {
            var context = new CustomerDatabaseContext();
            _readCustomerRepository = new ReadRepository<Customer>(context);
            _writeCustomerRepository = new WriteRepository<Customer>(context);
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
            var customer = new Customer()
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
                _writeCustomerRepository.Create(customer);
                ShowCustomers();
                radioButtonEditionOff.Checked = true;
                ClearTextBoxes();
                dataGridViewCustomers.ClearSelection();
            }
            catch (DbEntityValidationException dbEx)
            {
                var errors = "";
                errors = "Błędy zapisu: " + Environment.NewLine;
                errors = dbEx.EntityValidationErrors.SelectMany(validationErrors => validationErrors.ValidationErrors).Aggregate(errors, (current, validationError) => current + "Dotyczy: " + validationError.PropertyName + Environment.NewLine + "Rodzaj: " + validationError.ErrorMessage + Environment.NewLine + Environment.NewLine);
                MessageBox.Show(errors);
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Błąd zapisu: " + ex.Message);
            }
        }

        public void ShowCustomers()
        {
            dataGridViewCustomers.DataSource = null;
            dataGridViewCustomers.DataSource = _readCustomerRepository
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
                _writeCustomerRepository.Delete(_readCustomerRepository.GetById(int.Parse(customerId)));
                ShowCustomers();
                radioButtonEditionOff.Checked = true;
                ClearTextBoxes();
            }
            else
            {
                MessageBox.Show(@"Aby usunąc obiekt w pierwszej kolejności musisz jakiś wybrać.");
            }
        }

        private void buttonEdit_Click(object sender, EventArgs e)
        {
            if (dataGridViewCustomers.SelectedRows.Count > 0)
            {
                if (_searchingOn == false)
                {
                    var customerId = dataGridViewCustomers.SelectedRows[0].Cells["ID"].Value.ToString();
                    try
                    {
                        var customer = new Customer()
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
                        _writeCustomerRepository.Edit(_readCustomerRepository.GetById(int.Parse(customerId)), customer);
                        ShowCustomers();
                        radioButtonEditionOff.Checked = true;
                        dataGridViewCustomers.ClearSelection();
                    }
                    catch (DbEntityValidationException dbEx)
                    {
                        var errors = "";
                        errors = "Błędy edycji: " + Environment.NewLine;
                        errors = dbEx.EntityValidationErrors.SelectMany(validationErrors => validationErrors.ValidationErrors).Aggregate(errors, (current, validationError) => current + "Dotyczy: " + validationError.PropertyName + Environment.NewLine + "Rodzaj: " + validationError.ErrorMessage + Environment.NewLine + Environment.NewLine);
                        MessageBox.Show(errors);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(@"Błąd edycji: " + ex.Message);
                    }
                }
                else
                {
                    var customerId = dataGridViewCustomers.SelectedRows[0].Cells["ID"].Value.ToString();
                    var customerName = dataGridViewCustomers.SelectedRows[0].Cells["Nazwa"].Value.ToString();
                    var customerPhoneNumber = dataGridViewCustomers.SelectedRows[0].Cells["Telefon"].Value.ToString();
                    var customerEmail = dataGridViewCustomers.SelectedRows[0].Cells["Email"].Value.ToString();
                    var customerTrade = dataGridViewCustomers.SelectedRows[0].Cells["Branża"].Value.ToString();
                    var customerProvince =
                        dataGridViewCustomers.SelectedRows[0].Cells["Województwo"].Value.ToString();
                    var customerCity = dataGridViewCustomers.SelectedRows[0].Cells["Miasto"].Value.ToString();
                    var customerPostalCode =
                        dataGridViewCustomers.SelectedRows[0].Cells["Kod_pocztowy"].Value.ToString();
                    var customerStreet = dataGridViewCustomers.SelectedRows[0].Cells["Ulica"].Value.ToString();
                    try
                    {
                        var customer = new Customer()
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
                        _writeCustomerRepository.Edit(_readCustomerRepository.GetById(int.Parse(customerId)), customer);
                        radioButtonEditionOff.Checked = true;
                        dataGridViewCustomers.ClearSelection();
                    }
                    catch (DbEntityValidationException dbEx)
                    {
                        var errors = "";
                        errors = "Błędy edycji: " + Environment.NewLine;
                        errors = dbEx.EntityValidationErrors.SelectMany(validationErrors => validationErrors.ValidationErrors).Aggregate(errors, (current, validationError) => current + "Dotyczy: " + validationError.PropertyName + Environment.NewLine + "Rodzaj: " + validationError.ErrorMessage + Environment.NewLine + Environment.NewLine);
                        MessageBox.Show(errors);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(@"Błąd edycji: " + ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show(@"Aby edytować obiekt w pierwszej kolejności musisz jakiś wybrać.");
            }
        }

        private void radioButtonEditionOn_CheckedChanged(object sender, EventArgs e)
        {
            _editingOn = true;
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
            if (_searchingOn)
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
                _editingOn = false;
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
                if (_searchingOn == false)
                {

                    int id = (int)dataGridViewCustomers.SelectedRows[0].Cells["ID"].Value;
                    textBoxDescription.Text = _readCustomerRepository.GetById(id).Description ?? "";
                    textBoxComments.Text = _readCustomerRepository.GetById(id).Comments ?? "";
                    textBoxName.Text = dataGridViewCustomers.SelectedRows[0].Cells["Nazwa"].Value?.ToString() ?? "";
                    textBoxPhoneNumber.Text = dataGridViewCustomers.SelectedRows[0].Cells["Telefon"].Value?.ToString() ?? "";
                    textBoxEmail.Text = dataGridViewCustomers.SelectedRows[0].Cells["Email"].Value?.ToString() ?? "";
                    comboBoxTrade.Text = dataGridViewCustomers.SelectedRows[0].Cells["Branża"].Value?.ToString() ?? "";
                    textBoxCity.Text = dataGridViewCustomers.SelectedRows[0].Cells["Miasto"].Value?.ToString() ?? "";
                    textBoxPostalCode.Text = dataGridViewCustomers.SelectedRows[0].Cells["Kod_pocztowy"].Value?.ToString() ?? "";
                    textBoxStreet.Text = dataGridViewCustomers.SelectedRows[0].Cells["Ulica"].Value?.ToString() ?? "";
                    comboBoxProvince.Text = dataGridViewCustomers.SelectedRows[0].Cells["Województwo"].Value?.ToString() ?? "";
                }
                else
                {
                    var id = (int)dataGridViewCustomers.SelectedRows[0].Cells["ID"].Value;
                    textBoxDescription.Text = _readCustomerRepository.GetById(id).Description ?? "";
                    textBoxComments.Text = _readCustomerRepository.GetById(id).Comments ?? "";
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

            _searchingOn = false;
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

            _searchingOn = true;
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
            if (!_searchingOn) return;
            string buttonName;
            string option;
            if (sender.GetType().ToString() == "System.Windows.Forms.TextBox")
            {
                var textBox = (TextBox)sender;
                buttonName = textBox.Name;
                option = textBox.Text;
            }
            else
            {
                var comboBox = (ComboBox)sender;
                buttonName = comboBox.Name;
                option = comboBox.Text;
            }

            switch (buttonName)
            {
                case "textBoxName":
                    dataGridViewCustomers.DataSource = null;
                    dataGridViewCustomers.DataSource =
                        _readCustomerRepository.GetAll().Where(x => x.Name.Contains(option)).Select(x => new
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
                        _readCustomerRepository.GetAll().Where(x => x.PhoneNumber.Contains(option)).Select(x => new
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
                        _readCustomerRepository.GetAll().Where(x => x.Email.Contains(option)).Select(x => new
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
                case "ComboBoxTrade":
                    dataGridViewCustomers.DataSource = null;
                    dataGridViewCustomers.DataSource =
                        _readCustomerRepository.GetAll().Where(x => x.Trade.Contains(option)).Select(x => new
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
                case "ComboBoxProvince":
                    dataGridViewCustomers.DataSource = null;
                    dataGridViewCustomers.DataSource =
                        _readCustomerRepository.GetAll().Where(x => x.Province.Contains(option)).Select(x => new
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
                        _readCustomerRepository.GetAll().Where(x => x.City.Contains(option)).Select(x => new
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
                        _readCustomerRepository.GetAll().Where(x => x.PostalCode.Contains(option)).Select(x => new
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
                        _readCustomerRepository.GetAll().Where(x => x.Street.Contains(option)).Select(x => new
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

        private void predefinicjeBranżToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            var tradePredefinitions = new TradePredefinitions(comboBoxTrade);
            Visible = false;
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
            var provincePredefinitions = new ProvincePredefinitions(comboBoxProvince);
            Visible = false;
            provincePredefinitions.ShowDialog();
            if (provincePredefinitions.IsAccessible == false) Visible = true;

            comboBoxProvince.Items.Clear();
            foreach (var item in Settings.Default.ListProvincePredefinitions)
            {
                comboBoxProvince.Items.Add(item);
            }
        }

        private void CopyAlltoClipboard()
        {
            dataGridViewAllDataExportToExcel.DataSource = null;
            dataGridViewAllDataExportToExcel.DataSource = _readCustomerRepository
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
            var dataObj = dataGridViewAllDataExportToExcel.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void CopyAlltoClipboardWithoutCommentsAndDescriptions()
        {
            dataGridViewAllDataExportToExcel.DataSource = null;
            dataGridViewAllDataExportToExcel.DataSource = _readCustomerRepository
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
                MessageBox.Show(@"Czy obiekty w tabeli mają posiadać kolumny 'Opis' oraz 'Uwagi'?",
                    @"Wybierz wersję tabeli.", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
            {
                CopyAlltoClipboard();
                object misValue = System.Reflection.Missing.Value;
                var xlexcel = new Microsoft.Office.Interop.Excel.Application { Visible = true };
                var xlWorkBook = xlexcel.Workbooks.Add(misValue);
                var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.Item[1];
                var cr = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
                cr.Select();
                xlWorkSheet.PasteSpecial(cr, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
            else if (dialogResult == DialogResult.No)
            {
                CopyAlltoClipboardWithoutCommentsAndDescriptions();
                object misValue = System.Reflection.Missing.Value;
                var xlexcel = new Microsoft.Office.Interop.Excel.Application { Visible = true };
                var xlWorkBook = xlexcel.Workbooks.Add(misValue);
                var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.Item[1];
                var cr = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
                cr.Select();
                xlWorkSheet.PasteSpecial(cr, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
        }

        private void zamknijToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void dataGridViewCustomers_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var option = dataGridViewCustomers.Columns[e.ColumnIndex].Name;

            if (option == "ID")
            {
                dataGridViewCustomers.DataSource =
                    _readCustomerRepository.GetAll().OrderBy(x => x.ID).Select(x => new
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
            else if (option == "Nazwa")
            {
                dataGridViewCustomers.DataSource =
                    _readCustomerRepository.GetAll().OrderBy(x => x.Name).Select(x => new
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
            else if (option == "Telefon")
            {
                dataGridViewCustomers.DataSource =
                    _readCustomerRepository.GetAll().OrderBy(x => x.PhoneNumber).Select(x => new
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
            else if (option == "Email")
            {
                dataGridViewCustomers.DataSource =
                    _readCustomerRepository.GetAll().OrderBy(x => x.Email).Select(x => new
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
            else if (option == "Branża")
            {
                dataGridViewCustomers.DataSource =
                    _readCustomerRepository.GetAll().OrderBy(x => x.Trade).Select(x => new
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
            else if (option == "Województwo")
            {
                dataGridViewCustomers.DataSource =
                    _readCustomerRepository.GetAll().OrderBy(x => x.Province).Select(x => new
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
            else if (option == "Miasto")
            {
                dataGridViewCustomers.DataSource =
                    _readCustomerRepository.GetAll().OrderBy(x => x.City).Select(x => new
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
            else if (option == "Kod_pocztowy")
            {
                dataGridViewCustomers.DataSource =
                    _readCustomerRepository.GetAll().OrderBy(x => x.PostalCode).Select(x => new
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
            else if (option == "Ulica")
            {
                dataGridViewCustomers.DataSource =
                    _readCustomerRepository.GetAll().OrderBy(x => x.Street).Select(x => new
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
            Settings.Default.FormSize = Size;
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
                var tmpList = new List<string>();
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
                var tmpList = new List<string>();
                Settings.Default.ListProvincePredefinitions = tmpList;
                Settings.Default.Save();
            }
            if ((Settings.Default.FormSize.Width != 0) && (Settings.Default.FormSize.Height != 0))
            {
                Size = Settings.Default.FormSize;
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
