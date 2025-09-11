using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pg_Plaisio_App
{
    public partial class Form1 : Form
    {
        private const string fileName = "PG-PLAISIO 25.4.07 BOTH.xlsm";

        private ExcelUtilInterop _excelUtil;

        private static readonly DateTime expiryDate = new DateTime(2025, 9, 17);

        public Form1()
        {
            InitializeComponent();

            //if (CheckExpiry() == false)
            //    return;

            string pathExcel = Path.Combine(GetAppFolder(), fileName);

            _excelUtil = new ExcelUtilInterop(pathExcel);

            tabControl.SelectedIndex = 1;

            InitDataΧΟΝΔΡΙΚΗ();
        }

        public static bool CheckExpiry()
        {
            DateTime currentDate = DateTime.Now.Date;

            if (currentDate >= expiryDate)
            {
                MessageBox.Show("Vui lòng liên hệ để gia hạn sử dụng.");
                return false;
            }
            return true;
        }

        public static string GetAppFolder()
        {
            string location = Assembly.GetExecutingAssembly().Location;
            string dir = Path.GetDirectoryName(location);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            return dir;
        }

        private void InitDataΧΟΝΔΡΙΚΗ(string sheetName = "ΧΟΝΔΡΙΚΗ")
        {
            AddItemToComboBox(ComboBox1, sheetName, "K11:K274");

            AddItemToComboBox(ComboBox2, sheetName, "K11:K274");

            AddItemToComboBox(ComboBox3, sheetName, "R8:R9");

            AddItemToComboBox(ComboBox4, sheetName, "O12:O30");

            AddItemToComboBox(ComboBox5, sheetName, "AD4:AD12");

            AddItemToComboBox(ComboBox6, sheetName, "R14:R18");

            AddItemToComboBox(ComboBox7, sheetName, "Q35:Q100");

            AddItemToComboBox(ComboBox8, sheetName, "O106:O124");

            AddItemToComboBox(ComboBox9, sheetName, "Q105:Q110");

            AddItemToComboBox(ComboBox12, sheetName, "AB4:AB8");

            AddItemToComboBox(ComboBox10, sheetName, "AH4:AH7");

            AddItemToComboBox(ComboBox11, sheetName, "AD16:AD125");

            AddItemToComboBox(ComboBox13, sheetName, "W18:W28", true);

            TextBox1.Text = _excelUtil.GetCellValueAsString(sheetName, "B8");
            TextBox2.Text = _excelUtil.GetCellValueAsString(sheetName, "B10");
            TextBox3.Text = _excelUtil.GetCellValueAsString(sheetName, "B28");

            this.ComboBox1.SelectedValueChanged += new System.EventHandler(this.ComboBox1_SelectedValueChanged);
            this.ComboBox2.SelectedValueChanged += new System.EventHandler(this.ComboBox2_SelectedValueChanged);
            this.ComboBox3.SelectedValueChanged += new System.EventHandler(this.ComboBox3_SelectedValueChanged);
            this.ComboBox4.SelectedValueChanged += new System.EventHandler(this.ComboBox4_SelectedValueChanged);
            this.ComboBox5.SelectedValueChanged += new System.EventHandler(this.ComboBox5_SelectedValueChanged);
            this.ComboBox6.SelectedValueChanged += new System.EventHandler(this.ComboBox6_SelectedValueChanged);
            this.ComboBox7.SelectedValueChanged += new System.EventHandler(this.ComboBox7_SelectedValueChanged);
            this.ComboBox8.SelectedValueChanged += new System.EventHandler(this.ComboBox8_SelectedValueChanged);
            this.ComboBox9.SelectedValueChanged += new System.EventHandler(this.ComboBox9_SelectedValueChanged);

            this.ComboBox12.SelectedValueChanged += new System.EventHandler(this.ComboBox12_SelectedValueChanged);

            this.ComboBox10.SelectedValueChanged += new System.EventHandler(this.ComboBox10_SelectedValueChanged);

            this.ComboBox11.SelectedValueChanged += new System.EventHandler(this.ComboBox11_SelectedValueChanged);

            this.ComboBox13.SelectedValueChanged += new System.EventHandler(this.ComboBox13_SelectedValueChanged);
        }

        private void AddItemToComboBox(ComboBox comboBox, string sheetName, string rangeAddress, bool isFormat = false)
        {
            List<string> items = _excelUtil.GetRangeValues(sheetName, rangeAddress);
            comboBox.Items.Clear();

            foreach (var item in items)
            {
                if (string.IsNullOrEmpty(item))
                    continue;

                if (isFormat)
                {
                    comboBox.Items.Add(ReformatPercentageString(item));
                }
                else
                {
                    comboBox.Items.Add(item);
                }
            }

            if (comboBox.Items.Count > 0)
                comboBox.SelectedIndex = 0; // Chọn mục đầu tiên làm mục được chọn mặc định
        }

        public string ReformatPercentageString(string percentageString)
        {
            if (string.IsNullOrEmpty(percentageString))
            {
                return null;
            }

            string cleanedString = percentageString.Trim().Replace("%", "");

            if (double.TryParse(cleanedString, out double numericValue))
            {
                return numericValue.ToString("P2", CultureInfo.InvariantCulture);
            }

            return percentageString;
        }

        private void CommandButton5_Click(object sender, EventArgs e)
        {
            tabControl.SelectedIndex = 1;
        }

        private void CommandButton6_Click(object sender, EventArgs e)
        {
            tabControl.SelectedIndex = 2;
        }

        private void CommandButton7_Click(object sender, EventArgs e)
        {
            tabControl.SelectedIndex = 3;
        }

        private void CommandButton8_Click(object sender, EventArgs e)
        {
            tabControl.SelectedIndex = 4;
        }

        private void CommandButton9_Click(object sender, EventArgs e)
        {
            tabControl.SelectedIndex = 5;
        }

        private void CommandButton10_Click(object sender, EventArgs e)
        {
            tabControl.SelectedIndex = 6;
        }

        private void CommandButton11_Click(object sender, EventArgs e)
        {
            tabControl.SelectedIndex = 7;
        }

        private void CommandButton12_Click(object sender, EventArgs e)
        {
            tabControl.SelectedIndex = 8;
        }

        private void ComboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            string value = ComboBox1.Text;

            if ((string.IsNullOrEmpty(value)))
                return;

            bool startsWithZero = value.StartsWith("0");
            bool isNumeric = double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericVal);

            if (startsWithZero || !isNumeric)
            {
                // Xử lý như VĂN BẢN (TEXT)
                _excelUtil.SetFormatCellValue("ΧΟΝΔΡΙΚΗ", "B4", "@", value);
            }
            else
            {
                // Xử lý như một con SỐ (NUMBER)
                _excelUtil.SetFormatCellValue("ΧΟΝΔΡΙΚΗ", "B4", "General", numericVal);
            }

            UpdateLabel1();
        }

        private void ComboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            string value = ComboBox2.Text;
            if (string.IsNullOrEmpty(value)) return;

            bool startsWithZero = value.StartsWith("0");
            bool isNumeric = double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericVal);

            if (startsWithZero || !isNumeric)
            {
                _excelUtil.SetFormatCellValue("ΧΟΝΔΡΙΚΗ", "B6", "@", value);
            }
            else
            {
                _excelUtil.SetFormatCellValue("ΧΟΝΔΡΙΚΗ", "B6", "General", numericVal);
            }
            UpdateLabel1();
        }

        private void ComboBox3_SelectedValueChanged(object sender, EventArgs e)
        {
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B12", ComboBox3.Text);
            UpdateLabel1();
        }

        private void ComboBox4_SelectedValueChanged(object sender, EventArgs e)
        {
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B14", ComboBox4.Text);
            UpdateLabel1();
        }

        private void ComboBox5_SelectedValueChanged(object sender, EventArgs e)
        {
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B16", ComboBox5.Text);
            UpdateLabel1();
        }

        private void ComboBox6_SelectedValueChanged(object sender, EventArgs e)
        {
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B18", ComboBox6.Text);
            UpdateLabel1();
        }

        private void ComboBox7_SelectedValueChanged(object sender, EventArgs e)
        {
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B20", ComboBox7.Text);
            UpdateLabel1();
        }

        private void ComboBox8_SelectedValueChanged(object sender, EventArgs e)
        {
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B22", ComboBox8.Text);
            UpdateLabel1();
        }

        private void ComboBox9_SelectedValueChanged(object sender, EventArgs e)
        {
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B24", ComboBox9.Text);
            UpdateLabel1();
        }

        private void ComboBox10_SelectedValueChanged(object sender, EventArgs e)
        {
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B30", ComboBox10.Text);
            UpdateLabel1();
        }

        private void ComboBox11_SelectedValueChanged(object sender, EventArgs e)
        {
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B32", ComboBox11.Text);
            UpdateLabel1();
        }

        private void ComboBox12_SelectedValueChanged(object sender, EventArgs e)
        {
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B26", ComboBox12.Text);
            UpdateLabel1();
        }

        private void ComboBox13_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ComboBox13.SelectedIndex == -1)
                return;

            string selectedValue = ComboBox13.Text.Trim().Replace("%", "");

            if (double.TryParse(selectedValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedNumber))
            {
                double percentageValue = parsedNumber / 100;

                _excelUtil.SetFormatCellValue("ΧΟΝΔΡΙΚΗ", "C37", "0.00%", percentageValue);
            }
            else
            {
                MessageBox.Show("Μη έγκυρη τιμή!", "Σφάλμα", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            UpdateLabel1();
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            string value = TextBox1.Text;

            if (double.TryParse(value, out _) || string.IsNullOrEmpty(value))
            {
                _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B8", value);
            }
            UpdateLabel1();
        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {
            string value = TextBox2.Text;

            if (double.TryParse(value, out _) || string.IsNullOrEmpty(value))
            {
                _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B10", value);
            }
            UpdateLabel1();
        }

        private void TextBox3_TextChanged(object sender, EventArgs e)
        {
            string value = TextBox3.Text;

            if (double.TryParse(value, out _) || string.IsNullOrEmpty(value))
            {
                _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B28", value);
            }
            UpdateLabel1();
        }

        private void TextBoxInterger_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Cho phép số, dấu phẩy, dấu chấm, Backspace
            if (char.IsDigit(e.KeyChar) || e.KeyChar == ',' || e.KeyChar == '.'
                || e.KeyChar == (char)Keys.Back || e.KeyChar == (char)Keys.Delete)
            {
                // hợp lệ → không làm gì
            }
            else
            {
                // chặn ký tự không hợp lệ
                e.Handled = true;
            }
        }

        private void UpdateLabel1()
        {
            SetText(Label19, "C36");
            SetText(Label20, "C37", false);
            SetText(Label21, "C38");
            SetText(Label22, "C39");
            SetText(Label23, "C40");
            SetText(Label24, "C43");
            SetText(Label25, "C44");
            SetText(Label26, "C45");
            SetText(Label27, "C46");
            SetText(Label28, "C47");
            SetText(Label29, "C48");
            SetText(Label30, "C49");
            SetText(Label31, "C50");
            SetText(Label32, "C51");

            SetText(Label185, "K6");
            SetText(Label186, "D43", suffix: "m");
            SetText(Label187, "L6");
            SetText(Label188, "D44", suffix: "m");
            SetText(Label189, "D45");
            SetText(Label190, "D45");

            string selectedValue = (ComboBox10.Text ?? string.Empty)
                .Replace('\u00A0', ' ') // \u00A0 = non-breaking space
                .Trim();

            double val;

            if (selectedValue == "ΟΧΙ" || selectedValue == "ΧΑΡΤΙ")
            {
                val = _excelUtil.GetCellValueDouble("ΧΟΝΔΡΙΚΗ", "AI1");
            }
            else if (selectedValue == "ΚΑΜΒΑΣ ΠΟΛΥΕΣΤΕΡΙΚΟΣ" || selectedValue == "ΚΑΜΒΑΣ ΒΑΜΒΑΚΕΡΟΣ")
            {
                val = _excelUtil.GetCellValueDouble("ΧΟΝΔΡΙΚΗ", "AJ1");
            }
            else
            {
                val = _excelUtil.GetCellValueDouble("ΧΟΝΔΡΙΚΗ", "D50");
            }

            if (val != double.NaN)
                Label191.Text = val.ToString("0.00", CultureInfo.InvariantCulture);
            else
                Label191.Text = "Σφάλμα δεδομένων!";
        }

        private void SetText(System.Windows.Forms.Label label, string cellAddress, bool isNumber = true, string suffix = "", string sheetName = "ΧΟΝΔΡΙΚΗ")
        {
            string valueStr = _excelUtil.GetCellValueAsString(sheetName, cellAddress);

            if (double.TryParse(valueStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double value) && isNumber)
                valueStr = value.ToString("0.00", CultureInfo.InvariantCulture);

            label.Text = valueStr + suffix;
        }

        private void CommandButton1_Click(object sender, EventArgs e)
        {
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B4", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "K11"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B6", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "K11"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B8", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "J14"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B10", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "J15"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B12", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "R8"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B14", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "O12"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B16", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "AD4"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B18", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "R14"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B20", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "Q35"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B22", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "O106"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B24", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "Q105"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B26", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "AB4"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B28", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "J13"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B30", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "AH4"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "B32", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "AD16"));
            _excelUtil.SetCellValue("ΧΟΝΔΡΙΚΗ", "C37", _excelUtil.GetCellValueAsString("ΧΟΝΔΡΙΚΗ", "W18"));

            TextBox1.Text = "0";
            TextBox2.Text = "0";
            TextBox3.Text = "0";

            ComboBox1.SelectedIndex = 0;
            ComboBox2.SelectedIndex = 0;
            ComboBox3.SelectedIndex = 0;
            ComboBox4.SelectedIndex = 0;
            ComboBox5.SelectedIndex = 0;
            ComboBox6.SelectedIndex = 0;
            ComboBox7.SelectedIndex = 0;
            ComboBox8.SelectedIndex = 0;
            ComboBox9.SelectedIndex = 0;
            ComboBox12.SelectedIndex = 0;
            ComboBox10.SelectedIndex = 0;
            ComboBox11.SelectedIndex = 0;
            ComboBox13.SelectedIndex = 0;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            _excelUtil.Dispose();
        }
    }
}