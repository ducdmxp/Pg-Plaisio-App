using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
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

            if (CheckExpiry() == false)
                return;

            string pathExcel = Path.Combine(GetAppFolder(), fileName);

            _excelUtil = new ExcelUtilInterop(pathExcel);

            tabControl.SelectedIndex = 1;

            InitDataΧΟΝΔΡΙΚΗ();

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
            tabControl.SelectedIndex = 1; // Go to ΧΟΝΔΡΙΚΗ tab (Page1)
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
                MessageBox.Show("Ìç Ýãêõñç ôéìÞ!", "ÓöÜëìá", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            SetText(Label186, "D43");
            SetText(Label187, "L6");
            SetText(Label188, "D44");
            SetText(Label189, "D45");
            SetText(Label190, "D45");

            string selectedValue = (ComboBox12.Text ?? string.Empty)
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
                Label191.Text = "ÓöÜëìá äåäïìÝíùí!";
        }

        private void SetText(System.Windows.Forms.Label label, string cellAddress, bool isNumber = true, string sheetName = "ΧΟΝΔΡΙΚΗ")
        {
            string valueStr = _excelUtil.GetCellValueAsString(sheetName, cellAddress);

            if (double.TryParse(valueStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double value) && isNumber)
            {
                label.Text = value.ToString("0.00", CultureInfo.InvariantCulture);
            }
            else
            {
                label.Text = valueStr;
            }
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

    public class ExcelUtil
    {
        private XLWorkbook _xLWorkbook;

        private string ExcelFilePath { get; set; }

        public ExcelUtil(string filePath)
        {
            ExcelFilePath = filePath;

            ReadWorkbook(filePath);
        }

        public bool Save()
        {
            try
            {
                if (_xLWorkbook != null)
                {
                    _xLWorkbook.Save(new SaveOptions() { EvaluateFormulasBeforeSaving = true });

                    return true;
                }
            }
            catch (Exception)
            {
            }
            return false;
        }

        private void ReadWorkbook(string filePath)
        {
            try
            {
                _xLWorkbook = new XLWorkbook(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra một lỗi nghiêm trọng: {ex.Message}");
            }
        }

        public List<string> GetRangeValues(string worksheetName, string rangeAddress)
        {
            var values = new List<string>();

            if (_xLWorkbook == null)
            {
                Console.WriteLine($"Lỗi: Workbook chưa được khởi tạo.");
                return values;
            }

            try
            {
                var worksheet = _xLWorkbook.Worksheet(worksheetName);

                if (worksheet == null)
                {
                    Console.WriteLine($"Lỗi: Worksheet '{worksheetName}' không tồn tại.");
                    return values;
                }

                // Lấy vùng (range) theo địa chỉ được cung cấp
                var range = worksheet.Range(rangeAddress);

                if (range == null)
                {
                    Console.WriteLine($"Lỗi: Không thể tìm thấy vùng '{rangeAddress}'.");
                    return values;
                }

                // Duyệt qua từng ô trong vùng và thêm giá trị vào danh sách
                foreach (var cell in range.Cells())
                {
                    // GetFormattedString() để lấy giá trị dạng chuỗi đã được định dạng
                    values.Add(cell.GetFormattedString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra một lỗi nghiêm trọng: {ex.Message}");
                return values;
            }

            return values;
        }

        public double GetCellValueDouble(string worksheetName, string cellAddress)
        {
            try
            {
                if (_xLWorkbook == null)
                {
                    Console.WriteLine($"Lỗi: Workbook chưa được khởi tạo.");
                    return 0;
                }

                var worksheet = _xLWorkbook.Worksheet(worksheetName);

                if (worksheet == null)
                {
                    Console.WriteLine($"Lỗi: Worksheet '{worksheetName}' không tồn tại.");
                    return 0;
                }

                var cellVal = worksheet.Cell(cellAddress).GetValue<double>();

                return cellVal;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra một lỗi nghiêm trọng: {ex.Message}");
                return 0;
            }
        }

        public string GetCellValueAsString(string worksheetName, string cellAddress)
        {
            try
            {
                if (_xLWorkbook == null)
                {
                    Console.WriteLine($"Lỗi: Workbook chưa được khởi tạo.");
                    return string.Empty;
                }

                using (var workbook = new XLWorkbook(ExcelFilePath))
                {
                    var worksheet = workbook.Worksheet(worksheetName);

                    if (worksheet == null)
                    {
                        Console.WriteLine($"Lỗi: Worksheet '{worksheetName}' không tồn tại.");
                        return null; // Trả về null khi có lỗi
                    }

                    var cell = worksheet.Cell(cellAddress);

                    return cell.GetFormattedString();
                }

                //var worksheet = _xLWorkbook.Worksheet(worksheetName);

                //if (worksheet == null)
                //{
                //    Console.WriteLine($"Lỗi: Worksheet '{worksheetName}' không tồn tại.");
                //    return null; // Trả về null khi có lỗi
                //}

                //var cell = worksheet.Cell(cellAddress);

                //return cell.GetFormattedString();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra một lỗi nghiêm trọng: {ex.Message}");
                return null; // Trả về null khi có lỗi
            }
        }

        public bool SetFormatCellValue(string worksheetName, string cellAddress, string NumberFormat, object value)
        {
            try
            {
                using (var workbook = new XLWorkbook(ExcelFilePath))
                {
                    var worksheet = workbook.Worksheet(worksheetName);

                    if (worksheet == null)
                    {
                        Console.WriteLine($"Lỗi: Worksheet '{worksheetName}' không tồn tại.");
                        return false; // Trả về null khi có lỗi
                    }

                    worksheet.Cell(cellAddress).Style.NumberFormat.Format = NumberFormat;
                    worksheet.Cell(cellAddress).Value = XLCellValue.FromObject(value);

                    workbook.Save();
                }

                //if (_xLWorkbook == null)
                //{
                //    Console.WriteLine($"Lỗi: Workbook chưa được khởi tạo.");
                //    return false;
                //}

                //if (!_xLWorkbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet))
                //{
                //    Console.WriteLine($"Lỗi: Không tìm thấy worksheet có tên '{worksheetName}'.");
                //    return false;
                //}

                //worksheet.Cell(cellAddress).Style.NumberFormat.Format = NumberFormat;
                //worksheet.Cell(cellAddress).Value = XLCellValue.FromObject(value);

                //Save();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra lỗi: {ex.Message}");
                return false;
            }
        }

        public bool SetCellValue(string worksheetName, string cellAddress, object valueToSet, bool isSave = true)
        {
            try
            {
                if (!_xLWorkbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet))
                {
                    Console.WriteLine($"Lỗi: Không tìm thấy worksheet có tên '{worksheetName}'.");
                    return false;
                }

                worksheet.Cell(cellAddress).Value = XLCellValue.FromObject(valueToSet);

                if (isSave)
                    Save();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra lỗi: {ex.Message}");
                return false;
            }
        }
    }

    public class ExcelUtilInterop : IDisposable
    {
        private dynamic worksheet = null;
        private dynamic excelApp;
        private dynamic workbook;

        private bool isExcelOpen;
        private string ExcelFilePath { get; set; }

        public ExcelUtilInterop(string filePath)
        {
            ExcelFilePath = filePath;

            InitializeExcel();
            ReadWorkbook(filePath);
        }

        #region Khởi tạo và kết nối Excel

        /// <summary>
        /// Khởi tạo Excel Application
        /// </summary>
        private void InitializeExcel()
        {
            try
            {
                // Thử kết nối với Excel đang chạy trước
                try
                {
                    excelApp = (System.Windows.Forms.Application)Marshal.GetActiveObject("Excel.Application");
                    isExcelOpen = false;
                    Console.WriteLine("✅ Đã kết nối với Excel đang chạy");
                }
                catch
                {
                    // Nếu không có Excel nào chạy, tạo instance mới
                    Type excelType = Type.GetTypeFromProgID("Excel.Application");
                    excelApp = Activator.CreateInstance(excelType);
                    isExcelOpen = true;
                }

                excelApp.Visible = false;
                excelApp.DisplayAlerts = false; // Tắt cảnh báo
                excelApp.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Lỗi khởi tạo Excel: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Đọc workbook từ file
        /// </summary>
        private void ReadWorkbook(string filePath)
        {
            try
            {
                workbook = FindOpenWorkbook(excelApp, filePath);
                if (workbook == null)
                {
                    workbook = excelApp.Workbooks.Open(filePath,
                  UpdateLinks: false,
                  ReadOnly: false,
                  Format: Missing.Value,
                  Password: Missing.Value,
                  WriteResPassword: Missing.Value,
                  IgnoreReadOnlyRecommended: true,
                  Origin: Missing.Value,
                  Delimiter: Missing.Value,
                  Editable: false,
                  Notify: false,
                  Converter: Missing.Value,
                  AddToMru: false,
                  Local: false,
                  CorruptLoad: Missing.Value);
                }

                Console.WriteLine($"✅ Đã mở workbook: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Đã xảy ra một lỗi nghiêm trọng: {ex.Message}");
                throw;
            }
        }

        private static dynamic FindOpenWorkbook(dynamic excelApp, string filePath)
        {
            try
            {
                string fileName = Path.GetFileName(filePath);
                foreach (dynamic wb in excelApp.Workbooks)
                {
                    if (string.Equals(wb.name, fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        return wb;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error finding open workbook: {ex.Message}");
            }
            return null;
        }

        #endregion Khởi tạo và kết nối Excel

        #region Save Methods

        /// <summary>
        /// Lưu workbook
        /// </summary>
        public bool Save()
        {
            try
            {
                if (workbook != null)
                {
                    //workbook.Save();
                    //Console.WriteLine("✅ Đã lưu workbook");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Lỗi lưu file: {ex.Message}");
            }
            return false;
        }

        #endregion Save Methods

        #region Get Values Methods

        /// <summary>
        /// Lấy giá trị của một dải ô
        /// </summary>
        public List<string> GetRangeValues(string worksheetName, string rangeAddress)
        {
            var values = new List<string>();

            if (workbook == null)
            {
                Console.WriteLine($"❌ Lỗi: Workbook chưa được khởi tạo.");
                return values;
            }

            //dynamic worksheet = null;
            Range range = null;

            try
            {
                worksheet = GetWorksheet(worksheetName);
                if (worksheet == null)
                {
                    Console.WriteLine($"❌ Lỗi: Worksheet '{worksheetName}' không tồn tại.");
                    return values;
                }

                range = worksheet.Range[rangeAddress];
                if (range == null)
                {
                    Console.WriteLine($"❌ Lỗi: Không thể tìm thấy vùng '{rangeAddress}'.");
                    return values;
                }

                // Nếu là ô đơn
                if (range.Cells.Count == 1)
                {
                    object cellValue = range.Value2;
                    values.Add(cellValue?.ToString() ?? string.Empty);
                }
                else
                {
                    // Nếu là nhiều ô
                    object[,] rangeValues = range.Value2;
                    if (rangeValues != null)
                    {
                        int rows = rangeValues.GetLength(0);
                        int cols = rangeValues.GetLength(1);

                        for (int i = 1; i <= rows; i++)
                        {
                            for (int j = 1; j <= cols; j++)
                            {
                                object cellValue = rangeValues[i, j];
                                values.Add(cellValue?.ToString() ?? string.Empty);
                            }
                        }
                    }
                }

                Console.WriteLine($"✅ Đọc {values.Count} ô từ vùng {rangeAddress}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Đã xảy ra một lỗi nghiêm trọng: {ex.Message}");
            }
            finally
            {
                // Cleanup COM objects
                //if (range != null) Marshal.ReleaseComObject(range);
                //if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }

            return values;
        }

        /// <summary>
        /// Lấy giá trị ô dưới dạng double
        /// </summary>
        public double GetCellValueDouble(string worksheetName, string cellAddress)
        {
            //dynamic worksheet = null;
            Range cell = null;

            try
            {
                if (workbook == null)
                {
                    Console.WriteLine($"❌ Lỗi: Workbook chưa được khởi tạo.");
                    return 0;
                }

                worksheet = GetWorksheet(worksheetName);
                if (worksheet == null)
                {
                    Console.WriteLine($"❌ Lỗi: Worksheet '{worksheetName}' không tồn tại.");
                    return 0;
                }

                cell = worksheet.Range[cellAddress];
                object cellValue = cell.Value2;

                if (cellValue == null)
                    return 0;

                if (double.TryParse(cellValue.ToString(), out double result))
                {
                    Console.WriteLine($"✅ Đọc {cellAddress}: {result}");
                    return result;
                }

                Console.WriteLine($"⚠️ Không thể chuyển đổi '{cellValue}' thành double");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Đã xảy ra một lỗi nghiêm trọng: {ex.Message}");
                return 0;
            }
            finally
            {
                // Cleanup COM objects
                //if (cell != null) Marshal.ReleaseComObject(cell);
                //if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        /// <summary>
        /// Lấy giá trị ô dưới dạng string
        /// </summary>
        public string GetCellValueAsString(string worksheetName, string cellAddress)
        {
            //dynamic worksheet = null;
            Range cell = null;

            try
            {
                if (workbook == null)
                {
                    Console.WriteLine($"❌ Lỗi: Workbook chưa được khởi tạo.");
                    return string.Empty;
                }

                worksheet = GetWorksheet(worksheetName);
                if (worksheet == null)
                {
                    Console.WriteLine($"❌ Lỗi: Worksheet '{worksheetName}' không tồn tại.");
                    return null;
                }

                cell = worksheet.Range[cellAddress];
                object cellValue = cell.Value2;

                // Lấy giá trị đã format
                string formattedValue = cell.Text?.ToString() ?? cellValue?.ToString() ?? string.Empty;

                Console.WriteLine($"✅ Đọc {cellAddress}: '{formattedValue}'");
                return formattedValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Đã xảy ra một lỗi nghiêm trọng: {ex.Message}");
                return null;
            }
            finally
            {
                // Cleanup COM objects
                //if (cell != null) Marshal.ReleaseComObject(cell);
                //if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        #endregion Get Values Methods

        #region Set Values Methods

        /// <summary>
        /// Đặt format và giá trị cho ô
        /// </summary>
        public bool SetFormatCellValue(string worksheetName, string cellAddress, string numberFormat, object value)
        {
            //dynamic worksheet = null;
            Range cell = null;

            try
            {
                if (workbook == null)
                {
                    Console.WriteLine($"❌ Lỗi: Workbook chưa được khởi tạo.");
                    return false;
                }

                worksheet = GetWorksheet(worksheetName);
                if (worksheet == null)
                {
                    Console.WriteLine($"❌ Lỗi: Worksheet '{worksheetName}' không tồn tại.");
                    return false;
                }

                cell = worksheet.Range[cellAddress];

                // Đặt format trước
                cell.NumberFormat = numberFormat;

                // Đặt giá trị
                cell.Value2 = value;

                Console.WriteLine($"✅ Đã đặt format '{numberFormat}' và giá trị '{value}' cho {cellAddress}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Đã xảy ra lỗi: {ex.Message}");
                return false;
            }
            finally
            {
                // Cleanup COM objects
                //if (cell != null) Marshal.ReleaseComObject(cell);
                //if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        /// <summary>
        /// Đặt giá trị cho ô
        /// </summary>
        public bool SetCellValue(string worksheetName, string cellAddress, object valueToSet)
        {
            //dynamic worksheet = null;
            Range cell = null;

            try
            {
                if (workbook == null)
                {
                    Console.WriteLine($"❌ Lỗi: Workbook chưa được khởi tạo.");
                    return false;
                }

                worksheet = GetWorksheet(worksheetName);
                if (worksheet == null)
                {
                    Console.WriteLine($"❌ Lỗi: Không tìm thấy worksheet có tên '{worksheetName}'.");
                    return false;
                }

                cell = worksheet.Range[cellAddress];
                cell.Value2 = valueToSet;

                Console.WriteLine($"✅ Đã đặt giá trị '{valueToSet}' cho {cellAddress}");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Đã xảy ra lỗi: {ex.Message}");
                return false;
            }
            finally
            {
                // Cleanup COM objects
                //if (cell != null) Marshal.ReleaseComObject(cell);
                //if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        #endregion Set Values Methods

        #region Utility Methods

        /// <summary>
        /// Lấy worksheet theo tên
        /// </summary>
        private dynamic GetWorksheet(string worksheetName)
        {
            try
            {
                return workbook.Worksheets[worksheetName]; ;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Lỗi lấy worksheet: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Lấy danh sách tên worksheet
        /// </summary>
        public List<string> GetWorksheetNames()
        {
            var names = new List<string>();

            try
            {
                if (workbook != null)
                {
                    foreach (dynamic ws in workbook.Worksheets)
                    {
                        names.Add(ws.Name);
                        Marshal.ReleaseComObject(ws);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Lỗi lấy tên worksheet: {ex.Message}");
            }

            return names;
        }

        /// <summary>
        /// Tạo worksheet mới
        /// </summary>
        public bool CreateWorksheet(string worksheetName)
        {
            try
            {
                if (workbook == null) return false;

                dynamic newWorksheet = workbook.Worksheets.Add();
                newWorksheet.Name = worksheetName;

                Marshal.ReleaseComObject(newWorksheet);
                Console.WriteLine($"✅ Đã tạo worksheet: {worksheetName}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Lỗi tạo worksheet: {ex.Message}");
                return false;
            }
        }

        #endregion Utility Methods

        #region IDisposable Implementation

        private bool _disposed = false;

        /// <summary>
        /// Giải phóng tài nguyên COM
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                try
                {
                    if (worksheet != null) Marshal.ReleaseComObject(worksheet);

                    if (workbook != null)
                    {
                        workbook.Close(SaveChanges: false);
                    }

                    if (excelApp != null)
                    {
                        // Chỉ quit nếu chúng ta tạo Excel instance mới
                        if (isExcelOpen)
                        {
                            excelApp.Quit();
                        }
                    }

                    //if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                    //if (workbook != null) { Marshal.ReleaseComObject(workbook); }
                    //if (excelApp != null) { Marshal.ReleaseComObject(excelApp); }

                    worksheet = null;
                    workbook = null;
                    excelApp = null;

                    // Force garbage collection
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    Console.WriteLine("✅ Đã giải phóng tài nguyên Excel");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Lỗi cleanup: {ex.Message}");
                }

                _disposed = true;
            }
        }

        /// <summary>
        /// Finalizer
        /// </summary>
        ~ExcelUtilInterop()
        {
            Dispose(false);
        }

        #endregion IDisposable Implementation
    }
}