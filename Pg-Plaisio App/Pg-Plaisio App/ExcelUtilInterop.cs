using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Pg_Plaisio_App
{
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