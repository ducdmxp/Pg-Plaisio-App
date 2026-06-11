using Autodesk.Revit.DB;
using Microsoft.VisualBasic;
using Convert2DTo3D.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Convert2DTo3D.UI
{
    public partial class ScheduleSleeveFrm : System.Windows.Forms.Form
    {
        public string ScheduleName
        {
            get
            {
                return txtName.Text;
            }
            set
            {
                txtName.Text = value;
            }
        }

        public int SelectedType
        {
            get
            {
                if (rdbPipe.Checked)
                    return 0;
                else if (rdbCableTray.Checked)
                    return 1;
                else if (rdbDuct.Checked)
                    return 2;
                else if (rdbConduit.Checked)
                    return 3;
                else
                    return 0;
            }
            set
            {
                switch (value)
                {
                    case 0:
                        rdbPipe.Checked = true;
                        break;

                    case 1:
                        rdbCableTray.Checked = true;
                        break;

                    case 2:
                        rdbDuct.Checked = true;
                        break;

                    case 3:
                        rdbConduit.Checked = true;
                        break;

                    default:
                        rdbPipe.Checked = true;
                        break;
                }
            }
        }

        public BuiltInCategory SelectedBuiltInCategory
        {
            get
            {
                if (rdbPipe.Checked)
                    return BuiltInCategory.OST_PipeCurves;
                else if (rdbCableTray.Checked)
                    return BuiltInCategory.OST_CableTray;
                else if (rdbDuct.Checked)
                    return BuiltInCategory.OST_DuctCurves;
                else if (rdbConduit.Checked)
                    return BuiltInCategory.OST_Conduit;
                else
                    return 0;
            }
        }

        public FieldData HeaderLevel1
        {
            get
            {
                return cboLevel1.SelectedItem as FieldData;
            }
        }

        public FieldData HeaderLevel2
        {
            get
            {
                return cboLevel2.SelectedItem as FieldData;
            }
        }

        public FieldData HeaderLevel3
        {
            get
            {
                return cboLevel3.SelectedItem as FieldData;
            }
        }

        public FieldData HeaderLevel4
        {
            get
            {
                return cboLevel4.SelectedItem as FieldData;
            }
        }

        public ScheduleSleeveFrm()
        {
            InitializeComponent();

            InitData();
        }

        private void InitData()
        {
            GetSetting();
        }

        private void SaveSetting()
        {
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, "SelectedType", SelectedType.ToString());
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, "cboLevel1", cboLevel1.SelectedIndex.ToString());
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, "cboLevel2", cboLevel2.SelectedIndex.ToString());
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, "cboLevel3", cboLevel3.SelectedIndex.ToString());
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, "cboLevel4", cboLevel4.SelectedIndex.ToString());
        }

        private void GetSetting()
        {
            string selectedType = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, "SelectedType", "0");

            int.TryParse(selectedType, out int type);

            SelectedType = type;

            rdbCategory_CheckedChanged(null, null);
        }

        public void PreventCopy(object sender, KeyEventArgs e)
        {
            if (e.Control == true && e.KeyCode == Keys.V)
            {
                if (!double.TryParse(Clipboard.GetText(), out double result) ||
                    Clipboard.GetText().Contains(",") || Clipboard.GetText().Contains("-"))

                    e.SuppressKeyPress = true;
            }
        }

        public void Validate(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && (!char.IsDigit(e.KeyChar))
                && (e.KeyChar != '.'))//&& (e.KeyChar != '-'))
                e.Handled = true;

            if (sender is System.Windows.Forms.TextBox)
            {
                // only allow one decimal point
                if (e.KeyChar == '.' && (sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1)
                    e.Handled = true;

                // only allow minus sign at the beginning
                if (e.KeyChar == '-' && (sender as System.Windows.Forms.TextBox).SelectionStart > 0)
                    e.Handled = true;
            }
            else if (sender is System.Windows.Forms.ComboBox)
            {
                // only allow one decimal point
                if (e.KeyChar == '.' && (sender as System.Windows.Forms.ComboBox).Text.IndexOf('.') > -1)
                    e.Handled = true;

                // only allow minus sign at the beginning
                if (e.KeyChar == '-' && (sender as System.Windows.Forms.ComboBox).SelectionStart > 0)
                    e.Handled = true;
            }
        }

        private void cbo_DropDown(object sender, EventArgs e)
        {
            try
            {
                if (sender is System.Windows.Forms.ComboBox comboBox)
                {
                    object[] items = new object[comboBox.Items.Count];
                    comboBox.Items.CopyTo(items, 0);
                    comboBox.DropDownWidth = items.Select(obj => TextRenderer.MeasureText(comboBox.GetItemText(obj), comboBox.Font).Width).Max();
                }
            }
            catch { }
        }

        private List<Parameter> GetParamters()
        {
            try
            {
                List<Parameter> parameters = new List<Parameter>();

                Element elem = new FilteredElementCollector(Global.Doc)
                        .OfCategory(BuiltInCategory.OST_DuctAccessory)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(x => x.Symbol.FamilyName.Contains("Rec Duct Sleeve Opening Sleeve")
                               || x.Symbol.FamilyName.Contains("Round Duct Sleeve Opening Sleeve"))
                        .FirstOrDefault();

                if (elem == null)
                    return new List<Parameter>();

                //Parameter level = elem.get_Parameter(BuiltInParameter.ALL_MODEL_COST);
                //Parameter comment = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);

                parameters = elem.Parameters.Cast<Parameter>().OrderBy(x => x.Definition.Name).ToList();

                return parameters;
            }
            catch (Exception ex)
            {
                return new List<Parameter>();
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (IsValid())
            {
                SaveSetting();

                this.DialogResult = DialogResult.OK;
            }
        }

        private bool IsValid()
        {
            List<Autodesk.Revit.DB.View> lstSchedules = new FilteredElementCollector(Global.Doc)
                        .OfCategory(BuiltInCategory.OST_Views)
                        .OfClass(typeof(Autodesk.Revit.DB.View))
                        .Cast<Autodesk.Revit.DB.View>()
                        .Where(x => x.ViewType == ViewType.Schedule).ToList();

            if (string.IsNullOrEmpty(ScheduleName))
            {
                MessageBox.Show("Schedule name is null or empty");
                return false;
            }

            if (lstSchedules.Select(x => x.Name).Contains(ScheduleName))
            {
                MessageBox.Show("Schedule name is exits");
                return false;
            }

            return true;
        }

        private List<Parameter> parameters = new List<Parameter>();

        private void rdbCategory_CheckedChanged(object sender, EventArgs e)
        {
            parameters = GetParamters();

            cboLevel1.Items.Clear();

            List<FieldData> lstFieldLevel1 = new List<FieldData>();

            lstFieldLevel1.AddRange(parameters.Select(x => new FieldData(x)));

            foreach (var item in lstFieldLevel1)
            {
                cboLevel1.Items.Add(item);
            }

            cboLevel1.SelectedItem = lstFieldLevel1.FirstOrDefault();

            //
            List<FieldData> lstFieldLevel2 = new List<FieldData>() { new FieldData(null) };
            lstFieldLevel2.AddRange(parameters.Select(x => new FieldData(x)));

            foreach (var item in lstFieldLevel2)
            {
                cboLevel2.Items.Add(item);
                cboLevel3.Items.Add(item);
                cboLevel4.Items.Add(item);
            }

            cboLevel2.SelectedItem = lstFieldLevel2.FirstOrDefault();
            cboLevel3.SelectedItem = lstFieldLevel2.FirstOrDefault();
            cboLevel4.SelectedItem = lstFieldLevel2.FirstOrDefault();
        }

        private void cboLevel1_SelectedIndexChanged(object sender, EventArgs e)
        {
            FieldData fieldData = cboLevel1.SelectedItem as FieldData;

            if (fieldData == null)
            {
                cboLevel2.Enabled = false;
                return;
            }

            cboLevel2.Enabled = !fieldData.Name.Equals("None");
            cboLevel3.Enabled = !fieldData.Name.Equals("None");
            cboLevel4.Enabled = !fieldData.Name.Equals("None");
        }

        private void cboLevel2_SelectedIndexChanged(object sender, EventArgs e)
        {
            FieldData fieldData = cboLevel2.SelectedItem as FieldData;

            if (fieldData == null)
                return;

            cboLevel3.Enabled = !fieldData.Name.Equals("None");
            cboLevel4.Enabled = !fieldData.Name.Equals("None");
        }

        private void cboLevel3_SelectedIndexChanged(object sender, EventArgs e)
        {
            FieldData fieldData = cboLevel3.SelectedItem as FieldData;

            if (fieldData == null)
                return;

            cboLevel4.Enabled = !fieldData.Name.Equals("None");
        }

        private void cboLevel4_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
    }
}