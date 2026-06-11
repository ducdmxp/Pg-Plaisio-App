using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Microsoft.VisualBasic;
using Convert2DTo3D.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace CheckPanelProject.UI
{
    public partial class ConnectFCUFrm : System.Windows.Forms.Form
    {
        private List<ObjectItem> ListDuctTypeInput = new List<ObjectItem>();

        private List<ObjectItem> ListSystemTypeInput = new List<ObjectItem>();

        private List<ObjectItem> ListFamilys = new List<ObjectItem>();

        private List<ItemFamilySymbol> ListSimili = new List<ItemFamilySymbol>();

        private List<ItemFamilySymbol> ListHopGio = new List<ItemFamilySymbol>();

        public int ConnectionMode
        {
            get
            {
                if (rdbInput.Checked)
                    return 0;
                else if (rdbOutput.Checked)
                    return 1;
                else
                    return 2;
            }
            set
            {
                if (value == 0)
                    rdbInput.Checked = true;
                else if (value == 1)
                    rdbOutput.Checked = true;
                else
                    rdbAll.Checked = true;
            }
        }

        public int TypeConnectInput
        {
            get
            {
                return cboTypeConnectInput.SelectedIndex;
            }
            set
            {
                cboTypeConnectInput.SelectedIndex = value;
            }
        }

        public ElementId SystemTypeIdInput
        {
            get
            {
                return (cboSystemTypeInput.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public FamilySymbol SymbolSimiliIdInput
        {
            get
            {
                return (cboTypeSimiliInput.SelectedItem as ItemFamilySymbol).Symbol;
            }
        }

        public FamilySymbol SymboHopGioIdInput
        {
            get
            {
                return (cboTypeHopGioInput.SelectedItem as ItemFamilySymbol).Symbol;
            }
        }

        public ElementId DuctTypeIdInput
        {
            get
            {
                return (cboDuctTypeInput.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public double WidthInput
        {
            get
            {
                double width = 0.0;
                if (double.TryParse(cboWidthInput.Text, out width) == false)
                    return 0;
                return Math.Abs(width);
            }
            set
            {
                cboWidthInput.Text = value.ToString();
            }
        }

        public double HeightInput
        {
            get
            {
                double height = 0.0;
                if (double.TryParse(cboHeightInput.Text, out height) == false)
                    return 0;
                return Math.Abs(height);
            }
            set
            {
                cboHeightInput.Text = value.ToString();
            }
        }

        public double LenghtInput
        {
            get
            {
                double lenght = 0.0;
                if (double.TryParse(txtLenghtInput.Text, out lenght) == false)
                    return 0;

                return Math.Abs(lenght);
            }
            set
            {
                txtLenghtInput.Text = value.ToString();
            }
        }

        // Output

        public int TypeConnectOutput
        {
            get
            {
                return cboTypeConnectOutput.SelectedIndex;
            }
            set
            {
                cboTypeConnectOutput.SelectedIndex = value;
            }
        }

        public ElementId SystemTypeIdOutput
        {
            get
            {
                return (cboSystemTypeOutput.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public FamilySymbol SymbolSimiliIdOutput
        {
            get
            {
                return (cboTypeSimiliOutput.SelectedItem as ItemFamilySymbol).Symbol;
            }
        }

        public FamilySymbol SymboHopGioIdOutput
        {
            get
            {
                return (cboTypeHopGioOutput.SelectedItem as ItemFamilySymbol).Symbol;
            }
        }

        public ElementId DuctTypeIdOutput
        {
            get
            {
                return (cboDuctTypeOutput.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public double WidthOutput
        {
            get
            {
                double width = 0.0;
                if (double.TryParse(cboWidthOutput.Text, out width) == false)
                    return 0;
                return Math.Abs(width);
            }
            set
            {
                cboWidthOutput.Text = value.ToString();
            }
        }

        public double HeightOutput
        {
            get
            {
                double height = 0.0;
                if (double.TryParse(cboHeightOutput.Text, out height) == false)
                    return 0;
                return Math.Abs(height);
            }
            set
            {
                cboHeightOutput.Text = value.ToString();
            }
        }

        public double LenghtOutput
        {
            get
            {
                double lenght = 0.0;
                if (double.TryParse(txtLenghtOutput.Text, out lenght) == false)
                    return 0;

                return Math.Abs(lenght);
            }
            set
            {
                txtLenghtOutput.Text = value.ToString();
            }
        }

        //  private ConfigDiffuserConnectionData data = new ConfigDiffuserConnectionData();

        public ConnectFCUFrm()
        {
            InitializeComponent();
            InitData();

            // rdb_CheckedChanged(null, null);
        }

        public void InitData()
        {
            cboTypeConnectInput.SelectedIndex = 0;
            cboTypeConnectOutput.SelectedIndex = 0;

            // System Type
            List<MechanicalSystemType> mePSystems = new FilteredElementCollector(Global.UIDoc.Document)
                .OfClass(typeof(MechanicalSystemType))
                .Cast<MechanicalSystemType>()
                .OrderBy(x => x.Name).ToList();

            if (mePSystems != null && mePSystems.Count > 0)
            {
                ListSystemTypeInput = mePSystems.Select(x => new ObjectItem(x.Name, x.Id)).ToList();

                cboSystemTypeInput.Items.Clear();
                cboSystemTypeOutput.Items.Clear();

                foreach (var item in ListSystemTypeInput)
                {
                    cboSystemTypeInput.Items.Add(item);
                    cboSystemTypeOutput.Items.Add(item);
                }

                cboSystemTypeInput.SelectedIndex = 0;
                cboSystemTypeOutput.SelectedIndex = 0;
            }

            List<FamilySymbol> lstFamilySymbols = new FilteredElementCollector(Global.UIDoc.Document)
                        // .OfCategory(BuiltInCategory.OST_DuctAccessory)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory
                        || x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_MechanicalEquipment)
                        .OrderBy(x => x.Name).ToList();

            List<Family> lstFamilies = lstFamilySymbols.Select(x => x.Family)
                      .GroupBy(x => x.Id).Select(x => x.FirstOrDefault())
                      .OrderBy(x => x.Name).ToList();

            ListSimili = lstFamilySymbols.Select(x => new ItemFamilySymbol(x, false)).ToList();

            ListHopGio = lstFamilySymbols.Select(x => new ItemFamilySymbol(x, false)).ToList();

            if (lstFamilies.Count > 0)
            {
                cboFamilySimiliInput.Items.Clear();
                cboFamilyHopGioInput.Items.Clear();

                cboFamilySimiliOutput.Items.Clear();
                cboFamilyHopGioOutput.Items.Clear();

                ListFamilys = lstFamilies.Select(x => new ObjectItem(x.Name, x.Id)).ToList();

                foreach (var item in ListFamilys)
                {
                    cboFamilySimiliInput.Items.Add(item);
                    cboFamilyHopGioInput.Items.Add(item);

                    cboFamilySimiliOutput.Items.Add(item);
                    cboFamilyHopGioOutput.Items.Add(item);
                }
            }

            // Duct Type
            List<DuctType> ductTypes = new FilteredElementCollector(Global.UIDoc.Document)
                .OfClass(typeof(DuctType))
                .Cast<DuctType>()
                .Where(x => x.Shape == ConnectorProfileType.Rectangular)
                .OrderBy(x => x.Name).ToList();

            if (ductTypes != null && ductTypes.Count > 0)
            {
                ListDuctTypeInput = ductTypes.Select(x => new ObjectItem(x.Name, x.Id)).ToList();

                cboDuctTypeInput.Items.Clear();
                cboDuctTypeOutput.Items.Clear();

                foreach (var item in ListDuctTypeInput)
                {
                    cboDuctTypeInput.Items.Add(item);
                    cboDuctTypeOutput.Items.Add(item);
                }

                cboDuctTypeInput.SelectedIndex = 0;
                cboDuctTypeOutput.SelectedIndex = 0;
            }

            AddDuctSize();

            GetSettingData();
        }

        private void SaveSettingData()
        {
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, "ConnectionMode", ConnectionMode.ToString());

            //input
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeConnectInput.Name, cboTypeConnectInput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboSystemTypeInput.Name, cboSystemTypeInput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboFamilySimiliInput.Name, cboFamilySimiliInput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeSimiliInput.Name, cboTypeSimiliInput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboFamilyHopGioInput.Name, cboFamilyHopGioInput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeHopGioInput.Name, cboTypeHopGioInput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboDuctTypeInput.Name, cboDuctTypeInput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboWidthInput.Name, cboWidthInput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboHeightInput.Name, cboHeightInput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, txtLenghtInput.Name, txtLenghtInput.Text.ToString());

            //output

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeConnectOutput.Name, cboTypeConnectOutput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboSystemTypeOutput.Name, cboSystemTypeOutput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboFamilySimiliOutput.Name, cboFamilySimiliOutput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeSimiliOutput.Name, cboTypeSimiliOutput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboFamilyHopGioOutput.Name, cboFamilyHopGioOutput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeHopGioOutput.Name, cboTypeHopGioOutput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboDuctTypeOutput.Name, cboDuctTypeOutput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboWidthOutput.Name, cboWidthOutput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboHeightOutput.Name, cboHeightOutput.Text.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, txtLenghtOutput.Name, txtLenghtOutput.Text.ToString());
        }

        private void GetSettingData()
        {
            string value = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, "ConnectionMode", "2");
            if (int.TryParse(value, out int mode))
                ConnectionMode = mode;

            //input
            cboTypeConnectInput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeConnectInput.Name, "Duct");
            cboSystemTypeInput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboSystemTypeInput.Name, ListSystemTypeInput.FirstOrDefault()?.Name ?? string.Empty);
            cboFamilySimiliInput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboFamilySimiliInput.Name, ListFamilys.FirstOrDefault()?.Name ?? string.Empty);
            cboTypeSimiliInput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeSimiliInput.Name, ListSimili.FirstOrDefault()?.Name ?? string.Empty);
            cboFamilyHopGioInput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboFamilyHopGioInput.Name, ListFamilys.FirstOrDefault()?.Name ?? string.Empty);
            cboTypeHopGioInput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeHopGioInput.Name, ListHopGio.FirstOrDefault()?.Name ?? string.Empty);
            cboDuctTypeInput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboDuctTypeInput.Name, ListDuctTypeInput.FirstOrDefault()?.Name ?? string.Empty);
            cboWidthInput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboWidthInput.Name, "600");
            cboHeightInput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboHeightInput.Name, "300");
            txtLenghtInput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, txtLenghtInput.Name, "1000");

            //output
            cboTypeConnectOutput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeConnectOutput.Name, "Duct");
            cboSystemTypeOutput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboSystemTypeOutput.Name, ListSystemTypeInput.FirstOrDefault()?.Name ?? string.Empty);
            cboFamilySimiliOutput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboFamilySimiliOutput.Name, ListFamilys.FirstOrDefault()?.Name ?? string.Empty);
            cboTypeSimiliOutput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeSimiliOutput.Name, ListSimili.FirstOrDefault()?.Name ?? string.Empty);
            cboFamilyHopGioOutput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboFamilyHopGioOutput.Name, ListFamilys.FirstOrDefault()?.Name ?? string.Empty);
            cboTypeHopGioOutput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboTypeHopGioOutput.Name, ListHopGio.FirstOrDefault()?.Name ?? string.Empty);
            cboDuctTypeOutput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboDuctTypeOutput.Name, ListDuctTypeInput.FirstOrDefault()?.Name ?? string.Empty);
            cboWidthOutput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboWidthOutput.Name, "600");
            cboHeightOutput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboHeightOutput.Name, "300");
            txtLenghtOutput.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, txtLenghtOutput.Name, "1000");
        }

        /// <summary>
        /// Prevent invalid copy paste
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void PreventCopy(object sender, KeyEventArgs e)
        {
            if (e.Control == true && e.KeyCode == Keys.V)
            {
                if (!double.TryParse(Clipboard.GetText(), out double result) ||
                    Clipboard.GetText().Contains(",") || Clipboard.GetText().Contains("-"))

                    e.SuppressKeyPress = true;
            }
        }

        /// <summary>
        /// Validate textbox digit input
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        private bool IsError()
        {
            return false;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                if (!IsError())
                {
                    SaveSettingData();

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
            catch (Exception)
            {
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            if (btnPreview.Text == "Preview")
            {
                this.Size = this.MaximumSize;
                btnPreview.Text = "Preview >>";
            }
            else if (btnPreview.Text == "Preview >>")
            {
                this.Size = this.MinimumSize;
                btnPreview.Text = "Preview";
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

        private void rdb_CheckedChanged(object sender, EventArgs e)
        {
            grbInput.Enabled = (true == rdbInput.Checked || true == rdbAll.Checked);
            grbOutput.Enabled = (true == rdbOutput.Checked || true == rdbAll.Checked);
        }

        private void cboTypeConnectInput_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboFamilySimiliInput.Enabled = cboTypeConnectInput.SelectedIndex == 0;

            cboTypeSimiliInput.Enabled = cboTypeConnectInput.SelectedIndex == 0;

            cboFamilyHopGioInput.Enabled = cboTypeConnectInput.SelectedIndex == 0;

            cboTypeHopGioInput.Enabled = cboTypeConnectInput.SelectedIndex == 0;

            cboDuctTypeInput.Enabled = cboTypeConnectInput.SelectedIndex == 1;

            cboWidthInput.Enabled = cboTypeConnectInput.SelectedIndex == 1;

            cboHeightInput.Enabled = cboTypeConnectInput.SelectedIndex == 1;

            txtLenghtInput.Enabled = cboTypeConnectInput.SelectedIndex == 1;
        }

        private void AddDuctSize(DuctShape shape = DuctShape.Rectangular)
        {
            var settings = DuctSizeSettings.GetDuctSizeSettings(Global.UIDoc.Document);

            foreach (KeyValuePair<DuctShape, DuctSizes> keyPair in settings)
            {
                if (keyPair.Key != shape)
                    continue;

                if (keyPair.Key == DuctShape.Rectangular)
                {
                    cboWidthInput.Items.Clear();
                    cboHeightInput.Items.Clear();

                    cboWidthOutput.Items.Clear();
                    cboHeightOutput.Items.Clear();

                    foreach (MEPSize size in keyPair.Value)
                    {
                        var value = Math.Round(size.NominalDiameter * 304.8);

                        cboWidthInput.Items.Add(value);
                        cboHeightInput.Items.Add(value);

                        cboWidthOutput.Items.Add(value);
                        cboHeightOutput.Items.Add(value);
                    }

                    if (cboWidthInput.SelectedItem == null && cboWidthInput.Items.Count != 0)
                        cboWidthInput.SelectedIndex = 0;

                    if (cboHeightInput.SelectedItem == null && cboHeightInput.Items.Count != 0)
                        cboHeightInput.SelectedIndex = 0;

                    if (cboWidthOutput.SelectedItem == null && cboWidthOutput.Items.Count != 0)
                        cboWidthOutput.SelectedIndex = 0;

                    if (cboHeightOutput.SelectedItem == null && cboHeightOutput.Items.Count != 0)
                        cboHeightOutput.SelectedIndex = 0;
                }
            }
        }

        private void cboTypeConnectOutput_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboFamilySimiliOutput.Enabled = cboTypeConnectOutput.SelectedIndex == 0;

            cboTypeSimiliOutput.Enabled = cboTypeConnectOutput.SelectedIndex == 0;

            cboFamilyHopGioOutput.Enabled = cboTypeConnectOutput.SelectedIndex == 0;

            cboTypeHopGioOutput.Enabled = cboTypeConnectOutput.SelectedIndex == 0;

            cboDuctTypeOutput.Enabled = cboTypeConnectOutput.SelectedIndex == 1;

            cboWidthOutput.Enabled = cboTypeConnectOutput.SelectedIndex == 1;

            cboHeightOutput.Enabled = cboTypeConnectOutput.SelectedIndex == 1;

            txtLenghtOutput.Enabled = cboTypeConnectOutput.SelectedIndex == 1;
        }

        private void cboFamilySimiliInput_SelectedIndexChanged(object sender, EventArgs e)
        {
            ObjectItem selected = cboFamilySimiliInput.SelectedItem as ObjectItem;
            if (selected == null) return;

            List<ItemFamilySymbol> itemFamilySymbols = ListSimili.Where(x => x.Symbol.Family.Id == selected.ObjectId).ToList();

            if (itemFamilySymbols.Count <= 0) return;

            cboTypeSimiliInput.Items.Clear();

            foreach (var item in itemFamilySymbols)
            {
                cboTypeSimiliInput.Items.Add(item);
            }

            cboTypeSimiliInput.SelectedIndex = 0;
        }

        private void cboFamilySimiliOutput_SelectedIndexChanged(object sender, EventArgs e)
        {
            ObjectItem selected = cboFamilySimiliOutput.SelectedItem as ObjectItem;
            if (selected == null) return;

            List<ItemFamilySymbol> itemFamilySymbols = ListSimili.Where(x => x.Symbol.Family.Id == selected.ObjectId).ToList();

            if (itemFamilySymbols.Count <= 0) return;

            cboTypeSimiliOutput.Items.Clear();

            foreach (var item in itemFamilySymbols)
            {
                cboTypeSimiliOutput.Items.Add(item);
            }

            cboTypeSimiliOutput.SelectedIndex = 0;
        }

        private void cboFamilyHopGioInput_SelectedIndexChanged(object sender, EventArgs e)
        {
            ObjectItem selected = cboFamilyHopGioInput.SelectedItem as ObjectItem;
            if (selected == null) return;

            List<ItemFamilySymbol> itemFamilySymbols = ListHopGio.Where(x => x.Symbol.Family.Id == selected.ObjectId).ToList();

            if (itemFamilySymbols.Count <= 0) return;

            cboTypeHopGioInput.Items.Clear();

            foreach (var item in itemFamilySymbols)
            {
                cboTypeHopGioInput.Items.Add(item);
            }

            cboTypeHopGioInput.SelectedIndex = 0;
        }

        private void cboFamilyHopGioOutput_SelectedIndexChanged(object sender, EventArgs e)
        {
            ObjectItem selected = cboFamilyHopGioOutput.SelectedItem as ObjectItem;
            if (selected == null) return;

            List<ItemFamilySymbol> itemFamilySymbols = ListHopGio.Where(x => x.Symbol.Family.Id == selected.ObjectId).ToList();

            if (itemFamilySymbols.Count <= 0) return;

            cboTypeHopGioOutput.Items.Clear();

            foreach (var item in itemFamilySymbols)
            {
                cboTypeHopGioOutput.Items.Add(item);
            }

            cboTypeHopGioOutput.SelectedIndex = 0;
        }

        //public void MakeRequest(RequestId request)
        //{
        //    m_handler.Request.Make(request);
        //    m_exEvent.Raise();
        //}
    }

    public class ObjectItem
    {
        private string _name;
        private ElementId _objectId = ElementId.InvalidElementId;

        private string Guid = null;

        public ObjectItem(string name, ElementId objectId)
        {
            _name = name;
            _objectId = objectId;
        }

        public ObjectItem(string name, string guid)
        {
            _name = name;
            Guid = guid;
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public ElementId ObjectId
        {
            get { return _objectId; }
            set { _objectId = value; }
        }

        public override string ToString()
        {
            return _name;
        }
    }

    internal class ItemFamilySymbol
    {
        private string _name;

        private FamilySymbol _familySymbol;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public FamilySymbol Symbol
        {
            get { return _familySymbol; }
            set { _familySymbol = value; }
        }

        public ItemFamilySymbol(FamilySymbol symbol, bool isShowFamilyName = true)
        {
            Symbol = symbol;

            Name = (isShowFamilyName) ? symbol.FamilyName + " : " + symbol.Name : symbol.Name;
        }

        public override string ToString()
        {
            return Name.ToString();
        }
    }
}