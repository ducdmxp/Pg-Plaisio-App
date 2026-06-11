using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using CheckPanelProject.UI;
using Microsoft.VisualBasic;
using Convert2DTo3D.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Convert2DTo3D.UI
{
    public partial class ConnectWCFrm : System.Windows.Forms.Form
    {
        public int SelectedType
        {
            get
            {
                if (rdbType1.Checked)
                    return 0;
                else if (rdbType2.Checked)
                    return 1;
                else
                    return 2;
            }
            set
            {
                switch (value)
                {
                    case 0:
                        rdbType1.Checked = true;
                        break;

                    case 1:
                        rdbType2.Checked = true;
                        break;

                    case 2:
                        rdbType3.Checked = true;
                        break;

                    default:
                        rdbType1.Checked = true;
                        break;
                }
            }
        }

        public bool IsTee
        {
            get
            {
                return rdbTee.Checked;
            }
            set
            {
                rdbTee.Checked = value;
                rdbElbow.Checked = !value;
            }
        }

        public FamilySymbol SymbolCoren
        {
            get
            {
                return (cboCorenType.SelectedItem as ItemFamilySymbol).Symbol;
            }
        }

        public double Diameter
        {
            get
            {
                double diameter = 0.0;
                if (double.TryParse(cboPipeSize.Text, out diameter) == false)
                    return 0;
                return Math.Abs(diameter);
            }
        }

        public bool IsElbow45
        {
            get { return rdbCo45.Checked; }
            set
            {
                rdbCo45.Checked = value;
                rdbCo90.Checked = !value;
            }
        }

        public bool IsProject
        {
            get { return rdbProject.Checked; }
            set
            {
                rdbProject.Checked = value;
                rdbLinkRevit.Checked = !value;
            }
        }

        public bool IsPickWall
        {
            get { return rdbPickWall.Checked; }
            set
            {
                rdbPickWall.Checked = value;
                rdbPickLine.Checked = !value;
            }
        }

        public double Offset
        {
            get
            {
                if (rdbType1.Checked == false)
                    return 0.0;

                double offset = 0.0;
                if (double.TryParse(txtOffset.Text, out offset) == false)
                    return 0;

                if (false == chkOffsetBranch.Checked)
                    return 0;

                return Math.Abs(offset);
            }
            set
            {
                txtOffset.Text = value.ToString();
            }
        }

        public double OffsetLevel
        {
            get
            {
                double offset = 0.0;
                if (double.TryParse(txtOffsetLevel.Text, out offset) == false)
                    return 0;

                if (false == chkOffsetCoren.Checked)
                    return 0;

                return Math.Abs(offset);
            }
            set
            {
                txtOffsetLevel.Text = value.ToString();
            }
        }

        private List<ItemFamilySymbol> ListCorenType = new List<ItemFamilySymbol>();

        private List<ObjectItem> ListPipeType = new List<ObjectItem>();

        public ConnectWCFrm()
        {
            InitializeComponent();

            InitData();
        }

        private void InitData()
        {
            List<FamilySymbol> lstFamilySymbols = new FilteredElementCollector(Global.UIDoc.Document)
                       .OfCategory(BuiltInCategory.OST_PipeFitting)
                       .OfClass(typeof(FamilySymbol))
                       .Cast<FamilySymbol>()
                       .OrderBy(x => x.Name).ToList();

            if (lstFamilySymbols.Count > 0)
            {
                ListCorenType = lstFamilySymbols.Select(x => new ItemFamilySymbol(x)).ToList();

                foreach (var item in ListCorenType)
                {
                    cboCorenType.Items.Add(item);
                }

                cboCorenType.SelectedIndex = 0;
            }

            // Pipe Type
            cboPipeType.Items.Clear();
            List<PipeType> pipeTypes = new FilteredElementCollector(Global.UIDoc.Document)
                .OfClass(typeof(PipeType))
                .Cast<PipeType>()
                .OrderBy(x => x.Name).ToList();

            if (pipeTypes != null && pipeTypes.Count > 0)
            {
                foreach (PipeType type in pipeTypes)
                {
                    ObjectItem item = new ObjectItem(type.Name, type.Id);

                    ListPipeType.Add(item);
                }

                cboPipeType.DataSource = ListPipeType;
            }

            GetSetting();

            chkOffsetBranch_CheckedChanged(null, null);
            chkOffsetCoren_CheckedChanged(null, null);
        }

        private void SaveSetting()
        {
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, "SelectedType", SelectedType.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, rdbCo45.Name, rdbCo45.Checked.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, rdbTee.Name, rdbTee.Checked.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboCorenType.Name, cboCorenType.Text);
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboPipeType.Name, cboPipeType.Text);
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboPipeSize.Name, cboPipeSize.Text);

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, rdbProject.Name, rdbProject.Checked.ToString());
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, rdbPickWall.Name, rdbPickWall.Checked.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, chkOffsetBranch.Name, chkOffsetBranch.Checked.ToString());
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, chkOffsetCoren.Name, chkOffsetCoren.Checked.ToString());

            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, txtOffset.Name, txtOffset.Text.ToString());
            Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, txtOffsetLevel.Name, txtOffsetLevel.Text.ToString());
        }

        private void GetSetting()
        {
            string strSelectedType = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, "SelectedType", "0");

            int.TryParse(strSelectedType, out int selectedType);

            SelectedType = selectedType;

            //

            string strIsCo45 = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, rdbCo45.Name, "true");
            bool.TryParse(strIsCo45, out bool vIsCo45);
            IsElbow45 = vIsCo45;

            //

            string strIsTee = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, rdbTee.Name, "true");
            bool.TryParse(strIsTee, out bool vIsTee);
            IsTee = vIsTee;

            //

            cboCorenType.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboCorenType.Name, ListCorenType.FirstOrDefault()?.Name ?? string.Empty);

            //

            cboPipeType.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboPipeType.Name, ListPipeType.FirstOrDefault()?.Name ?? string.Empty);

            cboPipeSize.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, cboPipeSize.Name, cboPipeSize.Items[0].ToString() ?? string.Empty);

            //
            string strIsProject = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, rdbProject.Name, "true");
            bool.TryParse(strIsProject, out bool vIsProject);
            IsProject = vIsProject;

            string strIsPickWall = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, rdbPickWall.Name, "true");
            bool.TryParse(strIsPickWall, out bool vIsPickWall);
            IsPickWall = vIsPickWall;

            string strChkOffsetBranch = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, chkOffsetBranch.Name, "");

            bool.TryParse(strChkOffsetBranch, out bool vChkOffsetBranch);

            chkOffsetBranch.Checked = vChkOffsetBranch;

            string strChkOffsetCoren = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, chkOffsetCoren.Name, "");
            bool.TryParse(strChkOffsetCoren, out bool vChkOffsetCoren);
            chkOffsetCoren.Checked = vChkOffsetCoren;

            txtOffset.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, txtOffset.Name, "");
            txtOffsetLevel.Text = Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, this.Name, txtOffsetLevel.Name, "");
        }

        //public void MakeRequest(RequestId request)
        //{
        //    m_handler.Request.Make(request);
        //    m_exEvent.Raise();
        //}

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

        private void chkOffsetBranch_CheckedChanged(object sender, EventArgs e)
        {
            txtOffset.Enabled = chkOffsetBranch.Checked && chkOffsetBranch.Enabled;
        }

        private void chkOffsetCoren_CheckedChanged(object sender, EventArgs e)
        {
            txtOffsetLevel.Enabled = chkOffsetCoren.Checked;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            SaveSetting();

            this.DialogResult = DialogResult.OK;
        }

        private void cboPipeType_SelectedIndexChanged(object sender, EventArgs e)
        {
            ObjectItem objectItem = cboPipeType.SelectedItem as ObjectItem;

            if (objectItem == null)
                return;

            PipeType pipetype = Global.Doc.GetElement(objectItem.ObjectId) as PipeType;

            cboPipeSize.Items.Clear();

            List<double> doubles = GetSizeFromPipeType(Global.Doc, pipetype);
            foreach (var item in doubles)
            {
                cboPipeSize.Items.Add(item.ToString());
            }
            cboPipeSize.SelectedItem = doubles.FirstOrDefault().ToString();
        }

        public static List<double> GetSizeFromPipeType(Document doc, PipeType pipetype)
        {
            List<double> retval = new List<double>();
            if (doc == null || pipetype == null)
                return retval;

            try
            {
                if (pipetype.RoutingPreferenceManager != null)
                {
                    Segment CurrentSegment = null;
                    int count = pipetype.RoutingPreferenceManager.GetNumberOfRules(RoutingPreferenceRuleGroupType.Segments);

                    for (int i = 0; i < count; i++)
                    {
                        var rule = pipetype.RoutingPreferenceManager.GetRule(RoutingPreferenceRuleGroupType.Segments, i);

                        CurrentSegment = doc.GetElement(rule.MEPPartId) as PipeSegment;
                    }

                    if (CurrentSegment != null)
                    {
                        foreach (var mepsize in CurrentSegment.GetSizes())
                        {
                            if (mepsize == null)
                                continue;

                            double size = mepsize.NominalDiameter * 304.8;

                            size = Math.Round(size, 5);

                            retval.Add(size);
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            return retval;
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            if (btnPreview.Text == "Preview >>")
            {
                this.Size = this.MaximumSize;
                btnPreview.Text = "Preview <<";
            }
            else if (btnPreview.Text == "Preview <<")
            {
                this.Size = this.MinimumSize;
                btnPreview.Text = "Preview >>";
            }
        }

        private void rdbType_CheckedChanged(object sender, EventArgs e)
        {
            rdbTee.Enabled = rdbType1.Checked;
            rdbElbow.Enabled = rdbType1.Checked;

            chkOffsetBranch.Enabled = rdbType1.Checked;
            txtOffset.Enabled = rdbType1.Checked;
            grbAccessory.Enabled = rdbType1.Checked;
        }
    }
}