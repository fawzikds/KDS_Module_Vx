using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Windows.Forms;

using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using System.Xml.Linq;
//using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Diagnostics;

namespace KDS_Module_Vx
{
    public partial class insertSleeve_frm_CL : System.Windows.Forms.Form
    {

        // Declare an External Event Handeler, which  is our class2 (ButtonEE7Parameter). We need to activate it, see after initializeComponent below.
        //ButtonEE7Parameter myEE7Parameter;
        ExternalEvent myEE7ActionParameter;
        // end of declaration of External Event Handler

        public Document doc { get; set; }
        public insertSleeve_frm_CL()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void SysNm_lbl_Click(object sender, EventArgs e)
        {

        }

        private void OK_btn_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void Cancel_btn_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }



        private void SlctAllNone_chkbx_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < modelNames_chkLstBx.Items.Count; i++)
            {
                // I negate the current status... So if True, it becomes False, and vice versa
                modelNames_chkLstBx.SetItemChecked(i, !SlctAllNone_chkbx.Checked);
            }
        }

        private void modelNames_chkLstBx_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }  // End Of Class insertSleeve_frm
}  // End Of Name Space
