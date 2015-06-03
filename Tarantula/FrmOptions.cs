using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tarantula
{
    public partial class FrmOptions : Form
    {
        FrmMain frmMain;
        OptionsInfo optionsInfo;

        public FrmOptions(FrmMain frmMain, OptionsInfo optionsInfo)
        {
            this.frmMain = frmMain;
            this.optionsInfo = optionsInfo;

            InitializeComponent();
            txtNoOfLinesInCode.Text = optionsInfo.NumberOfCodeLines.ToString();
            txtTotalRec.Text = optionsInfo.NumberOfExecutions.ToString();
            txtTotalFailRuns.Text = optionsInfo.TotalNumberOfFailRuns.ToString();
            txtTotalPassRuns.Text = optionsInfo.TotalNumberOfPassRuns.ToString();

            cbDistanceMethod.SelectedIndex = (int)optionsInfo.distanceMethod;

            rbPercentage.Checked = !optionsInfo.isAverageThreshold;
            rbAverage.Checked = optionsInfo.isAverageThreshold;
            if (optionsInfo.isAverageThreshold)
            {
                nudFailThreshold.Value = nudPassThreshold.Value = 50;
            }
            else
            {
                nudPassThreshold.Value = (decimal)optionsInfo.PassThresholdPercent;
                nudFailThreshold.Value = (decimal)optionsInfo.FailThresholdPercent;
            }
            nudPassThreshold.Enabled = nudFailThreshold.Enabled = rbPercentage.Checked;

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            optionsInfo.distanceMethod = (DistanceMethod)cbDistanceMethod.SelectedIndex;

            optionsInfo.isAverageThreshold = rbAverage.Checked;
            optionsInfo.PassThresholdPercent=(int)nudPassThreshold.Value;
            optionsInfo.FailThresholdPercent=(int)nudFailThreshold.Value;

            string method = rbAverage.Text;
            if (rbPercentage.Checked) method = rbPercentage.Text;

            frmMain.SetOptionsInfo(optionsInfo, method);
            frmMain.SetLog("Options have been changed successfully...\r\nPass Threshold: " + optionsInfo.PassThreshold.ToString() + "\r\nFail Threshold: " + optionsInfo.FailThreshold.ToString() + "\r\nDistance Algorithm: " + optionsInfo.distanceMethod.ToString(), true);
            this.Close();
        }

        private void rbAverage_CheckedChanged(object sender, EventArgs e)
        {
            nudFailThreshold.Value = nudPassThreshold.Value = 50;
            nudPassThreshold.Enabled = nudFailThreshold.Enabled = rbPercentage.Checked;
        }
    }
}
