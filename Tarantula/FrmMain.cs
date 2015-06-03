using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel; 

namespace Tarantula
{
    public enum DistanceMethod { EuclideanDistance = 0, ManhattanDistance, HammingDistance };
    public struct OptionsInfo
    {
        public int NumberOfCodeLines;
        public int NumberOfExecutions;
        public int TotalNumberOfFailRuns;
        public int TotalNumberOfPassRuns;

        public DistanceMethod distanceMethod;
        public double PassThresholdPercent;
        public double FailThresholdPercent;

        public double PassThreshold;
        public double FailThreshold;

        public string TestCase;
        public string Version;
        public string FaultName;
        public long ActualFaultLineNo;

        public double ClassicalDiagScore;
        public double OurDiagScore;
        public bool isAverageThreshold;
    }

    public partial class FrmMain : Form
    {
        OptionsInfo optionsInfo;
        private void ResetOptionInfo()
        {
            optionsInfo.NumberOfCodeLines = 0;
            optionsInfo.NumberOfExecutions = 0;
            optionsInfo.TotalNumberOfFailRuns = 0;
            optionsInfo.TotalNumberOfPassRuns = 0;

            optionsInfo.distanceMethod = DistanceMethod.EuclideanDistance;

            optionsInfo.PassThresholdPercent = 50;
            optionsInfo.FailThresholdPercent = 50;

            optionsInfo.PassThreshold = 0;
            optionsInfo.FailThreshold = 0;

            optionsInfo.ClassicalDiagScore = optionsInfo.OurDiagScore = 0;

            optionsInfo.TestCase = optionsInfo.Version = optionsInfo.FaultName = string.Empty;

            optionsInfo.ActualFaultLineNo = 0;
            optionsInfo.ClassicalDiagScore = 0;
            optionsInfo.OurDiagScore = 0;
            optionsInfo.isAverageThreshold = true;
        }

        public void SetOptionsInfo(OptionsInfo optionsInfo, string ThresholdMethod)
        {
            this.optionsInfo = optionsInfo;
            lblThresholdMethodUsed.Text = ThresholdMethod;
        }

        private void setButtonsActivate(bool btnImportDataSt, bool btnImportMetaDataSt, bool btnProcessOptionsSt, bool btnProcessTarantulaSt, bool btnLoadResultSt, bool btnSaveResultsSt, bool btnNewSt)
        {
            btnImportData.Enabled = btnImportDataSt;
            btnImportMetaData.Enabled = btnImportMetaDataSt;
            btnProcessOptions.Enabled = btnProcessOptionsSt;
            btnProcessClassicTarantula.Enabled = btnProcessTarantulaSt;
            btnLoadResult.Enabled = btnLoadResultSt;
            btnSaveResults.Enabled = btnSaveResultsSt;
            btnNew.Enabled = btnNewSt;
        }

        public FrmMain()
        {
            InitializeComponent();

            setButtonsActivate(true, false, false, false, true, true, false);
            ResetOptionInfo();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void MakeNewOperation(bool ClearLog, bool ClearMetaData,bool ClearResults)
        {
            if (ClearLog) txtLog.Clear();

            dgvTrainingData.DataSource = null;
            dgvTrainingData.Rows.Clear();
            dgvTrainingData.Refresh();

            if (ClearMetaData)
            {
                dgvMetaData.DataSource = null;
                dgvMetaData.Rows.Clear();
                dgvMetaData.Refresh();
            }

            dgvDistancesAll.DataSource = null;
            dgvDistancesAll.Rows.Clear();
            dgvDistancesAll.Refresh();
           
            dgvDistancesShort.DataSource = null;
            dgvDistancesShort.Rows.Clear();
            dgvDistancesShort.Refresh();

            dgvHueSuspisiousClassic.DataSource = null;
            dgvHueSuspisiousClassic.Rows.Clear();
            dgvHueSuspisiousClassic.Refresh();

            dgvHueSuspisiousOur.DataSource = null;
            dgvHueSuspisiousOur.Rows.Clear();
            dgvHueSuspisiousOur.Refresh();

            if (ClearResults)
            {
                dgvResults.DataSource = null;
                dgvResults.Rows.Clear();
                dgvResults.Refresh();
            }
            
            txtCodeLineNumbers.Text = "0";
            txtNoOfExecutions.Text = "0";
            txtNoOfExecutions.Text = "0";
            txtPassRunsCount.Text = "0";
            txtFailRunsCount.Text = "0";

            txtActualFaultLineNo.Text = txtClassicalDiagnosisScore.Text = txtOurDiagnosisScore.Text = txtTestCase.Text = txtVersion.Text = txtFaultName.Text = string.Empty;

            txtNoShortenDistances.Text = txtNoTotalDistances.Text = "0";

            txtTotalNoEvaluations.Text = "0";
            txtBetterMethod.Text = "Both Equal";
            pbSad.Visible = pbHappy.Visible = false;

            ResetOptionInfo();
            lblMessage.Text = "-";
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("All operation and results will have been erased...\r\nAre you sure?", "New operation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                setButtonsActivate(true, false, false, false, true, true, false);
                MakeNewOperation(true, true, true);
            }
        }

        private void btnProcessOptions_Click(object sender, EventArgs e)
        {
            FrmOptions frmOptions = new FrmOptions(this, optionsInfo);
            frmOptions.ShowDialog(this);
        }

        public void SetLog(string logText, bool putSeperator)
        {
            txtLog.Text += logText + "\r\n";
            if (putSeperator) txtLog.Text += "=============================================================\r\n";
        }

        private void btnImportData_Click(object sender, EventArgs e)
        {
            long fileSize = 0;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.DataSet DtSet;
            System.Data.OleDb.OleDbDataAdapter MyCommand;

            openFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments.ToString();
            openFileDialog1.Filter = "Excel Files (*.xls,*.xlsx)|*.xls;*.xlsx";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.Multiselect = false;
            openFileDialog1.Title = "Import trainig data from excel file";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    tabsControl.SelectedIndex = 0;
                    if (dgvMetaData.Rows.Count <= 0 && dgvResults.Rows.Count <= 0) MakeNewOperation(true, false, true);
                    else MakeNewOperation(true, false, false);
                    tpProgressBar1.Value = 0;
                    //IsExecuted = false;
                    FileInfo fi = new FileInfo(openFileDialog1.FileName);
                    fileSize = fi.Length;

                    string excelObject = "Provider=Microsoft.{0}.OLEDB.{1};Data Source={2};Extended Properties=\"Excel {3};HDR=YES\"";

                    if (fi.Extension.Equals(".xls"))
                    {
                        excelObject = string.Format(excelObject, "Jet", "4.0", openFileDialog1.FileName, "8.0");
                    }
                    else if (fi.Extension.Equals(".xlsx"))
                    {
                        excelObject = string.Format(excelObject, "Ace", "12.0", openFileDialog1.FileName, "12.0");
                    }
                    //string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();

                    MyConnection = new System.Data.OleDb.OleDbConnection(excelObject);
                    MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection); //where results='passing'
                    MyCommand.TableMappings.Add("Table", "TestTable");
                    DtSet = new System.Data.DataSet();

                    MyCommand.Fill(DtSet);

                    dgvTrainingData.DataSource = DtSet.Tables[0];
                    MyConnection.Close();
                    txtNoOfExecutions.Text = dgvTrainingData.Rows.Count.ToString();
                    txtCodeLineNumbers.Text = dgvTrainingData.Rows[0].Cells[2].Value.ToString().Length.ToString();
                    optionsInfo.NumberOfExecutions = dgvTrainingData.Rows.Count;
                    optionsInfo.NumberOfCodeLines = dgvTrainingData.Rows[0].Cells[2].Value.ToString().Length;
                    int index = openFileDialog1.FileName.LastIndexOf('\\') + 1;
                    string filename = openFileDialog1.FileName.Substring(index);
                    index = filename.IndexOf('.');
                    string testCaseName = filename.Substring(0, index);
                    string temp = filename.Substring(index + 1);
                    index = temp.IndexOf('.');
                    string VersionName = temp.Substring(0, index);
                    temp = temp.Substring(index + 1);
                    index = temp.IndexOf('.');
                    string FaultName = temp.Substring(0, index);
                    txtTestCase.Text = optionsInfo.TestCase = testCaseName;
                    txtVersion.Text = optionsInfo.Version = VersionName;
                    txtFaultName.Text = optionsInfo.FaultName = FaultName;

                    dgvTrainingData.Columns[0].Width = 35;
                    dgvTrainingData.Columns[1].Width = 60;
                    dgvTrainingData.Columns[0].Frozen = true;
                    dgvTrainingData.Columns[1].Frozen = true;

                    tpProgressBar1.Value = 100;
                    lblProgressPercentage.Text = "100%";
                    SetLog(">>>> Data has been imported successfully...\r\nFile Name: " + openFileDialog1.FileName + "\r\nTest Case: " + testCaseName + "\r\nVersion: " + VersionName + "\r\nFault Nane: " + FaultName + "\r\nNumber of Executions: " + txtCodeLineNumbers.Text + ", Number of Code Lines: " + txtNoOfExecutions.Text, true);
                    lblMessage.Text = "Training data has been imported successfully...";
                    if (dgvMetaData.Rows.Count <= 0)
                    {
                        MessageBox.Show("Traning data has been imported successfully...\r\nPlease import meta data in order to complete the operation...", "Import Meta Data is needed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        setButtonsActivate(true, true, false, false, true, true, false);
                    }
                    tpProgressBar1.Value = 100;
                    lblProgressPercentage.Text = "100%";
                    optionsInfo.TotalNumberOfPassRuns = 0;
                    for (int counter = 0; counter < (dgvTrainingData.Rows.Count); counter++)
                    {
                        if (dgvTrainingData.Rows[counter].Cells[1].Value != null)
                        {
                            if (dgvTrainingData.Rows[counter].Cells[1].Value.ToString().ToUpper() == "passing".ToUpper())
                            {
                                optionsInfo.TotalNumberOfPassRuns++;
                            }
                        }
                    }
                    optionsInfo.TotalNumberOfFailRuns = optionsInfo.NumberOfExecutions - optionsInfo.TotalNumberOfPassRuns;
                    txtFailRunsCount.Text = optionsInfo.TotalNumberOfFailRuns.ToString();
                    txtPassRunsCount.Text = optionsInfo.TotalNumberOfPassRuns.ToString();

                    optionsInfo.ActualFaultLineNo = 0;
                    txtActualFaultLineNo.Text = string.Empty;
                    txtPassThreshold.Text = txtFailThreshold.Text = string.Empty;
                    for (int i = 0; i < dgvMetaData.Rows.Count; i++)
                    {
                        if (dgvMetaData.Rows[i].Cells[1].Value.ToString() == optionsInfo.TestCase)
                        {
                            if (dgvMetaData.Rows[i].Cells[2].Value.ToString() == optionsInfo.Version)
                            {
                                if (dgvMetaData.Rows[i].Cells[3].Value.ToString() == optionsInfo.FaultName)
                                {
                                    optionsInfo.ActualFaultLineNo = Convert.ToInt32(dgvMetaData.Rows[i].Cells[4].Value.ToString());
                                    txtActualFaultLineNo.Text = optionsInfo.ActualFaultLineNo.ToString();
                                    txtPassThreshold.Text = Math.Round(optionsInfo.PassThreshold,2).ToString();
                                    txtFailThreshold.Text = Math.Round(optionsInfo.FailThreshold,2).ToString();
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void btnImportMetaData_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.DataSet DtSet;
            System.Data.OleDb.OleDbDataAdapter MyCommand;

            openFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments.ToString();
            openFileDialog1.Filter = "Excel Files (*.xls,*.xlsx)|*.xls;*.xlsx";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.Multiselect = false;
            openFileDialog1.Title = "Import meta data from excel file";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    long fileSize = 0;

                    tabsControl.SelectedIndex = 1;
                    //MakeNewOperation();
                    tpProgressBar1.Value = 0;
                    //IsExecuted = false;
                    FileInfo fi = new FileInfo(openFileDialog1.FileName);
                    fileSize = fi.Length;

                    string excelObject = "Provider=Microsoft.{0}.OLEDB.{1};Data Source={2};Extended Properties=\"Excel {3};HDR=YES\"";

                    if (fi.Extension.Equals(".xls"))
                    {
                        excelObject = string.Format(excelObject, "Jet", "4.0", openFileDialog1.FileName, "8.0");
                    }
                    else if (fi.Extension.Equals(".xlsx"))
                    {
                        excelObject = string.Format(excelObject, "Ace", "12.0", openFileDialog1.FileName, "12.0");
                    }
                    //string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();

                    string condition = "WHERE ([Test Case]='" + optionsInfo.TestCase + "') and (Version='" + optionsInfo.Version + "') AND ([Fault Name]='" + optionsInfo.FaultName + "')";
                    MyConnection = new System.Data.OleDb.OleDbConnection(excelObject);
                    MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$] " + condition, MyConnection); //where results='passing'

                    MyCommand.TableMappings.Add("Table", "TestTable");
                    DtSet = new System.Data.DataSet();

                    MyCommand.Fill(DtSet);

                    if (DtSet.Tables[0].Rows.Count <= 0)
                    {
                        MessageBox.Show("Information related to the training dara has not found in the meta data.\r\nPlease try again...", "Error in finding meta data information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        setButtonsActivate(true, true, false, false, true, true, false);
                        MakeNewOperation(false, true, true);
                        SetLog(">>>> Information related to the training dara has not found in the meta data.\r\nPlease try again...", true);
                        tpProgressBar1.Value = 0;
                        tabsControl.SelectedIndex = 0;
                    }
                    else
                    {
                        optionsInfo.ActualFaultLineNo = Convert.ToInt32(DtSet.Tables[0].Rows[0][4].ToString());
                        txtActualFaultLineNo.Text = optionsInfo.ActualFaultLineNo.ToString();

                        txtPassThreshold.Text = optionsInfo.PassThreshold.ToString();
                        txtFailThreshold.Text = optionsInfo.FailThreshold.ToString();
                        MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection); //where results='passing'

                        MyCommand.TableMappings.Add("Table", "TestTable");
                        DtSet = new System.Data.DataSet();

                        MyCommand.Fill(DtSet);

                        dgvMetaData.DataSource = DtSet.Tables[0];
                        MyConnection.Close();

                        tpProgressBar1.Value = 100;
                        lblProgressPercentage.Text = "100%";

                        SetLog(">>>> Meta data has been imported successfully...\r\nActual Faulty Statement is in Line: ", true);
                        lblMessage.Text = "Meta data has been imported successfully...";
                        setButtonsActivate(true, true, true, true, true, true, true);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        //------------------------------------------------
        private void btnProcessClassicTarantula_Click(object sender, EventArgs e)
        {
            pbAnim.Visible = true;
            tabsControl.SelectedIndex = 3;

            dgvDistancesAll.DataSource = null;
            dgvDistancesAll.Rows.Clear();
            dgvDistancesAll.Refresh();

            dgvHueSuspisiousClassic.DataSource = null;
            dgvHueSuspisiousClassic.Rows.Clear();
            dgvHueSuspisiousClassic.Refresh();
            
            CheckForIllegalCrossThreadCalls = false;
            tpProgressBar1.Value = 0;
            bkWorkerClassicalTarantula.RunWorkerAsync();
            tpProgressBar1.Value = 100;
            lblProgressPercentage.Text = "100%";
        }

        public static int XOR(string a, string b)
        {
            int min, max;
            if (b.Length < a.Length)
            {
                min = b.Length;
                max = a.Length;
            }
            else
            {
                max = b.Length;
                min = a.Length;
            }
            int i = 0, r = 0;
            for (i = 0; i < min; i++)
            {
                r += (a[i] - 48) ^ (b[i] - 48);
            }

            for (int j = i; j < max; j++)
            {
                if (b.Length < a.Length) r += a[j] - 48;
                else r += b[j] - 48;

            }
            return r;
        }

        int FFCount, FPCount, TotalIterations;
        long[] PassedCount, FailedCount;

        int FFPrimeCount, FPPrimeCount;
        long[] PassedPrimeCount, FailedPrimeCount;
        double[] hue;
        double PMaxDistance, PMinDistance, FMaxDistance, FMinDistance;

        private void bkWorkerClassicalTarantula_DoWork(object sender, DoWorkEventArgs e)
        {
            bool[] removedList = new bool[optionsInfo.NumberOfExecutions];
            for (int i = 0; i < optionsInfo.NumberOfExecutions; i++) removedList[i] = false;

            PassedCount = new long[optionsInfo.NumberOfCodeLines];
            FailedCount = new long[optionsInfo.NumberOfCodeLines];
            hue = new double[optionsInfo.NumberOfCodeLines];
            FFCount = FPCount = TotalIterations = 0;
            PMaxDistance = FMaxDistance = 0;
            PMinDistance = FMinDistance = double.MaxValue;
            int perc = 0;
            lblMessage.Text = "Calculating Fail -> Fail and Fail -> Pass Distances...";
            double SumFF = 0, SumFP = 0;
            for (int i = 0; i < dgvTrainingData.Rows.Count; i++)
            {
                perc = i * 100 / dgvTrainingData.Rows.Count;
                bkWorkerClassicalTarantula.ReportProgress(perc);

                for (int s = 0; s < dgvTrainingData.Rows[i].Cells[2].Value.ToString().Length; s++)
                {
                    if (dgvTrainingData.Rows[i].Cells[1].Value.ToString().ToUpper() == "passing".ToUpper())
                    {
                        PassedCount[s] += dgvTrainingData.Rows[i].Cells[2].Value.ToString()[s] - 48;
                    }
                    else
                    {
                        FailedCount[s] += dgvTrainingData.Rows[i].Cells[2].Value.ToString()[s] - 48;
                    }
                }

                // --- section 2 ---
                if (dgvTrainingData.Rows[i].Cells[1].Value.ToString().ToUpper() == "failing".ToUpper())
                {
                    string v1 = dgvTrainingData.Rows[i].Cells[2].Value.ToString();
                    for (int j = 0; j < dgvTrainingData.Rows.Count; j++)
                    {
                        TotalIterations++;
                        string v2 = dgvTrainingData.Rows[j].Cells[2].Value.ToString();
                        if (dgvTrainingData.Rows[j].Cells[1].Value.ToString().ToUpper() == "passing".ToUpper())
                        {
                            if (!removedList[j])
                            {
                                int r = XOR(v1, v2);
                                double d = Math.Sqrt(r);
                                //if (d < optionsInfo.PassThreshold)
                                //{
                                if (d > PMaxDistance) PMaxDistance = d;
                                if (d < PMinDistance) PMinDistance = d;
                                FPCount++;
                                SumFP += d;
                                dgvDistancesAll.Rows.Add(dgvTrainingData.Rows[j].Cells[0].Value.ToString(), dgvTrainingData.Rows[j].Cells[1].Value.ToString(), d, dgvTrainingData.Rows[j].Cells[2].Value.ToString());
                                //}
                                removedList[j] = true;
                            }
                        }
                    }

                    for (int j = i; j < dgvTrainingData.Rows.Count; j++)
                    {
                        TotalIterations++;
                        string v2 = dgvTrainingData.Rows[j].Cells[2].Value.ToString();
                        if (dgvTrainingData.Rows[j].Cells[1].Value.ToString().ToUpper() == "failing".ToUpper())
                        {
                            if (!removedList[j])
                            {
                                int r = XOR(v1, v2);
                                double d = Math.Sqrt(r);
                                //if (d > optionsInfo.FailThreshold)
                                //{
                                if (d > FMaxDistance) FMaxDistance = d;
                                if (d < FMinDistance) FMinDistance = d;

                                FFCount++;
                                SumFF += d;
                                //distances.FFdistances.Add(d);
                                dgvDistancesAll.Rows.Add(dgvTrainingData.Rows[j].Cells[0].Value.ToString(), dgvTrainingData.Rows[j].Cells[1].Value.ToString(), d, dgvTrainingData.Rows[j].Cells[2].Value.ToString());
                                //}
                                removedList[j] = true;
                            }
                        }
                    }
                }
            }

            //--------- generate shortern distances list with respect to the values of thresholds
            dgvDistancesShort.Rows.Clear();
            dgvHueSuspisiousOur.Rows.Clear();

            //--- Set Thresholds ---
            if (FPCount == 0) optionsInfo.PassThreshold = 0;
            else optionsInfo.PassThreshold = Math.Round(SumFP / FPCount, 2);
            if (FFCount == 0) optionsInfo.FailThreshold = 0;
            else optionsInfo.FailThreshold = Math.Round(SumFF / FFCount, 2);

            if (!optionsInfo.isAverageThreshold)
            {
                optionsInfo.PassThreshold = Math.Round(optionsInfo.PassThreshold - (optionsInfo.PassThreshold - PMinDistance) * optionsInfo.PassThresholdPercent / 100.0, 2);
                optionsInfo.FailThreshold = Math.Round(optionsInfo.FailThreshold + (FMaxDistance - optionsInfo.FailThreshold) * optionsInfo.FailThresholdPercent / 100.0, 2);
            }

            txtPassThreshold.Text = optionsInfo.PassThreshold.ToString();
            txtFailThreshold.Text = optionsInfo.FailThreshold.ToString();

            for (int i = 0; i < dgvDistancesAll.Rows.Count; i++)
            {
                if (dgvDistancesAll.Rows[i].Cells[1].Value.ToString() == "failing" && Convert.ToDouble(dgvDistancesAll.Rows[i].Cells[2].Value) >= optionsInfo.FailThreshold)
                {
                    dgvDistancesShort.Rows.Add(dgvDistancesAll.Rows[i].Cells[0].Value, dgvDistancesAll.Rows[i].Cells[1].Value, dgvDistancesAll.Rows[i].Cells[2].Value, dgvDistancesAll.Rows[i].Cells[3].Value);
                }
                if (dgvDistancesAll.Rows[i].Cells[1].Value.ToString() == "passing" && Convert.ToDouble(dgvDistancesAll.Rows[i].Cells[2].Value) <= optionsInfo.PassThreshold)
                {
                    dgvDistancesShort.Rows.Add(dgvDistancesAll.Rows[i].Cells[0].Value, dgvDistancesAll.Rows[i].Cells[1].Value, dgvDistancesAll.Rows[i].Cells[2].Value, dgvDistancesAll.Rows[i].Cells[3].Value);
                }
            }
            txtNoShortenDistances.Text = dgvDistancesShort.Rows.Count.ToString();
            txtNoTotalDistances.Text = dgvDistancesAll.Rows.Count.ToString();
            //----------------- generate FFPrimeCount & FPPrimeCount ------------------------
            //dgvFailPrimeHue.Rows.Clear();
            FFPrimeCount = FPPrimeCount = 0;
            PassedPrimeCount = new long[optionsInfo.NumberOfCodeLines];
            FailedPrimeCount = new long[optionsInfo.NumberOfCodeLines];

            for (int i = 0; i < dgvDistancesShort.Rows.Count; i++)
            {
                if (dgvDistancesShort.Rows[i].Cells[1].Value.ToString().ToUpper() == "passing".ToUpper()) FPPrimeCount++;
                else FFPrimeCount++;

                for (int s = 0; s < dgvDistancesShort.Rows[i].Cells[3].Value.ToString().Length; s++)
                {
                    if (dgvDistancesShort.Rows[i].Cells[1].Value.ToString().ToUpper() == "passing".ToUpper())
                    {
                        PassedPrimeCount[s] += dgvDistancesShort.Rows[i].Cells[3].Value.ToString()[s] - 48;
                    }
                    else
                    {
                        FailedPrimeCount[s] += dgvDistancesShort.Rows[i].Cells[3].Value.ToString()[s] - 48;
                    }
                }
            }


            lblMessage.Text = "Calculating hue and suspisiousness of each statement...";
            for (int s = 0; s < dgvTrainingData.Rows[0].Cells[2].Value.ToString().Length; s++)
            {
                //---- classical tarantula hue and suspisious calc -------
                perc = s * 100 / dgvTrainingData.Rows[0].Cells[2].Value.ToString().Length;
                bkWorkerClassicalTarantula.ReportProgress(perc);

                bool isZero = false;
                if (optionsInfo.TotalNumberOfPassRuns == 0 || optionsInfo.TotalNumberOfFailRuns == 0) isZero = true;
                double p = (double)PassedCount[s] / optionsInfo.TotalNumberOfPassRuns;
                double denominator = (double)FailedCount[s] / optionsInfo.TotalNumberOfFailRuns + p;
                if (denominator == 0 || isZero) hue[s] = 0;
                else hue[s] = (p / denominator);

                /**/
                p = (double)FailedCount[s] / optionsInfo.TotalNumberOfFailRuns;
                denominator = (double)PassedCount[s] / optionsInfo.TotalNumberOfPassRuns + p;
                double q;
                if (denominator == 0 || isZero) q = 0;
                else q = (p / denominator);
                /**/

                dgvHueSuspisiousClassic.Rows.Add(s + 1, PassedCount[s], FailedCount[s], hue[s], q);//1 - hue[s]);

                //----- our tarantula Hue and Suspisiousness Calc ----------------
                isZero = false;
                if (FPPrimeCount == 0 || FFPrimeCount == 0) isZero = true;
                p = (double)PassedPrimeCount[s] / FPPrimeCount;
                denominator = (double)FailedPrimeCount[s] / FPPrimeCount + p;
                double h = 0;
                if (denominator == 0 || isZero) h = 0;
                else h = (p / denominator);

                /**/
                p = (double)FailedPrimeCount[s] / FPPrimeCount;
                denominator = (double)PassedPrimeCount[s] / FPPrimeCount + p;
                if (denominator == 0 || isZero) q = 0;
                else q = (p / denominator);
                /**/
                dgvHueSuspisiousOur.Rows.Add(s + 1, PassedPrimeCount[s], FailedPrimeCount[s], h, q);//1 - hue[s]);
            }

            //---------- classical tarantula rank -------------------
            lblMessage.Text = "Calculating ranks of each statement with respect to the suspisiousness value...";
            dgvHueSuspisiousClassic.Sort(dgvHueSuspisiousClassic.Columns[4], ListSortDirection.Descending);
            int si = 0, sj, sk = 0;
            long rank = 1;
            while (si < dgvHueSuspisiousClassic.Rows.Count)
            {
                //perc = si * 100 / dgvSource.Rows[0].Cells[2].Value.ToString().Length;
                //bkWorkerClassicalTarantula.ReportProgress(perc);
                int n = 0;
                for (sj = si + 1; sj < dgvHueSuspisiousClassic.Rows.Count; sj++)
                {
                    double a = Math.Round(Convert.ToDouble(dgvHueSuspisiousClassic.Rows[si].Cells[4].Value), 2);
                    double b = Math.Round(Convert.ToDouble(dgvHueSuspisiousClassic.Rows[sj].Cells[4].Value), 2);
                    if (a == b) n++;
                    else break;
                }
                for (; sk < sj; sk++)
                {
                    dgvHueSuspisiousClassic.Rows[sk].Cells[5].Value = rank + n;
                }
                si = sj;
                rank = rank + n + 1;
            }

            //---------- our tarantula rank -------------------
            dgvHueSuspisiousOur.Sort(dgvHueSuspisiousOur.Columns[4], ListSortDirection.Descending);
            si = sk = 0;
            rank = 1;
            while (si < dgvHueSuspisiousOur.Rows.Count)
            {
                //perc = si * 100 / dgvSource.Rows[0].Cells[2].Value.ToString().Length;
                //bkWorkerClassicalTarantula.ReportProgress(perc);
                int n = 0;
                for (sj = si + 1; sj < dgvHueSuspisiousOur.Rows.Count; sj++)
                {
                    double a = Math.Round(Convert.ToDouble(dgvHueSuspisiousOur.Rows[si].Cells[4].Value), 2);
                    double b = Math.Round(Convert.ToDouble(dgvHueSuspisiousOur.Rows[sj].Cells[4].Value), 2);
                    if (a == b) n++;
                    else break;
                }
                for (; sk < sj; sk++)
                {
                    dgvHueSuspisiousOur.Rows[sk].Cells[5].Value = rank + n;
                }
                si = sj;
                rank = rank + n + 1;
            }

            //--- classical Diagnosis Score ---
            dgvHueSuspisiousClassic.Sort(dgvHueSuspisiousClassic.Columns[5], ListSortDirection.Ascending);
            int row = 0;
            long sameValCount = 0;
            while (row < dgvHueSuspisiousClassic.Rows.Count)
            {
                if (Convert.ToInt32(dgvHueSuspisiousClassic.Rows[row].Cells[0].Value) == optionsInfo.ActualFaultLineNo)
                {
                    sameValCount = Convert.ToInt32(dgvHueSuspisiousClassic.Rows[row].Cells[5].Value);
                    break;
                }
                row++;
            }
            int row2 = row, rcount = 0;
            if (row2 < optionsInfo.NumberOfCodeLines)
            {
                while (sameValCount == Convert.ToInt32(dgvHueSuspisiousClassic.Rows[row2].Cells[0].Value))
                {
                    row2++;
                    if (row2 >= optionsInfo.NumberOfCodeLines) break;
                    rcount++;
                }
            }

            double v = Math.Round((double)(((row2 + rcount + 1) + row2) / 2.0), 0) + 1;
            optionsInfo.ClassicalDiagScore = Math.Round((1 - v / optionsInfo.NumberOfCodeLines) * 100, 2);
            txtClassicalDiagnosisScore.Text = optionsInfo.ClassicalDiagScore.ToString() + "%";
            dgvHueSuspisiousClassic.Sort(dgvHueSuspisiousClassic.Columns[0], ListSortDirection.Ascending);
            //lblMessage.Text = "Diagnosis Score Calculated...";

            //--- our Diagnosis Score ---
            dgvHueSuspisiousOur.Sort(dgvHueSuspisiousOur.Columns[5], ListSortDirection.Ascending);
            row = 0;
            sameValCount = 0;
            while (row < dgvHueSuspisiousOur.Rows.Count)
            {
                if (Convert.ToInt32(dgvHueSuspisiousOur.Rows[row].Cells[0].Value) == optionsInfo.ActualFaultLineNo)
                {
                    sameValCount = Convert.ToInt32(dgvHueSuspisiousOur.Rows[row].Cells[5].Value);
                    break;
                }
                row++;
            }
            row2 = row;
            rcount = 0;
            if (row2 < optionsInfo.NumberOfCodeLines)
            {
                while (sameValCount == Convert.ToInt32(dgvHueSuspisiousOur.Rows[row2].Cells[0].Value))
                {
                    row2++;
                    if (row2 >= optionsInfo.NumberOfCodeLines) break;
                    rcount++;
                }
            }

            v = Math.Round((double)(((row2 + rcount + 1) + row2) / 2.0), 0) + 1;
            optionsInfo.OurDiagScore = Math.Round((1 - v / optionsInfo.NumberOfCodeLines) * 100, 2);
            txtOurDiagnosisScore.Text = optionsInfo.OurDiagScore.ToString() + "%";
            dgvHueSuspisiousOur.Sort(dgvHueSuspisiousOur.Columns[0], ListSortDirection.Ascending);

        }

        private void bkWorkerClassicalTarantula_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            tpProgressBar1.Value = e.ProgressPercentage;
            lblProgressPercentage.Text = e.ProgressPercentage.ToString() + "%";
        }

        private void bkWorkerClassicalTarantula_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pbAnim.Visible = false;
            lblMessage.Text = "Tarantula Operation has been finished successfully...";
            SetLog("Classical Tarantula Process has been finished successfully...\r\nTotal Iterations: " + TotalIterations.ToString() + "\r\nNumber of F-F Distances: " + FFCount.ToString() + "\r\nNumber of F-P Distances: " + FPCount.ToString(), false);
            SetLog("Pass Threshold: " + optionsInfo.PassThreshold.ToString() + "\r\nFail Threshold: " + optionsInfo.FailThreshold.ToString(), false);
            SetLog("Total Passed: " + optionsInfo.TotalNumberOfPassRuns.ToString() + "\r\nTotal Failed: " + optionsInfo.TotalNumberOfFailRuns.ToString() + "\r\n\r\nDiagnosis Score: " + txtClassicalDiagnosisScore.Text, true);
            tpProgressBar1.Value = 100;
            lblProgressPercentage.Text = "100%";
            MessageBox.Show("Tarantula Process has been finished successfully...", "Tarantula Process Execution", MessageBoxButtons.OK, MessageBoxIcon.Information);
            tpProgressBar1.Value = 0;
            lblProgressPercentage.Text = "0%";

            ApplyResult();
        }

        private void ApplyResult()
        {
            if (dgvResults.Rows.Count > 0)
            {
                bool found = false;
                for (int i = 0; i < dgvResults.Rows.Count; i++)
                {
                    if (dgvResults.Rows[i].Cells[1].Value.ToString() == optionsInfo.TestCase && dgvResults.Rows[i].Cells[2].Value.ToString() == optionsInfo.Version && dgvResults.Rows[i].Cells[3].Value.ToString() == optionsInfo.FaultName)
                    {
                        found = true;
                        if (MessageBox.Show("The result for this evaluation is stored already.\r\nWould you like to replace new result?", "Replace new evaluation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                        {
                            dgvResults.Rows[i].Cells[4].Value = optionsInfo.ClassicalDiagScore;
                            dgvResults.Rows[i].Cells[5].Value = optionsInfo.OurDiagScore;
                            /*
                            if (optionsInfo.ClassicalDiagScore > optionsInfo.OurDiagScore) dgvResults.Rows[i].Cells[5].Value = "Classical";
                            else if (optionsInfo.ClassicalDiagScore < optionsInfo.OurDiagScore) dgvResults.Rows[i].Cells[5].Value = "Ours";
                            else dgvResults.Rows[i].Cells[5].Value = "Both Equal";
                             */ 
                        }
                        break;
                    }
                }

                if (!found)
                {
                    if (dgvResults.DataSource == null)
                    {
                        dgvResults.Rows.Add(dgvResults.Rows.Count + 1, optionsInfo.TestCase, optionsInfo.Version, optionsInfo.FaultName, optionsInfo.ClassicalDiagScore, optionsInfo.OurDiagScore, Math.Round(optionsInfo.PassThreshold, 2), Math.Round(optionsInfo.FailThreshold, 2), Convert.ToInt16(txtNoTotalDistances.Text), Convert.ToInt16(txtNoShortenDistances.Text));
                    }
                    else
                    {
                        DtSet.Tables[0].Rows.Add(dgvResults.Rows.Count + 1, optionsInfo.TestCase, optionsInfo.Version, optionsInfo.FaultName, optionsInfo.ClassicalDiagScore, optionsInfo.OurDiagScore, Math.Round(optionsInfo.PassThreshold, 2), Math.Round(optionsInfo.FailThreshold, 2), Convert.ToInt16(txtNoTotalDistances.Text), Convert.ToInt16(txtNoShortenDistances.Text));
                        dgvResults.DataSource = DtSet.Tables[0];
                    }
                }
            }
            else
            {
                if (dgvResults.DataSource == null)
                {
                    dgvResults.Rows.Add(1, optionsInfo.TestCase, optionsInfo.Version, optionsInfo.FaultName, optionsInfo.ClassicalDiagScore, optionsInfo.OurDiagScore, Math.Round(optionsInfo.PassThreshold, 2), Math.Round(optionsInfo.FailThreshold, 2), Convert.ToInt16(txtNoTotalDistances.Text), Convert.ToInt16(txtNoShortenDistances.Text));
                }
                else
                {
                    DtSet.Tables[0].Rows.Add(1, optionsInfo.TestCase, optionsInfo.Version, optionsInfo.FaultName, optionsInfo.ClassicalDiagScore, optionsInfo.OurDiagScore, Math.Round(optionsInfo.PassThreshold, 2), Math.Round(optionsInfo.FailThreshold, 2), Convert.ToInt16(txtNoTotalDistances.Text), Convert.ToInt16(txtNoShortenDistances.Text));
                    dgvResults.DataSource = DtSet.Tables[0];
                }
            }
            setEvaluationResults();
        }

        private void setEvaluationResults()
        {
            txtTotalNoEvaluations.Text = dgvResults.Rows.Count.ToString();
            double sumClassical=0, sumOurs=0, AvgClassical, AvgOurs;
            
            for (int i = 0; i < dgvResults.Rows.Count; i++)
            {
                sumClassical += Convert.ToDouble(dgvResults.Rows[i].Cells[4].Value);
                sumOurs += Convert.ToDouble(dgvResults.Rows[i].Cells[5].Value);
            }

            AvgClassical = Math.Round(sumClassical / dgvResults.Rows.Count, 2);
            AvgOurs = Math.Round(sumOurs / dgvResults.Rows.Count, 2);

            txtClassicalAvg.Text = AvgClassical.ToString() + "%";
            txtOursAvg.Text = AvgOurs.ToString() + "%";
            txtAvgDifference.Text = Math.Round(Math.Abs(AvgClassical - AvgOurs), 2).ToString();

            if (AvgOurs > AvgClassical)
            {
                pbHappy.Visible = true;
                pbSad.Visible = false;
                txtBetterMethod.Text = "Ours";
            }
            else if (AvgOurs < AvgClassical)
            {
                pbHappy.Visible = false;
                pbSad.Visible = true;
                txtBetterMethod.Text = "Classical";
            }
            else
            {
                pbHappy.Visible = false;
                pbSad.Visible = false;
                txtBetterMethod.Text = "Both Equal";
            }
        }

        //----------------------------- import data -------------------------
        private void btnClearLog_Click(object sender, EventArgs e)
        {
            tabsControl.SelectedIndex = 2;
            txtLog.Text = string.Empty;
        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            FrmAbout frm = new FrmAbout();
            frm.ShowDialog(this);
        }

        System.Data.DataSet DtSet;
        private void btnLoadResult_Click(object sender, EventArgs e)
        {
            long fileSize = 0;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.OleDb.OleDbDataAdapter MyCommand;

            openFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments.ToString();
            openFileDialog1.Filter = "Excel Files (*.xls,*.xlsx)|*.xls;*.xlsx";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.Multiselect = false;
            openFileDialog1.Title = "Import Results data from excel file";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    dgvResults.Columns.Clear();
                    tabsControl.SelectedIndex = 5;
                    tpProgressBar1.Value = 0;
                    //IsExecuted = false;
                    FileInfo fi = new FileInfo(openFileDialog1.FileName);
                    fileSize = fi.Length;

                    string excelObject = "Provider=Microsoft.{0}.OLEDB.{1};Data Source={2};Extended Properties=\"Excel {3};HDR=YES\"";

                    if (fi.Extension.Equals(".xls"))
                    {
                        excelObject = string.Format(excelObject, "Jet", "4.0", openFileDialog1.FileName, "8.0");
                    }
                    else if (fi.Extension.Equals(".xlsx"))
                    {
                        excelObject = string.Format(excelObject, "Ace", "12.0", openFileDialog1.FileName, "12.0");
                    }
                    //string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();

                    MyConnection = new System.Data.OleDb.OleDbConnection(excelObject);
                    MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection); //where results='passing'
                    MyCommand.TableMappings.Add("Table", "TestTable");
                    DtSet = new System.Data.DataSet();

                    MyCommand.Fill(DtSet);

                    dgvResults.DataSource = DtSet.Tables[0];
                    MyConnection.Close();

                    tpProgressBar1.Value = 100;
                    lblProgressPercentage.Text = "100%";
                    SetLog(">>>> Results information has been imported successfully...", true);
                    lblMessage.Text = "Results information has been imported successfully...";
                    bool isBtnTrainActive = false, isBtnMetaActive = false, isTarantulaBtnActive = false;
                    if (dgvTrainingData.Rows.Count > 0) isBtnTrainActive = true;
                    if (dgvMetaData.Rows.Count > 0) isBtnMetaActive = true;
                    if (isBtnTrainActive && isBtnMetaActive) isTarantulaBtnActive = true;
                    setButtonsActivate(true, isBtnTrainActive, isTarantulaBtnActive, isTarantulaBtnActive, true, true, true);
                    setEvaluationResults();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void btnSaveResults_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments.ToString();
            saveFileDialog1.Filter = "Excel Files (*.xls,*.xlsx)|*.xls;*.xlsx";
            saveFileDialog1.FilterIndex = 0;
            saveFileDialog1.Title = "Export Results data to the excel file";
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int i = 0;
                int j = 0;

                xlWorkSheet.Cells[1, 1] = "No";
                xlWorkSheet.Cells[1, 2] = "Test Case";
                xlWorkSheet.Cells[1, 3] = "Version";
                xlWorkSheet.Cells[1, 4] = "Fault Name";
                xlWorkSheet.Cells[1, 5] = "Classical Diagnousis Score";
                xlWorkSheet.Cells[1, 6] = "Our Diagnousis Score";
                xlWorkSheet.Cells[1, 7] = "Pass Threshold";
                xlWorkSheet.Cells[1, 8] = "Fail Threshold";
                xlWorkSheet.Cells[1, 9] = "Total Distances";
                xlWorkSheet.Cells[1, 10] = "Shorten Distances";
                for (i = 0; i < dgvResults.RowCount; i++)
                {
                    for (j = 0; j < dgvResults.ColumnCount; j++)
                    {
                        DataGridViewCell cell = dgvResults[j, i];
                        xlWorkSheet.Cells[i + 2, j + 1] = cell.Value;
                    }
                }

                xlWorkBook.SaveAs(saveFileDialog1.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                SetLog("Results has been exported to the Excel file...\r\nYou can find the file " + saveFileDialog1.FileName, true);
                lblMessage.Text = "Results has been exported to the Excel file...";
                MessageBox.Show("Results has been exported to the Excel file...\r\nYou can find the file "+saveFileDialog1.FileName);
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btnSaveLog_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments.ToString();
            saveFileDialog1.Filter = "Text Files (*.txt)|*.txt;";
            saveFileDialog1.FilterIndex = 0;
            saveFileDialog1.Title = "Save Activity Logs into the text file";
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                System.IO.File.WriteAllText(saveFileDialog1.FileName, txtLog.Text);
            }
        }

        private void dgvResults_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvResults.Rows.Count > 0)
            {
                txtTestCase.Text = dgvResults.CurrentRow.Cells[1].Value.ToString();
                txtVersion.Text = dgvResults.CurrentRow.Cells[2].Value.ToString();
                txtFaultName.Text = dgvResults.CurrentRow.Cells[3].Value.ToString();
                txtActualFaultLineNo.Text = "-";

                txtClassicalDiagnosisScore.Text = dgvResults.CurrentRow.Cells[4].Value.ToString();
                txtOurDiagnosisScore.Text = dgvResults.CurrentRow.Cells[5].Value.ToString();

                txtPassThreshold.Text = dgvResults.CurrentRow.Cells[6].Value.ToString();
                txtFailThreshold.Text = dgvResults.CurrentRow.Cells[7].Value.ToString();
            }
        }
    }
}
