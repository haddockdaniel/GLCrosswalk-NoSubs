using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            UpdateStatus("Creating New Accounts...", 0, 0);
            toolStripStatusLabel.Text = "Creating New Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();
           
            

            //Create New Accounts
            string SQL = @"Insert into ChartofAccounts(chtsysnbr, chtmainacct, chtsubacct, chtdesc, chtsubtotlevel, chtfinstmttype, chtsaftype, chtparencode, chtcomprescode,
            chtcashflowtype, chtsubacct1, chtsubacct2, chtsubacct3, chtsubacct4, chtsubacct5, chtsubacct6, chtsubacct7, chtsubacct8)
            select(select spnbrvalue from sysparam where spname = 'LastSysNbrChart') + rank() over(order by newacct) as Chtsysnbr, right('000000000' + cast(newacct as varchar(8)), 8),chtsubacct, left(min(newdesc), 30) as NewDesc, chtsubtotlevel, chtfinstmttype, chtsaftype, min(chtparencode), chtcomprescode,
            min(chtcashflowtype), (select isnull(min(coas1id),0) from coasubaccount1), 0, 0, 0, 0, 0, 0, 0
            from #tblcoa
            inner join chartofaccounts on oldsysnbr = chtsysnbr
                   where newsysnbr is null
            group by newacct, chtsubacct, chtsubtotlevel, chtfinstmttype,chtsaftype, chtcomprescode";

            _jurisUtility.ExecuteNonQueryCommand(0, SQL);


            SQL = @"update sysparam
            set spnbrvalue = (select max(chtsysnbr) from chartofaccounts) where spname = 'LastSysNbrChart'";

            _jurisUtility.ExecuteNonQueryCommand(0, SQL);


            SQL = @"Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyl)
            select(select spnbrvalue from sysparam where spname = 'LastSysNbrDocTree') + rank() over(order by chtsysnbr) as Chtsysnbr,'Y','2100','R','9',
            chtdesc, chtsysnbr
            from chartofaccounts
            where chtsysnbr not in (select dtkeyl from documenttree where dtdocclass = 2100 and dtdoctype = 'R' and dtparentid = 9)";

            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

           SQL= @"update sysparam
            set spnbrvalue = (select max(dtdocid) from documenttree) where spname = 'LastSysNbrDocTree'";

            _jurisUtility.ExecuteNonQueryCommand(0, SQL);


            SQL = @"Update #tblcoa
                set newsysnbr = chtsysnbr
                from chartofaccounts
                where chtmainacct = right('000000000' + cast(newacct as varchar(8)), 8) and newsysnbr is null";

                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
           
            //Update DataGrid With New Accounts

            string sql2 = "select * from #tblcoa order by newacct";
            DataSet coa = _jurisUtility.ExecuteSqlCommand(0, sql2);
            dataGridView1.DataSource = coa.Tables[0];

            UpdateStatus("Updating Chart of Accounts...", 1, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();

            //Update New Accounts and Accounts that will not have completely new data

            string sql = @"update  jebatchdetail
            set jebdaccount=newsysnbr from #tblcoa where jebdaccount=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            sql = @"update  voucherbatchdetail
        set VBDDiscAcct=newsysnbr from #tblcoa where VBDDiscAcct=oldsysnbr
        and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 2, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();

            sql = @"update  voucherbatchgldist
            set VBGGLAcct=newsysnbr from #tblcoa where VBGGLAcct=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  vouchergldist
            set VGLGLAcct=newsysnbr from #tblcoa where VGLGLAcct=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 3, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();


            sql = @"update  journalentry
            set jeaccount=newsysnbr from #tblcoa where jeaccount=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  jetemplatedetail
            set jetdaccount=newsysnbr from #tblcoa where jetdaccount=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  vendor
            set VenDiscAcct=newsysnbr from #tblcoa where VenDiscAcct=oldsysnbr
             and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 4, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();


            sql = @"update  vendor
            set VenDefaultDistAcct=newsysnbr from #tblcoa where VenDefaultDistAcct=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            sql = @"update apaccount
            set apaglacct=newsysnbr from #tblcoa where apaglacct=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            sql = @"update practclassglacct
            set pgaacct=newsysnbr from #tblcoa where pgaacct=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            sql = @"update OfficeGLAccount
            set ogaacct=newsysnbr from #tblcoa where ogaacct=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  voucher
            set VchDiscAcct=newsysnbr from #tblcoa where VchDiscAcct=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 5, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();


            sql = @"update  CRNonCliAlloc
            set CRNCreditAccount=newsysnbr from #tblcoa where CRNCreditAccount=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  ExpCodeGLAcct
            set ECGAAcct=newsysnbr from #tblcoa where ECGAAcct=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 6, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();







            sql = @"update  ExpDetailDist
            set EDDAccount=newsysnbr from #tblcoa where EDDAccount=oldsysnbr
             and statustype in (1,2,4)"; 

            _jurisUtility.ExecuteNonQueryCommand(0, sql);







            sql = @"update  BkAcctGLAcct
            set BGAAcct=newsysnbr from #tblcoa where BGAAcct=oldsysnbr  and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 7, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();


            sql = @"update  TimeDetailDist
            set TDDAccount=newsysnbr from #tblcoa where TDDAccount=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);



            sql = @"update  EmpGLAcct
             set EGAAcct=newsysnbr from #tblcoa where EGAAcct=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 8, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();



            sql = @"update  FSLayoutItem
            set FSLIChtSysNbr=newsysnbr from #tblcoa where FSLIChtSysNbr=oldsysnbr
            and statustype in (1,2,4)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            //Update Accounts that are moving to an existing account that is also moving


            sql = @"update jebatchdetail
            set jebdaccount = newsysnbr from #tblcoa where jebdaccount=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  jetemplatedetail
            set jetdaccount=newsysnbr from #tblcoa where jetdaccount=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 9, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();


            sql = @"update  voucherbatchdetail
            set VBDDiscAcct=newsysnbr from #tblcoa where VBDDiscAcct=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  voucherbatchgldist
            set VBGGLAcct=newsysnbr from #tblcoa where VBGGLAcct=oldsysnbr
             and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 10, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();

            sql = @"update apaccount
            set apaglacct=newsysnbr from #tblcoa where apaglacct=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  vouchergldist
            set VGLGLAcct=newsysnbr from #tblcoa where VGLGLAcct=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  journalentry
            set jeaccount=newsysnbr from #tblcoa where jeaccount=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 11, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();

            sql = @"update  vendor
            set VenDiscAcct=newsysnbr from #tblcoa where VenDiscAcct=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            sql = @"update practclassglacct
            set pgaacct=newsysnbr from #tblcoa where pgaacct=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update OfficeGLAccount
            set ogaacct=newsysnbr from #tblcoa where ogaacct=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  vendor
            set VenDefaultDistAcct=newsysnbr from #tblcoa where VenDefaultDistAcct=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 12, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();

            sql = @"update  voucher
            set VchDiscAcct=newsysnbr from #tblcoa where VchDiscAcct=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  CRNonCliAlloc
            set CRNCreditAccount=newsysnbr from #tblcoa where CRNCreditAccount=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 13, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();

            sql = @"update  ExpCodeGLAcct
            set ECGAAcct=newsysnbr from #tblcoa where ECGAAcct=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            sql = "select * from expdetaildist inner join #tblcoa on EDDAccount=oldsysnbr order by eddbatch, EDDRecNbr";
            DataSet dd = _jurisUtility.RecordsetFromSQL(sql);
            var result = dd.Tables[0]
                .AsEnumerable()
                .Where(myRow => myRow.Field<int>("RowNo") == 1);


            sql = @"update  ExpDetailDist
            set EDDAccount=newsysnbr from #tblcoa where EDDAccount=oldsysnbr
            and statustype in (3)";

            //select min(ogaacct) from officeglaccount
            //where ogatype=210 

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 14, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();


            sql = @"update  BkAcctGLAcct
            set BGAAcct=newsysnbr from #tblcoa where BGAAcct=oldsysnbr and statustype in (3)";


            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  TimeDetailDist
            set TDDAccount=newsysnbr from #tblcoa where TDDAccount=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart of Accounts...", 15, 25);
            toolStripStatusLabel.Text = "Updating Chart of Accounts...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();

            sql = @"update  EmpGLAcct
            set EGAAcct=newsysnbr from #tblcoa where EGAAcct=oldsysnbr
            and statustype in (3)";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update  FSLayoutItem
        set FSLIChtSysNbr=newsysnbr from #tblcoa where FSLIChtSysNbr=oldsysnbr
        and statustype in (3)";


            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating Chart Budgets...", 20, 25);
            toolStripStatusLabel.Text = "Updating Chart Budgets...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();

           string SqLCB = @"select newsysnbr,  chbperiod as prd, chbprdyear as yr, sum(chbnetchange) as NC, sum(chbbudget) as Bud
                    into #tblcb
                        from chartbudget
                    inner join #tblcoa on oldsysnbr=chbaccount
                     group by newsysnbr, chbperiod, chbprdyear";
            _jurisUtility.ExecuteNonQueryCommand(0, SqLCB);





            sql = "Delete from chartbudget";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            sql = @"Insert into ChartBudget(chbaccount, chbprdyear, chbperiod, chbnetchange, chbbudget)
                Select newsysnbr, yr, prd, nc, bud from #tblcb";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            sql = "drop table #tblcb";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            UpdateStatus("Account Clean Up...", 23, 25);
            toolStripStatusLabel.Text = "Account Clean Up...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();

            sql = @"update sysparam set sptxtvalue=cast(newsysnbr as varchar(10)) + ',0' from #tblcoa where spname='DefEmpGlAcct' and sptxtvalue=cast(oldsysnbr as varchar(10)) + ',0'";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            sql = @"update sysparam set sptxtvalue=cast(newsysnbr as varchar(10)) + ',0' from #tblcoa where spname='DefExpGlAcct' and sptxtvalue=cast(oldsysnbr as varchar(10)) + ',0'";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"update sysparam set sptxtvalue=cast(newsysnbr as varchar(10)) + ',0' from #tblcoa where spname='DefVenDiscAcct' and sptxtvalue=cast(oldsysnbr as varchar(10)) + ',0'";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            sql = @"delete  from chartofaccountchartcategory where chtsysnbr not in (Select newsysnbr from #tblcoa)";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = @"delete  from chartofaccounts where chtsysnbr not in (select newsysnbr from #tblcoa)";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = @"update  chartofaccounts set  chtsubacct1=0,chtdesc=left(newdesc,30) from #tblcoa where newsysnbr=chtsysnbr";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
           sql = @"delete from documenttree where dtdocclass='2100' and dtdoctype='R' and dtparentid=9 and dtkeyl not in (select chtsysnbr from chartofaccounts)";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            sql = "drop table #tblcoa";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            toolStripStatusLabel.Text = "Complete.";
            Cursor.Current = Cursors.Default;
            statusStrip.Refresh();
            UpdateStatus("Complete...", 25, 25);

            WriteLog("GL CROSSWALK CHART OF ACCOUNTS UPDATE " + DateTime.Now.ToString("MM/dd/yyyy"));

        }
        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
                labelCurrentStatus.Text = status;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                    labelCurrentStatus.Text = status;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                    labelCurrentStatus.Text = status;
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
        
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string fileName;
                fileName = dlg.FileName.ToString();
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection(@"provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes'");
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection);

                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                dataGridView1.DataSource = DtSet.Tables[0];
                MyConnection.Close();
            }
            
            toolStripStatusLabel.Text = "Getting Account Data...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();
            UpdateStatus("Getting Account Data...", 0, 0);


            string sql = "Create Table #oldcoa(oldsysnbr int, oldacct varchar(8), newacct varchar(8), newdesc varchar(100))";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            DataTable dvgSource = (DataTable)dataGridView1.DataSource;
            int numRows = dataGridView1.RowCount;
            int j = 1;
            foreach (DataRow dg in dvgSource.Rows)

            {
                string oldsys = dg["OldSysNbr"].ToString();
                string oldacct = dg["OldAccount"].ToString();
                string newacct = dg["NewAccount"].ToString();
                string newdesc = dg["NewDesc"].ToString();

                string isql = "Insert into #oldcoa values(" + oldsys.ToString() + ",'" + oldacct.ToString() + "','" + newacct.ToString() + "','" + newdesc.ToString() + "')";
                _jurisUtility.ExecuteNonQueryCommand(0, isql);

            }

            string jsql = @"select oldsysnbr, left(oldacct,MinLen) as oldacct, newacct, newdesc,chtsysnbr as newsysnbr,case when oldsysnbr=chtsysnbr then 1
                    when oldsysnbr<>chtsysnbr and left(oldacct,MinLen)  in (select newacct from #oldcoa) then 2
                    when oldsysnbr<>chtsysnbr and  left(oldacct,MinLen) not in (select newacct from #oldcoa) then 3 else 4 end as statustype
                    into #tblcoa 
                    from #oldcoa
                    left outer join (select min(Chtsysnbr) as Chtsysnbr, chtmainacct from chartofaccounts group by chtmainacct) CM on chtmainacct=right('00000000' + newacct,8), (select min(len(oldacct)) as MinLen from #oldcoa where len(oldacct)>1) ML";

            _jurisUtility.ExecuteNonQueryCommand(0, jsql);

            string dsql = "drop table #oldcoa";
            _jurisUtility.ExecuteNonQueryCommand(0, dsql);
            string sql2 = "select * from #tblcoa order by newacct";
            DataSet coa = _jurisUtility.ExecuteSqlCommand(0,sql2);
            dataGridView1.DataSource = coa.Tables[0];

            toolStripStatusLabel.Text = "Ready to Update...";
            Cursor.Current = Cursors.Default;
            statusStrip.Refresh();
            UpdateStatus("Ready to Update...", 0, 0);

        }

        private string getReportSQL()
        {
            string reportSQL = "";
            //if matter and billing timekeeper
            if (true)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,empinitials as CurrentBillingTimekeeper, 'DEF' as NewBillingTimekeeper" +
                        " from matter" +
                        " inner join client on matclinbr=clisysnbr" +
                        " inner join billto on matbillto=billtosysnbr" +
                        " inner join employee on empsysnbr=billtobillingatty" +
                        " where empinitials<>'ABC'";


            return reportSQL;
        }


    }
}
