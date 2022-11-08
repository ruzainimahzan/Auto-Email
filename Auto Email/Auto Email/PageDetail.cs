using Auto_Email_Class;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Auto_Email
{
    public partial class PageDetail : Form
    {

        string dbcon = ConfigurationManager.ConnectionStrings["dbAutoEmail"].ConnectionString;
        string sqlSelect = ConfigurationManager.AppSettings["SQLUserInfo"];
        
        private static Pop3Client client = new Pop3Client();
        DataSet dsEmail = new DataSet();
        public Email_Info _carryemailInfo;

        public PageDetail()
        {
            InitializeComponent();
        }

        private void PageDetail_Load(object sender, EventArgs e)
        {
            InitializeFieldName();
            LoadData();
            LoadDB();
            DBLoadEmail();
        }

        private void DBLoadEmail()
        {
            try
            {

                dsEmail.Clear();
                string sqlSelect = ConfigurationManager.AppSettings["SQLViewEmail"];
                SqlDataAdapter da = new SqlDataAdapter(sqlSelect, dbcon);
                da.TableMappings.Add("Table", "Email_Detail");

                // Create and fill the DataSet
                DataSet ds = new DataSet();
                da.Fill(ds);

                DataTable ClientInfo = ds.Tables["Email_Detail"];

                using (SqlConnection con = new SqlConnection(dbcon))
                {
                    using (SqlCommand cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLViewEmail"]))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (!dsEmail.Tables.Contains(ConfigurationManager.AppSettings["Email_Detail"]))
                                dsEmail.Tables.Add(ConfigurationManager.AppSettings["Email_Detail"]);

                            dsEmail.Tables[0].Load(sdr);

                        }
                        con.Close();
                    }

                    using (SqlCommand cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLUserInfo"]))
                    {
                        try
                        {
                            cmd.Connection = con;
                            con.Open();
                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                if (!dsEmail.Tables.Contains(ConfigurationManager.AppSettings["UserInfo"]))
                                    dsEmail.Tables.Add(ConfigurationManager.AppSettings["UserInfo"]);

                                dsEmail.Tables[3].Load(sdr);

                            }
                            con.Close();
                        }
                        catch (Exception ex)
                        {

                        }

                    }
                }
            
            }
            catch (Exception ex)
            {

            }



        }

        public void LoadData()
        {
            tbEID.Text = MainPage.UIDL; // MainPage.UIDL;
            tbRejectDateTime.Text = MainPage._RejectDateTime.ToString();
            tbSendDateTime.Text = MainPage._SendDate.ToString();
            tbSubject.Text = MainPage._Subject;
            tbSendFrom.Text = MainPage._SendFrom;
            tbSendFromIPAddress.Text = MainPage._SendFromIPAddress;
            tbRejectCategory.Text = MainPage._ReasonCategory;
            tbClientSendTo.Text = MainPage._SendTo;
            tbRejectReason.Text = MainPage._Reason;
            tbActionRequired.Text = MainPage._ActionRequired;
            tbActionTaken.Text = MainPage._ActionTaken;
            tbActionTakenBy.Text = MainPage._ActionTakenBy;
            tbActionDescription.Text = MainPage._ActionDescription;
            cbThisCase.Checked = MainPage._CaseStatus;

        }

        private void LoadDB()
        {
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlSelect, dbcon);
                da.TableMappings.Add("Table", "CLMAST");

                // Create and fill the DataSet
                DataSet ds = new DataSet();
                da.Fill(ds);

                DataTable ClientInfo = ds.Tables["CLMAST"];
                var ClientQuery = from d in ClientInfo.AsEnumerable()
                                  where d.Field<string>("LEMAIL") == MainPage._SendTo
                                  select new
                                  {
                                      //LACCT, LNAME, LTEL
                                      CliAcct = d.Field<string>("LACCT"),
                                      CliName = d.Field<string>("LNAME"),
                                      CliTel = d.Field<string>("LTEL")
                                  };

                foreach (var Client in ClientQuery)
                {
                    tbClientAcctNo.Text = Client.CliAcct;
                    tbClientName.Text = Client.CliName;
                    tbClientTel.Text = Client.CliTel;
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            MainPage mainpage = new MainPage();
            //mainpage.Show();
            mainpage.BringToFront();
        }

        private void InitializeFieldName()
        {
            tbEID.Text = "";
            tbRejectCategory.Text = "";
            tbRejectDateTime.Text = "";
            tbRejectReason.Text = "";
            tbSendDateTime.Text = "";
            tbSendFrom.Text = "";
            tbSendFromIPAddress.Text = "";
            tbClientSendTo.Text = "";
            tbSubject.Text = "";

            tbActionDateTime.Text = "";
            tbActionDescription.Text = "";
            tbActionRequired.Text = "";
            tbActionTaken.Text = "";
            tbActionTakenBy.Text = "";

            cbThisCase.Checked = false;
            cbAllCase.Checked = false;

            tbEID.ReadOnly = true;
            tbRejectCategory.ReadOnly = true;
            tbRejectDateTime.ReadOnly = true;
            tbRejectReason.ReadOnly = true;
            tbSendDateTime.ReadOnly = true;
            tbSendFrom.ReadOnly = true;
            tbSendFromIPAddress.ReadOnly = true;
            tbClientSendTo.ReadOnly = true;
            tbSubject.ReadOnly = true;
            tbActionDateTime.Enabled = false;

            timerCounter.Start();
        }

        private void UpdateToDB()
        {

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                PageDetail form2 = new PageDetail();
                SqlConnection con = new SqlConnection(dbcon);
                SqlCommand cmd;
                con.Open();

                //DataRow[] hasRow = dsEmail.Tables[ConfigurationManager.AppSettings["Email_Detail"]].Select("email_ID ='" + tbEID.Text.ToString() + "'");
                DataRow[] hasRow = dsEmail.Tables[0].Select("email_ID ='" + tbEID.Text.ToString() + "'");
                if (hasRow.Length != 0)
                {
                    cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLUpdateCLientIssue"], con);

                    cmd.Parameters.Add("@email_ID", SqlDbType.NVarChar, 25).Value = tbEID.Text.ToString() ?? (object)DBNull.Value;
                    //cmd.Parameters.Add("@sendDateTime", SqlDbType.DateTime).Value = _SendDate; //?? (object)DBNull.Value;
                    //cmd.Parameters.Add("@sendFromIPAddress", SqlDbType.VarChar, 20).Value = _SendFromIPAddress ?? (object)DBNull.Value;
                    //cmd.Parameters.Add("@sendFrom", SqlDbType.VarChar, 50).Value = _SendFrom ?? (object)DBNull.Value;
                    //cmd.Parameters.Add("@sendTo", SqlDbType.VarChar, 50).Value = _SendTo ?? (object)DBNull.Value;
                    //cmd.Parameters.Add("@subject", SqlDbType.NVarChar, 255).Value = _Subject ?? (object)DBNull.Value;
                    //cmd.Parameters.Add("@rejectDateTime", SqlDbType.DateTime).Value = _RejectDateTime;
                    //cmd.Parameters.Add("@rejectReason", SqlDbType.VarChar, 255).Value = _TrialReason ?? (object)DBNull.Value;
                    //cmd.Parameters.Add("@rejectReasonCategory", SqlDbType.VarChar, 50).Value = _ReasonCategory ?? (object)DBNull.Value;
                    cmd.Parameters.Add("@actionRequired", SqlDbType.VarChar, 255).Value = tbActionRequired.Text.ToString() ?? (object)DBNull.Value;
                    cmd.Parameters.Add("@actionTaken", SqlDbType.VarChar, 255).Value = tbActionTaken.Text.ToString() ?? (object)DBNull.Value;
                    cmd.Parameters.Add("@actionTakenDescription", SqlDbType.VarChar, 255).Value = tbActionDescription.Text.ToString() ?? (object)DBNull.Value;
                    cmd.Parameters.Add("@actionTakenBy", SqlDbType.VarChar, 100).Value = tbActionTakenBy.Text.ToString() ?? (object)DBNull.Value;
                    cmd.Parameters.Add("@actionTakenDateTime", SqlDbType.DateTime).Value = DateTime.Now;
                    cmd.Parameters.Add("@closeCase", SqlDbType.Bit).Value = cbThisCase.Checked;

                    cmd.CommandType = CommandType.Text;
                    int i = cmd.ExecuteNonQuery();

                   
                    
                }
                DialogResult dialog = MessageBox.Show("Record Successfully Updated", "Notification!", MessageBoxButtons.OK);
                if (dialog == DialogResult.OK)
                {
                    timerCounter.Stop();
                    this.Close();
                    MainPage mainpage = new MainPage();
                    //mainpage.Show();
                    form2.Dispose();
                    mainpage.BringToFront();

                    
                }

            }
            catch (Exception ex)
            {
                //MessageBoxButtons MsgBox = MessageBoxButtons.OK;
                //MessageBox.Show(ex.ToString() + "\r\n Email ID: " + UIDL, "Error", MsgBox);

            }
        }

        private void timerCounter_Tick(object sender, EventArgs e)
        {
            //Show current time
            DateTime datetime = DateTime.Now;
            //DateTime date = DateTime.Now.Date;
            tbActionDateTime.Text = datetime.ToString();// +" " + datetime.ToString();
            tbActionDateTime.CustomFormat = "dd/MM/yyyy H:mm:ss";
        }

        private void cbThisCase_Click(object sender, EventArgs e)
        {
            if (cbThisCase.Checked = true)
            {
                cbAllCase.Checked = false;
            }

        }

        private void cbAllCase_Click(object sender, EventArgs e)
        {
            if (cbAllCase.Checked = true)
            {
                cbThisCase.Checked = false;
            }
        }

        
    }
}
