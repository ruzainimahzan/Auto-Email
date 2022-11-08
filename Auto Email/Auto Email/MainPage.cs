using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using AccountCredential;
using System.Net.Security;
using System.Configuration;
using OpenPop.Mime.Header;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using Auto_Email_Class;
using System.Collections.Specialized;
using System.Net.Mail;
using System.Net.Mime;
using OpenPop.Mime.Decode;
using Mime;
using System.Diagnostics;
using System.Globalization;

namespace Auto_Email
{
	public partial class MainPage : Form
	{
		#region Declaration
		string dbcon = ConfigurationManager.ConnectionStrings["dbAutoEmail"].ConnectionString;
		private static Pop3Client client = new Pop3Client();
		DataSet dsEmail = new DataSet();
		DataSet AEiMS = new DataSet();
		DataTable dt = new DataTable();
        DataTable dtEmailLogin = new DataTable();
        DataTable dtRejectPattern = new DataTable();
        DataTable dtActionTaken = new DataTable();
		DataRow dRow;
		DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();
		MessageHeader headers;
		OpenPop.Mime.Message message;
		OpenPop.Mime.MessagePart messagePart, messagePart1, messagePart2, messagePart3 ;
		//Server Login Credential
		public static string _Hostname, _Username, _Password, _EmailAddress;
		public static int _PortNo;
        public static bool counter = false;
        public static bool ReloadRecord = false;
	   
		//DataSet dsUserInfo = new DataSet();
		string body, body1, body2;
		public static string UIDL, _SendTo, _SendFrom, _Subject, _Reason, _ReasonCategory, _ActionRequired , _SendFromIPAddress, _ActionTaken, _ActionTakenBy, _ActionDescription, _EmailType;
		public static DateTime _SendDate, _RejectDateTime;
		public static bool _CaseStatus = false;
		public static int messageCount = 0;

		public static string temp_UIDL, temp_SendTo, temp_SendFrom, temp_Subject, temp_Reason, temp_ReasonCategory, temp_ActionRequired, temp_SendFromIPAddress;
		public static DateTime temp_SendDate, temp_RejectDateTime;
		public static TimeSpan temp_SendTime;
		bool ChineseChar = false;

		//Not Use Yet
		//public Email_Info _emailInfo = new Email_Info();

		#endregion

		public MainPage()
		{
			InitializeComponent();
		}

		private void MainPage_Load(object sender, EventArgs e)
		{
            DeleteRecord();
			DBLoadEmail();
			ServerConnection();
			//InitiateFieldName();
			
			//RetrieveEmail();
			//comboxCaseStatus.Enabled = true;
			//client.Disconnect();
			
		}

		private static bool ValidateCertificate(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate, System.Security.Cryptography.X509Certificates.X509Chain chain, SslPolicyErrors sslPolicyErrors)
		{
			return true;
		}

		private void InitiateFieldName()
		{
			body = "";
			body1 = "";
			body2 = "";
			UIDL = "";
			_SendTo = "";
			_SendFrom = "" ; 
			_Subject = "" ;
			_Reason = "";
			_ReasonCategory = "";
			_ActionRequired = "";
			_SendFromIPAddress = "";
			_SendDate = DateTime.Now;
			_RejectDateTime = DateTime.Now;

			buttonColumn.HeaderText = "";
			buttonColumn.Name = "View";
			buttonColumn.UseColumnTextForButtonValue = true;

			//gvEmailList.Columns.Add(buttonColumn);
			if(gvEmailList.ColumnCount == 0)
			{
				dt.Columns.Add("Send Date", typeof(DateTime));
				dt.Columns.Add("Send To", typeof(string));
				dt.Columns.Add("Subject", typeof(string));
				dt.Columns.Add("Reject Date", typeof(DateTime));
				dt.Columns.Add("Reason Category", typeof(string));
				dt.Columns.Add("Action Acquired", typeof(string));
				dt.Columns.Add("mID", typeof(string));

				//Newly Added 02/05/2018
				dt.Columns.Add("Reason", typeof(string));
				dt.Columns.Add("Send From", typeof(string));
				dt.Columns.Add("Send From IP Address", typeof(string));

				gvEmailList.DataSource = dt;
				gvEmailList.Sort(gvEmailList.Columns["Send Date"], ListSortDirection.Descending);
				gvEmailList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
				gvEmailList.Columns["mID"].Visible = false;
				gvEmailList.Columns["Reason"].Visible = false;
				gvEmailList.Columns["Send From"].Visible = false;
				gvEmailList.Columns["Send From IP Address"].Visible = false;
				//gvEmailList.Columns["View"].Visible = false;
				gvEmailList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
			}
			

		}

		private void ServerConnection()
		{
            try
            {
               
                DeleteRecord();
				DataRow[] hasRow = dsEmail.Tables[ConfigurationManager.AppSettings["Server_Login"]].Select(); 
				foreach (DataRow Returned in hasRow)
				{
					//DBLoadEmail();
                    if (Returned.HasErrors == false)
                    {
                        _Hostname = Returned[1].ToString();
                        _EmailAddress = Returned[2].ToString();
                        _Username = Returned[3].ToString();
                        _Password = Returned[4].ToString();
                        _PortNo = Convert.ToInt16(Returned[5].ToString());

                        if (_Hostname != "" && _Username != "" && _Password != "" && _PortNo != 0)
                        {
                            client.Connect(_Hostname, _PortNo, true, 604800000, 604800000, ValidateCertificate);
                            client.Authenticate(_Username, _Password);

                            if (client.Connected == true)
                            {
                                //DBLoadEmail();
                                RetrieveEmail();
                                comboxCaseStatus.Enabled = true;
                            }

                        }
                        client.Disconnect();

                    }
				}
				tBLastTimeRun.Text = DateTime.Now.ToString();
				//dsEmail.Clear();
				timer1.Start();

            }
            catch (Exception ex)
            {
                CreateLogFile(ex.ToString());
            }

		}

		private void DBLoadEmail()
		{
			try
			{

				dsEmail.Clear();
				
				Stopwatch sw = Stopwatch.StartNew();
			   
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

					using (SqlCommand cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLRefReasonCategory"]))
					{
						cmd.Connection = con;
						con.Open();
						using (SqlDataReader sdr = cmd.ExecuteReader())
						{
							if (!dsEmail.Tables.Contains(ConfigurationManager.AppSettings["Ref_ReasonCategory"]))
								dsEmail.Tables.Add(ConfigurationManager.AppSettings["Ref_ReasonCategory"]);

							dsEmail.Tables[1].Load(sdr);

						}
						con.Close();
					}

					using (SqlCommand cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLRefAction"]))
					{
						cmd.Connection = con;
						con.Open();
						using (SqlDataReader sdr = cmd.ExecuteReader())
						{
							if (!dsEmail.Tables.Contains(ConfigurationManager.AppSettings["Ref_Action"]))
								dsEmail.Tables.Add(ConfigurationManager.AppSettings["Ref_Action"]);

							dsEmail.Tables[2].Load(sdr);

						}
						con.Close();
					}

					using (SqlCommand cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLUserInfo"]))
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

                   
                    using (SqlCommand cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLServerLogin"]))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (!dsEmail.Tables.Contains(ConfigurationManager.AppSettings["Server_Login"]))
                                dsEmail.Tables.Add(ConfigurationManager.AppSettings["Server_Login"]);

                            dsEmail.Tables[4].Load(sdr);

                        }
                        con.Close();
                    }

                    using (SqlCommand cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLViewAllEmail"]))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (!dsEmail.Tables.Contains(ConfigurationManager.AppSettings["All_Email"]))
                                dsEmail.Tables.Add(ConfigurationManager.AppSettings["All_Email"]);

                            dsEmail.Tables[5].Load(sdr);

                        }
                        con.Close();
                    }

                    
                    //if (counter == false)
                    //{
                        

                    //}
				}
				sw.Stop();
				long temp = sw.ElapsedMilliseconds;
                counter = true;

			}
			catch (Exception ex)
			{
				MessageBoxButtons MsgBox = MessageBoxButtons.OK;
				MessageBox.Show("Error: Load existing email from database into dataset \r\n" + ex.ToString(), "Error", MsgBox);

			}


		   
		}

		private void RetrieveEmail()
		{
			try
			{
				Stopwatch sw = Stopwatch.StartNew();

				//gvEmailList.Columns["Content"].Visible = false;

				gvEmailList.CellClick += new DataGridViewCellEventHandler(gvEmailList_CellContentClick);

				messageCount = client.GetMessageCount();

				List<OpenPop.Mime.Message> allMessages = new List<OpenPop.Mime.Message>(messageCount);

				if (dsEmail.Tables[ConfigurationManager.AppSettings["Email_Detail"]].Rows.Count == 0)
				//if (dsEmail.Tables[0].Rows.Count == 0)
				#region Getting Value for new email content
				{
					InitiateFieldName();
					for (int i = 1; i <= messageCount; i++)
					{
                        try
                        {
                            if(client.GetMessageUid(i).ToString() != "000001e35aa8e91f")
                            {
                                if ((client.GetMessageHeaders(i).ReplyTo == null))
                                {
                                    allMessages.Add(client.GetMessage(i));
                                    headers = client.GetMessageHeaders(i);
                                    message = client.GetMessage(i);
                                    loop_emailXtract(message);
                                    //EmailPartExtracted(message);
                                }
                                //allMessages.Add(client.GetMessage(i));
                                //Check you message is not null
                                //CheckUnicode(headers.Subject);
                                //Check you message is not null
                               

                                if ((_SendTo != "" && _Subject != "") && (_SendTo != null && _Subject != null))//&& (!UIDL.Contains("report")))
                                {
                                    SaveToDB();
                                }
                            }
                            
                        }
                        catch (Exception ex)
                        {
                            CreateLogFile(ex.ToString());
                        }

					}
				}
				#endregion
				else
				#region Add new email to existing database
				{
					DataRow[] temp_DRAEiMS;
					//DateTime FirstDateTime = Convert.ToDateTime(dsEmail.Tables[ConfigurationManager.AppSettings["Email_Detail"]].Rows[dsEmail.Tables[ConfigurationManager.AppSettings["tableName"]].Rows.Count - 1][3].ToString());

					string expression = "sendFrom ='" + _EmailAddress.ToString() + "'";
					string sortOrder = "sendDateTime DESC";
					temp_DRAEiMS = dsEmail.Tables[0].Select(expression, sortOrder);
                    if ((temp_DRAEiMS.Count() != 0 ) &&(temp_DRAEiMS[0].ItemArray[1] != null))
					{
						temp_SendDate = Convert.ToDateTime(temp_DRAEiMS[0].ItemArray[1]).Date;
						temp_SendTime = Convert.ToDateTime(temp_DRAEiMS[0].ItemArray[1]).TimeOfDay;
						DateTime temp_SendDateTime = Convert.ToDateTime(temp_DRAEiMS[0].ItemArray[1]);

						//string temp_Username = dsEmail.Tables[ConfigurationManager.AppSettings["Email_Detail"]].Rows[0][3].ToString().Split(new string[] { "@" }, StringSplitOptions.None).First();
						//string temp_sendFrom = _Username + "@" + _Hostname.Replace("mail", "");

						int ent = gvEmailList.ColumnCount;

						if (gvEmailList.Columns.Contains("Send Date") == false)
						{
							InitiateFieldName();
						}

						for (int i = messageCount; i >= 1; i--)
						{
                            try
                            {
                                #region
                               
                                headers = client.GetMessageHeaders(i);
                                if (client.GetMessageUid(i).ToString() != "000001e35aa8e91f")
                                {
                                    if (headers.ReplyTo == null)
                                    {
                                        CheckUnicode(client.GetMessage(i).Headers.Subject);
                                        headers = client.GetMessageHeaders(i);
                                        message = client.GetMessage(i);
                                        TimeSpan _tempTime = message.Headers.DateSent.TimeOfDay;
                                        DateTime _tempDate = message.Headers.DateSent.Date;
                                        DateTime _tempDateTime = message.Headers.DateSent;

                                        //if (message.Headers.DateSent >= temp_SendDate)

                                        //if ((temp_SendDate < _tempDate) && (temp_SendTime < _tempTime))
                                        if ((temp_SendDateTime < _tempDateTime))
                                        {

                                            headers = client.GetMessageHeaders(i);

                                            //Check you message is not null
                                            if (headers.Subject != null)
                                            {
                                                CheckUnicode(headers.Subject);
                                                loop_emailXtract(message);
                                                //EmailPartExtracted(message);
                                                
                                                //if ((_SendTo != "" && _Subject != "") && (_SendTo != null && _Subject != null))
                                                //{
                                                    SaveToDB();
                                                //}
                                            }
                                        }
                                        else
                                            break;

                                    }
                                
                                }
                                //string temp2 = client.GetMessageHeaders(i).Received;
                                #endregion

                            }
                            catch (Exception ex)
                            {
                                CreateLogFile(ex.ToString());

                            }
							
						}
					}
					else
					{
						int ent = gvEmailList.ColumnCount;

						if (gvEmailList.Columns.Contains("Send Date") == false)
						{
							InitiateFieldName();
						}

						for (int i = messageCount; i >= 1; i--)
						{
                            try
                            {
                                
                                if (client.GetMessageHeaders(i).ReplyTo == null)
                                {
                                    allMessages.Add(client.GetMessage(i));
                                    message = client.GetMessage(i);

                                    #region
                                    headers = client.GetMessageHeaders(i);
                                    //CheckUnicode(headers.Subject);
                                    loop_emailXtract(message);
                                    //EmailPartExtracted(message);

                                    //if ((_SendTo != "" && _Subject != "") && (_SendTo != null && _Subject != null) && (!UIDL.Contains("report")))
                                    //if ((_SendTo != "" && _Subject != "") && (_SendTo != null && _Subject != null))
                                    //{
                                        SaveToDB();
                                    //}
                                    #endregion
                                }
                                
                            }
                            catch (Exception ex)
                            {
                                CreateLogFile(ex.ToString());
                            }
						}
					}

					//dsEmail.Tables.Add(dt);
					int rowtotal = dt.Rows.Count;
				#endregion

				}
			}
			catch (Exception ex)
			{
                CreateLogFile(ex.ToString());
			   
			}
		}

        private void GenerateUIDL(string _tempUIDL)
        {
            UIDL = System.Guid.NewGuid().ToString();
            UIDL = UIDL.Replace("-", "");
            UIDL = UIDL.Substring(0, 25);
        }

		private void EmailPartExtracted(OpenPop.Mime.Message message)
		{
			try
			{
				UIDL = "";
				_SendTo = "";
				_SendFrom = "";
				_Subject = "";
				_Reason = "";
				_ReasonCategory = "";
				_ActionRequired = "";
				_SendFromIPAddress = "";
				_SendDate = DateTime.Now;
				_RejectDateTime = DateTime.Now;
                body = "";
                body1 = "";
                body2 = "";
				//string temphead = headers.Subject.ToString();

				if (headers.Subject != null)
				{
                  
                    //Filter by Email Sender
                    if (headers.From.Raw.Contains("Mail Delivery System"))
                    {
                        #region Mail Delivery System
                            if (headers.Subject.Contains("Undelivered Mail Returned to Sender"))
                            {
                                if (message.MessagePart.MessageParts != null)
                                {
                                    #region MessagePart

                                    OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                                    OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];
                                    OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];

                                    OpenPop.Mime.MessagePart conte = message.FindFirstPlainTextVersion();

                                    string body_content = messagePart.BodyEncoding.GetString(messagePart.Body);
                                    if (messagePart.Body != null)
                                        body = messagePart.BodyEncoding.GetString(messagePart.Body);
                                    if (messagePart1.Body != null)
                                        body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);
                                    if (messagePart2.Body != null)
                                        body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);

                                    EmailDetails(body1, body2);

                                    string _tempIPAddress = body2;
                                    int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                                    int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                    for (int j = ReceivedCount; j > 0; j--)
                                    {
                                        _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                    }
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                    _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                                    //Reject Date
                                    if (message.Headers.Date.Contains("-"))
                                    {
                                        string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                        _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                    }
                                    else if (message.Headers.Date.Contains("+"))
                                    {
                                        string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                        _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                    }

                                    //Actual Reject Reason
                                    if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
                                    {
                                        _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                                        _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                                        _Reason = _Reason.Replace("\r\n", "");
                                        _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();
                                        _Reason = _Reason.Replace("\r\n", "");


                                    }
                                    else if (body1.ToLower().Contains("diagnostic-code:") && body1.ToLower().Contains("x-unix"))
                                    {
                                        _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                        _Reason = _Reason.Split(new string[] { "\r\n550" }, StringSplitOptions.None).Last();

                                    }
                                    else if (body.Contains("The mail system"))
                                    {
                                        _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();
                                        if (_Reason.Contains(" Please"))
                                        {
                                            _Reason = _Reason.Split(new string[] { " Please" }, StringSplitOptions.None).First();
                                        }
                                        else
                                        {
                                            _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();
                                        }
                                    }
                                    else if ((body.Contains("procmail:")) || (body.Contains("permanent fatal errors")) || body.ToLower().Contains("warning"))
                                    {
                                        _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                    }
                                    else
                                    {
                                        _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();

                                    }

                                    RejectedReasonTrim(_Reason);

                                    //Reject Reason Category
                                    GetRejectReasonCategory();

                                    GetActionType();

                                    GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                    var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                                    dt.Rows.Add(newRow);

                                    #endregion
                                }
                                else
                                {
                                    #region HTML Version / Plain Text
                                    StringBuilder builder = new StringBuilder();
                                    OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                                    if (html != null)
                                    {
                                        // We found some plaintext!
                                        builder.Append(html.GetBodyAsText());
                                    }
                                    else
                                    {
                                        // Might include a part holding html instead

                                        OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                                        if (plainText != null)
                                        {
                                            // We found some html!
                                            builder.Append(plainText.GetBodyAsText());
                                        }
                                    }
                                    #endregion

                                }

                            }
                            else if(headers.Subject.Contains("Delivery Status Notification (Failure)"))
                            {
                                OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                                OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];
                                OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];

                                string body_content = messagePart.BodyEncoding.GetString(messagePart.Body);
                                if (messagePart.Body != null)
                                    body = messagePart.BodyEncoding.GetString(messagePart.Body);
                                if (messagePart1.Body != null)
                                    body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);
                                if (messagePart2.Body != null)
                                    body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);

                                EmailDetails(body1, body2);

                                _SendTo = body1.Split(new string[] { "\r\nAction" }, StringSplitOptions.None).First();
                                _SendTo = _SendTo.Split(new string[] { "rfc822;" }, StringSplitOptions.None).Last();
                                _SendTo = Regex.Replace(_SendTo, @"[<>]", "");

                                string _tempIPAddress = body2;
                                int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                                int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                for (int j = ReceivedCount; j > 0; j--)
                                {
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                }
                                _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                                //Reject Date
                                if (message.Headers.Date.Contains("-"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                }
                                else if (message.Headers.Date.Contains("+"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                }

                                //Actual Reject Reason
                                if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Replace("/r/n", "");
                                    _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();


                                }
                                else if (body1.ToLower().Contains("diagnostic-code:") && body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "\r\n550" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "\r\nuntie" }, StringSplitOptions.None).First();

                                }
                                else if (body.Contains("The mail system"))
                                {
                                    _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();
                                    if (_Reason.Contains(" Please"))
                                    {
                                        _Reason = _Reason.Split(new string[] { " Please" }, StringSplitOptions.None).First();
                                    }
                                    else
                                    {
                                        _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();
                                    }
                                }
                                else if ((body.Contains("procmail:")) || (body.Contains("permanent fatal errors")) || body.ToLower().Contains("warning"))
                                {
                                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                }
                                else
                                {
                                    _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();

                                }

                                RejectedReasonTrim(_Reason);

                                //Reject Reason Category
                                GetRejectReasonCategory();

                                GetActionType();

                                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                                dt.Rows.Add(newRow);

                            }

                        #endregion
                    }
                    else if (headers.From.Raw.Contains("Mail Delivery Subsystem"))
                    {
                        #region Mail Delivery Subsystem
                        if (headers.Subject.Contains("Returned mail: see transcript for details"))
                        {
                            if (message.MessagePart.MessageParts != null)
                            {
                                #region MessagePart

                                OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                                OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];
                                OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];

                                OpenPop.Mime.MessagePart conte = message.FindFirstPlainTextVersion();

                                string body_content = messagePart.BodyEncoding.GetString(messagePart.Body);
                                if (messagePart.Body != null)
                                    body = messagePart.BodyEncoding.GetString(messagePart.Body);
                                if (messagePart1.Body != null)
                                    body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);
                                if (messagePart2.Body != null)
                                    body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);

                                EmailDetails(body1, body2);

                                _SendTo = body.Split(new string[] { "\r\n    (reason:" }, StringSplitOptions.None).First();
                                _SendTo = _SendTo.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                _SendTo = Regex.Replace(_SendTo, @"[<>]", "");

                                string _tempIPAddress = body2;
                                int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                                int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                for (int j = ReceivedCount; j > 0; j--)
                                {
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                }
                                _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                                //Reject Date
                                if (message.Headers.Date.Contains("-"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                }
                                else if (message.Headers.Date.Contains("+"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                }

                                //Actual Reject Reason
                                if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Replace("/r/n", "");
                                    _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();


                                }
                                else if (body1.ToLower().Contains("diagnostic-code:") && body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "\r\n550" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "\r\nuntie" }, StringSplitOptions.None).First();

                                }
                                else if (body.Contains("The mail system"))
                                {
                                    _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();
                                    if (_Reason.Contains(" Please"))
                                    {
                                        _Reason = _Reason.Split(new string[] { " Please" }, StringSplitOptions.None).First();
                                    }
                                    else
                                    {
                                        _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();
                                    }
                                }
                                else if ((body.Contains("procmail:")) || (body.Contains("permanent fatal errors")) || body.ToLower().Contains("warning"))
                                {
                                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                }
                                else
                                {
                                    _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();

                                }

                                RejectedReasonTrim(_Reason);

                                //Reject Reason Category
                                GetRejectReasonCategory();

                                GetActionType();

                                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                                dt.Rows.Add(newRow);

                                #endregion
                            }
                            else
                            {
                                #region HTML Version / Plain Text
                                StringBuilder builder = new StringBuilder();
                                OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                                if (html != null)
                                {
                                    // We found some plaintext!
                                    builder.Append(html.GetBodyAsText());
                                }
                                else
                                {
                                    // Might include a part holding html instead

                                    OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                                    if (plainText != null)
                                    {
                                        // We found some html!
                                        builder.Append(plainText.GetBodyAsText());
                                    }
                                }
                                #endregion
                            }

                        }
                    #endregion
                    }
                    else if (headers.From.Raw.Contains("Mailer-daemon@yahoo.com"))
                    {
                        #region Mailer-daemon
                        if (headers.Subject.Contains("Delivery failure"))
                        {
                            if (message.MessagePart.MessageParts != null)
                            {
                                #region MessagePart

                                OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                                OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];
                                OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];

                                OpenPop.Mime.MessagePart conte = message.FindFirstPlainTextVersion();

                                string body_content = messagePart.BodyEncoding.GetString(messagePart.Body);
                                if (messagePart.Body != null)
                                    body = messagePart.BodyEncoding.GetString(messagePart.Body);
                                if (messagePart1.Body != null)
                                    body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);
                                if (messagePart2.Body != null)
                                    body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);

                                EmailDetails(body1, body2);

                                _SendTo = body.Split(new string[] { "\r\n    (reason:" }, StringSplitOptions.None).First();
                                _SendTo = _SendTo.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                _SendTo = Regex.Replace(_SendTo, @"[<>]", "");

                                string _tempIPAddress = body2;
                                int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                                int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                for (int j = ReceivedCount; j > 0; j--)
                                {
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                }
                                _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                                //Reject Date
                                if (message.Headers.Date.Contains("-"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                }
                                else if (message.Headers.Date.Contains("+"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                }

                                //Actual Reject Reason
                                if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Replace("/r/n", "");
                                    _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();


                                }
                                else if (body1.ToLower().Contains("diagnostic-code:") && body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "\r\n550" }, StringSplitOptions.None).Last();

                                }
                                else if (body.Contains("The mail system"))
                                {
                                    _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();
                                    if (_Reason.Contains(" Please"))
                                    {
                                        _Reason = _Reason.Split(new string[] { " Please" }, StringSplitOptions.None).First();
                                    }
                                    else
                                    {
                                        _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();
                                    }
                                }
                                else if ((body.Contains("procmail:")) || (body.Contains("permanent fatal errors")) || body.ToLower().Contains("warning"))
                                {
                                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                }
                                else
                                {
                                    _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();

                                }

                                RejectedReasonTrim(_Reason);

                                //Reject Reason Category
                                GetRejectReasonCategory();

                                GetActionType();

                                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                                dt.Rows.Add(newRow);

                                #endregion
                            }
                            else
                            {
                                #region HTML Version / Plain Text
                                StringBuilder builder = new StringBuilder();
                                OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                                if (html != null)
                                {
                                    // We found some plaintext!
                                    builder.Append(html.GetBodyAsText());
                                }
                                else
                                {
                                    // Might include a part holding html instead

                                    OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                                    if (plainText != null)
                                    {
                                        // We found some html!
                                        builder.Append(plainText.GetBodyAsText());
                                    }
                                }
                                #endregion
                            }

                        }
                        else if (headers.Subject.Contains("Failure Notice"))
                        {

                        }
                        else if (headers.Subject.Contains("failure notice"))
                        {

                        }
                        #endregion

                    }
                    else if (headers.From.Raw.Contains("postmaster@"))
                    {
                        #region postmaster
                        if (headers.Subject.Contains("Undeliverable:"))
                        {
                            if (message.MessagePart.MessageParts != null)
                            {
                                #region Message Part

                                string subs = headers.Subject.ToString();
                                int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                                if (ParMessagecount != 0)
                                {
                                    OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                                    OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                                    if (messagePart.Body != null)
                                        body = messagePart.BodyEncoding.GetString(messagePart.Body);
                                    if (messagePart1.Body != null)
                                        body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                                    if (ParMessagecount > 2)
                                    {
                                        OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                                        body2 = messagePart1.BodyEncoding.GetString(messagePart2.Body);
                                    }
                                }

                                EmailDetails(body1, body2);

                                if (headers.Subject.Contains("未傳遞的主旨"))
                                {
                                    _SendTo = body1.Split(new string[] { "Final-Recipient:" }, StringSplitOptions.None).Last();
                                    _SendTo = _SendTo.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                                    _SendTo = _SendTo.Split(new string[] { ";" }, StringSplitOptions.None).Last();

                                }
                                else if (body1.Contains("rfc822"))
                                {
                                    _SendTo = body1.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                                    _SendTo = _SendTo.Split(new string[] { "rfc822;" }, StringSplitOptions.None).Last();
                                }

                                string _tempIPAddress = body2;

                                if (headers.Subject.Contains("Undeliverable:"))
                                {
                                    int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                    for (int j = ReceivedCount; j > 0; j--)
                                    {
                                        _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nMessage-ID:" }, StringSplitOptions.None).First();
                                        _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                    }
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                    _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                                }
                                else
                                {
                                    int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                    for (int j = ReceivedCount; j > 0; j--)
                                    {
                                        _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                    }
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                    _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                                }
                                //int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;

                                //Reject Date
                                if (message.Headers.Date.Contains("-"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                }
                                else if (message.Headers.Date.Contains("+"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                }

                                //Actual Reject Reason
                                if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Replace("/r/n", "");
                                    _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();

                                }
                                else if (body1.ToLower().Contains("diagnostic-code:") && body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "\r\n550" }, StringSplitOptions.None).Last();

                                }
                                else if (body.Contains("The mail system"))
                                {
                                    _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();
                                    if (_Reason.Contains(" Please"))
                                    {
                                        _Reason = _Reason.Split(new string[] { " Please" }, StringSplitOptions.None).First();
                                    }
                                    else
                                    {
                                        _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();
                                    }
                                }
                                else if ((body.Contains("procmail:")) || (body.Contains("permanent fatal errors")) || body.ToLower().Contains("warning"))
                                {
                                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                }
                                else
                                {
                                    _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();

                                }

                                RejectedReasonTrim(_Reason);

                                //Reject Reason Category
                                GetRejectReasonCategory();

                                GetActionType();

                                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                                dt.Rows.Add(newRow);

                                #endregion
                            }
                            else if ((message.FindFirstHtmlVersion() != null) || (message.FindFirstPlainTextVersion() != null))
                            {
                                #region HTML version / Plain Text

                                StringBuilder builder = new StringBuilder();
                                string _tempSub = headers.Subject;

                                OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                                if (plainText != null)
                                {
                                    //plaintext version
                                    builder.Append(plainText.GetBodyAsText());
                                }
                                else
                                {
                                    //html instead
                                    OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                                    if (html != null)
                                    {
                                        //html version
                                        builder.Append(html.GetBodyAsText());
                                    }
                                }
                                string tempBody = builder.ToString();
                                string _tempSendFromIPAddress = tempBody;

                                int ReceivedCount = Regex.Matches(tempBody, "Received").Count;
                                for (int j = ReceivedCount; j > 0; j--)
                                {
                                    tempBody = tempBody.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                                }
                                _tempSendFromIPAddress = tempBody.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                _SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                                EmailDetails(body1, tempBody);

                                //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                if (message.Headers.Date.Contains("-"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                }
                                else if (message.Headers.Date.Contains("+"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                }

                                _Reason = builder.ToString().Split(new string[] { "\r\nDiagnostic " }, StringSplitOptions.None).First();
                                _Reason = _Reason.Split(new string[] { ">\r\n" }, StringSplitOptions.None).Last();
                                _Reason = _Reason.Replace("\r\n", " ");

                                RejectedReasonTrim(_Reason);
                                GetRejectReasonCategory();
                                GetActionType();
                                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                if (_SendTo != "" && _Subject != "")
                                {
                                    var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                                    dt.Rows.Add(newRow);

                                }

                                #endregion

                            }

                        }
                        else if (headers.Subject.Contains("Undeliverable Maill Returned to Sender"))
                        {
                            if (message.MessagePart.MessageParts != null)
                            {
                                #region Message Part

                                string subs = headers.Subject.ToString();
                                int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                                if (ParMessagecount != 0)
                                {
                                    OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                                    OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                                    if (messagePart.Body != null)
                                        body = messagePart.BodyEncoding.GetString(messagePart.Body);
                                    if (messagePart1.Body != null)
                                        body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                                    if (ParMessagecount > 2)
                                    {
                                        OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                                        body2 = messagePart1.BodyEncoding.GetString(messagePart2.Body);
                                    }
                                }

                                if (headers.Subject.Contains("退信")) // body.Contains("退信") || body1.Contains("退信") || body2.Contains("退信"))
                                {

                                }


                                EmailDetails(body1, body2);

                                if (headers.Subject.Contains("未傳遞的主旨"))
                                {
                                    _SendTo = body1.Split(new string[] { "Final-Recipient:" }, StringSplitOptions.None).Last();
                                    _SendTo = _SendTo.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                                    _SendTo = _SendTo.Split(new string[] { ";" }, StringSplitOptions.None).Last();

                                }

                                string _tempIPAddress = body2;

                                if (headers.Subject.Contains("Undeliverable:"))
                                {
                                    int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                    for (int j = ReceivedCount; j > 0; j--)
                                    {
                                        _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nMessage-ID:" }, StringSplitOptions.None).First();
                                        _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                    }
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                    _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                                }
                                else
                                {
                                    int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                    for (int j = ReceivedCount; j > 0; j--)
                                    {
                                        _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                    }
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                    _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                                }
                                //int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;

                                //Reject Date
                                if (message.Headers.Date.Contains("-"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                }
                                else if (message.Headers.Date.Contains("+"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                }

                                //Actual Reject Reason
                                if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Replace("/r/n", "");
                                    _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();

                                }
                                else if (body1.ToLower().Contains("diagnostic-code:") && body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "\r\n550" }, StringSplitOptions.None).Last();

                                }
                                else if (body.Contains("The mail system"))
                                {
                                    _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();
                                    if (_Reason.Contains(" Please"))
                                    {
                                        _Reason = _Reason.Split(new string[] { " Please" }, StringSplitOptions.None).First();
                                    }
                                    else
                                    {
                                        _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();
                                    }
                                }
                                else if ((body.Contains("procmail:")) || (body.Contains("permanent fatal errors")) || body.ToLower().Contains("warning"))
                                {
                                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                }
                                else
                                {
                                    _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();

                                }

                                RejectedReasonTrim(_Reason);

                                //Reject Reason Category
                                GetRejectReasonCategory();

                                GetActionType();

                                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                                dt.Rows.Add(newRow);



                                #endregion
                            }
                            else if ((message.FindFirstHtmlVersion() != null) || (message.FindFirstPlainTextVersion() != null))
                            {
                                #region HTML version / Plain Text

                                StringBuilder builder = new StringBuilder();
                                string _tempSub = headers.Subject;

                                OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                                if (plainText != null)
                                {
                                    //plaintext version
                                    builder.Append(plainText.GetBodyAsText());
                                }
                                else
                                {
                                    //html instead
                                    OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                                    if (html != null)
                                    {
                                        //html version
                                        builder.Append(html.GetBodyAsText());
                                    }
                                }
                                string tempBody = builder.ToString();
                                string _tempSendFromIPAddress = tempBody;

                                int ReceivedCount = Regex.Matches(tempBody, "Received").Count;
                                for (int j = ReceivedCount; j > 0; j--)
                                {
                                    tempBody = tempBody.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                                }
                                _tempSendFromIPAddress = tempBody.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                _SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                                EmailDetails(body1, body2);

                                //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                if (message.Headers.Date.Contains("-"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                }
                                else if (message.Headers.Date.Contains("+"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                }

                                _Reason = builder.ToString().Split(new string[] { "\r\nDiagnostic " }, StringSplitOptions.None).First();
                                _Reason = _Reason.Split(new string[] { ">\r\n" }, StringSplitOptions.None).Last();
                                _Reason = _Reason.Replace("\r\n", " ");

                                RejectedReasonTrim(_Reason);
                                GetRejectReasonCategory();
                                GetActionType();
                                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                if (_SendTo != "" && _Subject != "")
                                {
                                    var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                                    dt.Rows.Add(newRow);

                                }

                                #endregion

                            }

                        }
                        else if (headers.Subject.Contains("未傳遞的"))
                        {
                            int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                            if (ParMessagecount != 0)
                            {

                                OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                                OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                                string subject = headers.Subject.ToString();
                                if (messagePart.Body != null)
                                    body = messagePart.BodyEncoding.GetString(messagePart.Body);
                                if (messagePart1.Body != null)
                                    body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                                if (ParMessagecount > 2)
                                {
                                    OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                                    body2 = messagePart1.BodyEncoding.GetString(messagePart2.Body);
                                }

                                EmailDetails(body1, body2);

                                _SendTo = body1.Split(new string [] {"\r\nAction:"},StringSplitOptions.None).First();
                                _SendTo = _SendTo.Split(new string [] {"rfc822;"},StringSplitOptions.None).Last();

                                string _tempIPAddress = body2;
                                int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                                int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                for (int j = ReceivedCount; j > 0; j--)
                                {
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                }
                                _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                                //Reject Date
                                if (message.Headers.Date.Contains("-"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                }
                                else if (message.Headers.Date.Contains("+"))
                                {
                                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                }

                                if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "\r\n " }, StringSplitOptions.None).First();
                                    _Reason = _Reason.Split(new string[] { ";" }, StringSplitOptions.None).Last();


                                }
                                else if (body1.ToLower().Contains("diagnostic-code:") && body1.ToLower().Contains("x-unix"))
                                {
                                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "\r\n550" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { "\r\nuntie" }, StringSplitOptions.None).First();

                                }

                                RejectedReasonTrim(_Reason);

                                //Reject Reason Category
                                GetRejectReasonCategory();

                                GetActionType();

                                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                                dt.Rows.Add(newRow);
                                //_Reason = body1.

                            }

                        }
                       
                        else
                        {
                            #region Message Part
                            int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                            if (ParMessagecount != 0)
                            {

                                OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                                OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                                string subject = headers.Subject.ToString();
                                if (messagePart.Body != null)
                                    body = messagePart.BodyEncoding.GetString(messagePart.Body);
                                if (messagePart1.Body != null)
                                    body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                                if (ParMessagecount > 2)
                                {
                                    OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                                    body2 = messagePart1.BodyEncoding.GetString(messagePart2.Body);
                                }

                                if (body.Contains("黑名單") || body1.Contains("黑名單") || body2.Contains("黑名單"))
                                {

                                }
                                else if (body.Contains("拒收") || body1.Contains("拒收") || body2.Contains("拒收"))
                                {
                                    _Reason = body.Split(new string[] { ".<br><br>" }, StringSplitOptions.None).First();
                                    _Reason = _Reason.Split(new string[] { "<br>\r\n\r\n" }, StringSplitOptions.None).Last();
                                }
                                else if (body.Contains("退回") || body1.Contains("退回") || body2.Contains("退回"))
                                {
                                    _Reason = body.Split(new string[] { "ul_lst" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Split(new string[] { ".</li>\r\n" }, StringSplitOptions.None).First();
                                    _Reason = _Reason.Split(new string[] { "<li>" }, StringSplitOptions.None).Last();
                                    _Reason = _Reason.Replace("/<br>", "");

                                }
                                else if (body.Contains("退信") || body1.Contains("退信") || body2.Contains("退信"))
                                {

                                }

                                if (_Reason.Contains("黑名單") || _Reason.Contains("拒收") || _Reason.Contains("退回") || _Reason.Contains("退信"))
                                {
                                    EmailDetails(body1, body2);

                                    string _tempIPAddress = body2;
                                    int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                                    int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                    for (int j = ReceivedCount; j > 0; j--)
                                    {
                                        _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                    }
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                    _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                                    //Reject Date
                                    if (message.Headers.Date.Contains("-"))
                                    {
                                        string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                        _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                    }
                                    else if (message.Headers.Date.Contains("+"))
                                    {
                                        string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                        _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                    }

                                    RejectedReasonTrim(_Reason);

                                    //Reject Reason Category
                                    GetRejectReasonCategory();

                                    GetActionType();

                                    GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                    var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                                    dt.Rows.Add(newRow);
                                }
                            }
                            #endregion

                            else
                            {
                                #region HTML Version / Plain Text

                                #endregion

                            }

                        }
                        #endregion
                    }
                    else if ((headers.Subject.Contains("failure notice")||(headers.Subject.Contains("failure notice"))))
                    {
                        #region Failure Notice
                        if (message.MessagePart.MessageParts != null)
                        {
                            #region Message Part

                            string subs = headers.Subject.ToString();
                            int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                            if (ParMessagecount != 0)
                            {
                                OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                                OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                                if (messagePart.Body != null)
                                    body = messagePart.BodyEncoding.GetString(messagePart.Body);
                                if (messagePart1.Body != null)
                                    body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                                if (ParMessagecount > 2)
                                {
                                    OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                                    body2 = messagePart1.BodyEncoding.GetString(messagePart2.Body);
                                }
                            }

                            EmailDetails(body1, body2);

                            if (headers.Subject.Contains("未傳遞的主旨"))
                            {
                                _SendTo = body1.Split(new string[] { "Final-Recipient:" }, StringSplitOptions.None).Last();
                                _SendTo = _SendTo.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                                _SendTo = _SendTo.Split(new string[] { ";" }, StringSplitOptions.None).Last();

                            }

                            string _tempIPAddress = body2;

                            if (headers.Subject.Contains("Undeliverable:"))
                            {
                                int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                for (int j = ReceivedCount; j > 0; j--)
                                {
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nMessage-ID:" }, StringSplitOptions.None).First();
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                }
                                _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                            }
                            else
                            {
                                int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                for (int j = ReceivedCount; j > 0; j--)
                                {
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                                }
                                _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                            }

                            //Reject Date
                            if (message.Headers.Date.Contains("-"))
                            {
                                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                            }
                            else if (message.Headers.Date.Contains("+"))
                            {
                                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                            }

                            //Actual Reject Reason
                            if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
                            {
                                _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                                _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                                _Reason = _Reason.Replace("/r/n", "");
                                _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();

                            }
                            else if (body1.ToLower().Contains("diagnostic-code:") && body1.ToLower().Contains("x-unix"))
                            {
                                _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                _Reason = _Reason.Split(new string[] { "\r\n550" }, StringSplitOptions.None).Last();

                            }
                            else if (body.Contains("The mail system"))
                            {
                                _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();
                                if (_Reason.Contains(" Please"))
                                {
                                    _Reason = _Reason.Split(new string[] { " Please" }, StringSplitOptions.None).First();
                                }
                                else
                                {
                                    _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();
                                }
                            }
                            else if ((body.Contains("procmail:")) || (body.Contains("permanent fatal errors")) || body.ToLower().Contains("warning"))
                            {
                                _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                            }
                            else
                            {
                                _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();

                            }

                            RejectedReasonTrim(_Reason);

                            //Reject Reason Category
                            GetRejectReasonCategory();

                            GetActionType();

                            GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                            var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                            dt.Rows.Add(newRow);



                            #endregion
                        }
                        else if ((message.FindFirstHtmlVersion() != null) || (message.FindFirstPlainTextVersion() != null))
                        {
                            #region HTML version / Plain Text

                            StringBuilder builder = new StringBuilder();
                            string _tempSub = headers.Subject;

                            OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                            if (plainText != null)
                            {
                                //plaintext version
                                builder.Append(plainText.GetBodyAsText());
                            }
                            else
                            {
                                //html instead
                                OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                                if (html != null)
                                {
                                    //html version
                                    builder.Append(html.GetBodyAsText());
                                }
                            }
                            string tempBody = builder.ToString();
                            string _tempSendFromIPAddress = tempBody;

                            int ReceivedCount = Regex.Matches(tempBody, "Received").Count;
                            for (int j = ReceivedCount; j > 0; j--)
                            {
                                tempBody = tempBody.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                            }
                            _tempSendFromIPAddress = tempBody.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                            _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                            _SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                            EmailDetails(body1, tempBody);

                            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                            if (message.Headers.Date.Contains("-"))
                            {
                                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                            }
                            else if (message.Headers.Date.Contains("+"))
                            {
                                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                            }

                            if (builder.ToString().Contains("Diagnostic"))
                            {
                                _Reason = builder.ToString().Split(new string[] { "\r\nDiagnostic " }, StringSplitOptions.None).First();
                                _Reason = _Reason.Split(new string[] { ">\r\n" }, StringSplitOptions.None).Last();
                                _Reason = _Reason.Replace("\r\n", " ");
                            }
                            else
                            {
                                _Reason = builder.ToString();
                                _Reason = _Reason.Split(new string[] { "\r\n--- " },StringSplitOptions.None).First();
                                _Reason = _Reason.Split(new string[] { "Subject: " }, StringSplitOptions.None).Last();
                                _Reason = _Reason.Replace("//", "");

                                while (_Reason.Contains("\r\n"))
                                {
                                    _Reason = _Reason.Replace("\r\n", " ");

                                }

                            }
                            

                            RejectedReasonTrim(_Reason);
                            GetRejectReasonCategory();
                            GetActionType();
                            GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                            if (_SendTo != "" && _Subject != "")
                            {
                                var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                                dt.Rows.Add(newRow);

                            }

                            #endregion

                        }
                        #endregion

                    }

                    #region Remarked on 17/12/2018
                    //End of Filter by Email Sender
                    //==================================
                    //Filter by email Subject
                    //else if (message != null && (headers.Subject.Contains("Returned mail") || headers.Subject.Contains("Undelivered Mail")))
                    //{
                    //    #region Returned Mail / Undelivered Mail
                    //    if (message.MessagePart.MessageParts != null)
                    //    {
                    //        #region

                    //        OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                    //        OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];
                    //        OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];

                    //        OpenPop.Mime.MessagePart conte = message.FindFirstPlainTextVersion();

                    //        string body_content = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //        if (messagePart.Body != null)
                    //            body = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //        if (messagePart1.Body != null)
                    //            body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);
                    //        if (messagePart2.Body != null)
                    //            body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);

                    //        EmailDetails(body2);

                    //        _SendTo = body.Split(new string[] { "\r\n    (reason:" }, StringSplitOptions.None).First();
                    //        _SendTo = _SendTo.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                    //        _SendTo = Regex.Replace(_SendTo, @"[<>]", "");

                    //        string _tempIPAddress = body2;
                    //        int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                    //        int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                    //        for (int j = ReceivedCount; j > 0; j--)
                    //        {
                    //            _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                    //        }
                    //        _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //        _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //        _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                    //        //Reject Date
                    //        if (message.Headers.Date.Contains("-"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        }
                    //        else if (message.Headers.Date.Contains("+"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //        }

                    //        //Actual Reject Reason
                    //        if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
                    //        {
                    //            _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                    //            _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                    //            _Reason = _Reason.Replace("/r/n", "");
                    //            _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();


                    //        }
                    //        else if (body1.ToLower().Contains("diagnostic-code:") && body1.ToLower().Contains("x-unix"))
                    //        {
                    //            _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                    //            _Reason = _Reason.Split(new string[] { "\r\n550" }, StringSplitOptions.None).Last();

                    //        }
                    //        else if (body.Contains("The mail system"))
                    //        {
                    //            _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();
                    //            if (_Reason.Contains(" Please"))
                    //            {
                    //                _Reason = _Reason.Split(new string[] { " Please" }, StringSplitOptions.None).First();
                    //            }
                    //            else
                    //            {
                    //                _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();
                    //            }
                    //        }
                    //        else if ((body.Contains("procmail:")) || (body.Contains("permanent fatal errors")) || body.ToLower().Contains("warning"))
                    //        {
                    //            _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                    //        }
                    //        else
                    //        {
                    //            _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();

                    //        }

                    //        RejectedReasonTrim(_Reason);

                    //        //Reject Reason Category
                    //        GetRejectReasonCategory();

                    //        GetActionType();

                    //        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //        var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                    //        dt.Rows.Add(newRow);

                    //        #endregion
                    //    }
                    //    else
                    //    {
                    //        StringBuilder builder = new StringBuilder();
                    //        OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                    //        if (html != null)
                    //        {
                    //            // We found some plaintext!
                    //            builder.Append(html.GetBodyAsText());
                    //        }
                    //        else
                    //        {
                    //            // Might include a part holding html instead

                    //            OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    //            if (plainText != null)
                    //            {
                    //                // We found some html!
                    //                builder.Append(plainText.GetBodyAsText());
                    //            }
                    //        }
                    //    }
                    //    #endregion
                    //}
                    //else if (message != null && (headers.Subject.ToLower().Contains("failure notice") || headers.Subject.Contains("Failure Notice")))
                    //{
                    //    #region Failure Notice
                    //    if ((message.MessagePart.MessageParts != null) && ((message.FindFirstHtmlVersion() != null) || (message.FindFirstPlainTextVersion() != null)))
                    //    {
                    //        #region
                    //        StringBuilder builder = new StringBuilder();
                    //        string _tempSub = headers.Subject;

                    //        OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    //        if (plainText != null)
                    //        {
                    //            // We found some plaintext!
                    //            builder.Append(plainText.GetBodyAsText());
                    //        }
                    //        else
                    //        {
                    //            // Might include a part holding html instead
                    //            OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                    //            if (html != null)
                    //            {
                    //                // We found some html!
                    //                builder.Append(html.GetBodyAsText());
                    //            }
                    //        }
                    //        string tempBody = builder.ToString();
                    //        string _tempSendFromIPAddress = tempBody;

                    //        for (int j = 5; j > 0; j--)
                    //        {
                    //            tempBody = tempBody.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                    //        }
                    //        _tempSendFromIPAddress = tempBody.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //        _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //        _SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                    //        EmailDetails(tempBody);

                    //        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        if (message.Headers.Date.Contains("-"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        }
                    //        else if (message.Headers.Date.Contains("+"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //        }

                    //        if (builder.ToString().Contains("Diagnostic"))
                    //        {
                    //            _Reason = builder.ToString().Split(new string[] { "\r\nDiagnostic " }, StringSplitOptions.None).First();
                    //            _Reason = _Reason.Split(new string[] { "--- Below this line is a copy of the message." }, StringSplitOptions.None).First();
                    //            _Reason = _Reason.Split(new string[] { "Unfortunately, " }, StringSplitOptions.None).Last();
                    //            _Reason = _Reason.Split(new string[] { " You " }, StringSplitOptions.None).First();
                    //        }
                    //        else
                    //        {
                    //            _Reason = builder.ToString().Split(new string[] { "\r\n\r\n\r\n--- " }, StringSplitOptions.None).First();
                    //            _Reason = _Reason.Split(new string[] { "\r\n\r\n" }, StringSplitOptions.None).Last();
                    //            _Reason = _Reason.Replace("\r\n", " ");
                    //        }

                    //        RejectedReasonTrim(_Reason);
                    //        GetRejectReasonCategory();
                    //        GetActionType();
                    //        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //        if (_SendTo != "" && _Subject != "")
                    //        {
                    //            var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //            dt.Rows.Add(newRow);

                    //        }

                    //        #endregion
                    //    }
                    //    else if ((message.FindFirstHtmlVersion() != null) || (message.FindFirstPlainTextVersion() != null))
                    //    {
                    //        #region

                    //        //StringBuilder builder = new StringBuilder();
                    //        //string _tempSub = headers.Subject;

                    //        //OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    //        //if (plainText != null)
                    //        //{
                    //        //    // We found some plaintext!
                    //        //    builder.Append(plainText.GetBodyAsText());
                    //        //}
                    //        //else
                    //        //{
                    //        //    // Might include a part holding html instead
                    //        //    OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                    //        //    if (html != null)
                    //        //    {
                    //        //        // We found some html!
                    //        //        builder.Append(html.GetBodyAsText());
                    //        //    }
                    //        //}
                    //        //string tempBody = builder.ToString();
                    //        //string _tempSendFromIPAddress = tempBody;

                    //        //int ReceivedCount = Regex.Matches(tempBody, "Received").Count;
                    //        //for (int j = ReceivedCount; j > 0; j--)
                    //        //{
                    //        //    tempBody = tempBody.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                    //        //}
                    //        //_tempSendFromIPAddress = tempBody.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //        //_tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //        //_SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                    //        //EmailDetails(tempBody);

                    //        ////_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        //if (message.Headers.Date.Contains("-"))
                    //        //{
                    //        //    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //        //    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //        //    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        //}
                    //        //else if (message.Headers.Date.Contains("+"))
                    //        //{
                    //        //    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //        //    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //        //}

                    //        //_Reason = builder.ToString().Split(new string[] { "\r\nDiagnostic " }, StringSplitOptions.None).First();
                    //        //_Reason = _Reason.Split(new string[] { "--- Below this line is a copy of the message." }, StringSplitOptions.None).First();
                    //        //_Reason = _Reason.Split(new string[] { "Unfortunately, " }, StringSplitOptions.None).Last();
                    //        //_Reason = _Reason.Split(new string[] { " You " }, StringSplitOptions.None).First();


                    //        //RejectedReasonTrim(_Reason);
                    //        //GetRejectReasonCategory();
                    //        //GetActionType();
                    //        //GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //        //if (_SendTo != "" && _Subject != "")
                    //        //{
                    //        //    var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //        //    dt.Rows.Add(newRow);

                    //        //}

                    //        #endregion
                    //    }
                    //    #endregion
                    //}
                    //else
                    //{

                    //    if (message.MessagePart.MessageParts != null) // && ((message.FindFirstHtmlVersion() != null) || (message.FindFirstPlainTextVersion() != null)))
                    //    {
                    //        #region
                    //        string subs = headers.Subject.ToString();
                    //        int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                    //        if (ParMessagecount != 0)
                    //        {
                    //            OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                    //            OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                    //            if (messagePart.Body != null)
                    //                body = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //            if (messagePart1.Body != null)
                    //                body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                    //            if (ParMessagecount > 2)
                    //            {
                    //                OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                    //                body2 = messagePart1.BodyEncoding.GetString(messagePart2.Body);
                    //            }

                    //            if (body.Contains("黑名單") || body1.Contains("黑名單") || body2.Contains("黑名單"))
                    //            {

                    //            }
                    //            else if (body.Contains("拒收") || body1.Contains("拒收") || body2.Contains("拒收"))
                    //            {
                    //                _Reason = body.Split(new string[] { ".<br><br>" }, StringSplitOptions.None).First();
                    //                _Reason = _Reason.Split(new string[] { "<br>\r\n\r\n" }, StringSplitOptions.None).Last();
                    //            }
                    //            else if (body.Contains("退回") || body1.Contains("退回") || body2.Contains("退回"))
                    //            {

                    //            }
                    //            else if (body.Contains("退信") || body1.Contains("退信") || body2.Contains("退信"))
                    //            {

                    //            }

                    //            if (_Reason.Contains("黑名單") || _Reason.Contains("拒收") || _Reason.Contains("退回") || _Reason.Contains("退信"))
                    //            {
                    //                EmailDetails(body2);

                    //                string _tempIPAddress = body2;
                    //                int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                    //                int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                    //                for (int j = ReceivedCount; j > 0; j--)
                    //                {
                    //                    _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                    //                }
                    //                _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //                _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //                _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                    //                //Reject Date
                    //                if (message.Headers.Date.Contains("-"))
                    //                {
                    //                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //                    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //                }
                    //                else if (message.Headers.Date.Contains("+"))
                    //                {
                    //                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //                }

                    //                RejectedReasonTrim(_Reason);

                    //                //Reject Reason Category
                    //                GetRejectReasonCategory();

                    //                GetActionType();

                    //                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //                var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                    //                dt.Rows.Add(newRow);


                    //            }

                    //        }

                    //        #endregion

                    //    }
                    //    else if ((message.FindFirstHtmlVersion() != null) || (message.FindFirstPlainTextVersion() != null))
                    //    {
                    //        StringBuilder builder = new StringBuilder();
                    //        string _tempSub = headers.Subject;

                    //        OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    //        if (plainText != null)
                    //        {
                    //            // We found some plaintext!
                    //            builder.Append(plainText.GetBodyAsText());
                    //        }
                    //        else
                    //        {
                    //            // Might include a part holding html instead
                    //            OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                    //            if (html != null)
                    //            {
                    //                // We found some html!
                    //                builder.Append(html.GetBodyAsText());
                    //            }
                    //        }
                    //        string tempBody = builder.ToString();

                    //        if (tempBody.Contains("黑名單"))
                    //        {

                    //        }
                    //        else if (tempBody.Contains("拒收"))
                    //        {

                    //        }
                    //        else if (tempBody.Contains("退回"))
                    //        {

                    //        }
                    //        else if (tempBody.Contains("退信"))
                    //        {


                    //        }

                    //    }

                    //}
#endregion
                    //End of Filter by Email Subject
                    //==================================
                    //Filter by email Subject
                    
                    //Remark on 14/12/2018
                    #region Previously used subject to filter email
                    //if (message != null && (headers.Subject.Contains("Returned mail") || headers.Subject.Contains("Undelivered Mail") || headers.Subject.Contains("Warning")))
                    //{
                    //    if (message.MessagePart.MessageParts != null)
                    //    {
                    //        #region
                    //        StringBuilder builder = new StringBuilder();
                    //        OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                    //        if (html != null)
                    //        {
                    //            // We found some plaintext!
                    //            builder.Append(html.GetBodyAsText());
                    //        }
                    //        else
                    //        {
                    //            // Might include a part holding html instead

                    //            OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    //            if (plainText != null)
                    //            {
                    //                // We found some html!
                    //                builder.Append(plainText.GetBodyAsText());
                    //            }
                    //        }
                    //        string tempBody = builder.ToString();

                    //        OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                    //        OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];
                    //        OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];

                    //        OpenPop.Mime.MessagePart conte = message.FindFirstPlainTextVersion();

                    //        string body_content = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //        if (messagePart.Body != null)
                    //            body = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //        if (messagePart1.Body != null)
                    //            body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);
                    //        if (messagePart2.Body != null)
                    //            body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);

                    //        if (body2.Contains("estatement@phillip.com.hk"))
                    //        {

                    //        }

                    //        string _TempSendDate = body2.Split(new string[] { "Date:" }, StringSplitOptions.None).Last();
                    //        _SendDate = Convert.ToDateTime(_TempSendDate.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        if(_SendDate == Convert.ToDateTime("23/11/2018").Date)
                    //        {

                    //        }

                    //        string sender = headers.From.Address;
                    //        string _tempSub = headers.Subject;

                    //        if (body.Contains("73"))
                    //        {

                    //        }

                    //        EmailDetails(body2);

                    //        string _tempIPAddress = body2;
                    //        int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                    //        int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                    //        for (int j = ReceivedCount; j > 0; j--)
                    //        {
                    //            _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                    //        }
                    //        _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //        _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //        _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                    //        //Reject Date
                    //        if (message.Headers.Date.Contains("-"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        }
                    //        else if (message.Headers.Date.Contains("+"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //        }

                    //        //Actual Reject Reason
                    //        if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
                    //        {
                    //            _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                    //            _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                    //            _Reason = _Reason.Replace("/r/n", "");
                    //            _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();


                    //        }
                    //        else if (body1.ToLower().Contains("diagnostic-code:") && body1.ToLower().Contains("x-unix"))
                    //        {
                    //            _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                    //            _Reason = _Reason.Split(new string[] { "\r\n550" }, StringSplitOptions.None).Last();

                    //        }
                    //        else if (body.Contains("The mail system"))
                    //        {
                    //            _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();
                    //            if (_Reason.Contains(" Please"))
                    //            {
                    //                _Reason = _Reason.Split(new string[] { " Please" }, StringSplitOptions.None).First();
                    //            }
                    //            else
                    //            {
                    //                _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();
                    //            }
                    //        }
                    //        else if ((body.Contains("procmail:")) || (body.Contains("permanent fatal errors")) || body.ToLower().Contains("warning"))
                    //        {
                    //            _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                    //        }
                    //        else
                    //        {
                    //            _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();

                    //        }

                    //        if(_SendTo == "chukkwanwong@phillip.com.hk")
                    //        {
                            
                    //        }

                    //        RejectedReasonTrim(_Reason);

                    //        //Reject Reason Category
                    //        GetRejectReasonCategory();

                    //        GetActionType();

                    //        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //        var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                    //        dt.Rows.Add(newRow);

                    //        #endregion
                    //    }
                    //}
                    //else if (headers.Subject.Contains("Delivery Status Notification (Failure)"))
                    //{
                    //    if (message.MessagePart.MessageParts != null)
                    //    {
                    //        #region
                    //        StringBuilder builder = new StringBuilder();
                    //        string tempBody = "", _tempSendFromIPAddress;
                    //        int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                    //        if (ParMessagecount != 0)
                    //        {
                    //            OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                    //            OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];
                    //            OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];

                    //            OpenPop.Mime.MessagePart conte = message.FindFirstPlainTextVersion();

                    //            //string body_content = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //            if (messagePart.Body != null)
                    //                body = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //            if (messagePart1.Body != null)
                    //                body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);
                    //            if (messagePart2.Body != null)
                    //                body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);

                    //            string _TempSendDate = body2.Split(new string[] { "Date:" }, StringSplitOptions.None).Last();
                    //            _SendDate = Convert.ToDateTime(_TempSendDate.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //            if (_SendDate == Convert.ToDateTime("23/11/2018").Date)
                    //            {

                    //            }
                    //            //UIDL = message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First();
                    //            ////UIDL = UIDL.Substring(13);
                    //            //UIDL = Regex.Replace(UIDL, @"[.-]", "");
                    //            //UIDL = (UIDL.Length > 25) ? UIDL.Substring(UIDL.Length - 25, 25) : UIDL;
                    //            //UIDL = System.Guid.NewGuid().ToString();
                    //            //UIDL = Regex.Replace(UIDL, @"[.-]", "");
                    //            //UIDL = (UIDL.Length > 25) ? UIDL.Substring(0, 25) : UIDL;

                    //            EmailDetails(body2);

                    //            string _tempIPAddress = body2;
                    //            int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                    //            for (int j = 5; j > 0; j--)
                    //            {
                    //                _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                    //            }
                    //            _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //            _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //            _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                    //            //Reject Date
                    //            if (message.Headers.Date.Contains("-"))
                    //            {
                    //                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //                //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //            }
                    //            else if (message.Headers.Date.Contains("+"))
                    //            {
                    //                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //            }



                    //            //Actual Reject Reason
                    //            if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
                    //            {
                    //                _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                    //                _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                    //                _Reason = _Reason.Replace("/r/n", "");
                    //                _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();


                    //            }
                    //            else if (body.Contains("The mail system"))
                    //            {
                    //                _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();
                    //                if (_Reason.Contains(" Please"))
                    //                {
                    //                    _Reason = _Reason.Split(new string[] { " Please" }, StringSplitOptions.None).First();
                    //                }
                    //                else
                    //                {
                    //                    _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();
                    //                }
                    //            }
                    //            else if ((body.Contains("procmail:")) || (body.Contains("permanent fatal errors")) || body.ToLower().Contains("warning"))
                    //            {
                    //                _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                    //            }
                    //            else
                    //            {
                    //                _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();

                    //            }

                    //            RejectedReasonTrim(_Reason);

                    //            //Reject Reason Category
                    //            GetRejectReasonCategory();

                    //            GetActionType();

                    //        }
                    //        else
                    //        {
                    //            string _tempSub = headers.Subject;

                    //            OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    //            if (plainText != null)
                    //            {
                    //                // We found some plaintext!
                    //                builder.Append(plainText.GetBodyAsText());
                    //            }
                    //            else
                    //            {
                    //                // Might include a part holding html instead
                    //                OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                    //                if (html != null)
                    //                {
                    //                    // We found some html!
                    //                    builder.Append(html.GetBodyAsText());
                    //                }
                    //            }
                    //            tempBody = builder.ToString();
                    //            _tempSendFromIPAddress = tempBody;

                    //            //UIDL = tempBody.Split(new string[] { "Message-ID: <" }, StringSplitOptions.None).Last();
                    //            //UIDL = UIDL.Split(new string[] { ">\r\n" }, StringSplitOptions.None).First();
                    //            //UIDL = UIDL.Split(new string[] { "@" }, StringSplitOptions.None).First();
                    //            //UIDL = (UIDL.Length > 25) ? UIDL.Substring(UIDL.Length - 25, 25) : UIDL;
                    //            UIDL = System.Guid.NewGuid().ToString();
                    //            UIDL = Regex.Replace(UIDL, @"[.-]", "");
                    //            UIDL = (UIDL.Length > 25) ? UIDL.Substring(0, 25) : UIDL;

                    //            for (int j = 5; j > 0; j--)
                    //            {
                    //                tempBody = tempBody.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                    //            }
                    //            _tempSendFromIPAddress = tempBody.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //            _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //            _SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                    //            EmailDetails(tempBody);

                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.DateSent); //   .Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //            if (message.Headers.Date.Contains("-"))
                    //            {
                    //                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //                //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //            }
                    //            else if (message.Headers.Date.Contains("+"))
                    //            {
                    //                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //            }

                    //            _Reason = builder.ToString().Split(new string[] { ". " }, StringSplitOptions.None).First();
                    //            _Reason = _Reason.Split(new string[] { "\r\n\r\n" }, StringSplitOptions.None).Last();
                    //            _Reason = _Reason.Replace("\r\n", " ");

                    //            RejectedReasonTrim(_Reason);
                    //            GetRejectReasonCategory();
                    //            GetActionType();
                    //        }

                    //        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //        if (_SendTo != "" && _Subject != "")
                    //        {
                    //            var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //            dt.Rows.Add(newRow);

                    //        }
                    //        #endregion
                    //    }

                    //}
                    ////Bounce Email Subject: Delivery Failure
                    //else if (headers.Subject.ToLower().Contains("delivery failure") || headers.Subject.Contains("Delivery Status Notification"))
                    //{
                    //    if (message.MessagePart.MessageParts != null)
                    //    {
                    //        #region
                    //        int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);

                    //        OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                    //        OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                    //        if (messagePart.Body != null)
                    //            body = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //        if (messagePart1.Body != null)
                    //            body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                    //        if (ParMessagecount > 2)
                    //        {
                    //            OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                    //            if (messagePart2.Body != null)
                    //                body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);

                    //        }

                    //        string _TempSendDate = body2.Split(new string[] { "Date:" }, StringSplitOptions.None).Last();
                    //        _SendDate = Convert.ToDateTime(_TempSendDate.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        if (_SendDate == Convert.ToDateTime("23/11/2018").Date)
                    //        {

                    //        }
                    //        //UIDL = message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First();
                    //        ////UIDL = UIDL.Substring(13);
                    //        //UIDL = Regex.Replace(UIDL, @"[.-@]", "");
                    //        //UIDL = (UIDL.Length > 25) ? UIDL.Substring(UIDL.Length - 25, 25) : UIDL;
                    //        //UIDL = System.Guid.NewGuid().ToString();
                    //        //UIDL = Regex.Replace(UIDL, @"[.-]", "");
                    //        //UIDL = (UIDL.Length > 25) ? UIDL.Substring(0, 25) : UIDL;
                    //        EmailDetails(body2);

                    //        string _tempIPAddress = body2;
                    //        int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                    //        for (int j = 5; j > 0; j--)
                    //        {
                    //            _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                    //        }
                    //        _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //        _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //        _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                    //        //Reject Date
                    //        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        if (message.Headers.Date.Contains("-"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        }
                    //        else if (message.Headers.Date.Contains("+"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //        }


                    //        //Actual Reject Reason
                    //        if (body1.Contains("Diagnostic-Code:"))
                    //        {
                    //            _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                    //            _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                    //            //_Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();

                    //        }

                    //        RejectedReasonTrim(_Reason);

                    //        //Reject Reason Category
                    //        GetRejectReasonCategory();

                    //        GetActionType();

                    //        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //        var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //        dt.Rows.Add(newRow);

                    //        #endregion
                    //    }
                    //}
                    ////Bounce Email Subject: Contain Chinese Character
                    //else if (headers.Subject.Contains("Failure Notice"))
                    //{
                    //    if (message.MessagePart.MessageParts != null)
                    //    {
                    //        #region
                    //        StringBuilder builder = new StringBuilder();
                    //        string _tempSub = headers.Subject;

                    //        OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    //        if (plainText != null)
                    //        {
                    //            // We found some plaintext!
                    //            builder.Append(plainText.GetBodyAsText());
                    //        }
                    //        else
                    //        {
                    //            // Might include a part holding html instead
                    //            OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                    //            if (html != null)
                    //            {
                    //                // We found some html!
                    //                builder.Append(html.GetBodyAsText());
                    //            }
                    //        }
                    //        string tempBody = builder.ToString();
                    //        string _tempSendFromIPAddress = tempBody;

                    //        //if (tempBody.Contains("Message-ID"))
                    //        //    UIDL = tempBody.Split(new string[] { "Message-ID: <" }, StringSplitOptions.None).Last();
                    //        //else
                    //        //    UIDL = tempBody.Split(new string[] { "Message-Id: <" }, StringSplitOptions.None).Last();
                    //        //UIDL = UIDL.Split(new string[] { ">\r\n" }, StringSplitOptions.None).First();
                    //        //UIDL = UIDL.Split(new string[] { "@" }, StringSplitOptions.None).First();
                    //        //UIDL = Regex.Replace(UIDL, @"[.-]", "");
                    //        //UIDL = (UIDL.Length > 25) ? UIDL.Substring(UIDL.Length - 25, 25) : UIDL;
                    //        //UIDL = System.Guid.NewGuid().ToString();
                    //        //UIDL = Regex.Replace(UIDL, @"[.-]", "");
                    //        //UIDL = (UIDL.Length > 25) ? UIDL.Substring(0, 25) : UIDL;

                    //        for (int j = 5; j > 0; j--)
                    //        {
                    //            tempBody = tempBody.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                    //        }
                    //        _tempSendFromIPAddress = tempBody.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //        _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //        _SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                    //        EmailDetails(tempBody);

                    //        if (_SendTo == "ylwong111@yahoo.com.hk")
                    //        {
                    //        }

                    //        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        if (message.Headers.Date.Contains("-"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        }
                    //        else if (message.Headers.Date.Contains("+"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //        }

                    //        _Reason = builder.ToString().Split(new string[] { "\r\n\r\n\r\n--- " }, StringSplitOptions.None).First();
                    //        _Reason = _Reason.Split(new string[] { "\r\n\r\n" }, StringSplitOptions.None).Last();
                    //        _Reason = _Reason.Replace("\r\n", " ");

                    //        RejectedReasonTrim(_Reason);
                    //        GetRejectReasonCategory();
                    //        GetActionType();
                    //        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //        if (_SendTo != "" && _Subject != "")
                    //        {
                    //            var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //            dt.Rows.Add(newRow);

                    //        }

                    //        #endregion
                    //    }
                    //    else if((message.FindFirstHtmlVersion() != null)||(message.FindFirstPlainTextVersion() != null))
                    //    {
                    //        #region

                    //        StringBuilder builder = new StringBuilder();
                    //        string _tempSub = headers.Subject;

                    //        OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    //        if (plainText != null)
                    //        {
                    //            // We found some plaintext!
                    //            builder.Append(plainText.GetBodyAsText());
                    //        }
                    //        else
                    //        {
                    //            // Might include a part holding html instead
                    //            OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                    //            if (html != null)
                    //            {
                    //                // We found some html!
                    //                builder.Append(html.GetBodyAsText());
                    //            }
                    //        }
                    //        string tempBody = builder.ToString();
                    //        string _tempSendFromIPAddress = tempBody;

                    //        int ReceivedCount = Regex.Matches(tempBody, "Received").Count;
                    //        for (int j = ReceivedCount; j > 0; j--)
                    //        {
                    //            tempBody = tempBody.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                    //        }
                    //        _tempSendFromIPAddress = tempBody.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //        _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //        _SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                    //        EmailDetails(tempBody);

                    //        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        if (message.Headers.Date.Contains("-"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        }
                    //        else if (message.Headers.Date.Contains("+"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //        }

                    //        _Reason = builder.ToString().Split(new string[] { "\r\nDiagnostic " }, StringSplitOptions.None).First();
                    //        _Reason = _Reason.Split(new string[] { "--- Below this line is a copy of the message." }, StringSplitOptions.None).First();
                    //        _Reason = _Reason.Split(new string[] { "Unfortunately, " }, StringSplitOptions.None).Last();
                    //        _Reason = _Reason.Split(new string[] { " You " }, StringSplitOptions.None).First();
                           

                    //        RejectedReasonTrim(_Reason);
                    //        GetRejectReasonCategory();
                    //        GetActionType();
                    //        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //        if (_SendTo != "" && _Subject != "")
                    //        {
                    //            var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //            dt.Rows.Add(newRow);

                    //        }

                    //        #endregion
                    //    }
                    //}
                   
                    //else if (headers.Subject.Contains("Undeliverable: Daily"))
                    //{
                    //    if (message.MessagePart.MessageParts != null)
                    //    {
                    //        #region
                    //        int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);

                    //        OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                    //        OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                    //        if (messagePart.Body != null)
                    //            body = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //        if (messagePart1.Body != null)
                    //            body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);
                    //        if (ParMessagecount > 2)
                    //        {
                    //            OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                    //            if (messagePart2.Body != null)
                    //                body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);
                    //        }

                    //        string _TempSendDate = body2.Split(new string[] { "Date:" }, StringSplitOptions.None).Last();
                    //        _SendDate = Convert.ToDateTime(_TempSendDate.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        if (_SendDate == Convert.ToDateTime("23/11/2018").Date)
                    //        {

                    //        }

                    //        EmailDetails(body2);

                    //        string _tempIPAddress = body2;
                    //        int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                    //        for (int j = 5; j > 0; j--)
                    //        {
                    //            _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                    //        }
                    //        _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //        _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //        _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                    //        //Reject Date
                    //        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        if (message.Headers.Date.Contains("-"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        }
                    //        else if (message.Headers.Date.Contains("+"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //        }


                    //        //Actual Reject Reason
                    //        if (body1.Contains("Diagnostic-Code:"))
                    //        {
                    //            if (body1.Contains("Stage: CreateMessage"))
                    //            {
                    //                _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                    //                _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).First();
                    //                _Reason = _Reason.Split(new string[] { ";" }, StringSplitOptions.None).Last();

                    //            }
                    //            else
                    //            {
                    //                _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                    //                _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();

                    //            }

                    //            //_Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();

                    //        }

                    //        RejectedReasonTrim(_Reason);

                    //        //Reject Reason Category
                    //        GetRejectReasonCategory();

                    //        GetActionType();
                    //        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //        var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //        dt.Rows.Add(newRow);

                    //        #endregion
                    //    }

                    //}
                    //else if (headers.Subject.Contains("Undeliverable"))
                    //{
                    //    if (message.MessagePart.MessageParts != null)
                    //    {
                    //        #region
                    //        StringBuilder builder = new StringBuilder();
                    //        string _tempSub = headers.Subject;

                    //        OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    //        if (plainText != null)
                    //        {
                    //            //plaintext version
                    //            builder.Append(plainText.GetBodyAsText());
                    //        }
                    //        else
                    //        {
                    //            //html instead
                    //            OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                    //            if (html != null)
                    //            {
                    //                //html version
                    //                builder.Append(html.GetBodyAsText());
                    //            }
                    //        }
                    //        string tempBody = builder.ToString();
                    //        string _tempSendFromIPAddress = tempBody;

                    //        int ReceivedCount = Regex.Matches(tempBody, "Received").Count;
                    //        for (int j = ReceivedCount; j > 0; j--)
                    //        {
                    //            tempBody = tempBody.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                    //        }
                    //        _tempSendFromIPAddress = tempBody.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //        _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //        _SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                    //        EmailDetails(tempBody);

                    //        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        if (message.Headers.Date.Contains("-"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        }
                    //        else if (message.Headers.Date.Contains("+"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //        }

                    //        _Reason = builder.ToString().Split(new string[] { "\r\nDiagnostic " }, StringSplitOptions.None).First();
                    //        _Reason = _Reason.Split(new string[] { ">\r\n" }, StringSplitOptions.None).Last();
                    //        _Reason = _Reason.Replace("\r\n", " ");

                    //        RejectedReasonTrim(_Reason);
                    //        GetRejectReasonCategory();
                    //        GetActionType();
                    //        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //        if (_SendTo != "" && _Subject != "")
                    //        {
                    //            var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //            dt.Rows.Add(newRow);

                    //        }

                    //        #endregion
                    //    }

                    //}
                    //else if ((ChineseChar == false) && !headers.Subject.Contains("Automated Reply") && !headers.Subject.Contains("qq.com"))
                    //{
                    //    if (message.MessagePart.MessageParts != null)
                    //    {
                    //        #region
                    //        string _tempSub = headers.Subject;
                    //        //var msgPart = null;
                    //        int ParMessagecount = 0;


                    //        if (message.MessagePart.MessageParts != null)
                    //        {
                    //            ParMessagecount = message.MessagePart.MessageParts.Count;

                    //            if (ParMessagecount > 0)
                    //            {
                    //                OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                    //                if (message.MessagePart.Body != null)
                    //                {
                    //                    body = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //                }
                    //            }

                    //            if (ParMessagecount > 1)
                    //            {
                    //                OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];
                    //                if (message.MessagePart.Body != null)
                    //                {
                    //                    body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);
                    //                }
                    //            }

                    //            if (ParMessagecount > 2)
                    //            {
                    //                OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                    //                if (message.MessagePart.Body != null)
                    //                {
                    //                    body2 = messagePart1.BodyEncoding.GetString(messagePart2.Body);
                    //                }
                    //            }

                    //            if ((body2 == null || body2 == "") && (body1.Contains("[")))
                    //            {
                    //                body1 = body1.Split(new string[] { "[" }, StringSplitOptions.None).Last();
                    //                _SendFromIPAddress = body1.Split(new string[] { "])" }, StringSplitOptions.None).First();

                    //                EmailDetails(body1);

                    //            }
                    //            else if ((body2 != null || body2 != "") && (body2.Contains("[")))
                    //            {
                    //                string _tempSendFromIPAddress = body2;
                    //                int count = Regex.Matches(body2, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                    //                int ReceivedCount = Regex.Matches(_tempSendFromIPAddress, "Received").Count;
                    //                for (int j = ReceivedCount; j > 0; j--)
                    //                {
                    //                    _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                    //                }
                    //                _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //                _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //                _SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                    //                EmailDetails(body2);

                    //            }

                    //            //Reject Date Time
                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //            if (message.Headers.Date.Contains("-"))
                    //            {
                    //                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //                //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //            }
                    //            else if (message.Headers.Date.Contains("+"))
                    //            {
                    //                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //            }

                    //            if ((body1.Contains("Diagnostic-Code:") || body2.Contains("Diagnostic-Code:") )&& !body1.Contains("73"))
                    //            {
                    //                _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                    //                _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                    //                _Reason = _Reason.Replace("/r/n", "");
                    //                _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();

                    //            }
                    //            else if (body.ToLower().Contains("reason"))
                    //            {
                    //                _Reason = body.Split(new string[] { "reason: " }, StringSplitOptions.None).Last();
                    //                _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();

                    //                //_Reason = body.Split(new string[] { ":\r\n" }, StringSplitOptions.None).Last();

                    //            }
                    //            else if (body.ToLower().Contains("because"))
                    //            {
                    //                _Reason = body.Split(new string[] { "because:\r\n\r\n" }, StringSplitOptions.None).Last();
                    //                _Reason = _Reason.Replace("  ", "");

                    //            }
                    //            else
                    //            {
                    //                _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                    //                _Reason = _Reason.Replace("\r\n", ". ");

                    //            }

                    //            RejectedReasonTrim(_Reason);
                    //            GetRejectReasonCategory();

                    //            GetActionType();
                    //            GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //            if (_SendTo != "" && _Subject != "")
                    //            {
                    //                var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //                dt.Rows.Add(newRow);

                    //            }

                    //        }

                    //        #endregion
                    //    }

                    //}
                    //else if (headers.Subject.Contains("qq.com"))
                    //{
                    //    if (message.MessagePart.MessageParts != null)
                    //    {
                    //        #region

                    //        int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);

                    //        OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                    //        OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                    //        if (messagePart.Body != null)
                    //            body = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //        if (messagePart1.Body != null)
                    //            body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);
                    //        if (ParMessagecount > 2)
                    //        {
                    //            OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                    //            if (messagePart2.Body != null)
                    //                body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);

                    //        }

                    //        EmailDetails(body1);

                    //        string _tempIPAddress = body2;
                    //        int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                    //        for (int j = 5; j > 0; j--)
                    //        {
                    //            _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();

                    //        }
                    //        _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //        _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //        _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                    //        //Reject Date
                    //        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        if (message.Headers.Date.Contains("-"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //        }
                    //        else if (message.Headers.Date.Contains("+"))
                    //        {
                    //            string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //            _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //        }

                    //        //Actual Reject Reason
                    //        if (body1.Contains("Diagnostic-Code:"))
                    //        {
                    //            if (body1.Contains("Stage: CreateMessage"))
                    //            {
                    //                _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                    //                _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).First();
                    //                _Reason = _Reason.Split(new string[] { ";" }, StringSplitOptions.None).Last();

                    //            }
                    //            else
                    //            {
                    //                _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                    //                _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();

                    //            }

                    //            //_Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();

                    //        }
                    //        else
                    //        {
                    //            string tempoReason = body.Split(new string[] { "字节<br>" }, StringSplitOptions.None).Last();
                    //            tempoReason = tempoReason.Split(new string[] { ".<br><br>" }, StringSplitOptions.None).First();
                    //            _Reason = tempoReason.Replace("&lt;", " ");
                    //            _Reason = _Reason.Replace("&gt;,", "");
                    //            _Reason = _Reason.Replace("\r\n", "");

                    //        }

                    //        RejectedReasonTrim(_Reason);

                    //        //Reject Reason Category
                    //        GetRejectReasonCategory();

                    //        GetActionType();
                    //        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());

                    //        var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //        dt.Rows.Add(newRow);

                    //        #endregion
                    //    }

                    //}
                    //else if (headers.Subject.Contains("自动回复:提示: 登入交易平台提示"))
                    //{
                    //    if (message.MessagePart.MessageParts != null)
                    //    {
                    //        #region
                    //        _Subject = headers.Subject.ToString();
                    //        StringBuilder builder = new StringBuilder();
                    //        string _tempSub = headers.Subject;

                    //        OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    //        if (plainText != null)
                    //        {
                    //            // We found some plaintext!
                    //            builder.Append(plainText.GetBodyAsText());
                    //        }
                    //        else
                    //        {
                    //            // Might include a part holding html instead
                    //            OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                    //            if (html != null)
                    //            {
                    //                // We found some html!
                    //                builder.Append(html.GetBodyAsText());
                    //            }
                    //        }
                    //        string tempBody = builder.ToString();
                    //        string _tempSendFromIPAddress = tempBody;

                    //        #endregion
                    //    }
                    //}
                    
                    //else if (headers.Subject.Contains("Re: Mobile Deposit"))
                    //{
                    //    if (message.MessagePart.MessageParts != null)
                    //    {
                    //        #region
                    //        _Subject = headers.Subject.ToString();

                    //        //int PartMsg = message.MessagePart.MessageParts.Count;
                    //        if (message.MessagePart.MessageParts != null)
                    //        {
                    //            int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);

                    //            if (ParMessagecount != 0)
                    //            {
                    //                OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                    //                OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                    //                if (messagePart.Body != null)
                    //                    body = messagePart.BodyEncoding.GetString(messagePart.Body);
                    //                if (messagePart1.Body != null)
                    //                    body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                    //                if (ParMessagecount > 2)
                    //                {
                    //                    OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                    //                    body2 = messagePart1.BodyEncoding.GetString(messagePart2.Body);
                    //                }

                    //                if ((body2 == null || body2 == "") && (body1.Contains("[")))
                    //                {
                    //                    body1 = body1.Split(new string[] { "[" }, StringSplitOptions.None).Last();
                    //                    _SendFromIPAddress = body1.Split(new string[] { "])" }, StringSplitOptions.None).First();

                    //                    EmailDetails(body1);

                    //                }
                    //                else if ((body2 != null || body2 != "") && (body2.Contains("[")))
                    //                {
                    //                    string _tempSendFromIPAddress = body2;
                    //                    int count = Regex.Matches(body2, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                    //                    for (int j = 5; j > 0; j--)
                    //                    {
                    //                        body2 = body2.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                    //                    }
                    //                    _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //                    _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //                    _SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                    //                    EmailDetails(body2);

                    //                }
                    //                else
                    //                {

                    //                }

                    //                //Email ID
                    //                //UIDL = message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First();
                    //                ////UIDL = UIDL.Substring(13);
                    //                //UIDL = Regex.Replace(UIDL, @"[.-]", "");
                    //                //UIDL = (UIDL.Length > 25) ? UIDL.Substring(UIDL.Length - 25, 25) : UIDL;
                    //                //UIDL = System.Guid.NewGuid().ToString();
                    //                //UIDL = Regex.Replace(UIDL, @"[.-]", "");
                    //                //UIDL = (UIDL.Length > 25) ? UIDL.Substring(0, 25) : UIDL;

                    //                //Reject Date Time
                    //                //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //                if (message.Headers.Date.Contains("-"))
                    //                {
                    //                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //                    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //                }
                    //                else if (message.Headers.Date.Contains("+"))
                    //                {
                    //                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //                }

                    //                if (body1.Contains("Diagnostic-Code:") || body2.Contains("Diagnostic-Code:"))
                    //                {
                    //                    _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                    //                    _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                    //                    _Reason = _Reason.Replace("/r/n", "");
                    //                    _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();

                    //                }
                    //                else if (body.ToLower().Contains("reason"))
                    //                {
                    //                    _Reason = body.Split(new string[] { ":\r\n" }, StringSplitOptions.None).Last();

                    //                }
                    //                else if (body.ToLower().Contains("because"))
                    //                {
                    //                    _Reason = body.Split(new string[] { "because:\r\n\r\n" }, StringSplitOptions.None).Last();
                    //                    _Reason = _Reason.Replace("  ", "");

                    //                }
                    //                else
                    //                {
                    //                    _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                    //                    _Reason = _Reason.Replace("\r\n", ". ");

                    //                }

                    //                RejectedReasonTrim(_Reason);
                    //                GetRejectReasonCategory();

                    //                GetActionType();
                    //                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //                if (_SendTo != "" && _Subject != "")
                    //                {
                    //                    var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //                    dt.Rows.Add(newRow);

                    //                }
                    //            }
                    //        }
                    //        else
                    //        {


                    //            StringBuilder builder = new StringBuilder();
                    //            string _tempSub = headers.Subject;

                    //            OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    //            if (plainText != null)
                    //            {
                    //                // We found some plaintext!
                    //                builder.Append(plainText.GetBodyAsText());
                    //            }
                    //            else
                    //            {
                    //                // Might include a part holding html instead
                    //                OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                    //                if (html != null)
                    //                {
                    //                    // We found some html!
                    //                    builder.Append(html.GetBodyAsText());
                    //                }
                    //            }
                    //            string tempBody = builder.ToString();
                    //            string _tempSendFromIPAddress = tempBody;

                    //            //UIDL = tempBody.Split(new string[] { "Message-ID: <" }, StringSplitOptions.None).Last();
                    //            //UIDL = UIDL.Split(new string[] { ">\r\n" }, StringSplitOptions.None).First();
                    //            //UIDL = UIDL.Split(new string[] { "@" }, StringSplitOptions.None).First();
                    //            //UIDL = System.Guid.NewGuid().ToString();
                    //            //UIDL = Regex.Replace(UIDL, @"[.-]", "");
                    //            //UIDL = (UIDL.Length > 25) ? UIDL.Substring(0, 25) : UIDL;

                    //            for (int j = 5; j > 0; j--)
                    //            {
                    //                tempBody = tempBody.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                    //            }
                    //            _tempSendFromIPAddress = tempBody.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                    //            _tempSendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                    //            _SendFromIPAddress = _tempSendFromIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                    //            EmailDetails(tempBody);

                    //            //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //            if (message.Headers.Date.Contains("-"))
                    //            {
                    //                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                    //                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                    //                //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    //            }
                    //            else if (message.Headers.Date.Contains("+"))
                    //            {
                    //                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                    //                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                    //            }

                    //            _Reason = builder.ToString().Split(new string[] { "\r\n\r\n\r\n--- " }, StringSplitOptions.None).First();
                    //            _Reason = _Reason.Split(new string[] { "\r\n\r\n" }, StringSplitOptions.None).Last();
                    //            _Reason = _Reason.Replace("\r\n", " ");

                    //            RejectedReasonTrim(_Reason);
                    //            GetRejectReasonCategory();
                    //            GetActionType();
                    //            GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                    //            if (_SendTo != "" && _Subject != "")
                    //            {
                    //                var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };
                    //                dt.Rows.Add(newRow);

                    //            }

                    //        }
                    //        #endregion
                    //    }

                    //}
                    
                    #endregion

                }
			}
			catch (Exception ex)
			{
                CreateLogFile(ex.ToString());
			}
		}

        #region Extracting email content
        private void EmailDetails(string _emailContent, string _emailContent2)
		{
			try
			{
                if (_emailContent2 == null || _emailContent2 == "")
                {
                    _emailContent2 = _emailContent;

                }
                if ((!_emailContent2.ToLower().Contains("date")) || (!_emailContent2.ToLower().Contains("from")) || (!_emailContent2.ToLower().Contains("to")))
				{
					_SendDate = DateTime.Now;
					_SendFrom = "";
					_SendTo = "";
					_Subject = "";
				}
				else
				{
                    //_SendDate = DateTime.Now;
                    _SendFrom = "";
                    _SendTo = "";
                    _Subject = "";

                    if (_emailContent2.Contains("Thread-Topic"))
                    {
                        _Subject = _emailContent2.Split(new string[] { "Thread-Topic: " }, StringSplitOptions.None).Last();
                        _Subject = _Subject.Split(new string[] { "\r\nFrom:" }, StringSplitOptions.None).First();
                    }
                    else if (_emailContent2.Contains("Q?"))
                    {
                        _Subject = _emailContent2.Split(new string[] { "Subject: " }, StringSplitOptions.None).Last();
                        _Subject = _Subject.Split(new string[] { "\r\n" }, StringSplitOptions.None).First();
                        _Subject = _Subject.Split(new string[] { "Q?" }, StringSplitOptions.None).Last();
                        _Subject = _Subject.Split(new string[] { "?=" }, StringSplitOptions.None).First();

                    }
                    else
                    {
                        _Subject = headers.Subject;
                        _Subject = _emailContent2.Split(new string[] { "Subject: " }, StringSplitOptions.None).Last();
                        _Subject = _Subject.Split(new string[] { "\r\n" }, StringSplitOptions.None).First();
                    }

					//Send Date
                    if (_emailContent2 == _emailContent)
                    {
                        string _TempSendDate = _emailContent2.Split(new string[] { "Date:" }, StringSplitOptions.None).Last();
                        _SendDate = Convert.ToDateTime(_TempSendDate.Split(new string[] { " +" }, StringSplitOptions.None).First());

                    }
                    else
                    {
                        string _TempSendDate = _emailContent2.Split(new string[] { "Date:" }, StringSplitOptions.None).Last();
                        _SendDate = Convert.ToDateTime(_TempSendDate.Split(new string[] { " +" }, StringSplitOptions.None).First());
                    }
                    

					//Send From
					if (_emailContent2.Split(new string[] { "From: " }, StringSplitOptions.None).Last().Contains("=?big5?Q?"))
					{
						_SendFrom = _emailContent2.Split(new string[] { "From: " }, StringSplitOptions.None).Last();
						_SendFrom = _SendFrom.Split(new string[] { ">\r\n" }, StringSplitOptions.None).First();
						_SendFrom = _SendFrom.Split(new string[] { "<" }, StringSplitOptions.None).Last();

					}
					else
					{
						_SendFrom = _emailContent2.Split(new string[] { "From: " }, StringSplitOptions.None).Last();
						_SendFrom = _SendFrom.Split(new string[] { "\r\n" }, StringSplitOptions.None).First();
						_SendFrom = Regex.Replace(_SendFrom, @"[<>]", "");
					}


					//Send To
                    if (headers.Subject.Contains("系统退信/Systems bounce"))
                    {
                        _SendTo = _emailContent.Split(new string[] { "rfc822;" }, StringSplitOptions.None).Last();
                        _SendTo = _SendTo.Split(new string[] { "\r\n" }, StringSplitOptions.None).First();
                        _SendTo = Regex.Replace(_SendTo, @"[<>]", "");

                    }
                    else
                    {
                        _SendTo = _emailContent2.Split(new string[] { "To: " }, StringSplitOptions.None).Last();
                        _SendTo = _SendTo.Split(new string[] { "\r\n" }, StringSplitOptions.None).First();
                        _SendTo = Regex.Replace(_SendTo, @"[<> ]", "");

                    }
                    

					//Subject
					if (_emailContent2.Contains("Thread-Topic"))
					{
						_Subject = _emailContent2.Split(new string[] { "Thread-Topic: " }, StringSplitOptions.None).Last();
						_Subject = _Subject.Split(new string[] { "\r\nFrom:" }, StringSplitOptions.None).First();
					}
					else if (_emailContent2.Contains("Q?"))
					{
						_Subject = _emailContent2.Split(new string[] { "Subject: " }, StringSplitOptions.None).Last();
						_Subject = _Subject.Split(new string[] { "\r\n" }, StringSplitOptions.None).First();
						_Subject = _Subject.Split(new string[] { "Q?" }, StringSplitOptions.None).Last();
						_Subject = _Subject.Split(new string[] { "?=" }, StringSplitOptions.None).First();

					}
					else
					{
						_Subject = _emailContent2.Split(new string[] { "Subject: " }, StringSplitOptions.None).Last();
						_Subject = _Subject.Split(new string[] { "\r\n" }, StringSplitOptions.None).First();
					}

					_Subject = EncodedWord.Decode(_Subject); // Encode using OPENPOP.net MIME Class

				}
			}
			catch (Exception ex)
			{
                CreateLogFile(ex.ToString());
                //MessageBoxButtons MsgBox = MessageBoxButtons.OK;
                //MessageBox.Show("Error: Extracting Email Details \r\n" + ex.ToString(), "Error", MsgBox);
			}
		   

		}

        private void method_Reason(string body, string body1)
        {
            if (body1.ToLower().Contains("diagnostic-code:") && !body1.ToLower().Contains("x-unix"))
            {
                _Reason = body1.Split(new string[] { "Diagnostic-Code:" }, StringSplitOptions.None).Last();
                _Reason = _Reason.Split(new string[] { "; " }, StringSplitOptions.None).Last();
                _Reason = _Reason.Replace("\r\n", "");
                _Reason = _Reason.Split(new string[] { "(" }, StringSplitOptions.None).First();
                _Reason = _Reason.Replace("\r\n", "");


            }
            else if (body1.ToLower().Contains("diagnostic-code:") && body1.ToLower().Contains("x-unix"))
            {
                _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                _Reason = _Reason.Split(new string[] { "\r\n550" }, StringSplitOptions.None).Last();
                _Reason = _Reason.Split(new string[] { "\r\nuntie" }, StringSplitOptions.None).First();

            }
            else if (body.Contains("The mail system"))
            {
                _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();
                if (_Reason.Contains(" Please"))
                {
                    _Reason = _Reason.Split(new string[] { " Please" }, StringSplitOptions.None).First();
                }
                else
                {
                    _Reason = _Reason.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();
                }
            }
            else if ((body.Contains("procmail:")) || (body.Contains("permanent fatal errors")) || body.ToLower().Contains("warning"))
            {
                _Reason = body.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
            }
            else if (headers.Subject.Contains("系统退信/Systems bounce"))
            {

            }
            else
            {
                _Reason = body.Split(new string[] { ">: " }, StringSplitOptions.None).Last();

            }

            RejectedReasonTrim(_Reason);

            //Reject Reason Category
            GetRejectReasonCategory();

            GetActionType();

        }

        //## Added 20190109
        private void loop_emailXtract(OpenPop.Mime.Message message)
        {
            try
            {
                UIDL = "";
                _SendTo = "";
                _SendFrom = "";
                _Subject = "";
                _Reason = "";
                _ReasonCategory = "";
                _ActionRequired = "";
                _SendFromIPAddress = "";
                _SendDate = DateTime.Now;
                _RejectDateTime = DateTime.Now;
                _EmailType = "";
                body = "";
                body1 = "";
                body2 = "";

                //Start New Added Code 20190108
                var newRow = new string[] {};
                int emailsenderInitiate = 0;
                string[] EmailSenderArray = ConfigurationManager.AppSettings["EmailSender"].Split(',').Select(s => s.Trim()).ToArray();
                List<string> listEmailSender = new List<string>(EmailSenderArray);
                int emailsenderCounter = listEmailSender.Count;
                //string _temp = listEmailSender[emailsenderCounter].ToString();

                int emailSubjectInitiate = 0;
                string[] EmailSubjectArray = ConfigurationManager.AppSettings["EmailSubject"].Split(',').Select(s => s.Trim()).ToArray();
                List<string> listEmailSubject = new List<string>(EmailSubjectArray);
                int emailSubjectCounter = listEmailSubject.Count;

                #region Trial

                //string subject_temp = headers.Subject.ToString();
                //if (!headers.Subject.Contains("Reminder") && !headers.Subject.Contains("AUTO:") && !headers.Subject.Contains("Out of office") && !headers.Subject.Contains("Monthly Statement") && !headers.Subject.Contains("renamed") && !headers.Subject.Contains("Verify your mailbox to avoid account suspension!") && !headers.Subject.Contains("AutoReply:") && !headers.Subject.Contains("自动回复") && !headers.Subject.Contains("Automated Reply")) 
                //{
                //    int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                //    if (ParMessagecount != 0)
                //    {
                //        OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                //        OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                //        if (messagePart.Body != null)
                //            body = messagePart.BodyEncoding.GetString(messagePart.Body);
                //        if (messagePart1.Body != null)
                //            body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                //        if (ParMessagecount > 2)
                //        {
                //            OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                //            body2 = messagePart1.BodyEncoding.GetString(messagePart2.Body);
                //        }

                //        if (body1.Contains("Action: failed") || headers.Subject.Contains("系统退信"))
                //        {
                //            EmailDetails(body1, body2);

                //            method_Reason(body, body1);

                //            #region Get Recipient Address
                //            if (headers.Subject.Contains("Delivery Status Notification (Failure)") || (headers.Subject.Contains("未傳遞的")) || headers.Subject.Contains("系统退信") || headers.Subject.Contains("Undelivered Mail Returned to Sender")) // || (body1.Contains("rfc822")))
                //            {
                //                _SendTo = body1.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                //                _SendTo = _SendTo.Split(new string[] { "rfc822;" }, StringSplitOptions.None).Last();
                //            }
                //            else if (headers.Subject.Contains("未傳遞的主旨"))
                //            {
                //                _SendTo = body1.Split(new string[] { "Final-Recipient:" }, StringSplitOptions.None).Last();
                //                _SendTo = _SendTo.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                //                _SendTo = _SendTo.Split(new string[] { ";" }, StringSplitOptions.None).Last();

                //            }
                //            else
                //            {
                //                _SendTo = body.Split(new string[] { "\r\n    (reason:" }, StringSplitOptions.None).First();
                //                _SendTo = _SendTo.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                //                _SendTo = Regex.Replace(_SendTo, @"[<>]", "");
                //            }

                //            if (_SendTo.Contains("This is the mail system at host mx.phillip.com.hk."))
                //            {

                //            }
                //            #endregion Recipient

                //            #region Get IP Address of email sender

                //            string _tempIPAddress = body2;
                //            int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                //            int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                //            for (int j = ReceivedCount; j > 0; j--)
                //            {
                //                _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                //            }
                //            _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                //            _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                //            _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                //            #endregion end IP Address of email sender

                //            #region Get Reject Date
                //            if (message.Headers.Date.Contains("-"))
                //            {
                //                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                //                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                //                //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                //            }
                //            else if (message.Headers.Date.Contains("+"))
                //            {
                //                string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                //                _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                //            }
                //            #endregion End Reject Date

                //        }
                //        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                //        var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                //        dt.Rows.Add(newRow);

                  

                //    }

                //}

                #endregion

                #region Get Rejected Email TEMP REMARK
                //if (EmailSenderArray.Any(headers.From.Raw.Contains))
                //{
                //        if (EmailSubjectArray.Any(headers.Subject.Contains))
                //        {
                //            int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                //            if (ParMessagecount != 0)
                //            {
                //                OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                //                OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                //                if (messagePart.Body != null)
                //                    body = messagePart.BodyEncoding.GetString(messagePart.Body);
                //                if (messagePart1.Body != null)
                //                    body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                //                if (ParMessagecount > 2)
                //                {
                //                    OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                //                    body2 = messagePart1.BodyEncoding.GetString(messagePart2.Body);
                //                }

                //                if (body.Contains("Action: failed") || headers.Subject.Contains("系统退信"))
                //                {

                //                }
                //                else if (body1.Contains("Action: failed") || headers.Subject.Contains("系统退信"))
                //                {

                //                }
                //                else if (body2.Contains("Action: failed") || headers.Subject.Contains("系统退信"))
                //                {

                //                }

                //                EmailDetails(body1, body2);

                //                method_Reason(body, body1);

                //                #region Get Recipient Address
                //                if (headers.Subject.Contains("Delivery Status Notification (Failure)") || (headers.Subject.Contains("未傳遞的")) || headers.Subject.Contains("系统退信") || headers.Subject.Contains("Undelivered Mail Returned to Sender")) // || (body1.Contains("rfc822")))
                //                {
                //                    _SendTo = body1.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                //                    _SendTo = _SendTo.Split(new string[] { "rfc822;" }, StringSplitOptions.None).Last();
                //                }
                //                else if (headers.Subject.Contains("未傳遞的主旨"))
                //                {
                //                    _SendTo = body1.Split(new string[] { "Final-Recipient:" }, StringSplitOptions.None).Last();
                //                    _SendTo = _SendTo.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                //                    _SendTo = _SendTo.Split(new string[] { ";" }, StringSplitOptions.None).Last();

                //                }
                //                else
                //                {
                //                    _SendTo = body.Split(new string[] { "\r\n    (reason:" }, StringSplitOptions.None).First();
                //                    _SendTo = _SendTo.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                //                    _SendTo = Regex.Replace(_SendTo, @"[<>]", "");
                //                }

                //                if (_SendTo.Contains("This is the mail system at host mx.phillip.com.hk."))
                //                {

                //                }
                //                #endregion Recipient

                //                #region Get IP Address of email sender

                //                string _tempIPAddress = body2;
                //                int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                //                int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                //                for (int j = ReceivedCount; j > 0; j--)
                //                {
                //                    _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                //                }
                //                _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                //                _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                //                _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                //                #endregion end IP Address of email sender

                //                #region Get Reject Date
                //                if (message.Headers.Date.Contains("-"))
                //                {
                //                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                //                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                //                    //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                //                }
                //                else if (message.Headers.Date.Contains("+"))
                //                {
                //                    string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                //                    _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                //                }
                //                #endregion End Reject Date

                //            }

                //            RejectedReasonTrim(_Reason);

                //            //Reject Reason Category
                //            GetRejectReasonCategory();

                //            GetActionType();

                //            GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                //            var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                //            dt.Rows.Add(newRow);

                //        }
                //}
                

                #endregion

                #region Loop to Get Rejected Email Code
                for (emailsenderInitiate = 0; emailsenderInitiate < emailsenderCounter; emailsenderInitiate++)
                {
                    //if (headers.From.Raw.Contains(listEmailSender[emailsenderInitiate]))
                    if ((EmailSenderArray.Any(headers.From.Raw.Contains)) || (EmailSubjectArray.Any(headers.Subject.Contains)))
                    {
                        for (emailSubjectInitiate = 0; emailSubjectInitiate < emailSubjectCounter; emailSubjectInitiate++)
                        {
                            //if (headers.Subject.Contains(listEmailSubject[emailSubjectInitiate]))
                            if (EmailSubjectArray.Any(headers.Subject.Contains))
                            {
                                int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                                if (ParMessagecount != 0)
                                {

                                    OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                                    

                                    if (messagePart.Body != null)
                                        body = messagePart.BodyEncoding.GetString(messagePart.Body);

                                    if (ParMessagecount > 1)
                                    {
                                        OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];
                                        if (messagePart1.Body != null)
                                            body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);
                                    }

                                    if (ParMessagecount > 2)
                                    {
                                        OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                                        body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);
                                    }

                                    EmailDetails(body1, body2);

                                    method_Reason(body, body1);

                                    if (headers.Subject.Contains("Undelivered Mail Returned to Sender"))
                                    {
                                        string daateeee_temp = Convert.ToString(headers.DateSent);

                                    }

                                    #region Get Recipient Address
                                    if (headers.From.Address.Contains("postmaster@"))
                                    {
                                        _SendTo = body2.Split(new string[] { ">\r\nDate: " }, StringSplitOptions.None).First();
                                        _SendTo = _SendTo.Split(new string[] { "To: <" }, StringSplitOptions.None).Last();

                                    }
                                    else if (headers.Subject.Contains("Delivery Status Notification (Failure)") || (headers.Subject.Contains("未傳遞的")) || headers.Subject.Contains("系统退信") || headers.Subject.Contains("Undelivered Mail Returned to Sender")) // || (body1.Contains("rfc822")))
                                    {
                                        _SendTo = body1.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                                        _SendTo = _SendTo.Split(new string[] { "rfc822;" }, StringSplitOptions.None).Last();
                                    }
                                    else if (headers.Subject.Contains("未傳遞的主旨"))
                                    {
                                        _SendTo = body1.Split(new string[] { "Final-Recipient:" }, StringSplitOptions.None).Last();
                                        _SendTo = _SendTo.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                                        _SendTo = _SendTo.Split(new string[] { ";" }, StringSplitOptions.None).Last();

                                    }
                                    else if (headers.Subject.Contains("Mail delivery failed: returning message to sender"))
                                    {
                                        _SendTo = body1.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).Last();
                                        _SendTo = _SendTo.Split(new string[] { "Status:" }, StringSplitOptions.None).First();
                                        _SendTo = _SendTo.Split(new string[] { "rfc822;" }, StringSplitOptions.None).Last();
                                    }
                                    else
                                    {
                                        _SendTo = body.Split(new string[] { "\r\n    (reason:" }, StringSplitOptions.None).First();
                                        _SendTo = _SendTo.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                        _SendTo = Regex.Replace(_SendTo, @"[<>]", "");
                                    }

                                    #endregion Recipient

                                    #region Get IP Address of email sender

                                    string _tempIPAddress = body2;
                                    int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                                    int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "Message-ID:" }, StringSplitOptions.None).First();
                                    for (int j = ReceivedCount; j > 0; j--)
                                    {
                                        _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                                    }
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                    _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                                    #endregion end IP Address of email sender

                                    #region Get Reject Date
                                    if (message.Headers.Date.Contains("-"))
                                    {
                                        string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                        _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                    }
                                    else if (message.Headers.Date.Contains("+"))
                                    {
                                        string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                        _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                    }
                                    #endregion End Reject Date

                                    _EmailType = "Rejected Email";

                                }

                                RejectedReasonTrim(_Reason);

                                //Reject Reason Category
                                GetRejectReasonCategory();

                                GetActionType();

                                if (_Subject == "登入交易平台提示 Reminder: Trading Platform Login Notification (X7533XX by POEMS Web)")
                                {
                                }

                                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                                //dt.Rows.Add(newRow);
                                break;


                            }
                            else
                            {
                                int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                                if (ParMessagecount != 0)
                                {

                                    OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                                    OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                                    if (messagePart.Body != null)
                                        body = messagePart.BodyEncoding.GetString(messagePart.Body);
                                    if (messagePart1.Body != null)
                                        body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                                    if (ParMessagecount > 2)
                                    {
                                        OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                                        body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);
                                    }

                                    EmailDetails(body1, body2);

                                    method_Reason(body, body1);

                                    if (headers.Subject.Contains("Undelivered Mail Returned to Sender"))
                                    {
                                        string daateeee_temp = Convert.ToString(headers.DateSent);

                                    }

                                    #region Get Recipient Address
                                    if (headers.From.Address.Contains("postmaster@"))
                                    {
                                        _SendTo = body2.Split(new string[] { ">\r\nDate: " }, StringSplitOptions.None).First();
                                        _SendTo = _SendTo.Split(new string[] { "To: <" }, StringSplitOptions.None).Last();

                                    }
                                    else if (headers.Subject.Contains("Delivery Status Notification (Failure)") || (headers.Subject.Contains("未傳遞的")) || headers.Subject.Contains("系统退信") || headers.Subject.Contains("Undelivered Mail Returned to Sender")) // || (body1.Contains("rfc822")))
                                    {
                                        _SendTo = body1.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                                        _SendTo = _SendTo.Split(new string[] { "rfc822;" }, StringSplitOptions.None).Last();
                                    }
                                    else if (headers.Subject.Contains("未傳遞的主旨"))
                                    {
                                        _SendTo = body1.Split(new string[] { "Final-Recipient:" }, StringSplitOptions.None).Last();
                                        _SendTo = _SendTo.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                                        _SendTo = _SendTo.Split(new string[] { ";" }, StringSplitOptions.None).Last();

                                    }
                                    else if (headers.Subject.Contains("Mail delivery failed: returning message to sender"))
                                    {
                                        _SendTo = body1.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).Last();
                                        _SendTo = _SendTo.Split(new string[] { "Status:" }, StringSplitOptions.None).First();
                                        _SendTo = _SendTo.Split(new string[] { "rfc822;" }, StringSplitOptions.None).Last();
                                    }
                                    else
                                    {
                                        _SendTo = body.Split(new string[] { "\r\n    (reason:" }, StringSplitOptions.None).First();
                                        _SendTo = _SendTo.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                                        _SendTo = Regex.Replace(_SendTo, @"[<>]", "");
                                    }

                                    #endregion Recipient

                                    #region Get IP Address of email sender

                                    string _tempIPAddress = body2;
                                    int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                                    int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "Message-ID:" }, StringSplitOptions.None).First();
                                    for (int j = ReceivedCount; j > 0; j--)
                                    {
                                        _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                                    }
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                                    _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                                    _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();
                                    _SendFromIPAddress = _SendFromIPAddress.Split(new string[] { "(" }, StringSplitOptions.None).Last();
                                    _SendFromIPAddress = _SendFromIPAddress.Split(new string[] { ")\r\n" }, StringSplitOptions.None).First();

                                    #endregion end IP Address of email sender

                                    #region Get Reject Date
                                    if (message.Headers.Date.Contains("-"))
                                    {
                                        string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                                        _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                                        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                                    }
                                    else if (message.Headers.Date.Contains("+"))
                                    {
                                        string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                                        _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                                    }
                                    #endregion End Reject Date

                                    _EmailType = "Rejected Email";

                                }

                            }
                        }
                    }
                }
                dt.Rows.Add(newRow);
                #endregion

                #region Get All Email for Record

                _Subject = headers.Subject;
                if (!EmailSubjectArray.Any(headers.Subject.Contains))
                {
                    if (message.MessagePart.MessageParts == null || headers.Subject.Contains("不在辦公室"))
                    {
                        _SendTo = headers.To[0].Address;
                        _RejectDateTime = headers.DateSent;
                        _SendDate = headers.DateSent;
                        _SendFrom = headers.From.Address;
                        _Subject = headers.Subject;
                        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());

                    }
                    else
                    {
                        for (emailSubjectInitiate = 0; emailSubjectInitiate < emailSubjectCounter; emailSubjectInitiate++)
                        {
                            if (!headers.Subject.Contains(listEmailSubject[emailSubjectInitiate]))
                                if (!EmailSubjectArray.Any(headers.Subject.Contains))
                                {
                                    if (message.MessagePart.MessageParts != null)
                                    {

                                        int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                                        if (ParMessagecount != 0)
                                        {
                                            OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];

                                            if (messagePart.Body != null)
                                                body = messagePart.BodyEncoding.GetString(messagePart.Body);

                                            if (ParMessagecount > 1)
                                            {
                                                OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];
                                                if (messagePart1.Body != null)
                                                    body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                                            }


                                            if (ParMessagecount > 2)
                                            {
                                                OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                                                body2 = messagePart2.BodyEncoding.GetString(messagePart2.Body);
                                            }

                                            StringBuilder builder = new StringBuilder();
                                            OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                                            if (html != null)
                                            {
                                                // We found some plaintext!
                                                builder.Append(html.GetBodyAsText());
                                            }
                                            else
                                            {
                                                // Might include a part holding html instead
                                                OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                                                if (plainText != null)
                                                {
                                                    // We found some html!
                                                    builder.Append(plainText.GetBodyAsText());
                                                }
                                            }

                                            string tempo = builder.ToString();

                                            EmailDetails(body, body2);

                                            if (_SendTo == null || _SendTo == "" && _SendFrom == null || _SendFrom == "")
                                            {
                                                _SendTo = headers.To[0].Address.ToString();
                                                _Subject = headers.Subject.ToString();
                                                _SendFrom = headers.From.Address.ToString();
                                                _RejectDateTime = headers.DateSent;
                                                _SendDate = headers.DateSent;
                                                _Reason = "";

                                            }

                                            //method_Reason(body, body1);

                                            #region Get Recipient Address
                                            _SendTo = headers.To[0].Address.ToString();
                                            _Subject = headers.Subject.ToString();
                                            _SendFrom = headers.From.Address.ToString();
                                            _SendDate = headers.DateSent;
                                            _Reason = "";

                                            #endregion Recipient

                                        }

                                        GetRejectReasonCategory();

                                        GetActionType();

                                        GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                                        //newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                                        break;
                                    }
                                }
                        }
                    }
                }

                #endregion

                #region while Remark
                //while (emailsenderInitiate < emailsenderCounter)
                //{
                //    string _tempEmailSubject = headers.To.ToString();

                //    if (headers.From.Raw.Contains(listEmailSubject[emailsenderInitiate]))
                //    {
                //        while (emailSubjectInitiate < emailSubjectCounter)
                //        {
                //            if (headers.Subject.Contains(listEmailSubject[emailSubjectInitiate]))
                //            {
                //                int ParMessagecount = Convert.ToInt32(message.MessagePart.MessageParts.Count);
                //                if (ParMessagecount != 0)
                //                {
                //                    GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());

                //                    OpenPop.Mime.MessagePart messagePart = message.MessagePart.MessageParts[0];
                //                    OpenPop.Mime.MessagePart messagePart1 = message.MessagePart.MessageParts[1];

                //                    if (messagePart.Body != null)
                //                        body = messagePart.BodyEncoding.GetString(messagePart.Body);
                //                    if (messagePart1.Body != null)
                //                        body1 = messagePart1.BodyEncoding.GetString(messagePart1.Body);

                //                    if (ParMessagecount > 2)
                //                    {
                //                        OpenPop.Mime.MessagePart messagePart2 = message.MessagePart.MessageParts[2];
                //                        body2 = messagePart1.BodyEncoding.GetString(messagePart2.Body);
                //                    }


                //if (body.Contains("黑名單") || body1.Contains("黑名單") || body2.Contains("黑名單"))
                //{

                //}
                //else if (body.Contains("拒收") || body1.Contains("拒收") || body2.Contains("拒收"))
                //{

                //}
                //else if (body.Contains("退回") || body1.Contains("退回") || body2.Contains("退回"))
                //{

                //}
                //else if (body.Contains("退信") || body1.Contains("退信") || body2.Contains("退信"))
                //{

                //}

                //                    EmailDetails(body2);

                //                    method_Reason(body, body1);

                //                    #region Get Recipient Address
                //                    if (headers.Subject.Contains("Delivery Status Notification (Failure)") || (headers.Subject.Contains("未傳遞的")) || (body1.Contains("rfc822")))
                //                    {
                //                        _SendTo = body1.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                //                        _SendTo = _SendTo.Split(new string[] { "rfc822;" }, StringSplitOptions.None).Last();
                //                    }
                //                    else if (headers.Subject.Contains("未傳遞的主旨"))
                //                    {
                //                        _SendTo = body1.Split(new string[] { "Final-Recipient:" }, StringSplitOptions.None).Last();
                //                        _SendTo = _SendTo.Split(new string[] { "\r\nAction:" }, StringSplitOptions.None).First();
                //                        _SendTo = _SendTo.Split(new string[] { ";" }, StringSplitOptions.None).Last();

                //                    }
                //                    else
                //                    {
                //                        _SendTo = body.Split(new string[] { "\r\n    (reason:" }, StringSplitOptions.None).First();
                //                        _SendTo = _SendTo.Split(new string[] { "-----\r\n" }, StringSplitOptions.None).Last();
                //                        _SendTo = Regex.Replace(_SendTo, @"[<>]", "");
                //                    }
                //                    #endregion Recipient

                //                    #region Get IP Address of email sender

                //                    string _tempIPAddress = body2;
                //                    int count = Regex.Matches(_tempIPAddress, "\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}").Count;
                //                    int ReceivedCount = Regex.Matches(_tempIPAddress, "Received").Count;
                //                    for (int j = ReceivedCount; j > 0; j--)
                //                    {
                //                        _tempIPAddress = _tempIPAddress.Split(new string[] { "\r\nReceived" }, StringSplitOptions.None).Last();
                //                    }
                //                    _tempIPAddress = _tempIPAddress.Split(new string[] { "+0800" }, StringSplitOptions.None).First();
                //                    _tempIPAddress = _tempIPAddress.Split(new string[] { "([" }, StringSplitOptions.None).Last();
                //                    _SendFromIPAddress = _tempIPAddress.Split(new string[] { "])" }, StringSplitOptions.None).First();

                //                    #endregion end IP Address of email sender

                //                    #region Get Reject Date
                //                    if (message.Headers.Date.Contains("-"))
                //                    {
                //                        string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " -" }, StringSplitOptions.None).First();
                //                        _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));
                //                        //_RejectDateTime = Convert.ToDateTime(message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First());
                //                    }
                //                    else if (message.Headers.Date.Contains("+"))
                //                    {
                //                        string _tempoRejectDateTime = message.Headers.Date.Split(new string[] { " +" }, StringSplitOptions.None).First();
                //                        _RejectDateTime = Convert.ToDateTime(string.Format(String.Format("{0:d/M/yyyy HH:mm:ss}", _tempoRejectDateTime)));

                //                    }
                //                    #endregion End Reject Date


                //                }

                //                RejectedReasonTrim(_Reason);

                //                //Reject Reason Category
                //                GetRejectReasonCategory();

                //                GetActionType();

                //                GenerateUIDL(message.Headers.MessageId.Split(new string[] { "@" }, StringSplitOptions.None).First());
                //                var newRow = new string[] { _SendDate.ToString(), _SendTo, _Subject, _RejectDateTime.ToString(), _ReasonCategory, _ActionRequired, UIDL, _Reason, _SendFrom, _SendFromIPAddress };

                //                dt.Rows.Add(newRow);



                //            }
                //            //emailSubjectInitiate++;
                //        }
                //    }
                //    //if(emailsenderCounter.count = emailsenderInitiate[message.MessagePart.MessageParts.Count))
                //    ///emailsenderInitiate++;
                //}
                #endregion

                //End Recent Added
            }
            catch (Exception ex)
            {
                CreateLogFile(ex.ToString());
            }

        }

        #endregion

        private string GetRejectReasonCategory()
		{
			_Reason = Regex.Replace(_Reason, @"['\r\n]", "").ToLower().Trim();
		  
			DataRow[] hasRow = dsEmail.Tables[ConfigurationManager.AppSettings["Ref_ReasonCategory"]].Select(); //.Select("reason LIKE '%" + _Reason.ToLower().ToString() + "%'");
			foreach (DataRow Returned in hasRow)
				{
					if (_Reason.Contains(Returned[1].ToString()))
					{
						_ReasonCategory = Returned[2].ToString();
						break;
					}
				}
			return _ReasonCategory;
		}

		private string GetActionType()
		{
			
			DataRow[] hasRow = dsEmail.Tables[ConfigurationManager.AppSettings["Ref_Action"]].Select(); //.Select("reason LIKE '%" + _Reason.ToLower().ToString() + "%'");
			foreach (DataRow Returned in hasRow)
			{
				if (_ReasonCategory.Contains(Returned[1].ToString()))
				{
					_ActionRequired = Returned[2].ToString();
					break;
				}

			}

			return _ActionRequired; 
		}

		private void SaveToDB()
		{
			try
			{
                bool recordExist = false;
                string[] EmailSubjectArray = ConfigurationManager.AppSettings["EmailSubject"].Split(',').Select(s => s.Trim()).ToArray();

				SqlConnection con = new SqlConnection(dbcon);
				SqlCommand cmd;
				con.Open();

                using (SqlConnection conn = new SqlConnection(dbcon))
                {
                    #region Save Rejected Email
                    
                    if (EmailSubjectArray.Any(headers.Subject.Contains))
                    {
                        
                        using (SqlCommand cmdd = new SqlCommand("SELECT * from " + ConfigurationManager.AppSettings["Email_Detail"] + " where email_ID = '" + UIDL.ToString() + "'"))  // AND sendFrom = '" + _SendFrom.ToString() + "'"))
                        {
                            DataRow[] DataRow = dsEmail.Tables[0].Select("email_ID ='" + UIDL + "'");
                            if (DataRow.Length >= 1)
                            {
                                do
                                {
                                    UIDL = System.Guid.NewGuid().ToString();
                                    UIDL = UIDL.Replace("-", "");
                                    UIDL = UIDL.Substring(0, 25);

                                    DataRow = dsEmail.Tables[0].Select("email_ID ='" + UIDL + "'");
                                } while (DataRow.Length >= 1);

                            }
                            cmdd.Connection = conn;
                            conn.Open();
                            using (SqlDataReader sdr = cmdd.ExecuteReader())
                            {
                                //Check Record Exist
                                //int i = Convert.ToInt16(sdr.Read());
                                //if (i == 0)
                                //{
                                cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLInsert"], con);

                                cmd.Parameters.Add("@sendFrom", SqlDbType.NVarChar, 50).Value = _SendFrom ?? (object)DBNull.Value;
                                cmd.Parameters.Add("@sendDateTime", SqlDbType.DateTime).Value = _RejectDateTime; //?? (object)DBNull.Value;
                                cmd.Parameters.Add("@sendFromIPAddress", SqlDbType.NVarChar, 20).Value = _SendFromIPAddress ?? (object)DBNull.Value;
                                cmd.Parameters.Add("@sendTo", SqlDbType.NVarChar, 50).Value = _SendTo ?? (object)DBNull.Value;
                                cmd.Parameters.Add("@subject", SqlDbType.NVarChar, 255).Value = _Subject ?? (object)DBNull.Value;
                                cmd.Parameters.Add("@rejectDateTime", SqlDbType.DateTime).Value = _SendDate;
                                cmd.Parameters.Add("@rejectReason", SqlDbType.NVarChar, 255).Value = _Reason ?? (object)DBNull.Value;
                                cmd.Parameters.Add("@rejectReasonCategory", SqlDbType.NVarChar, 50).Value = _ReasonCategory ?? (object)DBNull.Value;
                                cmd.Parameters.Add("@actionRequired", SqlDbType.NVarChar, 255).Value = _ActionRequired ?? (object)DBNull.Value;
                                cmd.Parameters.Add("@actionTaken", SqlDbType.NVarChar, 255).Value = "" ?? (object)DBNull.Value;
                                cmd.Parameters.Add("@actionTakenDescription", SqlDbType.NVarChar, 255).Value = "" ?? (object)DBNull.Value;
                                cmd.Parameters.Add("@actionTakenBy", SqlDbType.NVarChar, 100).Value = "" ?? (object)DBNull.Value;
                                cmd.Parameters.Add("@actionTakenDateTime", SqlDbType.DateTime).Value = DateTime.Now;
                                cmd.Parameters.Add("@closeCase", SqlDbType.Bit).Value = false;
                                cmd.Parameters.Add("@email_ID", SqlDbType.NVarChar, 25).Value = UIDL.ToString() ?? (object)DBNull.Value;

                                cmd.ExecuteNonQuery();
                                //}
                            }
                            conn.Close();
                        }
                    }
                    #endregion

                    #region Save All Email
                        using (SqlCommand cmd_allEmail = new SqlCommand("SELECT * from " + ConfigurationManager.AppSettings["All_Email"] + " where email_ID = '" + UIDL.ToString() + "'"))  // AND sendFrom = '" + _SendFrom.ToString() + "'"))
                        {
                            DataRow[] DataRow = dsEmail.Tables[5].Select("email_ID ='" + UIDL + "'");
                            if (DataRow.Length >= 1)
                            {
                                do
                                {
                                    UIDL = System.Guid.NewGuid().ToString();
                                    UIDL = UIDL.Replace("-", "");
                                    UIDL = UIDL.Substring(0, 25);

                                    DataRow = dsEmail.Tables[5].Select("email_ID ='" + UIDL + "'");
                                } while (DataRow.Length >= 1);

                            }

                            if (_EmailType == "" || _EmailType == null)
                            {
                                _EmailType = "Non-Rejected Email";
                            }


                            DateTime _tempoDate = _RejectDateTime.Date.Date;
                            DateTime _tempoDate2 = _SendDate.Date;

                            if (_RejectDateTime == DateTime.Now || _SendDate == DateTime.Now)
                            {

                            }
                            if (UIDL == "")
                            {

                            }

                            if (_Subject != "" || _Subject != null && _SendFrom != "" || _SendFrom != null && _SendTo != "" || _SendTo != null) //&& UIDL == ""
                            {
                                cmd_allEmail.Connection = conn;
                                conn.Open();
                                using (SqlDataReader sdr = cmd_allEmail.ExecuteReader())
                                {
                                    cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLInsert_AllEmail"], con);

                                    cmd.Parameters.Add("@sendFrom", SqlDbType.NVarChar, 50).Value = _SendFrom ?? (object)DBNull.Value;
                                    cmd.Parameters.Add("@sendDateTime", SqlDbType.DateTime).Value = _RejectDateTime; //?? (object)DBNull.Value;
                                    cmd.Parameters.Add("@sendTo", SqlDbType.NVarChar, 50).Value = _SendTo ?? (object)DBNull.Value;
                                    cmd.Parameters.Add("@subject", SqlDbType.NVarChar, 255).Value = _Subject ?? (object)DBNull.Value;
                                    cmd.Parameters.Add("@emailType", SqlDbType.NVarChar, 50).Value = _EmailType ?? (object)DBNull.Value;
                                    cmd.Parameters.Add("@email_ID", SqlDbType.NVarChar, 25).Value = UIDL.ToString() ?? (object)DBNull.Value;
                                    cmd.ExecuteNonQuery();
                                }
                                conn.Close();

                                if (_RejectDateTime == DateTime.Now || _SendDate == DateTime.Now)
                                {

                                }

                            }
                            
                        }

                    #endregion
                }

			}
			catch (Exception ex)
			{
                CreateLogFile(ex.ToString());
                //MessageBoxButtons MsgBox = MessageBoxButtons.OK;
                //MessageBox.Show("Error: Save Record to Database \r\n " + ex.ToString() + "\r\n Email ID: " + UIDL + "\r\n Date:" + _SendDate, "Error", MsgBox);

			}
		}

        private void DeleteRecord()
        {
            try
            {

                DateTime DelDate = DateTime.Now.AddDays(-1).Date;
                DateTime delDate = Convert.ToDateTime(DateTime.Today.ToShortDateString());

                SqlConnection con = new SqlConnection(dbcon);
                SqlCommand cmd;
                con.Open();

                cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLDelRecord"], con);
                cmd.Parameters.AddWithValue("@DelDate", delDate);
                cmd.ExecuteNonQuery();

                //bool recordExist = false;

                //SqlConnection con = new SqlConnection(dbcon);
                //SqlCommand cmd;
                //con.Open();
                
                //string DelDate = DateTime.Now.AddDays(-1).Date;

                //using (SqlConnection conn = new SqlConnection(dbcon))
                //{
                //    using (SqlCommand cmdd = new SqlCommand())  // AND sendFrom = '" + _SendFrom.ToString() + "'"))
                //    {
                //        cmdd.Connection = conn;
                //        conn.Open();
                //        using (SqlDataReader sdr = cmdd.ExecuteReader())
                //        {
                           
                //            cmd = new SqlCommand(ConfigurationManager.AppSettings["SQLDelRecord" + DateTime.Today.], con);

                //            cmd.ExecuteNonQuery();
                //            //}
                //        }
                //        conn.Close();
                //    }
                //}

            }
            catch (Exception ex)
            {
                CreateLogFile(ex.ToString());
            }

        }

		private string RejectedReasonTrim(string _TempRejectedReason)
		{
			_TempRejectedReason = Regex.Replace(_TempRejectedReason, @"\d<-#", "");
			_Reason = _TempRejectedReason;

			if (_TempRejectedReason.Contains("(in"))
			{
				_Reason = _Reason.Split(new string[] { "(in" }, StringSplitOptions.None).First().ToString();
			}

			if (_TempRejectedReason.Contains("(n"))
			{
				_Reason = _Reason.Split(new string[] { "(not" }, StringSplitOptions.None).First().ToString();
			}

			if (_TempRejectedReason.Contains("(#"))
			{
				_Reason = _Reason.Split(new string[] { "(#" }, StringSplitOptions.None).First().ToString();
			}

			_Reason = _Reason.Replace("#", "").ToString();

			//if (_TempRejectedReason.Contains(SearchCond.GlobalSymbolNewLine) || _TempRejectedReason.Contains("...") || _TempRejectedReason.Contains(">>> DATA <<<") || _TempRejectedReason.Contains("     ") || _TempRejectedReason.Contains("dd") || _TempRejectedReason.Contains("\r\n    "))
			//{
			#region Error Code

			if (_TempRejectedReason.Contains("451"))
				_Reason = _Reason.Replace("451", "").ToString();
				//_Reason = _Reason.Split(new string[] { "451" }, StringSplitOptions.None).Last();

			if (_TempRejectedReason.Contains("452-4.2.2") || _TempRejectedReason.Contains("452 4.2.2"))
				_Reason = _Reason.Replace("4.2.2", "").ToString();
				//_Reason = _Reason.Split(new string[] { "4.2.2" }, StringSplitOptions.None).Last();

            if (_TempRejectedReason.Contains("452-"))
                _Reason = _Reason.Replace("452- ", "").ToString();

            if (_TempRejectedReason.Contains("452"))
                _Reason = _Reason.Replace("452 ", "").ToString();

            if (_TempRejectedReason.Contains("422"))
                _Reason = _Reason.Replace("422 ", "").ToString();

			_Reason = _Reason.Replace("550", "").ToString();
			//if (_TempRejectedReason.Contains("550"))
			//    _RejectedReason = _RejectedReason.Split(new string[] { "550" }, StringSplitOptions.None).Last();

			if (_TempRejectedReason.Contains("5.5.0"))
				_Reason = _Reason.Replace("5.5.0", "").ToString();
				//_Reason = _Reason.Split(new string[] { "5.5.0" }, StringSplitOptions.None).Last();

			if (_TempRejectedReason.Contains("5.1.1"))
			{
				_Reason = _Reason.Replace("-5.1.1", "").ToString();
				_Reason = _Reason.Replace("5.1.1", "").ToString();
				//_Reason = _Reason.Split(new string[] { "5.1.1" }, StringSplitOptions.None).Last();
			}
			   
			if (_TempRejectedReason.Contains("5.1.0"))
				_Reason = _Reason.Replace("5.1.0", "").ToString();
				//_Reason = _Reason.Split(new string[] { "5.1.0" }, StringSplitOptions.None).Last();

			if (_TempRejectedReason.Contains("5.2.1"))
				_Reason = _Reason.Replace("5.2.1", "").ToString();
				//_Reason = _Reason.Split(new string[] { "5.2.1" }, StringSplitOptions.None).Last();

			if (_TempRejectedReason.Contains("5.4.1"))
				_Reason = _Reason.Replace("5.4.1", "").ToString(); 
			//_Reason = _Reason.Split(new string[] { "5.4.1" }, StringSplitOptions.None).Last();

            if (_TempRejectedReason.Contains("5.4.0"))
                _Reason = _Reason.Replace("5.4.0 ", "").ToString();

            if (_TempRejectedReason.Contains("5.4.6"))
                _Reason = _Reason.Replace("5.4.6 ", "").ToString(); 

			if (_TempRejectedReason.Contains("552"))
				_Reason = _Reason.Replace("552", "").ToString();
				//_Reason = _Reason.Split(new string[] { "552" }, StringSplitOptions.None).Last();

			if (_TempRejectedReason.Contains("5.3.0"))
				_Reason = _Reason.Replace("5.3.0", "").ToString();
				//_Reason = _Reason.Replace("5.3.0", "").ToString();

			if (_TempRejectedReason.Contains("554"))
			{
				_Reason = _Reason.Split(new string[] { "delivery error:" }, StringSplitOptions.None).Last();
				_Reason = _Reason.Split(new string[] { "554" }, StringSplitOptions.None).Last();
				_Reason = _Reason.Replace("5.0.0", "").ToString();
			}

			_Reason = _Reason.Replace("5.1.2", "").ToString();
            _Reason = _Reason.Replace("-5.2.2", "").ToString();
			_Reason = _Reason.Replace("5.0.0", "").ToString();
			_Reason = _Reason.Replace("553", "").ToString();
            _Reason = _Reason.Replace("500 ", "").ToString();
			_Reason = _Reason.Replace("5.7.1", "").ToString();


			#endregion

			if (_TempRejectedReason.Contains(" -"))
				_Reason = _Reason.Replace(" -", "").ToString(); //_Reason = _Reason.Split(new string[] { " -" }, StringSplitOptions.None).First();

			if (_TempRejectedReason.Contains("\r\n"))
				_Reason = _Reason.Replace("\r\n", " ").ToString();

			if (_TempRejectedReason.Contains("..."))
				_Reason = _Reason.Replace("...", "").ToString();

			if (_TempRejectedReason.Contains(">>> DATA <<<")) //|| tbErrorReason.Text.Contains("<<<"))
				_Reason = _Reason.Replace(">>> DATA <<<", "").ToString();

			if (_TempRejectedReason.Contains("<<<"))
				_Reason = _Reason.Replace(" <<<", "").ToString();

			if (_TempRejectedReason.Contains("\r\n    "))
				_Reason = _Reason.Replace("\r\n    ", " ").ToString();

			if (_TempRejectedReason.Contains("    "))
				_Reason = _Reason.Replace("     ", " ").ToString();

			if (_TempRejectedReason.Contains(" dd"))
			{
				_Reason = _Reason.Replace(" dd\r\n", "").ToString();
				_Reason = _Reason.Replace(" dd ", "").ToString();
				_Reason = _Reason.Split(new string[] { "[" }, StringSplitOptions.None).First();
			}
			
			if (_TempRejectedReason.Contains("delivery temporarily suspended:"))
				_Reason = _Reason.Replace("delivery temporarily suspended:", "").ToString();

			if (_TempRejectedReason.Contains(":"))
				_Reason = _Reason.Replace(":", "").ToString();

			if (_TempRejectedReason.Contains("type=A"))
				_Reason = _Reason.Split(new string[] { "type=A" }, StringSplitOptions.None).First();

			if (_TempRejectedReason.Contains("Last-Attempt-Date"))
				_Reason = _Reason.Split(new string[] { "Last-Attempt-Date" }, StringSplitOptions.None).First();

			if (_TempRejectedReason.Contains("Please"))
				_Reason = _Reason.Split(new string[] { "Please" }, StringSplitOptions.None).First();

			//}
			return _Reason;

		}
	   
		private void gvEmailList_CellContentClick(object sender, DataGridViewCellEventArgs e)
		{
			//InitiateFieldName();
            //int row = e.RowIndex;
            //PageDetail form2 = new PageDetail();
            //DataTable EmailList;
            //if (e.RowIndex >= 0 && e.RowIndex != null)
            //{

            //    DataGridViewRow selectedRow = gvEmailList.Rows[gvEmailList.SelectedCells[0].RowIndex];
				
            //    UIDL = selectedRow.Cells[6].Value.ToString();
            //    EmailList = dsEmail.Tables[0];


            //    var query = EmailList.Select(string.Format("email_ID ='{0}' ", UIDL.ToString())); // ("email_ID = " + UIDL.ToString());

            //    foreach (var c in query)
            //    {
            //        //dt.Rows.Add(new object[] { c.ItemArray[1], c.ItemArray[4], c.ItemArray[5], c.ItemArray[6], c.ItemArray[8], c.ItemArray[9], c.ItemArray[0], c.ItemArray[7], c.ItemArray[3], c.ItemArray[2] });
            //        _SendFrom = c.ItemArray[3].ToString();
            //        _SendFromIPAddress = c.ItemArray[2].ToString();
            //        _Subject = c.ItemArray[5].ToString();
            //        _RejectDateTime = Convert.ToDateTime(c.ItemArray[6]);
            //        _ReasonCategory = c.ItemArray[8].ToString();
            //        _Reason = c.ItemArray[7].ToString();
            //        _SendDate = Convert.ToDateTime(c.ItemArray[1]);
            //        _SendTo = c.ItemArray[4].ToString();
            //        _ActionRequired = c.ItemArray[9].ToString();
            //        _ActionTaken = c.ItemArray[10].ToString();
            //        _ActionDescription = c.ItemArray[11].ToString();
            //        _ActionTakenBy = c.ItemArray[12].ToString();
            //        _CaseStatus = Convert.ToBoolean(c.ItemArray[14]);
            //    }

            //    form2.Show();
			//}
		 
		   
		}

		private bool CheckUnicode(string subject)
		{
            if (subject != null && subject != "")
            {
                string word = subject.Trim();
                //bool containUnicode = false;
                for (int x = 0; x < word.Length; x++)
                {
                    if (char.GetUnicodeCategory(word[x]) == UnicodeCategory.OtherLetter)
                    {
                        ChineseChar = true;
                        return ChineseChar;
                    }
                }
               

            }
            ChineseChar = false;
            return ChineseChar;
			

		}

		private void comboxCaseStatus_SelectedIndexChanged(object sender, EventArgs e)
		{
			//InitiateFieldName();
			DBLoadEmail();
			DataTable EmailList;
			
			int counter;
			if (dsEmail.Tables["Email_Details"] == null || dsEmail.Tables["Email_Details"].Rows.Count == 0)
			{
				EmailList = dsEmail.Tables[0];
				counter = dsEmail.Tables[0].Rows.Count;
			}
			else
			{
				EmailList = dsEmail.Tables["Email_Details"];
				counter = dsEmail.Tables["Email_Details"].Rows.Count;
			}

			int index = comboxCaseStatus.SelectedIndex;

			switch (index)
			{
				case 0:

					//gvEmailList.DataSource = EmailList;
					var query = EmailList.Select();

					foreach (var c in query)
					{
						dt.Rows.Add(new object[] { c.ItemArray[1], c.ItemArray[4], c.ItemArray[5], c.ItemArray[6], c.ItemArray[8], c.ItemArray[9], c.ItemArray[0], c.ItemArray[7], c.ItemArray[3], c.ItemArray[2]});
					}

					break;

				case 1:
					query = EmailList.Select("closeCase <> 1");

					foreach (var c in query)
					{

						dt.Rows.Add(new object[] { c.ItemArray[1], c.ItemArray[4], c.ItemArray[5], c.ItemArray[6], c.ItemArray[8], c.ItemArray[9], c.ItemArray[0], c.ItemArray[7], c.ItemArray[3], c.ItemArray[2] });
						
					}
					
					break;

				case 2:
					query = EmailList.Select("closeCase <> 0");
				   
					foreach (var c in query)
					{
						dt.Rows.Add(new object[] { c.ItemArray[1], c.ItemArray[4], c.ItemArray[5], c.ItemArray[6], c.ItemArray[8], c.ItemArray[9], c.ItemArray[0], c.ItemArray[7], c.ItemArray[3], c.ItemArray[2] });
					}

					break;
			}
			//if (gvEmailList.RowCount <= 0)
			//{
			//    //InitiateFieldName();
			//}

		}

		private void button1_Click(object sender, EventArgs e)
		{
			tabControl1.SelectTab("tabPageEmailList");
		}

		private void timer1_Tick(object sender, EventArgs e)
		{
			timer1.Stop();
		   // DeleteRecord();
			DBLoadEmail();
			ServerConnection();
            dt.Clear();
            dsEmail.Clear();
		}

		private void CreateLogFile(string error)
		{
            if (!error.Contains("big5_tw"))
            {
                string fileName = string.Format("RejectedEmailLog-{0:yyyy-MM-dd}.txt", DateTime.Now);
                string filePath = ConfigurationManager.AppSettings["LogPath"] + fileName;
                filePath = Regex.Replace(filePath, @"[/]", "").ToString();

                var lineCount = 0;
                string _tempError = error.Split(new string[] { ". " }, StringSplitOptions.None).First();

                if (!File.Exists(filePath))
                {
                    using (FileStream fs = File.Create(filePath)) ; { }
                    // Create a file to write to.
                    string createText = "=============== ERROR DETAIL =============== " + Environment.NewLine;
                    File.WriteAllText(filePath, createText);
                }

                using (var file = new StreamReader(filePath))
                {
                    while (file.ReadLine() != null)
                    {
                        lineCount++;
                    }
                }
                List<string> lines = System.IO.File.ReadAllLines(filePath).ToList<string>();

                if (_Subject == null) _Subject = "";
                if (_SendFrom == null) _SendFrom = "";


                string errorDetail = ("Run Time: " + DateTime.Now.TimeOfDay + "\r\nSend Date: " + _SendDate.ToString() + "\r\nReject Time: " + _RejectDateTime.ToString() + "\r\nSubject: " + _Subject + "\r\nSend From: " + _SendFrom.ToString() + "\r\nError Cause: " + error.ToString()).ToString() + " \r\n================================================== ";
                File.AppendAllText(filePath, errorDetail + Environment.NewLine);
            }
		}

        #region Rejected Email Pattern Tab
        // Display all rejected email pattern & action required.
        // Allow user to add new rejected email pattern & action required
        private void InitiateFieldnameEmailRejectPattern()
        {
            txtActionCategory.Text = "";
            txtActionRequired.Text = "";
            txtReasonPattern.Text = "";
            txtRejectCategory.Text = "";

            buttonColumn.HeaderText = "";
            buttonColumn.Name = "View";
            buttonColumn.UseColumnTextForButtonValue = true;

            //gvEmailList.Columns.Add(buttonColumn);
            if (dgvRejectReasonPattern.ColumnCount == 0)
            {
                dtRejectPattern.Columns.Add("Code", typeof(string));
                dtRejectPattern.Columns.Add("Category", typeof(string));
                dtRejectPattern.Columns.Add("Reason", typeof(string));

                dgvRejectReasonPattern.DataSource = dtRejectPattern;
                dgvRejectReasonPattern.Sort(dgvRejectReasonPattern.Columns["Code"], ListSortDirection.Ascending);
                dgvRejectReasonPattern.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                dgvRejectReasonPattern.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            }
            if (dgvActionRequired.ColumnCount == 0)
            {
                dtActionTaken.Columns.Add("Code", typeof(string));
                dtActionTaken.Columns.Add("Category", typeof(string));
                dtActionTaken.Columns.Add("Action Required", typeof(string));

                dgvActionRequired.DataSource = dtActionTaken;
                dgvActionRequired.Sort(dgvActionRequired.Columns["Code"], ListSortDirection.Ascending);
                dgvActionRequired.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                dgvActionRequired.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process_tabPageRejectPattern();
           
            tabControl1.SelectTab("tabPageRejectPattern");
        }

        private void Process_tabPageRejectPattern()
        {
            dtActionTaken.Rows.Clear();
            dtRejectPattern.Rows.Clear();
            DBLoadEmail();
            InitiateFieldnameEmailRejectPattern();
            LoadRejectedPattern_ActionRequired();

        }

        private void LoadRejectedPattern_ActionRequired()
        {
            DataRow[] hasRowRejectPattern = dsEmail.Tables[1].Select();
            foreach (DataRow Returned in hasRowRejectPattern)
            {
                var newRow = new string[] { Returned[0].ToString(), Returned[2].ToString(), Returned[1].ToString() };
                dtRejectPattern.Rows.Add(newRow);
            }

            DataRow[] hasRowActionTaken = dsEmail.Tables[2].Select();
            foreach (DataRow Returned in hasRowActionTaken)
            {
                var newRow = new string[] { Returned[0].ToString(), Returned[1].ToString(), Returned[2].ToString() };
                dtActionTaken.Rows.Add(newRow);
            }

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtActionCategory.Text != "")
                {
                    DBLoadEmail();
                    int actionCtr = 1, reasonCtr = 1;
                    CultureInfo cultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
                    TextInfo textInfo = cultureInfo.TextInfo;

                    SqlConnection con = new SqlConnection(dbcon);
                    SqlCommand cmd;
                    con.Open();

                    reasonCtr = dsEmail.Tables[1].Rows.Count + 1;

                    actionCtr = dsEmail.Tables[2].Rows.Count + 1;

                    string tempReasonPattern = txtReasonPattern.Text.ToString().ToLower();
                    string tempRejectCategory = textInfo.ToTitleCase(txtRejectCategory.Text);
                    string tempActionRequired = textInfo.ToTitleCase(txtActionRequired.Text);
                    string tempActRejectCat = textInfo.ToTitleCase(txtRejectCategory.Text);

                    //DBLoadEmail();
                    DataRow[] ReasonRow = dsEmail.Tables[1].Select("reason ='" + tempReasonPattern + "'");
                    if (ReasonRow.Length <= 0)
                    {
                        cmd = new SqlCommand(ConfigurationManager.AppSettings["Reason_Insert"], con);

                        cmd.Parameters.Add("@code", SqlDbType.NVarChar, 10).Value = reasonCtr.ToString().Trim() ?? (object)DBNull.Value;
                        cmd.Parameters.Add("@reason", SqlDbType.VarChar, 255).Value = txtReasonPattern.Text.ToString().ToLower() ?? (object)DBNull.Value;
                        cmd.Parameters.Add("@category", SqlDbType.VarChar, 100).Value = textInfo.ToTitleCase(txtRejectCategory.Text) ?? (object)DBNull.Value;

                        cmd.CommandType = CommandType.Text;
                        int i = cmd.ExecuteNonQuery();
                    }

                    DataRow[] ActionRow = dsEmail.Tables[2].Select("category ='" + tempActRejectCat + "'");
                    if (ActionRow.Length <= 0)
                    {
                        cmd = new SqlCommand(ConfigurationManager.AppSettings["Action_Insert"], con);

                        cmd.Parameters.Add("@code", SqlDbType.NVarChar, 10).Value = actionCtr.ToString().Trim() ?? (object)DBNull.Value;
                        cmd.Parameters.Add("@action", SqlDbType.VarChar, 255).Value = textInfo.ToTitleCase(txtActionRequired.Text) ?? (object)DBNull.Value;
                        cmd.Parameters.Add("@category", SqlDbType.VarChar, 100).Value = textInfo.ToTitleCase(txtActionCategory.Text) ?? (object)DBNull.Value;

                        cmd.CommandType = CommandType.Text;
                        int i = cmd.ExecuteNonQuery();
                    }
                    Process_tabPageRejectPattern();

                }
                else
                {
                    MessageBoxButtons MsgBox = MessageBoxButtons.OK;
                    MessageBox.Show("Error: Incomplete rejected email pattern or category", "Error", MsgBox);

                }
               
            }
            catch (Exception ex)
            {
                CreateLogFile(ex.ToString());
                //MessageBoxButtons MsgBox = MessageBoxButtons.OK;
                //MessageBox.Show("Error: Failed to Save Rejected Email Pattern & Action \r\n" + ex.ToString(), "Error", MsgBox);
            }

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtActionCategory.Text = "";
            txtActionRequired.Text = "";
            txtReasonPattern.Text = "";
            txtRejectCategory.Text = "";
        }

        private void txtRejectCategory_Leave(object sender, EventArgs e)
        {
            txtActionCategory.Text = txtRejectCategory.Text;
        }

        #endregion

        #region Email Address Login Tab
        // Display all existing email login info from database to the program.
        // Allow user to add new email login information
        private void btnSaveEmailLogin_Click(object sender, EventArgs e)
        {

            if ((tbLoginEmailAddress.Text != "" ) && (tbLoginHostname.Text != "") && (tbLoginPassword.Text != "" ) && (tbLoginPortNo.Text != "" ) && (tbLoginUsername.Text != ""))
            {
                int emailLoginCtr = 1;
                CultureInfo cultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
                TextInfo textInfo = cultureInfo.TextInfo;

                SqlConnection con = new SqlConnection(dbcon);
                SqlCommand cmd;
                con.Open();

                emailLoginCtr = dsEmail.Tables[4].Rows.Count + 1;

                DBLoadEmail();
                DataRow[] ReasonRow = dsEmail.Tables[4].Select("id ='" + emailLoginCtr + "'");
                if (ReasonRow.Length <= 0)
                {
                    cmd = new SqlCommand(ConfigurationManager.AppSettings["EmailLogin_Insert"], con);

                    cmd.Parameters.Add("@id", SqlDbType.Int, 0).Value = emailLoginCtr;
                    cmd.Parameters.Add("@hostname", SqlDbType.NVarChar, 50).Value = tbLoginHostname.Text;
                    cmd.Parameters.Add("@emailAddress", SqlDbType.NVarChar, 50).Value = tbLoginEmailAddress.Text;
                    cmd.Parameters.Add("@username", SqlDbType.NVarChar, 50).Value = tbLoginUsername.Text;
                    cmd.Parameters.Add("@password", SqlDbType.NVarChar, 50).Value = tbLoginPassword.Text;
                    cmd.Parameters.Add("@portNo", SqlDbType.NVarChar, 10).Value = tbLoginPortNo.Text;

                    cmd.CommandType = CommandType.Text;
                    int i = cmd.ExecuteNonQuery();
                }

                MessageBoxButtons MsgBox = MessageBoxButtons.OK;
                MessageBox.Show("Record successfully saved", "Success", MsgBox);

                dtEmailLogin.Rows.Clear();
                tabControl1.SelectTab("tabPageEmailLogin");
                DBLoadEmail();
                InitiateFieldnameEmailLogin();
                LoadEmailLoginRecord();
            }
            else
            {
                MessageBoxButtons MsgBox = MessageBoxButtons.OK;
                MessageBox.Show("Incomplete email address login information.", "Error", MsgBox);
            }
        }
        
        private void InitiateFieldnameEmailLogin()
        {
            tbLoginEmailAddress.Text = "";
            tbLoginHostname.Text = "";
            tbLoginPassword.Text = "";
            tbLoginPortNo.Text = "";
            tbLoginUsername.Text = "";

            buttonColumn.HeaderText = "";
            buttonColumn.Name = "View";
            buttonColumn.UseColumnTextForButtonValue = true;

            //gvEmailList.Columns.Add(buttonColumn);
            if (dgvEmailLogin.ColumnCount == 0)
            {
                dtEmailLogin.Columns.Add("ID", typeof(int));
                dtEmailLogin.Columns.Add("Hostname", typeof(string));
                dtEmailLogin.Columns.Add("Email Address", typeof(string));
                dtEmailLogin.Columns.Add("Username", typeof(string));
                dtEmailLogin.Columns.Add("Password", typeof(string));
                dtEmailLogin.Columns.Add("Port No", typeof(string));

                dgvEmailLogin.DataSource = dtEmailLogin;
                dgvEmailLogin.Sort(dgvEmailLogin.Columns["ID"], ListSortDirection.Ascending);
                dgvEmailLogin.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                dgvEmailLogin.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }

        }

        private void btnClearEmailLogin_Click(object sender, EventArgs e)
        {
            tbLoginEmailAddress.Text = "";
            tbLoginHostname.Text = "";
            tbLoginPassword.Text = "";
            tbLoginPortNo.Text = "";
            tbLoginUsername.Text = "";
        }

        private void LoadEmailLoginRecord()
        {
            DataRow[] hasRow = dsEmail.Tables[4].Select();
            foreach (DataRow Returned in hasRow)
            {
                var newRow = new string[] { Returned[0].ToString(), Returned[1].ToString(), Returned[2].ToString(), Returned[3].ToString(), Returned[4].ToString(), Returned[5].ToString() };
                dtEmailLogin.Rows.Add(newRow);
            }
        }

        private void btnEmailLogin_Click_1(object sender, EventArgs e)
        {
            dtEmailLogin.Rows.Clear();
            tabControl1.SelectTab("tabPageEmailLogin");
            DBLoadEmail();
            InitiateFieldnameEmailLogin();
            LoadEmailLoginRecord();
        }

        private void btnClearEmailLogin_Click_1(object sender, EventArgs e)
        {
            tbLoginEmailAddress.Text = "";
            tbLoginHostname.Text = "";
            tbLoginPassword.Text = "";
            tbLoginPortNo.Text = "";
            tbLoginUsername.Text = "";
        }
       
        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (tbLoginEmailAddress.Text != "")
                {
                    MessageBoxButtons MsgBox = MessageBoxButtons.YesNo;
                    DialogResult userAnswer = MessageBox.Show("Are you sure you want to delete this record?", "Alert", MsgBox);
                    if (userAnswer == DialogResult.Yes)
                    {
                        string Del_ID = lbl_LoginID.Text.ToString();
                        SqlConnection con = new SqlConnection(dbcon);
                        SqlCommand cmd;
                        con.Open();

                        cmd = new SqlCommand(ConfigurationManager.AppSettings["EmailLogin_Delete"], con);
                        cmd.Parameters.AddWithValue("@id", Del_ID);
                        cmd.ExecuteNonQuery();

                        dtEmailLogin.Rows.Clear();
                        tabControl1.SelectTab("tabPageEmailLogin");
                        DBLoadEmail();
                        InitiateFieldnameEmailLogin();
                        LoadEmailLoginRecord();

                    }
                    else
                    {

                    }

                }
                else
                {
                    MessageBoxButtons MsgBox = MessageBoxButtons.OK;
                    MessageBox.Show("Please select 1 record", "Alert", MsgBox);

                }
               

            }
            catch (Exception ex)
            {
                MessageBoxButtons MsgBox = MessageBoxButtons.OK;
                MessageBox.Show("Error: Delete Todays Record \r\n" + ex.ToString(), "Error", MsgBox);

            }

        }

        private void dgvEmailLogin_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            string counter;
            DataTable EmailLogin;
            if (e.RowIndex >= 0 && e.RowIndex != null)
            {

                DataGridViewRow selectedRow = dgvEmailLogin.Rows[dgvEmailLogin.SelectedCells[0].RowIndex];

                counter = selectedRow.Cells[0].Value.ToString();
                EmailLogin = dsEmail.Tables[4];


                var query = EmailLogin.Select(string.Format("id ='{0}' ", counter)); // ("email_ID = " + UIDL.ToString());

                foreach (var c in query)
                {
                    tbLoginEmailAddress.Text = c.ItemArray[2].ToString();
                    tbLoginHostname.Text = c.ItemArray[1].ToString();
                    tbLoginPassword.Text = c.ItemArray[4].ToString();
                    tbLoginPortNo.Text = c.ItemArray[5].ToString();
                    tbLoginUsername.Text = c.ItemArray[3].ToString();
                    lbl_LoginID.Text = c.ItemArray[0].ToString();

                }

            }

        }

        private void dgvEmailLogin_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == 4 && e.Value != null)
            {
                e.Value = new String('*', e.Value.ToString().Length);
            }
        }

        #endregion

        private void gvEmailList_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
			PageDetail form2 = new PageDetail();
			DataTable EmailList;
            if (e.RowIndex >= 0 && e.RowIndex != null)
            {

                DataGridViewRow selectedRow = gvEmailList.Rows[gvEmailList.SelectedCells[0].RowIndex];

                UIDL = selectedRow.Cells[6].Value.ToString();
                EmailList = dsEmail.Tables[0];


                var query = EmailList.Select(string.Format("email_ID ='{0}' ", UIDL.ToString())); // ("email_ID = " + UIDL.ToString());

                foreach (var c in query)
                {
                    //dt.Rows.Add(new object[] { c.ItemArray[1], c.ItemArray[4], c.ItemArray[5], c.ItemArray[6], c.ItemArray[8], c.ItemArray[9], c.ItemArray[0], c.ItemArray[7], c.ItemArray[3], c.ItemArray[2] });
                    _SendFrom = c.ItemArray[3].ToString();
                    _SendFromIPAddress = c.ItemArray[2].ToString();
                    _Subject = c.ItemArray[5].ToString();
                    _RejectDateTime = Convert.ToDateTime(c.ItemArray[6]);
                    _ReasonCategory = c.ItemArray[8].ToString();
                    _Reason = c.ItemArray[7].ToString();
                    _SendDate = Convert.ToDateTime(c.ItemArray[1]);
                    _SendTo = c.ItemArray[4].ToString();
                    _ActionRequired = c.ItemArray[9].ToString();
                    _ActionTaken = c.ItemArray[10].ToString();
                    _ActionDescription = c.ItemArray[11].ToString();
                    _ActionTakenBy = c.ItemArray[12].ToString();
                    _CaseStatus = Convert.ToBoolean(c.ItemArray[14]);
                }

                form2.Show();
            }

        }

        private void btnEmailSearch_Click(object sender, EventArgs e)
        {
            //DBLoadEmail();
            //DataTable EmailList;

            //if (dsEmail.Tables["Email_Details"] == null || dsEmail.Tables["Email_Details"].Rows.Count == 0)
            //{
            //    EmailList = dsEmail.Tables[0];
            //}
            //else
            //{
            //    EmailList = dsEmail.Tables["Email_Details"];
            //}

            //var filter = "sendDateTime >='" + Convert.ToDateTime(dtPickerDateFrom.ToString()) + " 00:00:00' AND sendDateTime < '" + Convert.ToDateTime(dtPickerDateTo.ToString()) + " 23:59:59'"";
            //var sort = "sendDateTime DESC";
            //var query = EmailList.Select();

            //foreach (var c in query)
            //{
            //    dt.Rows.Add(new object[] { c.ItemArray[1], c.ItemArray[4], c.ItemArray[5], c.ItemArray[6], c.ItemArray[8], c.ItemArray[9], c.ItemArray[0], c.ItemArray[7], c.ItemArray[3], c.ItemArray[2] });

            //}

        }

    }
	
}
