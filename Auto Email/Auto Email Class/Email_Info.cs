using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Auto_Email_Class
{
    public class Email_Info
    {
        public string UIDL { get; set; }
        public string _SendTo { get; set; }
        public string _SendFrom { get; set; }
        public string _Subject { get; set; }
        public string _Reason { get; set; }
        public string _TrialReason { get; set; }
        public string _ReasonCategory { get; set; }
        public string _ActionRequired { get; set; }
        public string _SendFromIPAddress { get; set; }
        public DateTime _SendDate { get; set; }
        public DateTime _RejectDateTime { get; set; }

        public struct FieldName
        {
            public string UIDL;
            public string _SendTo;
            public string _SendFrom ;
            public string _Subject ;
            public string _Reason ;
            public string _TrialReason ;
            public string _ReasonCategory;
            public string _ActionRequired;
            public string _SendFromIPAddress ;
            public DateTime _SendDate ;
            public DateTime _RejectDateTime ;
        }

        public FieldName InitializeFieldname()
        {
            FieldName _EmailInfo = new FieldName();
            UIDL = "";
            _SendTo = "";
            _SendFrom = "";
            _Subject = "";
            _Reason = "";
            _TrialReason = "";
            _ReasonCategory = "";
            _ActionRequired = "";
            _SendFromIPAddress = "";
            _SendDate = DateTime.Now.AddDays(-15);
            _RejectDateTime = DateTime.Now.AddDays(-15);

            //tbActionDateTime.Text = "";
            //tbActionDescription.Text = "";
            //tbActionRequired.Text = "";
            //tbActionTaken.Text = "";
            //tbActionTakenBy.Text = "";
            //cbCaseStatus.Checked = false;

            return _EmailInfo;


        }



        
    }

}
