#region UsingDirectives
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace addcontactsfrommail
{
    /// <summary>
    /// Class object for stored contact info
    /// </summary>
    public class ContactInfo
    {
        /// <summary>
        /// sender name value
        /// </summary>
        public string SenderName { get; set; }

        /// <summary>
        /// sender mail adress value
        /// </summary>
        public string SenderMail { get; set; }
    }
}
