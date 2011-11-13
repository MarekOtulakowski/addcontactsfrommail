#region UsingDirectives
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MSOutlook = Microsoft.Office.Interop.Outlook;
using System.Linq.Expressions;
using System.Runtime.InteropServices; 
#endregion

namespace addcontactsfrommail
{
    public partial class F_Main : Form
    {
        #region Variables
        /// <summary>
        /// application MS Outlook
        /// </summary>
        MSOutlook.Application oApp = null;

        /// <summary>
        /// namespace MS Outlook
        /// </summary>
        MSOutlook.NameSpace oNS = null;

        /// <summary>
        /// folder object MS Outlook
        /// </summary>
        MSOutlook.MAPIFolder folder = null;

        /// <summary>
        /// mail object MS Outlook
        /// </summary>
        MSOutlook.MailItem mail = null;

        /// <summary>
        /// helping variable for check selected folder from user (browse folder it's true value)
        /// </summary>
        bool selectedFolder = false;

        /// <summary>
        /// helping variable for user choosed read sender or recipient mail (browse folder question)
        /// </summary>
        bool senderContactsFromSelectedFolder = true;

        /// <summary>
        /// helping variable for abort thread (BW_worker.CancelAsync() not break thread)
        /// </summary>
        private bool cancelThread = false;

        /// <summary>
        /// BackgroundWorker for background thread import contacts, becouse imported contact freeze mail form
        /// </summary>
        private System.ComponentModel.BackgroundWorker BW_worker;

        /// <summary>
        /// count all folder and subfolder for sected source folder
        /// </summary>
        int sourceFolderToSearchCount = 1;

        /// <summary>
        /// count all reading mail
        /// </summary>
        int countAllmail = 0;

        /// <summary>
        /// helping variable for actual count mail in current folder
        /// </summary>
        int sourceEmailCount = 0;

        /// <summary>
        /// contact object MS Outlook
        /// </summary>
        MSOutlook.ContactItem nContact = null;

        /// <summary>
        /// constans string for check folder type
        /// </summary>
        const string defaultNameClass = "IPM.Note"; 
        #endregion

        #region Properties
        /// <summary>
        /// Get source folder to search count
        /// </summary>
        /// <returns>sourceFolderToSearchCount</returns>
        public int GetSourceFolderToSearchCount()
        {
            return sourceFolderToSearchCount;
        }

        /// <summary>
        /// Get MS Outlook mail object (for test)
        /// </summary>
        /// <returns>mail MS object</returns>
        public MSOutlook.MailItem GetMSMail()
        {
            try
            {
                if (folder != null)
                    if (folder.Items.Count > 0)
                    {
                        if (mail != null) mail = null;
                        mail = folder.Items.GetFirst() as MSOutlook.MailItem;
                    }
            }
            catch { }

            return mail;
        }

        /// <summary>
        /// Get MS Outlook folder object (for test)
        /// </summary>
        /// <returns>folder MS object</returns>
        public MSOutlook.MAPIFolder GetMSOutlookFolder()
        {
            return folder;
        }

        /// <summary>
        /// Get MS Outlook application (for test)
        /// </summary>
        /// <returns>oApp applictation MS object</returns>
        public MSOutlook.Application GetMSOutlookApplication()
        {
            return oApp;
        }

        /// <summary>
        /// Get MS Outlook namespace (for test)
        /// </summary>
        /// <returns>oNS namespace MS object</returns>
        public MSOutlook.NameSpace GetMSOutlookNamespace()
        {
            return oNS;
        } 
        #endregion

        #region DefaultConstructor
        public F_Main()
        {
            InitializeComponent();

            this.BW_worker.WorkerReportsProgress = true;
            this.BW_worker.WorkerSupportsCancellation = true;
            this.BW_worker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BW_worker_DoWork);
            this.BW_worker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BW_worker_RunWorkerCompleted);
            this.BW_worker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.BW_worker_ProgressChanged);
        } 
        #endregion

        /// <summary>
        /// Return actual list of contacts (mail and name)
        /// from default MS Outlook folder
        /// </summary>
        /// <returns>list actual contacts</returns>
        private List<ContactInfo> GetActualContactsList()
        {
            List<ContactInfo> listContacts = new List<ContactInfo>();

            MSOutlook.MAPIFolder oContactsFolder = oNS.GetDefaultFolder(MSOutlook.OlDefaultFolders.olFolderContacts) as MSOutlook.MAPIFolder;

            foreach (MSOutlook.ContactItem oContact in oContactsFolder.Items)
            {
                ContactInfo newContact = new ContactInfo();
                newContact.SenderName = oContact.Email1DisplayName;
                newContact.SenderMail = oContact.Email1Address;
                listContacts.Add(newContact);
            }

            oContactsFolder = null;
            return listContacts;
        }
        
        /// <summary>
        /// Add contact from parametr folder
        /// </summary>
        /// <param name="oFolder">MAPIFolder object</param>
        /// <returns>result true if all success</returns>
        private bool AddContactsForOneFolder(MSOutlook.MAPIFolder oFolder)
        {
            bool result = false;

            try
            {              
                foreach (var oMail in oFolder.Items)
                {
                    ++sourceEmailCount;
                    if (!(oMail is MSOutlook.MailItem)) continue;
                    
                    int value = (int)(((float)sourceEmailCount / countAllmail) * 100);
                    
                    //for prevent error
                    if (value > 100)
                        value = 100;
                    if (value < 0)
                        value = 0;

                    BW_worker.ReportProgress(value);

                    if (cancelThread) return true;

                    List<ContactInfo> listActualContacts = GetActualContactsList();
                    if (oMail == null) break;
                    try
                    {
                        if (((MSOutlook.MailItem)oMail).SenderEmailAddress == null) continue;
                        string noInfo = "(" +
                                        sourceEmailCount.ToString() + 
                                        "/" + 
                                        countAllmail.ToString() +
                                        ") ";
                        this.UIThread(() => this.LB_result.Items.Add(noInfo +
                            "Check for email subject: " +
                            ((MSOutlook.MailItem)oMail).Subject));
                        this.UIThread(() => this.LB_result.SelectedIndex = this.LB_result.Items.Count - 1);
                    }
                    catch
                    {
                        continue;
                    }
                    ContactInfo oCon = listActualContacts.Find(s => s.SenderMail == ((MSOutlook.MailItem)oMail).SenderEmailAddress);

                    if (oCon == null)
                    {
                        if (nContact != null) nContact = null;
                        nContact = oApp.CreateItem(MSOutlook.OlItemType.olContactItem) as MSOutlook.ContactItem;

                        if (RB_receiveFolder.Checked)
                        {
                            if (((MSOutlook.MailItem)oMail).SenderEmailAddress == null) continue;
                            if (((MSOutlook.MailItem)oMail).SenderEmailAddress.Trim() == string.Empty) continue;

                            nContact.Email1Address = ((MSOutlook.MailItem)oMail).SenderEmailAddress;
                            if (((MSOutlook.MailItem)oMail).SenderName != null)
                            {
                                string name = ((MSOutlook.MailItem)oMail).SenderName;
                                name = RenameContactConditions(name);
                                nContact.FullName = name;
                            }
                            else
                            {
                                string name = ((MSOutlook.MailItem)oMail).SenderEmailAddress;
                                name = RenameContactConditions(name);
                                nContact.FullName = name;
                            }
                            nContact.Save();

                            AddNewLineToLB_result(nContact);
                            nContact = null;
                        }
                        else
                        {
                            if (senderContactsFromSelectedFolder)
                            {
                                if (((MSOutlook.MailItem)oMail).SenderEmailAddress == null) continue;
                                if (((MSOutlook.MailItem)oMail).SenderEmailAddress.Trim() == string.Empty) continue;

                                nContact.Email1Address = ((MSOutlook.MailItem)oMail).SenderEmailAddress;

                                if (((MSOutlook.MailItem)oMail).SenderName != null)
                                {
                                    string name = ((MSOutlook.MailItem)oMail).SenderName;
                                    name = RenameContactConditions(name);
                                    nContact.FullName = name;
                                }
                                else
                                {
                                    string name = ((MSOutlook.MailItem)oMail).SenderEmailAddress;
                                    name = RenameContactConditions(name);
                                    nContact.FullName = name;
                                }
                                nContact.Save();

                                AddNewLineToLB_result(nContact);
                            }
                            else
                            {
                                MSOutlook.Recipients recipients = ((MSOutlook.MailItem)oMail).Recipients as MSOutlook.Recipients;
                                foreach (MSOutlook.Recipient oRecipient in recipients)
                                {
                                    if (cancelThread) return true;

                                    ContactInfo oCon2 = listActualContacts.Find(s => s.SenderMail == oRecipient.Address);
                                    if (oCon2 == null)
                                    {
                                        if (nContact != null) nContact = null;

                                        if (oRecipient == null) continue;
                                        if (oRecipient.Address == null) continue;
                                        if (oRecipient.Address.Trim() == string.Empty) continue;

                                        nContact = oApp.CreateItem(MSOutlook.OlItemType.olContactItem) as MSOutlook.ContactItem;

                                        nContact.Email1Address = oRecipient.Address;
                                        if (oRecipient.Name != null)
                                        {
                                            string name = oRecipient.Name;
                                            name = RenameContactConditions(name);
                                            nContact.FullName = name;
                                        }
                                        else
                                        {
                                            string name = oRecipient.Address;
                                            name = RenameContactConditions(name);
                                            nContact.FullName = name;
                                        }
                                        nContact.Save();

                                        AddNewLineToLB_result(nContact);
                                        nContact = null;
                                    }
                                }

                                recipients = null;
                            }
                        }
                    }

                    if (RB_sendFolder.Checked)
                    {
                        MSOutlook.Recipients recipients = ((MSOutlook.MailItem)oMail).Recipients as MSOutlook.Recipients;
                        foreach (MSOutlook.Recipient oRecipient in recipients)
                        {
                            if (cancelThread) return true;

                            if (oRecipient.Address == null) continue;
                            ContactInfo oCon2 = listActualContacts.Find(s => s.SenderMail == oRecipient.Address);
                            if (oCon2 == null)
                            {
                                if (oRecipient == null) continue;
                                if (oRecipient.Address == null) continue;
                                if (oRecipient.Address.Trim() == string.Empty) continue;

                                if (nContact != null) nContact = null;
                                nContact = oApp.CreateItem(MSOutlook.OlItemType.olContactItem) as MSOutlook.ContactItem;

                                nContact.Email1Address = oRecipient.Address;
                                if (oRecipient.Name != null)
                                {
                                    string name = oRecipient.Name;
                                    name = RenameContactConditions(name);
                                    nContact.FullName = name;
                                }
                                else
                                {
                                    string name = oRecipient.Address;
                                    name = RenameContactConditions(name);
                                    nContact.FullName = name;
                                }
                                    
                                nContact.Save();

                                AddNewLineToLB_result(nContact);
                            }
                        }

                        recipients = null;
                    }

                    if (nContact != null) nContact = null;
                }

                if (oFolder != null) oFolder = null;

                result = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error AddContactFromFolder: " + ex.Message,
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }

            return result;
        }

        /// <summary>
        /// Add new linet to LB_result controls
        /// </summary>
        /// <param name="nContact">Outlook Contact</param>
        private void AddNewLineToLB_result(MSOutlook.ContactItem nContact)
        {
            string message = "Adding new Contact (" +
                             nContact.Email1DisplayName +
                             ") " +
                             nContact.Email1Address;
            this.UIThread(() => this.LB_result.Items.Add(message));
            this.UIThread(() => this.LB_result.SelectedIndex = this.LB_result.Items.Count - 1);
        }

        /// <summary>
        /// First rename contact name conditions
        /// </summary>
        /// <param name="name">source name</param>
        /// <returns>changed name</returns>
        public string RenameContactConditions(string name)
        {
            name = name.Replace("'", "");

            if (name.Contains("@"))
                name = name.Substring(0, name.IndexOf("@"));

            if (name.Contains("."))
                if (name.IndexOf(".") == name.LastIndexOf("."))
                    name = RenameContactName(name);

            if (name.Contains("-"))
                if (name.IndexOf("-") == name.LastIndexOf("-"))
                    name = RenameContactName(name);

            if (name.Contains("_"))
                if (name.IndexOf("_") == name.LastIndexOf("_"))
                    name = RenameContactName(name);

            if (name.Contains(" "))
                if (name.IndexOf(" ") == name.LastIndexOf(" "))
                    name = RenameContactName(name);

            return name;
        }

        /// <summary>
        /// Second rename contact name
        /// </summary>
        /// <param name="name">source name</param>
        /// <returns>return name</returns>
        private static string RenameContactName(string name)
        {
            name = name.ToLower();

            try
            {
                bool containsSign = false;
                if (name.Contains("."))
                {
                    if (name.Substring(name.LastIndexOf("."), 
                                       name.Length - name.LastIndexOf(".")) == ".")
                        containsSign = true;
                }
                if (name.Contains("-"))
                {
                    if (name.Substring(name.LastIndexOf("-"),
                                       name.Length - name.LastIndexOf("-")) == "-")
                        containsSign = true;
                }
                if (name.Contains("_"))
                {
                    if (name.Substring(name.LastIndexOf("_"),
                                       name.Length - name.LastIndexOf("_")) == "_")
                        containsSign = true;
                }

                if (containsSign)
                    return name;

                string fName = name.Substring(0, 1);
                fName = fName.ToUpper();
                string lName = string.Empty;
                name = fName + name.Substring(1, name.Length - 1);
                if (name.Contains("."))
                {
                    string sign = ".";
                    RenameForSingleSign(ref name, ref lName, ref sign);
                }
                else if (name.Contains(" "))
                {
                    string sign = " ";
                    RenameForSingleSign(ref name, ref lName, ref sign);
                }
                else if (name.Contains("-"))
                {
                    string sign = "-";
                    RenameForSingleSign(ref name, ref lName, ref sign);
                }
                else if (name.Contains("_"))
                {
                    string sign = "_";
                    RenameForSingleSign(ref name, ref lName, ref sign);
                } 
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error RenameContactName: " + ex.Message,
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
            return name;
        }

        /// <summary>
        /// Rename name for specified sign
        /// </summary>
        /// <param name="name">source name</param>
        /// <param name="lName">last name first letter</param>
        /// <param name="sign">specified sign</param>
        private static void RenameForSingleSign(ref string name, ref string lName, ref string sign)
        {
            lName = name.Substring(name.IndexOf(sign) + 1, 1);
            lName = lName.ToUpper();
            string f2Name = lName + name.Substring(name.IndexOf(sign) + 2, name.Length - name.IndexOf(sign) - 2);
            string l2Name = name.Substring(0, name.IndexOf(sign));
            name = l2Name + " " + f2Name;
        }
       
        /// <summary>
        /// Read subfolder for test and to count folder
        /// </summary>
        /// <param name="nameFolder">source folder</param>
        private void ReadSubfolderTest(MSOutlook.MAPIFolder nameFolder)
        {
            if (nameFolder.DefaultMessageClass == defaultNameClass)
                countAllmail += nameFolder.Items.Count;
            ++sourceFolderToSearchCount;
            Console.WriteLine(nameFolder.FullFolderPath);

            if (nameFolder.Folders.Count > 0)
                //recursive call
                foreach (MSOutlook.MAPIFolder nFolder in nameFolder.Folders)
                     ReadSubfolderTest(nFolder);
        }

        /// <summary>
        /// Read folder for test and count folders
        /// </summary>
        public void ReadAllFolderRecursiveTest()
        {
            countAllmail = folder.Items.Count;
            sourceFolderToSearchCount = 1;
            Console.WriteLine(folder.FullFolderPath);

            if (folder.Folders.Count > 0)
                foreach (MSOutlook.MAPIFolder fol2 in folder.Folders)
                    ReadSubfolderTest(fol2);
        }

        /// <summary>
        /// Add contact from source folder
        /// </summary>
        /// <param name="folder2">source folder</param>
        private void AddFromFolder(MSOutlook.MAPIFolder folder2)
        {
            if (cancelThread) return;

            this.UIThread(() => this.LB_result.Items.Add("Start search contacts for folder: " + folder2.FullFolderPath));

            if (folder2.DefaultMessageClass == defaultNameClass)
                AddContactsForOneFolder(folder2);

            foreach (MSOutlook.MAPIFolder oFolder2 in folder2.Folders)
            {
                if (oFolder2.DefaultMessageClass == defaultNameClass)
                    AddContactsForOneFolder(oFolder2);

                if (oFolder2.Folders.Count > 0)
                    //recursive call
                    foreach (MSOutlook.MAPIFolder oFolder3 in folder2.Folders)
                        AddFromFolder(oFolder3);
            }
        }

        /// <summary>
        /// Add contact from source folder (root selected folder)
        /// </summary>
        /// <param name="allSubFolders">recursive subfolders</param>
        /// <returns>true if all success</returns>
        private bool AddContacts(bool allSubFolders)
        {
            bool result = false;

            if (cancelThread) return true;

            try
            {
                if (!folder.FullFolderPath.Substring(2, folder.FullFolderPath.Length - 2).Contains("\\"))
                {
                    this.UIThread(() => this.LB_result.Items.Add("Cannot continue becouse selected folder is root folder\nPlease select source folder again, correctly"));
                    return false;
                }

                if (allSubFolders)
                {
                    this.UIThread(() => this.LB_result.Items.Add("Start search contacts for folder: " + folder.FullFolderPath));

                    if (folder.Folders.Count > 0)
                        AddFromFolder(folder);
                    else
                        AddContactsForOneFolder(folder);
                }
                else
                {
                    this.UIThread(() => this.LB_result.Items.Add("Start search contacts for folder: " + folder.FullFolderPath));
                    AddContactsForOneFolder(folder);                    
                }
                
                result = true;
            }
            catch (Exception ex)
            {
                this.UIThread(() => this.LB_result.Items.Add("Stop search contacts"));
                if (!allSubFolders)
                {
                    MessageBox.Show("Error AddContacts: " + ex.Message,
                                    "Error",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                }
                else
                {
                    string message = "Error AddContacts: " + ex.Message;
                    this.UIThread(() => this.LB_result.Items.Add(message));
                }
            }

            return result;
        }

        /// <summary>
        /// Method to add contact call from BW_worker thread
        /// </summary>
        private void AddContact()
        {
            sourceEmailCount = 0;
            bool allSubfolders = CB_selectedFolderWithAllSubfolders.Checked;
            if (!allSubfolders)
                countAllmail = folder.Items.Count;
            else
                ReadAllFolderRecursiveTest();

            if (RB_sendFolder.Checked)
                AddContacts(allSubfolders);
            else if (RB_receiveFolder.Checked)
                AddContacts(allSubfolders);
            if (selectedFolder)
            {
                //reset selectedFolder default value
                selectedFolder = false;
                AddContacts(allSubfolders);
            }
        }

        /// <summary>
        /// Event for click add contact button
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void B_addContacts_Click(object sender, EventArgs e)
        {
            PB_progress.Value = 0;
            BW_worker.RunWorkerAsync();
        }

        /// <summary>
        /// Event for click clear result list
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void B_clearUpperList_Click(object sender, EventArgs e)
        {
            ClearListResult();
        }

        /// <summary>
        /// Function call from clear list result button
        /// </summary>
        public void ClearListResult()
        {
            LB_result.Items.Clear();
        }

        /// <summary>
        /// Helping function for test adding lines to LB_result list
        /// </summary>
        /// <param name="listLines">initial list to add</param>
        public void SetLinesToLB_result(List<string> listLines)
        {
            foreach (string oLine in listLines)
                LB_result.Items.Add(oLine);   
        }

        /// <summary>
        /// Helping function to test for get actual items line from list result
        /// </summary>
        /// <returns>return items list lines</returns>
        public List<string> GetLinesFromLB_result()
        {
            List<string> listLines = new List<string>();

            if (LB_result.Items.Count > 0)
                foreach (string oLine in LB_result.Items)
                    listLines.Add(oLine);

            return listLines;
        }

        /// <summary>
        /// Event for click browse folder
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void B_otherFolder_Click(object sender, EventArgs e)
        {
            bool selectedReceiveFolder = false;
            if (RB_receiveFolder.Checked) selectedReceiveFolder = true;
            if (RB_sendFolder.Checked) selectedReceiveFolder = false;
            RB_sendFolder.Checked = false;
            RB_receiveFolder.Checked = false;

            if (folder != null) folder = null;
            folder = oApp.Session.PickFolder() as MSOutlook.MAPIFolder;
            if (folder != null)
            {
                if (folder.DefaultMessageClass != defaultNameClass)
                {
                    MessageBox.Show("If you want continue, select select correctly folder!",
                                    "Wrong choosed folder",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Stop);
                    RestoreSelected(selectedReceiveFolder);
                    return;
                }
                else
                {
                    selectedFolder = true;   
                    DialogResult dr2 = MessageBox.Show("Import contacts from:\nSENDER address mail click (Yes) or\nfrom all RECIPIENS mail address click (No)",
                                                  "Question choosed contacts source",
                                                  MessageBoxButtons.YesNo,
                                                  MessageBoxIcon.Question);
                    if (dr2 == DialogResult.Yes)
                        senderContactsFromSelectedFolder = true;
                    else
                        senderContactsFromSelectedFolder = false;

                    L_folderInfo.Text = folder.FullFolderPath;                        
                }
            }
            else
            {
                MessageBox.Show("Not selected source folder!\nPlease select again correclty",
                                "Wrong choosed folder",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                RestoreSelected(selectedReceiveFolder);
                return;
            }            
        }

        /// <summary>
        /// Helping function for restoring early selected RadioButton boxes
        /// </summary>
        /// <param name="selectedReceiveFolder">early selected RadioButton checked</param>
        private void RestoreSelected(bool selectedReceiveFolder)
        {
            if (selectedReceiveFolder)
            {
                RB_sendFolder.Checked = false;
                RB_receiveFolder.Checked = true;
            }
            else
            {
                RB_sendFolder.Checked = true;
                RB_receiveFolder.Checked = false;
            }
        }

        /// <summary>
        /// Event for first running program
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void F_Main_Load(object sender, EventArgs e)
        {
            InitializeOutlook();

            bool result = true;
            if (oApp == null) result = false;
            if (oNS == null) result = false;
            if (!result)
            {
                MessageBox.Show("Cannot initalize Outlook, break program!",
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                Application.Exit();
            }

            InitializeFolderSentMail();
            L_folderInfo.Text = folder.FullFolderPath;

            ////test only
            //TestNUnit test = new TestNUnit();
            //test.CreateForm();
            //test.RenameContactsName();

            //test.InitializeOutlook();
            //test.InitializeFolderInboxMail();
            //test.GetFirstContactInfoForFirstMailInboxMail();
            
            //test.AddItemsToListResult();
            ////test.ClearListResult();
        }

        /// <summary>
        /// Initialize Outlook, wrapping to MAPI MS Outlook
        /// </summary>
        public void InitializeOutlook()
        {
            oApp = new MSOutlook.Application();
            oNS = oApp.GetNamespace("MAPI");
        }

        /// <summary>
        /// Initialize check default send mail folder
        /// </summary>
        public void InitializeFolderSentMail()
        {
            if (folder != null) folder = null;
            folder = oNS.GetDefaultFolder(MSOutlook.OlDefaultFolders.olFolderSentMail) as MSOutlook.MAPIFolder;
        }

        /// <summary>
        /// Initialize check default inbox mail folder
        /// </summary>
        public void InitializeFolderInboxMail()
        {
            if (folder != null) folder = null;
            folder = oNS.GetDefaultFolder(MSOutlook.OlDefaultFolders.olFolderInbox) as MSOutlook.MAPIFolder;
        }

        /// <summary>
        /// Helping function for read first contact from actual contact from MS Outlook
        /// to diagnostic test
        /// </summary>
        public void GetFirstContactInfo()
        {
            if (folder.Items.Count > 0)
            {
                MSOutlook.MailItem mail = folder.Items.GetFirst() as MSOutlook.MailItem;
                Console.WriteLine("Sender name = " + mail.SenderName);
                Console.WriteLine("Sender mail = " + mail.SenderEmailAddress);
            }
            else
                Console.WriteLine(folder.FullFolderPath + " not contains any mail!");
        }

        /// <summary>
        /// Event for changed selected RadioButton send folder
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void RB_sendFolder_CheckedChanged(object sender, EventArgs e)
        {
            InitializeFolderSentMail();
            L_folderInfo.Text = folder.FullFolderPath;
        }

        /// <summary>
        /// Event for changed selected RadioButton inbox folder
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void RB_receiveFolder_CheckedChanged(object sender, EventArgs e)
        {
            InitializeFolderInboxMail();
            L_folderInfo.Text = folder.FullFolderPath;
        }

        /// <summary>
        /// Event for click abort button, abort import contacts
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void B_abort_Click(object sender, EventArgs e)
        {
            cancelThread = true;
            BW_worker.CancelAsync();
        }

        /// <summary>
        /// Event for start imported contacts
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void BW_worker_DoWork(object sender, DoWorkEventArgs e)
        {
            cancelThread = false;
            AddContact();
        }

        /// <summary>
        /// Event for progress BW_Worker thread
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void BW_worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.UIThread(() => this.PB_progress.Value = e.ProgressPercentage);
        }

        /// <summary>
        /// Event for finished imported contacts, or abort imported
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void BW_worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            PB_progress.Value = PB_progress.Maximum;
            MessageBox.Show("Finished adding Contacts",
                            "Add Contacts result",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

        /// <summary>
        /// Event for click homepage project link
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void LL_projectWebsite_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://addcontactsfrommail.codeplex.com/");
        }

        /// <summary>
        /// Event for closing program
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void F_Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (oApp != null)
            {
                //free COM object
                Marshal.ReleaseComObject(oNS);
                GC.SuppressFinalize(oNS);
                oNS = null;

                Marshal.ReleaseComObject(oApp);
                GC.SuppressFinalize(oApp);
                oApp = null;
            }
        }
    }

    //from http://stackoverflow.com/questions/661561/how-to-update-gui-from-another-thread-in-c
    public static class ControlExtensions
    {
        /// <summary>
        /// Executes the Action asynchronously on the UI thread, does not block execution on the calling thread.
        /// </summary>
        /// <param name="control"></param>
        /// <param name="code"></param>
        public static void UIThread(this Control @this, Action code)
        {
            if (@this.InvokeRequired)
                @this.BeginInvoke(code);
            else
                code.Invoke();
        }
    }
}
