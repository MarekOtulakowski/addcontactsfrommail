#region UsingDirectives
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using MSOutlook = Microsoft.Office.Interop.Outlook;
#endregion

namespace addcontactsfrommail
{
    //reference to
    //NUnit (net framework 3,5) launchpad.net/nunit-3.0/trunk/2.9.3/+download/NUnit-2.9.3.msi

    //GUI
    //NUnit GUI from http://launchpad.net/nunitv2/2.5/2.5.10/+download/NUnit-2.5.10.11092.msi
    //configuration in VS2008 express http://www.dijksterhuis.org/setting-up-nunit-for-c-unit-testing-with-visual-studio-c-express-2008/

    /// <summary>
    /// Main class for all test
    /// </summary>
    [TestFixture]
    public class TestNUnit
    {
        #region Variables
        MainWindow form;
        #endregion

        /// <summary>
        /// On start test initial form
        /// </summary>
        [TestFixtureSetUp]
        public void CreateForm()
        {
            form = new MainWindow();
            Assert.AreNotEqual(null, form);
        }

        /// <summary>
        /// test for adding correctly result list
        /// </summary>
        [Test]
        public void AddItemsToListResult()
        {
            List<string> listLines = new List<string>();
            listLines.Add("lines one");
            listLines.Add("lines two");
            listLines.Add("lines three");
            form.SetLinesToLB_result(listLines);
            int linesCount = form.GetLinesFromLB_result().Count;

            Assert.AreEqual(3, linesCount);
        }

        /// <summary>
        /// test for clear corrctly result list
        /// </summary>
        [Test]
        public void ClearListResult()
        {
            form.ClearListResult();
            int linesCount = form.GetLinesFromLB_result().Count;
            Assert.AreEqual(0, linesCount);
        }

        /// <summary>
        /// test for initialize wrapper for MS Outlook
        /// </summary>
        [Test]
        public void InitializeOutlook()
        {
            form.InitializeOutlook();

            object oApp = form.GetMSOutlookApplication();
            object oNS = form.GetMSOutlookNamespace();
            Assert.AreNotEqual(null, oApp);
            Assert.AreNotEqual(null, oNS);
        }

        /// <summary>
        /// test for initialize default folder send mail
        /// </summary>
        /// <returns>folder send mail</returns>
        public MSOutlook.MAPIFolder InitializeFolderSentMail()
        {
            InitializeOutlook();
            form.InitializeFolderSentMail();
            MSOutlook.MAPIFolder oFolder = form.GetMSOutlookFolder();
            Assert.AreNotEqual(null, oFolder);
            return oFolder;
        }

        /// <summary>
        /// test for initialize default folder inbox mail
        /// </summary>
        /// <returns>folder inbox mail</returns>
        public MSOutlook.MAPIFolder InitializeFolderInboxMail()
        {
            InitializeOutlook();
            form.InitializeFolderInboxMail();
            MSOutlook.MAPIFolder oFolder = form.GetMSOutlookFolder();
            Assert.AreNotEqual(null, oFolder);
            return oFolder;
        }

        /// <summary>
        /// test for correctly read default send mail info
        /// </summary>
        [Test]
        public void DisplayFolderSentMailInfo()
        {
            MSOutlook.MAPIFolder oFolder = InitializeFolderSentMail();
            if (oFolder != null)
            {
                //info for test
                Console.WriteLine("Store ID = " + oFolder.StoreID);
                Console.WriteLine("ID = " + oFolder.EntryID);
                Console.WriteLine("Folder path = " + oFolder.FolderPath);
                Console.WriteLine("Full folder path = " + oFolder.FullFolderPath);
                Console.WriteLine("Folder name = " + oFolder.Name);
                Console.WriteLine("Items folder count = " + oFolder.Items.Count);
            }
            else
            {
                Console.WriteLine("Folder is null");
            }
            Assert.AreNotEqual(null, oFolder);
        }

        /// <summary>
        /// test for correctly read default inbox mail info
        /// </summary>
        [Test]
        public void DisplayFolderInboxMailInfo()
        {
            MSOutlook.MAPIFolder oFolder = InitializeFolderInboxMail();
            if (oFolder != null)
            {
                //info for test
                Console.WriteLine("Store ID = " + oFolder.StoreID);
                Console.WriteLine("ID = " + oFolder.EntryID);
                Console.WriteLine("Folder path = " + oFolder.FolderPath);
                Console.WriteLine("Full folder path = " + oFolder.FullFolderPath);
                Console.WriteLine("Folder name = " + oFolder.Name);
                Console.WriteLine("Items folder count = " + oFolder.Items.Count);
            }
            else
            {
                Console.WriteLine("Folder is null");
            }
            Assert.AreNotEqual(null, oFolder);
        }

        /// <summary>
        /// test for correctly read first contact from MS Outlook default send mail folder
        /// </summary>
        [Test]
        public void GetFirstContactInfoForFirstMailSenderMail()
        {
            InitializeOutlook();
            InitializeFolderSentMail();
            form.GetFirstContactInfo();
            MSOutlook.MailItem mail = form.GetMSMail();
            Assert.AreNotEqual(null, mail);
        }

        /// <summary>
        /// test for correctly read first contact from MS Outlook default inbox mail folder
        /// </summary>
        [Test]
        public void GetFirstContactInfoForFirstMailInboxMail()
        {
            InitializeOutlook();
            InitializeFolderInboxMail();
            form.GetFirstContactInfo();
            MSOutlook.MailItem mail = form.GetMSMail();
            Assert.AreNotEqual(null, mail);
        }

        /// <summary>
        /// test for checking correctly changed name contacts
        /// </summary>
        [Test]
        public void RenameContactsName()
        {
            Assert.AreEqual("John Smith",
                            form.RenameContactConditions("john.smith@domail.com"));
            Assert.AreEqual("John Montana",
                            form.RenameContactConditions("john montana"));
            Assert.AreEqual("Jerry Kovalsky",
                            form.RenameContactConditions("jeRRy.KovaLskY"));
            Assert.AreEqual("Jerry Kovalsky",
                            form.RenameContactConditions("jeRRy-KovaLskY"));
            Assert.AreEqual("Jerry Kovalsky",
                            form.RenameContactConditions("jeRRy_KovaLskY"));
            Assert.AreEqual("a.b.c",
                            form.RenameContactConditions("a.b.c@gmail.com"));
            Assert.AreEqual("Jerry Novak",
                            form.RenameContactConditions("jerry-Novak@gmail.com"));
            Assert.AreEqual("Natally Herman",
                            form.RenameContactConditions("NataLLY_HerMaN@example.com"));
            Assert.AreEqual("Allan Test",
                            form.RenameContactConditions("\"Allan Test\""));
        }

        /// <summary>
        /// test for read all folders for initial directory (default inbox mail)
        /// </summary>
        [Test]
        public void ReadRecursiveFullNameFolders()
        {
            InitializeOutlook();
            InitializeFolderInboxMail();
            form.ReadAllFolderRecursiveTest();
            Assert.AreEqual(true, true);
            Console.WriteLine("All folder count: " + form.GetSourceFolderToSearchCount().ToString());
        }
    }
}
