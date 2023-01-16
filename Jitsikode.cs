using System;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private Office.CommandBar buttonBar;
        private Office.CommandBarPopup newMenuBar;
        private Office.CommandBarButton randomLinkButton;
        private string domain = "https://meet.ballerup.dk/";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CreateRibbon();
        }

        private void CreateRibbon()
        {
            buttonBar = Application.ActiveExplorer().CommandBars.ActiveMenuBar;
            newMenuBar = (Office.CommandBarPopup)buttonBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, true);
            newMenuBar.Caption = "Random Link";
            randomLinkButton = (Office.CommandBarButton)newMenuBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
            randomLinkButton.Caption = "Nyt Jitsi-møde";
            randomLinkButton.FaceId = 939;
            randomLinkButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(RandomLinkButton_Click);
        }

        private void RandomLinkButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            Random rnd = new Random();
            string endLink = "";
            for (int i = 0; i < 10; i++)
            {
                endLink += rnd.Next(0, 10);
            }
            string randomLink = domain + endLink;
            mailItem.HTMLBody = "Du er hermed inviteret til et Jitsi-møde: " + randomLink;
            mailItem.Subject = "Invitering til Jitsi-møde";
            mailItem.Display(false);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_
