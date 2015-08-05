using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace MoveRule
{
    public partial class ThisAddIn
    {
        Outlook.Folder folder;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            folder = Application.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
            folder.BeforeItemMove += Folder_BeforeItemMove;
        }

        private bool isRulepresent(Outlook.Rules rules,string mailaddress)
        {
            foreach (Outlook.Rule r in rules)
            {
                if (r.RuleType == Outlook.OlRuleType.olRuleReceive)
                {
                    Outlook.RuleConditions rcs = r.Conditions;
                    foreach (Outlook.RuleCondition rc in rcs)
                    {
                        if (rc.ConditionType == Outlook.OlRuleConditionType.olConditionFrom)
                        {
                            Outlook.ToOrFromRuleCondition tf = r.Conditions.From;
                            Outlook.Recipients recipients = tf.Recipients;
                            foreach (Outlook.Recipient recipient in recipients)
                            {
                                if (recipient.Address.Equals(mailaddress))
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
            }
            return false;
        }

        private void Folder_BeforeItemMove(object Item, Outlook.MAPIFolder MoveTo, ref bool Cancel)
        {
            Outlook.MailItem mi = Item as Outlook.MailItem;
            Outlook.Rules rules = Application.Session.DefaultStore.GetRules();

            

            if (!isRulepresent(rules,mi.SenderEmailAddress))
            {
                String Message = "Sollen Nachrichten des Absenders " + mi.SenderEmailAddress + " zukünftig im Ordner "+MoveTo.Name+" abgelegt werden?";
                String Rulename = "Nachrichten von " + mi.SenderEmailAddress + " in Ordner " + MoveTo.Name + " verschieben";
                DialogResult res = MessageBox.Show(Message, "MoveRule", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    Outlook.Rule rule = rules.Create(Rulename, Outlook.OlRuleType.olRuleReceive);
                    rule.Conditions.From.Recipients.Add(mi.SenderEmailAddress);
                    rule.Conditions.From.Recipients.ResolveAll();
                    rule.Conditions.From.Enabled = true;

                    rule.Actions.MoveToFolder.Folder = MoveTo;
                    rule.Actions.MoveToFolder.Enabled = true;
                    rules.Save(true);
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    ausgeführt werden muss, wenn Outlook geschlossen wird, informieren Sie sich unter http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
