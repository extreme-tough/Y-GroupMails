using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace Y_GroupMails
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            string memberURL,targetURL;
            string profileURL, email1, email2, username="";
            int nStart,recCount;
            string tempEmail;
            DateTime navStarted;

            OleDbConnection oCon;
            OleDbCommand oCmd;

            StatusText.Clear();

            StatusText.AppendText(DateTime.Now +  " : Opening database..." + Environment.NewLine);

            string path = Application.StartupPath;
            oCon = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + path + @"\data.mdb");
            oCon.Open();

            StatusText.AppendText(DateTime.Now +  " : Deleting existing records..." + Environment.NewLine);
            oCmd = new OleDbCommand();
            oCmd.Connection = oCon;
            oCmd.CommandText = "DELETE FROM Members";
            //oCmd.ExecuteNonQuery();

            webBrowser1.ScriptErrorsSuppressed = true;
            webBrowser2.ScriptErrorsSuppressed = true;

            string CSVPath = Application.StartupPath + @"\output.csv";
            string MDBPath = Application.StartupPath + @"\output.mdb";


            File.WriteAllText(CSVPath, "Name,Email1,Email2"+Environment.NewLine);
            navStarted = DateTime.Now;
            StatusText.AppendText(DateTime.Now +  " : Trying to login..." + Environment.NewLine);
            webBrowser1.Navigate("https://login.yahoo.com/config/login");
            while (webBrowser1.ReadyState!=WebBrowserReadyState.Complete)
            {
                Application.DoEvents();
                if (DateTime.Now >= navStarted.AddSeconds(150))
                {
                    webBrowser1.Stop();
                    break;
                }
            }
            StatusText.AppendText(DateTime.Now +  " : Login in progress..." + Environment.NewLine);
            webBrowser1.Document.GetElementById("username").SetAttribute("value", LoginID.Text);
            webBrowser1.Document.GetElementById("passwd").SetAttribute("value", Password.Text);
            webBrowser1.Document.GetElementById(".save").InvokeMember("click");

            while (!webBrowser1.Url.ToString().Contains("my.yahoo.com"))
            {
                Application.DoEvents();
            }
            while (webBrowser1.ReadyState != WebBrowserReadyState.Complete)
            {
                try
                {
                    if (webBrowser1.Document.Body.InnerText.Contains("Hi, "))
                        break;
                }
                catch { }
                Application.DoEvents();
                if (DateTime.Now >= navStarted.AddSeconds(150))
                {
                    webBrowser1.Stop();
                    break;
                }
            }
            StatusText.AppendText(DateTime.Now +  " : Login successful" + Environment.NewLine);


            if (GroupURL.Text.EndsWith("/"))
                memberURL = GroupURL.Text + "members";
            else
                memberURL = GroupURL.Text + "/members";

            nStart = 1;

            while (true)
            {
                profileURL = "";
                recCount = 0;
                StatusText.AppendText(DateTime.Now +  " : Starting from entry... " + nStart.ToString() + Environment.NewLine);
                targetURL = memberURL + "?start=" + nStart;
                navStarted = DateTime.Now;
                webBrowser1.Navigate(targetURL);
                while (webBrowser1.ReadyState != WebBrowserReadyState.Complete)
                {
                    Application.DoEvents();
                    if (DateTime.Now >= navStarted.AddSeconds(150))
                    {
                        webBrowser1.Stop();
                        break;
                    }

                }
                
                foreach(HtmlElement oElem in webBrowser1.Document.GetElementsByTagName("TD"))
                {
                    if (oElem.OuterHtml.StartsWith("\r\n<TD class=info"))
                    {
                        //Identified a new name
                        username = "";
                        if (oElem.GetElementsByTagName("A").Count == 0) break;
                        profileURL = oElem.GetElementsByTagName("A")[0].GetAttribute("href");
                        recCount++;
                        StatusText.AppendText(DateTime.Now +  " : Collecting name from " + profileURL + Environment.NewLine);
                        navStarted = DateTime.Now;
                        webBrowser2.Navigate(profileURL);
                        while (webBrowser2.ReadyState != WebBrowserReadyState.Complete)
                        {
                            Application.DoEvents();
                            if (DateTime.Now >= navStarted.AddSeconds(150))
                            {
                                webBrowser2.Stop();
                                break;
                            }
                        }
                        foreach (HtmlElement oElemIn in webBrowser2.Document.GetElementsByTagName("DIV"))
                        {
                            if (oElemIn.OuterHtml.StartsWith("\r\n<DIV class=relation"))
                            {
                                username = oElemIn.Children[0].InnerText;
                            }
                        }
                    }


                    if (oElem.OuterHtml.StartsWith("\r\n<TD class=\"email ygrp-nowrap"))
                    {
                        //Identified a new name
                        if (!oElem.GetElementsByTagName("A")[0].InnerText.Contains("@"))
                            tempEmail = oElem.GetElementsByTagName("A")[0].GetAttribute("title");
                        else
                            tempEmail = oElem.GetElementsByTagName("A")[0].InnerText;

                        tempEmail = tempEmail.Substring(0, tempEmail.IndexOf("@"));
                        email1 = tempEmail + "@yahoo.com.br";
                        email2 = tempEmail + "@yahoo.com";
                        File.AppendAllText(CSVPath, username + "," + email1 + "," + email2+Environment.NewLine);
                        StatusText.AppendText(DateTime.Now +  " : Inserting records to db... " + Environment.NewLine);
                        if (StatusText.Lines.Length > 100) StatusText.Clear();
                        oCmd.CommandText = "INSERT INTO Members(MemberName,Email1,Email2) VALUES ('" + username + "','" + email1 + "','" + email2 + "')";
                        try
                        {
                            oCmd.ExecuteNonQuery();
                        }
                        catch
                        {
                            try
                            {
                                oCmd.CommandText = "UPDATE Members SET MemberName='" + username + "' WHERE email1 = '" + email1 + "'";
                                oCmd.ExecuteNonQuery();
                            }
                            catch
                            {
                                StatusText.AppendText("Error in updating");
                            }
                        }

                    }
                }

                if (recCount == 0)
                    break;
                
                nStart += 10;
            }
            MessageBox.Show("Process complete");
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
