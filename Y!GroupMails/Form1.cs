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

        private int getLastTable()
        {
            OleDbConnection oCon;
            OleDbCommand oCmd;
            oCon = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + Application.StartupPath + @"\data.mdb");
            oCon.Open();

            

            int tableNum = 1;

            while (true)
            {
                oCmd = new OleDbCommand();
                oCmd.Connection = oCon;

                oCmd.CommandText = "SELECT * FROM Table" + tableNum;
                try
                {
                    oCmd.ExecuteReader();
                }
                catch
                {
                    oCon.Close();
                    return tableNum-1;
                }
                tableNum = tableNum + 1;
                
            }
        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            string memberURL,targetURL;
            string profileURL, email1, email2, username="";
            int nStart,recCount;
            string tempEmail;
            string groupName;


            string url = GroupURL.Text.EndsWith("/") ? GroupURL.Text.Substring(0, GroupURL.Text.Length - 1) : GroupURL.Text;
            string[] urlParts = url.Split(new char[] { '/' }, StringSplitOptions.None);
            groupName =  urlParts[urlParts.Length - 1];

            Uri n = new Uri(GroupURL.Text);
            int k = n.Fragment.Length;

            DateTime navStarted;

            OleDbConnection oCon;
            OleDbCommand oCmd;

            string path = Application.StartupPath;
            oCon = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + path + @"\data.mdb");
            oCon.Open();


            StatusText.Clear();

            int LastTable = getLastTable();


                StatusText.AppendText(DateTime.Now + " : Creating new table..." + Environment.NewLine);

                oCmd = new OleDbCommand();
                oCmd.Connection = oCon;
                oCmd.CommandText = "CREATE TABLE Table" + (LastTable+1) + "  ( [ID] AUTOINCREMENT Primary key, GroupName text(70), Membername Text(70), Email1  Text(90), Email2 Text(90))  ";
                oCmd.ExecuteNonQuery();
                oCmd.CommandText = "CREATE UNIQUE INDEX Email ON Table" + (LastTable + 1)  + "(Email1)";
                oCmd.ExecuteNonQuery();
            

            StatusText.AppendText(DateTime.Now +  " : Opening database..." + Environment.NewLine);




            webBrowser1.ScriptErrorsSuppressed = true;
            webBrowser2.ScriptErrorsSuppressed = true;

            string CSVPath = Application.StartupPath + @"\output.csv";
            string MDBPath = Application.StartupPath + @"\output.mdb";


            File.WriteAllText(CSVPath, "GroupName, Name,Email1,Email2"+Environment.NewLine);
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

            oCmd = new OleDbCommand();
            oCmd.Connection = oCon;

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
                                break;
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
                        File.AppendAllText(CSVPath, groupName + "," + username + "," + email1 + "," + email2 + Environment.NewLine);
                        StatusText.AppendText(DateTime.Now +  " : Inserting records to db... " + Environment.NewLine);
                        if (StatusText.Lines.Length > 100) StatusText.Clear();

                        if (LastTable == 0)
                        {
                            oCmd.CommandText = "INSERT INTO Table1(GroupName,MemberName,Email1,Email2) VALUES ('" + groupName + "','" + username + "','" + email1 + "','" + email2 + "')";
                            try
                            {
                                oCmd.ExecuteNonQuery();
                            }
                            catch (Exception ex)
                            {
                                string a = "";
                            }
                        }

                        oCmd.CommandText = "INSERT INTO Members(GroupName, MemberName,Email1,Email2) VALUES ('" + groupName + "','" + username + "','" + email1 + "','" + email2 + "')";
                        try
                        {
                            oCmd.ExecuteNonQuery();
                            //If succcess it is a new record
                            try
                            {
                                oCmd.CommandText = "INSERT INTO Table" + (LastTable + 1) + "(GroupName, MemberName,Email1,Email2) VALUES ('" + groupName + "','" + username + "','" + email1 + "','" + email2 + "')";
                                oCmd.ExecuteNonQuery();
                            }
                            catch
                            {
                                StatusText.AppendText("Serius Error in updating");
                            }
                        }
                        catch (Exception ex)
                        {
                            //Insertion failed
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
