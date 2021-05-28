using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Configuration;
using System.Data.SqlClient;
using System.Net.Mail;
namespace BC_SFTP_LOG
{

    class Program
    {

        static void Main(string[] args)
        {
            string status = "Success";
            var inputDirectory = new DirectoryInfo(@"E:\Interface\outbound\Log1");
            var myFile = inputDirectory.GetFiles().OrderByDescending(f => f.LastWriteTime).First();
            string Fpath = @"E:\Interface\outbound\Log\" + myFile.ToString();
            string mailBody = "";
            String EX_Occured = "N";
            SqlConnection cnn1;
            string connectionString1 = null;
            string sql = null, sqlEX = null;
            connectionString1 = ConfigurationManager.ConnectionStrings["Cnn"].ConnectionString;
            cnn1 = new SqlConnection(connectionString1);
            cnn1.Open();
            sql = "select [FILE_NAME],[START_TIME],[END_TIME],[STATUS] from BC_BIW_SFTP_LOG where [UPDATED_DATETIME] IS NULL";
            SqlDataAdapter dscmd = new SqlDataAdapter(sql, cnn1);
            System.Data.DataTable ds_mail = new System.Data.DataTable();
            dscmd.Fill(ds_mail);
            if (ds_mail.Rows.Count > 0)
            {
                for (int i = 0; i < ds_mail.Rows.Count; i++)
                {
                    if (i == 0)
                        mailBody = "<table border=1><b><tr bgcolor='#81BEF7'><td>S.No</td><td>FILE_NAME</td><td>START_TIME</td><td>END_TIME</td><td>STATUS</td></tr></b><tr><td>" + (i + 1).ToString() + "</td><td>" + ds_mail.Rows[i][0].ToString() + "</td><td>" + ds_mail.Rows[i][1].ToString() + "</td><td>" + ds_mail.Rows[i][2].ToString() + "</td><td>" + ds_mail.Rows[i][3].ToString() + "</td></tr>";
                    else
                        mailBody = mailBody + "<tr><td>" + (i + 1).ToString() + "</td><td>" + ds_mail.Rows[i][0].ToString() + "</td><td>" + ds_mail.Rows[i][1].ToString() + "</td><td>" + ds_mail.Rows[i][2].ToString() + "</td><td>" + ds_mail.Rows[i][3].ToString() + "</td></tr>";
                }
                //    //sqlEX = "select INTERFACE_NAME,DISTRIBUTOR,SKU,INV_NUMBER,BILL_DATE,DELV_DATE,ERROR_DETAILS,[DATETIME] from [BC_I_EX_ERRORLOG_EXTRACT] where INTERFACE_NAME='BalanceStock' and [UPDATED_DATETIME] IS NULL";
                //    //SqlDataAdapter dscmdEX = new SqlDataAdapter(sqlEX, cnn);
                //    //DataTable ds_mailEX = new DataTable();
                //    //dscmdEX.Fill(ds_mailEX);
                //    //if (ds_mailEX.Rows.Count > 0)
                //    //{
                //    //    EX_Occured = "Y";
                //    //    exporttoexcel("BalanceStock");
                //    //}

                //    //if (EX_Occured == "Y")
                //    //    new I_MailUtility().SendErrorReportBIW(mailBody + "</table>", "CSDP-BIW interface - BalanceStock Scheduler Progress Details " + DateTime.Now, 1, ExcelFilePath, "", schedulerstart.ToString());
                //    //else
                //    //new I_MailUtility().SendErrorReportBIW(mailBody + "</table>", "CSDP Interface - DRS BalanceStock Scheduler Progress Details " + DateTime.Now, 1, ExcelFilePath, "", "6 AM");

                //    string s1Year1 = DateTime.Now.Year.ToString();
                //    string s1Month1 = DateTime.Now.Month.ToString().PadLeft(2, '0');
                //    string s1Day1 = DateTime.Now.Day.ToString().PadLeft(2, '0');
                //    string s1ErrorTime1 = s1Day1 + "-" + s1Month1 + "-" + s1Year1;
                //    DateTime dt = DateTime.Now;
                //    string from = System.Configuration.ConfigurationSettings.AppSettings["Mail_From"].ToString(),
                //        to = System.Configuration.ConfigurationSettings.AppSettings["Mail_To"].ToString(),
                //        copy = System.Configuration.ConfigurationSettings.AppSettings["Mail_Copy"].ToString(),

                //        //to = "Lal.Samansiri@unilever.com,Lavan.Harshanga@unilever.com,Gayan.Abeysingha@unilever.com,Tusira.Dilnath@unilever.com,Madura.Perera@unilever.com,Suraj.Perera@unilever.com,Chandana.Liyanage@unilever.com",
                //        //copy = "Amarnath.Gopinath@unilever.com,Raneesh.Rajeevan@unilever.com",

                //        subject = "SFTP Process Status Generated On: " + s1ErrorTime1, body, filePath;
                //    //string Attachment_path = System.Configuration.ConfigurationSettings.AppSettings["Attachment_path"].ToString();
                //    bool isHtmlBody;
                //    //string FileNameFileName = "SL-CSDP-BIW-Sal-Ext-" + dt.Day + dt.Month + dt.Year + dt.Hour + dt.Minute + dt.Second + ".csv";

                //    string server = "10.212.18.70";
                //    int mailPort = 25;

                //    string connectionString = null;
                //    //string sql = null, sqlEX = null;

                //    //string FileName = decimal.Parse(GetDataTable(sql).Rows[0]["SELL_FACTOR1"].ToString());

                //    string returnMsg = "";
                //    MailMessage mailMsg = new MailMessage();

                //    //set from Address
                //    MailAddress mailAddress = new MailAddress(from);
                //    mailMsg.From = mailAddress;
                //    //set to Adress
                //    mailMsg.To.Add(to);
                //    // Set Message subject
                //    mailMsg.Subject = subject;
                //    //set mail cc
                //    if (copy != "")
                //        mailMsg.CC.Add(copy);
                //    // Set Message Body
                //    mailMsg.IsBodyHtml = true;
                //    filePath = @"E:\Interface\outbound\Log\";
                //    DirectoryInfo dir = new DirectoryInfo(filePath);

                //    dir.Refresh();
                //    filePath = filePath + myFile;
                //    if (filePath != "")
                //    {
                //        Attachment attach3 = new Attachment(filePath, "application/vnd.ms-excel");
                //        mailMsg.Attachments.Add(attach3);
                //    }

                //    string v = subject.Replace("NGDMS Interface - ", "");
                //    string v1 = v.Replace(" Error", "");
                //    string htmlBody, htmlBody2, htmlBody3, htmlBody4;
                //    htmlBody = "<html>Dear Team,<br/><br/></html>";

                //    htmlBody2 = "<html>Please find the attachment of SFTP Process Status file and below details for your reference, which was taken on " + s1Day1 + "/" + s1Month1 + "/" + s1Year1 + "  at " + DateTime.Now.ToLongTimeString().ToString() + ".<br/><br/></html>";
                //    htmlBody4 = "<html><br/>Regards, <br/></html>CSDP Interface Support Team.<br/></html>";
                //    //<br/><br/>The information contained in this electronic message and any attachments to this message are intended for the exclusive use of the addressee(s) and may contain proprietary, confidential or privileged information. <br/>If you are not the intended recipient, you should not disseminate, distribute or copy this e-mail. Please notify the sender immediately and destroy all copies of this message and any attachments.<br/></html>";
                //    mailMsg.Body = htmlBody + "  " + htmlBody2 + mailBody + "</table>" + htmlBody4;


                //    //set message body format
                //    //mailMsg.IsBodyHtml = isHtmlBody;
                //    // System.IO.File.WriteAllText(@"c:\abc.xlsx", attachmentStream);
                //    //byte[] data = GetData(attachmentStream);


                //    //if (attachmentStream != null)//Define mail attachment.
                //    //{
                //    //    ms.Position = 0;//explicitly set the starting position of the MemoryStream
                //    //Attachment attach = new Attachment(filePath, "application/vnd.ms-excel");
                //    //mailMsg.Attachments.Add(attach);
                //    //}


                //    //set exchange server
                //    SmtpClient smtpClient = new SmtpClient(server);
                //    smtpClient.Send(mailMsg);
                //    mailMsg.Dispose();
                //    //return returnMsg;


                Insert_SFTPLog();
            }
            else
            {
                SFTPLOGREPORT();
            }
        }
        private static void SFTPLOGREPORT()
        {
            string filepath = System.Configuration.ConfigurationSettings.AppSettings["LogFile_Path"].ToString();
            var inputDirectory = new DirectoryInfo(filepath);
            var myFile = inputDirectory.GetFiles().OrderByDescending(f => f.LastWriteTime).First();

            string s1Year1 = DateTime.Now.Year.ToString();
            string s1Month1 = DateTime.Now.Month.ToString().PadLeft(2, '0');
            string s1Day1 = DateTime.Now.Day.ToString().PadLeft(2, '0');
            string s1ErrorTime1 = s1Day1 + "-" + s1Month1 + "-" + s1Year1;
            DateTime dt = DateTime.Now;
            string from = System.Configuration.ConfigurationSettings.AppSettings["Mail_From"].ToString(),
                to = System.Configuration.ConfigurationSettings.AppSettings["Mail_To"].ToString(),
                copy = System.Configuration.ConfigurationSettings.AppSettings["Mail_Copy"].ToString(),

                //to = "Lal.Samansiri@unilever.com,Lavan.Harshanga@unilever.com,Gayan.Abeysingha@unilever.com,Tusira.Dilnath@unilever.com,Madura.Perera@unilever.com,Suraj.Perera@unilever.com,Chandana.Liyanage@unilever.com",
                //copy = "Amarnath.Gopinath@unilever.com,Raneesh.Rajeevan@unilever.com",

                subject = "SFTP Process Status - Generated On: " + s1ErrorTime1;
            //string Attachment_path = System.Configuration.ConfigurationSettings.AppSettings["Attachment_path"].ToString();
            bool isHtmlBody;
       
            string server = System.Configuration.ConfigurationSettings.AppSettings["Mail_Server"].ToString();
            int mailPort = 25;

            string connectionString = null;
            
            string returnMsg = "";
            MailMessage mailMsg = new MailMessage();

            //set from Address
            MailAddress mailAddress = new MailAddress(from);
            mailMsg.From = mailAddress;
            //set to Adress
            mailMsg.To.Add(to);
            // Set Message subject
            mailMsg.Subject = subject;
            //set mail cc
            if (copy != "")
                mailMsg.CC.Add(copy);
            // Set Message Body
            mailMsg.IsBodyHtml = true;
       
            DirectoryInfo dir = new DirectoryInfo(filepath);

            dir.Refresh();
            filepath = filepath + myFile;
            if (filepath != "")
            {
                Attachment attach3 = new Attachment(filepath, "application/vnd.ms-excel");
                mailMsg.Attachments.Add(attach3);
            }

            string v = subject.Replace("NGDMS Interface - ", "");
            string v1 = v.Replace(" Error", "");
            string htmlBody, htmlBody2, htmlBody4;
            htmlBody = "<html>Dear Team,<br/><br/></html>";

            htmlBody2 = "<html>Please find the attachment of SFTP Process Status file and below details for your reference, which was taken on " + s1Day1 + "/" + s1Month1 + "/" + s1Year1 + "  at " + DateTime.Now.ToLongTimeString().ToString() + ".<br/><br/></html>";
            htmlBody4 = "<html><br/>Regards, <br/></html>CSDP Interface Support Team.<br/></html>";
            //<br/><br/>The information contained in this electronic message and any attachments to this message are intended for the exclusive use of the addressee(s) and may contain proprietary, confidential or privileged information. <br/>If you are not the intended recipient, you should not disseminate, distribute or copy this e-mail. Please notify the sender immediately and destroy all copies of this message and any attachments.<br/></html>";
            mailMsg.Body = htmlBody + "  " + htmlBody2 + "</table>" + htmlBody4;


            //set message body format
            //mailMsg.IsBodyHtml = isHtmlBody;
            // System.IO.File.WriteAllText(@"c:\abc.xlsx", attachmentStream);
            //byte[] data = GetData(attachmentStream);


            //if (attachmentStream != null)//Define mail attachment.
            //{
            //    ms.Position = 0;//explicitly set the starting position of the MemoryStream
            //Attachment attach = new Attachment(filePath, "application/vnd.ms-excel");
            //mailMsg.Attachments.Add(attach);
            //}


            //set exchange server
            SmtpClient smtpClient = new SmtpClient(server);
            smtpClient.Send(mailMsg);
            mailMsg.Dispose();


        }
        private static void Insert_SFTPLog()
        {
            //var lineCount = 0;
            string status = "Success";
            var inputDirectory = new DirectoryInfo(@"E:\Interface\outbound\Log");
            var myFile = inputDirectory.GetFiles().OrderByDescending(f => f.LastWriteTime).First();
            string Fpath = @"E:\Interface\outbound\Log\" + myFile.ToString();
            StreamReader reader = File.OpenText(Fpath);
            string line;
            string RecordType;
            string filename = "", starttime = "";
            // statusnew = "", endtimenew = ""; string stringToSearch = "-> remote"; string message = "", messagenew = "";, starttimenew = "", Filenanem = "", 
            while ((line = reader.ReadLine()) != null)
            {
                //lineCount++;
                //if (lineCount == 31)
                //{

                //    line = reader.ReadLine();
                //    string[] items1 = line.Split('/');
                //    message = items1[5].ToString();
                //    string input = message;
                //    var a = input.Split(' '); statusnew = a[a.Length - 1];
                //    if (statusnew == "") { Filenanem = a[a.Length - 2]; } else { Filenanem = a[a.Length - 4]; };
                //    if (statusnew == "OK") { statusnew = "SUCCESS"; } else { statusnew = "FAILURE"; }
                //    line = reader.ReadLine();
                //    bool containserr = line.Contains("ERROR:");
                //    if (containserr == true)
                //    {

                //        endtimenew = starttimenew;
                //        starttimenew = endtimenew;

                //    }
                //    else
                //    {
                //        string[] items22 = line.Split(' ');
                //        endtimenew = items22[13].ToString() + " " + items22[14].ToString();
                //        starttimenew = endtimenew;
                //    }
                //}
                //if (lineCount == 31)
                //{
                //    line = reader.ReadLine();

                //    bool containserr = line.Contains("ERROR:");
                //    if (containserr == true)
                //    {

                //        endtimenew = endtimenew;
                //        starttimenew = endtimenew;

                //    }
                //    else
                //    {
                //        line = reader.ReadLine(); string[] items1 = line.Split('/');
                //        endtimenew = items1[13].ToString() + " " + items1[14].ToString();
                //        starttimenew = endtimenew;
                //    }
                //}
                //if (lineCount == 32)
                //{

                //    line = reader.ReadLine();
                //    string[] items1 = line.Split('/');
                //    message = items1[5].ToString();
                //    string input = message;
                //    var a = input.Split(' '); statusnew = a[a.Length - 1];
                //    if (statusnew == "") { Filenanem = a[a.Length - 2]; } else { Filenanem = a[a.Length - 4]; };
                //    if (statusnew == "OK") { statusnew = "SUCCESS"; } else { statusnew = "FAILURE"; }
                //    line = reader.ReadLine();
                //    bool containserr = line.Contains("ERROR:");
                //    if (containserr == true)
                //    {

                //        endtimenew = starttimenew;
                //        starttimenew = endtimenew;

                //    }
                //    else
                //    {
                //        string[] items22 = line.Split(' ');
                //        endtimenew = items22[10].ToString() + " " + items22[11].ToString();
                //        //starttimenew = endtimenew;
                //    }
                //}
                //if (lineCount == 33)
                //{

                //    line = reader.ReadLine();
                //    string[] items1 = line.Split('/');
                //    message = items1[5].ToString();
                //    string input = message;
                //    var a = input.Split(' '); statusnew = a[a.Length - 1];
                //    if (statusnew == "") { Filenanem = a[a.Length - 2]; } else { Filenanem = a[a.Length - 4]; };
                //    if (statusnew == "OK") { statusnew = "SUCCESS"; } else { statusnew = "FAILURE"; }
                //    line = reader.ReadLine();
                //    bool containserr = line.Contains("ERROR:");
                //    if (containserr == true)
                //    {

                //        endtimenew = starttimenew;
                //        starttimenew = endtimenew;

                //    }
                //    else
                //    {

                //        string[] items22 = line.Split(' ');
                //        endtimenew = items22[10].ToString() + " " + items22[12].ToString();
                //        //starttimenew = endtimenew;
                //    }
                //}
                //if (lineCount == 34)
                //{

                //    line = reader.ReadLine();
                //    string[] items1 = line.Split('/');
                //    message = items1[5].ToString();
                //    string input = message;
                //    var a = input.Split(' '); statusnew = a[a.Length - 1];
                //    if (statusnew == "") { Filenanem = a[a.Length - 2]; } else { Filenanem = a[a.Length - 4]; };
                //    if (statusnew == "OK") { statusnew = "SUCCESS"; } else { statusnew = "FAILURE"; }
                //    line = reader.ReadLine();
                //    bool containserr = line.Contains("ERROR:");
                //    if (containserr == true)
                //    {

                //        endtimenew = starttimenew;
                //        starttimenew = endtimenew;

                //    }
                //    else
                //    {
                //        string[] items22 = line.Split(' ');
                //        endtimenew = items22[13].ToString() + " " + items22[14].ToString();
                //        //starttimenew = endtimenew;
                //    }
                //}
                //if (lineCount == 35)
                //{

                //    line = reader.ReadLine();
                //    string[] items1 = line.Split('/');
                //    message = items1[5].ToString();
                //    string input = message;
                //    var a = input.Split(' '); statusnew = a[a.Length - 1];
                //    if (statusnew == "") { Filenanem = a[a.Length - 2]; } else { Filenanem = a[a.Length - 4]; };
                //    if (statusnew == "OK") { statusnew = "SUCCESS"; } else { statusnew = "FAILURE"; }
                //    line = reader.ReadLine();
                //    bool containserr = line.Contains("ERROR:");
                //    if (containserr == true)
                //    {

                //        endtimenew = starttimenew;
                //        starttimenew = endtimenew;

                //    }
                //    else
                //    {
                //        string[] items22 = line.Split(' ');
                //        endtimenew = items22[13].ToString() + " " + items22[14].ToString();
                //        //starttimenew = endtimenew;
                //    }
                //}
                //if (lineCount == 36)
                //{

                //    line = reader.ReadLine();
                //    string[] items1 = line.Split('/');
                //    message = items1[5].ToString();
                //    string input = message;
                //    var a = input.Split(' '); statusnew = a[a.Length - 1];
                //    if (statusnew == "") { Filenanem = a[a.Length - 2]; } else { Filenanem = a[a.Length - 4]; };
                //    if (statusnew == "OK") { statusnew = "SUCCESS"; } else { statusnew = "FAILURE"; }
                //    line = reader.ReadLine();
                //    bool containserr = line.Contains("ERROR:");
                //    if (containserr == true)
                //    {

                //        endtimenew = starttimenew;
                //        starttimenew = endtimenew;

                //    }
                //    else
                //    {
                //        string[] items22 = line.Split(' ');
                //        endtimenew = items22[10].ToString() + " " + items22[11].ToString();
                //starttimenew = endtimenew;
                //    }
                //}
                //if (statusnew != "" && Filenanem != "" && endtimenew != "")
                //{
                //    SqlConnection cnn11;
                //    string connectionString11 = null;
                //    connectionString11 = ConfigurationManager.ConnectionStrings["Cnn"].ConnectionString;
                //    //connectionString = "data source=servername;initial catalog=databasename;user id=username;password=password;";
                //    cnn11 = new SqlConnection(connectionString11);
                //    cnn11.Open();

                //    string insert_reconcildata1 = "insert into BC_BIW_SFTP_LOG(FILE_NAME,START_TIME,END_TIME,STATUS) values('" + Filenanem + "','" + starttimenew + "','" + endtimenew + "','" + statusnew + "')";
                //    SqlCommand cmd1 = new SqlCommand(insert_reconcildata1, cnn11);
                //    cmd1.ExecuteNonQuery();
                //    cnn11.Close();
                //    starttimenew = endtimenew; statusnew = ""; Filenanem = ""; endtimenew = "";
                //}
                //int i = reader.ReadLine().Split('|').Length;
                //int j = i;

                string[] items = line.Split(','); 
                string endtime = "", message = "";
                //string endtime = "", message = ""; message = items[32].ToString();
                //int myInteger = int.Parse(items[1]); // Here's your integer.
                if (items.Length > 6)
                {


                    endtime = items[7].ToString();
                    message = items[10].ToString();

                    bool containserr = message.Contains("Error");
                    if (containserr == true)
                    {
                        SqlConnection cnn11;
                        string connectionString11 = null;
                        connectionString11 = ConfigurationManager.ConnectionStrings["Cnn"].ConnectionString;
                        //connectionString = "data source=servername;initial catalog=databasename;user id=username;password=password;";
                        cnn11 = new SqlConnection(connectionString11);
                        cnn11.Open();

                        string insert_reconcildata1 = "insert into BC_BIW_SFTP_LOG(FILE_NAME,START_TIME,END_TIME,STATUS) values('" + filename + "','" + starttime + "','" + endtime + "','" + message + "')";
                        SqlCommand cmd1 = new SqlCommand(insert_reconcildata1, cnn11);
                        cmd1.ExecuteNonQuery();
                        cnn11.Close();
                        status = "Failure";
                    }




                    bool contains = message.Contains("Task ended");
                    if (contains == true)
                    {
                        message = message;
                    }
                    else
                    {



                        bool contains5 = message.Contains("START SFTP Task");
                        if (contains5 == true)
                        {
                            starttime = items[6].ToString();
                            message = "";
                        }






                        bool contains1 = message.Contains("rendered");
                        if (contains1 == true)
                        {
                            bool contains2 = message.Contains("SL");
                            if (contains2 == true)
                            {
                                string source = message;
                                string split = @"push\";

                                // over-the-lazy-dog
                                string result = source.Substring(source.IndexOf(split) + split.Length);

                                filename = result;
                                message = "";
                            }
                        }
                        else
                        {
                            message = "";
                        }
                    }

                    // Now let's find the path.
                    if (message != "" && message != "message")
                    {

                        String St = message;


                        int pTo = St.LastIndexOf("SL");


                        //        //bool contains2 = message.Contains("SL");
                        //        //if (contains2 == true)
                        //        //{
                        //        //    string source = message;
                        //        //    string split = @"push\";

                        //        //    // over-the-lazy-dog
                        //        //    string result = source.Substring(source.IndexOf(split) + split.Length);

                        //        //    filename = result;
                        //        //}

                        string path = null;
                        SqlConnection cnn;
                        string connectionString = null;
                        connectionString = ConfigurationManager.ConnectionStrings["Cnn"].ConnectionString;
                        //connectionString = "data source=servername;initial catalog=databasename;user id=username;password=password;";
                        cnn = new SqlConnection(connectionString);
                        cnn.Open();
                        string insert_reconcildata = "insert into BC_BIW_SFTP_LOG(FILE_NAME,START_TIME,END_TIME,STATUS) values('" + filename + "','" + starttime + "','" + endtime + "','" + message + "')";
                        SqlCommand cmd = new SqlCommand(insert_reconcildata, cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                        bool contains3 = message.Contains("Task ended");
                        if (contains3 == true)
                        {
                            //string insert_reconcildata1 = "update BC_BIW_SFTP_LOG set FILE_NAME='" + filename + "' where UPDATED_DATETIME IS NULL and FILE_NAME='' ";
                            //SqlCommand cmd1 = new SqlCommand(insert_reconcildata1, cnn);
                            //cmd1.ExecuteNonQuery();
                            filename = "";
                            starttime = "";
                        }


                        }

                     }




                        //foreach (string item in items)
                        //{
                        //    if (item.StartsWith("item\\") && item.EndsWith(".ddj"))
                        //    {
                        //        path = item;
                        //    }
                        //}



                        // At this point, `myInteger` and `path` contain the values we want
                        // for the current line. We can then store those values or print them,
                        // or anything else we like.
                    }

                    string mailBody = "";
                    String EX_Occured = "N";
                    SqlConnection cnn1;
                    string connectionString1 = null;
                    string sql = null, sqlEX = null;
                    connectionString1 = ConfigurationManager.ConnectionStrings["Cnn"].ConnectionString;
                    cnn1 = new SqlConnection(connectionString1);
                    cnn1.Open();
                    sql = "select [FILE_NAME],[START_TIME],[END_TIME],[STATUS] from BC_BIW_SFTP_LOG where [UPDATED_DATETIME] IS NULL";
                    SqlDataAdapter dscmd = new SqlDataAdapter(sql, cnn1);
                    System.Data.DataTable ds_mail = new System.Data.DataTable();
                    dscmd.Fill(ds_mail);
                    if (ds_mail.Rows.Count > 0)
                    {
                        for ( int i = 0; i < ds_mail.Rows.Count; i++)
                        {
                            if (i == 0)
                                mailBody = "<table border=1><b><tr bgcolor='#81BEF7'><td>S.No</td><td>FILE_NAME</td><td>START_TIME</td><td>END_TIME</td><td>STATUS</td></tr></b><tr><td>" + (i + 1).ToString() + "</td><td>" + ds_mail.Rows[i][0].ToString() + "</td><td>" + ds_mail.Rows[i][1].ToString() + "</td><td>" + ds_mail.Rows[i][2].ToString() + "</td><td>" + ds_mail.Rows[i][3].ToString() + "</td></tr>";
                            else
                                mailBody = mailBody + "<tr><td>" + (i + 1).ToString() + "</td><td>" + ds_mail.Rows[i][0].ToString() + "</td><td>" + ds_mail.Rows[i][1].ToString() + "</td><td>" + ds_mail.Rows[i][2].ToString() + "</td><td>" + ds_mail.Rows[i][3].ToString() + "</td></tr>";
                        }
                        //sqlEX = "select INTERFACE_NAME,DISTRIBUTOR,SKU,INV_NUMBER,BILL_DATE,DELV_DATE,ERROR_DETAILS,[DATETIME] from [BC_I_EX_ERRORLOG_EXTRACT] where INTERFACE_NAME='BalanceStock' and [UPDATED_DATETIME] IS NULL";
                        //SqlDataAdapter dscmdEX = new SqlDataAdapter(sqlEX, cnn);
                        //DataTable ds_mailEX = new DataTable();
                        //dscmdEX.Fill(ds_mailEX);
                        //if (ds_mailEX.Rows.Count > 0)
                        //{
                        //    EX_Occured = "Y";
                        //    exporttoexcel("BalanceStock");
                        //}

                        //if (EX_Occured == "Y")
                        //    new I_MailUtility().SendErrorReportBIW(mailBody + "</table>", "CSDP-BIW interface - BalanceStock Scheduler Progress Details " + DateTime.Now, 1, ExcelFilePath, "", schedulerstart.ToString());
                        //else
                        //new I_MailUtility().SendErrorReportBIW(mailBody + "</table>", "CSDP Interface - DRS BalanceStock Scheduler Progress Details " + DateTime.Now, 1, ExcelFilePath, "", "6 AM");

                        string s1Year1 = DateTime.Now.Year.ToString();
                        string s1Month1 = DateTime.Now.Month.ToString().PadLeft(2, '0');
                        string s1Day1 = DateTime.Now.Day.ToString().PadLeft(2, '0');
                        string s1ErrorTime1 = s1Day1 + "-" + s1Month1 + "-" + s1Year1;
                        DateTime dt = DateTime.Now;
                        string from = System.Configuration.ConfigurationSettings.AppSettings["Mail_From"].ToString(),
                            to = System.Configuration.ConfigurationSettings.AppSettings["Mail_To"].ToString(),
                            copy = System.Configuration.ConfigurationSettings.AppSettings["Mail_Copy"].ToString(),

                            //to = "Lal.Samansiri@unilever.com,Lavan.Harshanga@unilever.com,Gayan.Abeysingha@unilever.com,Tusira.Dilnath@unilever.com,Madura.Perera@unilever.com,Suraj.Perera@unilever.com,Chandana.Liyanage@unilever.com",
                            //copy = "Amarnath.Gopinath@unilever.com,Raneesh.Rajeevan@unilever.com",

                            subject = "SFTP Process Status - " + status + " Generated On: " + s1ErrorTime1, body, filePath;
                        //string Attachment_path = System.Configuration.ConfigurationSettings.AppSettings["Attachment_path"].ToString();
                        bool isHtmlBody;
                        //string FileNameFileName = "SL-CSDP-BIW-Sal-Ext-" + dt.Day + dt.Month + dt.Year + dt.Hour + dt.Minute + dt.Second + ".csv";

                        string server = System.Configuration.ConfigurationSettings.AppSettings["Mail_Server"].ToString();
                        int mailPort = 25;

                        string connectionString = null;
                        //string sql = null, sqlEX = null;

                        //string FileName = decimal.Parse(GetDataTable(sql).Rows[0]["SELL_FACTOR1"].ToString());

                        string returnMsg = "";
                        MailMessage mailMsg = new MailMessage();

                        //set from Address
                        MailAddress mailAddress = new MailAddress(from);
                        mailMsg.From = mailAddress;
                        //set to Adress
                        mailMsg.To.Add(to);
                        // Set Message subject
                        mailMsg.Subject = subject;
                        //set mail cc
                        if (copy != "")
                            mailMsg.CC.Add(copy);
                        // Set Message Body
                        mailMsg.IsBodyHtml = true;
                        filePath = @"\\10.21.36.10\Interface\outbound\Log\";
                        DirectoryInfo dir = new DirectoryInfo(filePath);

                        dir.Refresh();
                        filePath = filePath + myFile;
                        if (filePath != "")
                        {
                            Attachment attach3 = new Attachment(filePath, "application/vnd.ms-excel");
                            mailMsg.Attachments.Add(attach3);
                        }

                        string v = subject.Replace("NGDMS Interface - ", "");
                        string v1 = v.Replace(" Error", "");
                        string htmlBody, htmlBody2, htmlBody3, htmlBody4;
                        htmlBody = "<html>Dear Team,<br/><br/></html>";

                        htmlBody2 = "<html>Please find the attachment of SFTP Process Status file and below details for your reference, which was taken on " + s1Day1 + "/" + s1Month1 + "/" + s1Year1 + "  at " + DateTime.Now.ToLongTimeString().ToString() + ".<br/><br/></html>";
                        htmlBody4 = "<html><br/>Regards, <br/></html>CSDP Interface Support Team.<br/></html>";
                        //<br/><br/>The information contained in this electronic message and any attachments to this message are intended for the exclusive use of the addressee(s) and may contain proprietary, confidential or privileged information. <br/>If you are not the intended recipient, you should not disseminate, distribute or copy this e-mail. Please notify the sender immediately and destroy all copies of this message and any attachments.<br/></html>";
                        mailMsg.Body = htmlBody + "  " + htmlBody2 + mailBody + "</table>" + htmlBody4;


                        //set message body format
                        //mailMsg.IsBodyHtml = isHtmlBody;
                        // System.IO.File.WriteAllText(@"c:\abc.xlsx", attachmentStream);
                        //byte[] data = GetData(attachmentStream);


                        //if (attachmentStream != null)//Define mail attachment.
                        //{
                        //    ms.Position = 0;//explicitly set the starting position of the MemoryStream
                        //Attachment attach = new Attachment(filePath, "application/vnd.ms-excel");
                        //mailMsg.Attachments.Add(attach);
                        //}


                        //set exchange server
                        SmtpClient smtpClient = new SmtpClient(server);
                        smtpClient.Send(mailMsg);
                        mailMsg.Dispose();
                        //return returnMsg;

                    }
                    string update_errorlog;

                    update_errorlog = "update BC_BIW_SFTP_LOG SET UPDATED_DATETIME = @DATE   where UPDATED_DATETIME is null";
                    SqlCommand cmdUDRS = new SqlCommand(update_errorlog, cnn1);
                    cmdUDRS.Parameters.AddWithValue("@DATE", DateTime.Now);
                    cmdUDRS.ExecuteNonQuery();

                    //DBInteraction dbClass_log = new DBInteraction();
                    //dbClass_log.InsertToDB("update BC_BIW_DRS_BalanceStock_LOG SET UPDATED_DATETIME = getdate()  where UPDATED_DATETIME is null ");
                    cnn1.Close();

                    //FileData temp = new FileData();
                    //foreach (var line in File.ReadLines("filepath.txt").Skip(1))
                    //{
                    //    var tempLine = line.Split('\t');
                    //    temp.Column2 = tempLine[1];
                    //    temp.Column12 = tempLine[11];
                    //    temp.Column45 = tempLine[44];
                    //    filedata.Add(temp);
                    //}
                }
            }

        }
    

