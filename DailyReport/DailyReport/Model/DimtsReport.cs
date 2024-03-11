using DailyReport.Model;
using MySql.Data.MySqlClient;
using ProtoBuf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DailyReport.Model
{
    public class DimtsReport
    {
        #region variable     
        string fromPassword, fromEmail, computerDetails = string.Empty; bool sendMail; bool Clflag = false; int millisec1;
        string Mail = ConfigurationSettings.AppSettings["Mail"]; bool KClflag = false; int _millisec1; bool MClflag = false; int Mmillisec1;
        //string FileSourcePath = ConfigurationSettings.AppSettings["FileSourcePath"];
        DateTime lastPbGeneratedTime = DateTime.Now.AddSeconds(-29); string[] FtpUpload = ConfigurationSettings.AppSettings["Ftp"].Split('/');
        string LogPath, logPath = string.Empty; bool DalFlag = false; DateTime startTime; string _time = DateTime.Now.ToString("HH:mm:ss");
        List<FeedUpdationTime> _updttime = new List<FeedUpdationTime>(); DateTime dtThirMinute; double _seconds;
        object _Report = new object(); object _lock = new object(); bool Emailflag = false; bool _thSec = false;
        int ST = Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["StartTime"]);
        int ET = Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["EndTime"]);
        string ConnectionInsert = ConfigurationSettings.AppSettings["ConnectionInsert"];
        string bkptablename = string.Empty;
        string Database = ConfigurationSettings.AppSettings["Database"];
        List<BufferData> _StopTimes = new List<BufferData>();
        string insertDumpValues = string.Empty;
        string[] mailingaddresto = Convert.ToString(ConfigurationSettings.AppSettings["MailToAddress"]).Split(',');

        #endregion
        public DimtsReport()
        {
            try
            {
                string[] mailDetails = Mail.Split('/');
                if (mailDetails.Length.Equals(2))
                {
                    fromEmail = mailDetails[0] + "@gmail.com";
                    fromPassword = mailDetails[1];
                    sendMail = true;
                }
                else
                {
                    sendMail = false;
                }
                string hostName = Dns.GetHostName();
                string myIP = Dns.GetHostByName(hostName).AddressList[0].ToString();
                computerDetails = hostName + " [" + myIP + "]";
            }
            catch (Exception ex) { }
        }
        public void StartVoidProcess()
        {
            try
            {
                Thread[] objThread = new Thread[1];
                objThread[0] = new Thread(new ThreadStart(DIMTSDataDownload));
                //objThread[1] = new Thread(new ThreadStart(CheckRailTransit));
                foreach (Thread myThread in objThread)
                {
                    myThread.Start();
                }
            }
            catch { }

        }

        public void WriteLog(string logMessage, string foldername)
        {

            try
            {
                LogPath = ConfigurationSettings.AppSettings["LogPath"];

                logPath = LogPath + "\\Logged\\" + foldername;
                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }
                //logPath = LogPath + "Logged\\" + foldername + "\\Log_" + DateTime.Now.ToString("dd_MM_yyyy") + ".txt";
                logPath = logPath+"\\Log_" + DateTime.Now.ToString("dd_MM_yyyy") + ".txt";
                if (!File.Exists(logPath))
                {
                    FileStream fileStream = File.Create(logPath);
                    fileStream.Close();

                }
                using (StreamWriter txtWriter = File.AppendText(logPath))
                {
                    txtWriter.WriteLine(DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "\t" + logMessage);
                    if (!logMessage.Contains("Processing"))
                        txtWriter.WriteLine("-------------------------------------------------------------------------");
                }
            }

            catch (Exception ex)
            {

            }
        }
        public void DIMTSDataDownload()
        {
            TimeSpan start = new TimeSpan(23, 59, 59); //10 o'clock
            TimeSpan end = new TimeSpan(05, 0, 0); //12 o'clock
            TimeSpan now = DateTime.Now.TimeOfDay;

            try
            {
                while (true)
                {
                    now = DateTime.Now.TimeOfDay;
                    if ((now < start) && (now > end))
                    {
                        lock (_Report)
                        {

                            string filepath = string.Empty;
                            try
                            {
                                if (Emailflag == false)
                                {
                                    dtThirMinute = CurrentToThirtyMinute();
                                }
                                WriteLog("Feed  hit time " + DateTime.Now.ToString() + "", "UlrHittime");
                                startTime = DateTime.Now;
                                filepath = System.Configuration.ConfigurationManager.AppSettings["DIMTSURL"].ToString().Trim();
                                string fileuploadedpath = System.Configuration.ConfigurationManager.AppSettings["FileCopypath"].ToString().Trim();
                                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                string fileUpload = AppDomain.CurrentDomain.BaseDirectory + "\\DIMTS\\VehiclePositions.pb";
                                string directory = Path.GetDirectoryName(fileUpload);
                                if (!(Directory.Exists(directory)))
                                { Directory.CreateDirectory(directory); }
                                using (var client = new WebClient())
                                {
                                    client.DownloadFile(filepath, fileUpload);
                                }
                                File.Copy(fileUpload, fileuploadedpath, true);
                                WebRequest req = HttpWebRequest.Create(filepath);
                                string hittime = DateTime.Now.ToString("s");
                                TransitRealtime.FeedMessage feed = Serializer.Deserialize<TransitRealtime.FeedMessage>(req.GetResponse().GetResponseStream());
                                startTime = DateTime.Now;
                                insertDumpValues = String.Empty;
                                if (feed.Entities.Count > 0)
                                {
                                    Emailflag = false;
                                    if (feed.Entities.Count > 0)
                                    {
                                        _StopTimes = new List<BufferData>();
                                        foreach (TransitRealtime.FeedEntity entity in feed.Entities)
                                        {
                                            try
                                            {
                                                var realtime = entity;
                                                if (realtime != null)
                                                {
                                                    _StopTimes.Add(new BufferData() { Bus_reg_no = entity.Id, 
                                                        route_name = entity.Vehicle.Trip.RouteId,
                                                        Latitude = entity.Vehicle.Position.Latitude,
                                                        Longitude = entity.Vehicle.Position.Longitude,
                                                        velocity = entity.Vehicle.Position.Speed.ToString(), 
                                                        timestamp = (int)entity.Vehicle.Timestamp, 
                                                        api_hit_time = hittime,                                                        
                                                        local_timestamp = (HelperClass.UnixTimeStampToLocalDateTime(entity.Vehicle.Timestamp.ToString())).ToString("s"),
                                                        feed_header_timestamp= (int)feed.Header.Timestamp                                                   });
                                                }

                                            }
                                            catch (Exception ex)
                                            {
                                                //HelperClass.HelperClass.ExceptionLog(ex.Source, ex.Message, ex.StackTrace, "GenerateDataFeedUrl");
                                            }
                                        }
                                    }
                                    CheckCurrentDataTable("agencydata");
                                    foreach (var lists in _StopTimes)
                                    {
                                        insertDumpValues = insertDumpValues.Equals(string.Empty) ? "('" + lists.api_hit_time + "','" + lists.route_name + "','" + lists.Latitude + "','" + lists.Longitude + "','" + lists.Bus_reg_no + "','" + lists.velocity + "','" + lists.timestamp + "','" + lists.local_timestamp + "','" + lists.feed_header_timestamp + "')\n"
                                                              : insertDumpValues + ",\n" + "('" + lists.api_hit_time +"','" + lists.route_name + "','" + lists.Latitude + "','" + lists.Longitude + "','" + lists.Bus_reg_no + "','" + lists.velocity + "','" + lists.timestamp + "','" + lists.local_timestamp + "','" + lists.feed_header_timestamp + "')";
                                    }
                                    InsertRecordTable("INSERT INTO " + bkptablename + "(api_hit_time,route_name,Latitude,Longitude,Bus_reg_no,velocity,timestamp,local_timestamp,feed_header_timestamp) " +
                                                                                   "values " + insertDumpValues);
                                }
                                else
                                {
                                    Emailflag = true;
                                    TimeSpan _thSdelayMin = dtThirMinute - DateTime.UtcNow;
                                    if (_thSdelayMin < TimeSpan.Zero)
                                    {
                                        Emailflag = false;
                                        _seconds = TimeSpan.Parse(_time).TotalSeconds;
                                        SendMailAsync("Data is not coming from DIMTS agency side. Please check the link :- " + filepath + "", mailingaddresto, "WARNING ON SERVER 2.21");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                SendMailAsync("Process Failure.\t"+ ex+"\t", mailingaddresto, "WARNING ON SERVER 2.21");
                            }
                            //download api feed every 32 seconds::keep delay still next downloading time
                            int delayThread = Convert.ToInt32(ConfigurationManager.AppSettings["RST"]) - (int)((DateTime.Now - startTime).TotalSeconds);
                            if (delayThread > 0)
                                Thread.Sleep(1000 * delayThread);
                        }
                    }
                }
            }
            catch (Exception ex) { }

        }

        public void SendMailAsync(string body, string[] emailIDs, string mailType = "")
        {
            try
            {
                if (!sendMail)
                    return;
                body = body + "<p />Thanks,<br />Lepton - RT Service <br />-<br /><i> <u>Sent from " + computerDetails + "</u></i>";
                string Subject = mailType.Equals("INFO") ? "RT Service Started & Stopped Status. !" + mailType : "Data is not coming from delhi agency side. " + mailType;
                foreach (string emailID in emailIDs)
                {
                    var fromAddress = new MailAddress(fromEmail);
                    var toAddress = new MailAddress(emailID);
                    SmtpClient smtp = new SmtpClient
                    {
                        Host = "smtp.gmail.com",
                        Port = 587,
                        EnableSsl = true,
                        DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network,
                        UseDefaultCredentials = false,
                        Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
                    };
                    var message = new MailMessage(fromAddress, toAddress)
                    {
                        Subject = Subject,
                        Body = body,
                        IsBodyHtml = true,
                    };
                    smtp.Send(message);
                }
            }
            catch (Exception ex)
            {

            }
        }
        FileStream fs; Stream strm;
        void UploadFeedinFTP(string filename)
        {

            try
            {

                if (FtpUpload.Length >= 2)
                {
                    FileInfo fileInf = new FileInfo(filename);
                    string ftppath = ConfigurationManager.AppSettings["ftpfilepath"].ToString();
                    FtpWebRequest reqFTP;
                    reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(ftppath + fileInf.Name));
                    reqFTP.Credentials = new NetworkCredential(FtpUpload[0], FtpUpload[1]);
                    reqFTP.KeepAlive = false;
                    reqFTP.Method = WebRequestMethods.Ftp.UploadFile;
                    //reqFTP.UsePassive = false;
                    reqFTP.UseBinary = true;

                    reqFTP.ContentLength = fileInf.Length;
                    int buffLength = 1024 * 1024;
                    byte[] buff = new byte[buffLength];
                    int contentLen;

                    try
                    {
                        fs = fileInf.OpenRead();
                        strm = reqFTP.GetRequestStream();
                        try
                        {
                            contentLen = fs.Read(buff, 0, buffLength);
                            double totalReadBytesCount = 0;
                            while (contentLen != 0)
                            {

                                strm.Write(buff, 0, contentLen);
                                contentLen = fs.Read(buff, 0, buffLength);
                                totalReadBytesCount += contentLen;
                                var progress = totalReadBytesCount * 70 / fs.Length;
                            }
                        }
                        catch (Exception ex) { WriteLog(ex.Source.ToString() + "::" + ex.Message + "::" + ex.StackTrace.ToString(), "uploadfeedinftp"); }
                        strm.Close();
                        fs.Close();
                    }
                    catch (WebException e)
                    {
                        strm.Close();
                        fs.Close();
                        String status = ((FtpWebResponse)e.Response).StatusDescription;
                        // helperclass.exceptionlog(e.source, e.message, e.stacktrace, "uploadfeedinftp");
                        WriteLog(e.Source.ToString() + "::" + e.Message + "::" + e.StackTrace.ToString(), "uploadfeedinftp");
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        DateTime CurrentToThirtyMinute()
        {
            DateTime now = DateTime.UtcNow,
            result = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second);
            return result.AddMinutes(15);
        }

        public void InsertRecordTable(string Insrtqrt)
        {
            try
            {
                using (var connection = new MySqlConnection(ConnectionInsert))
                {
                    using (var command = connection.CreateCommand())
                    {
                        connection.Open();
                        try
                        {
                            command.CommandText = Insrtqrt;
                            command.CommandTimeout = 180;
                            command.ExecuteNonQuery();
                            connection.Close();
                        }
                        catch { connection.Close(); }
                    }
                }
            }
            catch (Exception ex)
            {
                // HelperClass.HelperClass.ExceptionLog(ex.Source, ex.Message, ex.StackTrace, "InsertRecordTable");
            }
        }

        private void CheckCurrentDataTable(string TableName)
        {
            try
            {
                string Tabledate = DateTime.Now.ToString("yyyy_MM_dd", CultureInfo.InvariantCulture);
                if (TableName != string.Empty)
                {
                    TableName = TableName + "_" + Tabledate;
                    bkptablename = TableName;
                    using (var connection = new MySqlConnection(ConnectionInsert))
                    {
                        using (var command = connection.CreateCommand())
                        {
                            connection.Open();
                            string strQuery = string.Empty;
                            strQuery = "SELECT table_name FROM information_schema.tables WHERE table_schema = '" + Database + "' AND table_name = '" + TableName.ToLower() + "'";
                            MySqlDataAdapter da = new MySqlDataAdapter(strQuery, ConnectionInsert);
                            DataSet ds = new DataSet();
                            da.Fill(ds);
                            try
                            {
                                if (ds.Tables[0].Rows[0][0].ToString() == "0")
                                {
                                    string strq = string.Empty;
                                    strq = "CREATE TABLE  " + TableName.ToLower() + " (api_hit_time varchar(100),route_name varchar(30),Latitude varchar (100),Longitude varchar (100),Bus_reg_no varchar(100),velocity varchar(100),timestamp varchar(100),local_timestamp varchar(100),feed_header_timestamp varchar(100))";
                                    using (MySqlCommand cmd = new MySqlCommand(strq, connection))
                                    {
                                        WriteLog(strq, "Tableqry");
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }
                            catch
                            {
                                string strq = string.Empty;
                                strq = "CREATE TABLE  " + TableName.ToLower() + " (api_hit_time varchar(100),route_name varchar(30),Latitude varchar (100),Longitude varchar (100),Bus_reg_no varchar(100),velocity varchar(100),timestamp varchar(100),local_timestamp varchar(100),feed_header_timestamp varchar(100))";
                                using (MySqlCommand cmd = new MySqlCommand(strq, connection))
                                {
                                    WriteLog(strq, "Tableqry");
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            connection.Close();
                        }
                    }
                }
            }
            catch (Exception ex) { }
        }
    }

    public class FeedUpdationTime
    {
        public DateTime datetime { get; set; }
        public bool FlagValue { get; set; }
    }

    public class BufferData
    {
        public string Bus_reg_no { get; set; }
        public string route_name { get; set; }
        public double? Latitude { get; set; }
        public double? Longitude { get; set; }
        public string velocity { get; set; }
        public int timestamp { get; set; }
        public string local_timestamp { get; set; }
        public string api_hit_time { get; set; }
        public int feed_header_timestamp { get; set; }
    }

}
