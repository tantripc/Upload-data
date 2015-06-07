using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Net;
using System.Threading;
using System.Windows.Forms;

namespace synSef
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            button1.Enabled = false;
        }

        DataTable db = null;
        Thread th = null;
        Thread th1 = null;
        Thread th2 = null;
        int row1 = 0;
        int row2 = 0;
        int tientrinh1 = 0;
        int tientrinh2 = 0;

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = openDialog();
                db = GetSheet(txtFileName.Text, 0);
                row1 = db.Rows.Count / 2;
                row2 = db.Rows.Count;
                progressBar1.Value = 0;
                progressBar1.Maximum = row2;
                progressBar1.Minimum = 0;
                if (db != null)
                {

                    th = new System.Threading.Thread(new System.Threading.ThreadStart(CapNhatTong));
                    th.Start();

                    timer1.Start();
                    btnChon.Enabled = false;
                    button1.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Bảng dữ liệu không tồn tại!");
                }

            }
            catch (Exception exx) { SendLog(exx.Message); }
        }

        void CapNhatTong()
        {
            th1 = new System.Threading.Thread(new System.Threading.ThreadStart(Update1));
            th2 = new System.Threading.Thread(new System.Threading.ThreadStart(Update2));

            th1.Start();
            th2.Start();
        }


        void Update1()
        {
            int l = 0;
            for (int i = row1-1; i >= 0; i--)
            {
                DataRow row = db.Rows[i];
                clsBill oBill = new clsBill();
                oBill.Index = row[0].ToString();
                oBill.SE = row[2].ToString(); 
                if (!string.IsNullOrEmpty(oBill.SE))
                {
                    oBill.SE = "SE" + oBill.SE.Replace("-", "").Replace("'", "");
                    oBill.Atr_where = row[3].ToString().Replace("'", "");
                    oBill.Atr_when = row[11].ToString().Replace("'", "");
                    oBill.Atr_when = oBill.Atr_when.Trim() != "" ? oBill.Atr_when.Replace("H", ":") : oBill.Atr_when;
                    oBill.Atr_who = row[12].ToString().Replace("'", "");
                    string nextH = "";
                    if (oBill.Atr_when.Trim() != "" && oBill.Atr_when.Contains(":"))
                    {
                        if (oBill.Atr_when.Substring(oBill.Atr_when.IndexOf(":") + 1, 1) == " ")
                        {
                            nextH = "00";
                            oBill.Atr_when = oBill.Atr_when.Insert(oBill.Atr_when.IndexOf(":") + 1, nextH);
                        }
                    }
                    string q1 = "insert ";
                    string q2 = "into sef_customer(SE,atr_where,atr_when,atr_who, ngayNhap) "
                                    + "values('" + oBill.SE + "','" + oBill.Atr_where + "','" + oBill.Atr_when + "','" + oBill.Atr_who + "','" + DateTime.Now.ToString() + "')";

                    string u1 = "update ";
                    string u2 = "sef_customer set atr_where='" + oBill.Atr_where + "',"
                    + " atr_when='" + oBill.Atr_when + "',"
                    + " atr_who='" + oBill.Atr_who + "' where SE='" + oBill.SE + "'";
                    //MessageBox.Show("SE: " + oBill.SE + "\nNoi Nhan: " + oBill.Atr_where+"\nMaMK: "+oBill.Atr_customer+"\nDien Thoai :" + oBill.Atr_phone+"\nGhi Chu: " + oBill.Atr_note+"\nNgay nhan: " + oBill.Atr_when+"\nNguoi Nhan: " + oBill.Atr_who);
                    //MessageBox.Show("SE: " + oBill.SE + "\nNoi Nhan: " + oBill.Atr_where+"\nMaMK: "+oBill.Atr_customer+"\nDien Thoai :" + oBill.Atr_phone+"\nGhi Chu: " + oBill.Atr_note+"\nNgay nhan: " + oBill.Atr_when+"\nNguoi Nhan: " + oBill.Atr_who);
                    q2 = q2.Replace("#", ""); u2 = u2.Replace("#", "");

                    if (l <= 100)
                    {
                        if (SendUpdate1(q1, q2, oBill.Index).ToLower() == "false")
                        {
                            SendUpdate1(u1, u2, oBill.Index).ToLower();
                            l++;
                        }
                    }
                    else
                    {
                        SendUpdate1(u1, u2, oBill.Index);
                    }
                    tientrinh1++;                
                }                     
            }
        }
        void Update2()
        {
            int k = 0;
            for (int i = row2-1; i >= row1; i--)
            {
                DataRow row = db.Rows[i];
                clsBill oBill = new clsBill();
                oBill.Index = row[0].ToString();
                oBill.SE = row[2].ToString(); 
                if (!string.IsNullOrEmpty(oBill.SE))
                {
                    oBill.SE = "SE" + oBill.SE.Replace("-", "").Replace("'", "");
                    oBill.Atr_where = row[3].ToString().Replace("'", "");
                    oBill.Atr_when = row[11].ToString().Replace("'", "");
                    oBill.Atr_when = oBill.Atr_when.Trim() != "" ? oBill.Atr_when.Replace("H", ":") : oBill.Atr_when;
                    oBill.Atr_who = row[12].ToString().Replace("'", "");
                    string nextH = "";
                    if (oBill.Atr_when.Trim() != "" && oBill.Atr_when.Contains(":"))
                    {
                        if (oBill.Atr_when.Substring(oBill.Atr_when.IndexOf(":") + 1, 1) == " ")
                        {
                            nextH = "00";
                            oBill.Atr_when = oBill.Atr_when.Insert(oBill.Atr_when.IndexOf(":") + 1, nextH);
                        }
                    }
                    string q1 = "insert ";
                    string q2 = "into sef_customer(SE,atr_where,atr_when,atr_who, ngayNhap) "
                                     + "values('" + oBill.SE + "','" + oBill.Atr_where + "','" + oBill.Atr_when + "','" + oBill.Atr_who + "','" + DateTime.Now.ToString() + "')";

                    string u1 = "update ";
                    string u2 = "sef_customer set atr_where='" + oBill.Atr_where + "',"
                    + " atr_when='" + oBill.Atr_when + "',"
                    + " atr_who='" + oBill.Atr_who + "' where SE='" + oBill.SE + "'";
                    //MessageBox.Show("SE: " + oBill.SE + "\nNoi Nhan: " + oBill.Atr_where+"\nMaMK: "+oBill.Atr_customer+"\nDien Thoai :" + oBill.Atr_phone+"\nGhi Chu: " + oBill.Atr_note+"\nNgay nhan: " + oBill.Atr_when+"\nNguoi Nhan: " + oBill.Atr_who);
                    //MessageBox.Show("SE: " + oBill.SE + "\nNoi Nhan: " + oBill.Atr_where+"\nMaMK: "+oBill.Atr_customer+"\nDien Thoai :" + oBill.Atr_phone+"\nGhi Chu: " + oBill.Atr_note+"\nNgay nhan: " + oBill.Atr_when+"\nNguoi Nhan: " + oBill.Atr_who);
                    q2 = q2.Replace("#", ""); u2 = u2.Replace("#", "");

                    if (k <= 100)
                    {
                        if (SendUpdate2(q1, q2, oBill.Index).ToLower() == "false")
                        {
                            SendUpdate2(u1, u2, oBill.Index);
                            k++;
                        }
                        else
                            k = 0;
                    }
                    else
                    {
                        SendUpdate2(u1, u2, oBill.Index).ToLower();
                    }
                    tientrinh2++;
                }                
            }
        }
        string SendUpdate1(string q1, string q2, string index)
        {
            string data = "";
            ///// Tao Request
            HttpWebRequest oRequest = (HttpWebRequest)WebRequest.Create("http://sefsaigon.com/import.php?q1=" + q1 + "&q2=" + q2);
            oRequest.Method = "GET";
            oRequest.Timeout = int.MaxValue;
            try
            {
                HttpWebResponse oResponse = (HttpWebResponse)oRequest.GetResponse();
                Stream oStream = oResponse.GetResponseStream();
                StreamReader reader = new StreamReader(oStream);
                data = reader.ReadToEnd();
                //Thread.Sleep(10);
                oStream.Close();
                oResponse.Close();
                //if (data == "false")
                //{
                //    clsLogg logg = new clsLogg();
                //    logg.day = DateTime.Now.ToString();
                //    logg.er = data;
                //    logg.query = q1 + q2;
                //    logg.index = index;

                //    StreamWriter wt = new StreamWriter(Application.StartupPath + "\\Logg.txt", true);

                //    wt.WriteLine(logg.day);
                //    wt.WriteLine(logg.er);
                //    wt.WriteLine(logg.query);
                //    wt.WriteLine(logg.index);
                //    wt.WriteLine("---------------");
                //    wt.Close();
                //    wt.Dispose();
                //}
            }
            catch (Exception exx)
            {
                SendLog(exx.Message);
            }
            return data;
        }


        string SendUpdate2(string q1, string q2, string index)
        {
            string data1 = "";
            HttpWebRequest oRequest1 = (HttpWebRequest)WebRequest.Create("http://sefsaigon.com/import.php?q1=" + q1 + "&q2=" + q2);
            oRequest1.Method = "GET";
            oRequest1.Timeout = int.MaxValue;
            try
            {
                HttpWebResponse oResponse1 = (HttpWebResponse)oRequest1.GetResponse();
                Stream oStream1 = oResponse1.GetResponseStream();
                StreamReader reader1 = new StreamReader(oStream1);
                data1 = reader1.ReadToEnd();
                //MessageBox.Show(data1);
                oStream1.Close();
                oResponse1.Close();
                //if (data1 == "false")
                //{
                //    clsLogg logg = new clsLogg();
                //    logg.day = DateTime.Now.ToString();
                //    logg.er = data1;
                //    logg.query = q1 + q2;
                //    logg.index = index;
                //    StreamWriter wt = new StreamWriter(Application.StartupPath + "\\Logg.txt", true);
                //    wt.WriteLine(logg.day);
                //    wt.WriteLine(logg.er);
                //    wt.WriteLine(logg.query);
                //    wt.WriteLine(logg.index);
                //    wt.WriteLine("---------------");
                //    wt.Close();
                //    wt.Dispose();
                //}
            }
            catch (Exception exx)
            {
                SendLog(exx.Message);
            }
            return data1;
        }

        DialogResult openDialog()
        {
            OpenFileDialog openfiledialog = new OpenFileDialog
            {
                Filter = "Excel 97-2003 files (*.xls)|*.xls",
                InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyComputer),
                Multiselect = false,
                Title = "Chọn tập tin Excel",
                CheckFileExists = true,
                CheckPathExists = true
            };
            DialogResult dr = openfiledialog.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.Abort || dr == System.Windows.Forms.DialogResult.Cancel || dr == System.Windows.Forms.DialogResult.Ignore)
            {
                return DialogResult.Cancel;
            }
            else
            {
                txtFileName.Text = openfiledialog.FileName;
                return DialogResult.OK;
            }
        }

        DataTable GetSheet(string fileName, int index)
        {
            try
            {
                string sheetName = GetSheetNames(fileName)[index];
                string strConn = "";// "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + fileName + "';Extended Properties=Excel 8.0;";
                strConn = "provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + fileName + "';Extended Properties=Excel 8.0;";
                OleDbConnection objConn = new OleDbConnection(strConn);
                try
                {
                    objConn.Open();
                }
                catch
                {
                    strConn = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + fileName + "';Extended Properties=Excel 8.0;";
                    objConn.ConnectionString = strConn;
                    objConn.Open();
                }
                string query = "select * from [" + sheetName + "]";
                OleDbCommand cmd = new OleDbCommand(query, objConn);
                DataSet ds = new DataSet();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(ds, sheetName);
                objConn.Close();

                return ds.Tables[0];
            }
            catch (Exception exx)
            {
                SendLog(exx.Message);
                return null;
            }
        }
        String[] GetSheetNames(string fileName)
        {
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;
            try
            {
                string strConn = "";
                strConn = "provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + fileName + "';Extended Properties=Excel 8.0;";

                objConn = new OleDbConnection(strConn);
                try
                {
                    objConn.Open();
                }
                catch (Exception exx)
                {
                    SendLog(exx.Message);
                    strConn = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + fileName + "';Extended Properties=Excel 8.0;";
                    objConn.ConnectionString = strConn;
                    objConn.Open();
                }
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dt == null)
                    return null;
                string[] excelsheet = new string[dt.Rows.Count];
                int i = 0;
                foreach (DataRow row in dt.Rows)
                {
                    excelsheet[i] = row["TABLE_NAME"].ToString();
                    i++;
                }
                return excelsheet;
            }
            catch (Exception exx)
            {
                SendLog(exx.Message);
                return null;
            }
            finally
            {
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        void SendLog(string er)
        {

            HttpWebRequest oRequest1 = (HttpWebRequest)WebRequest.Create("http://sefsaigon.com/handle_log.php?er=" + er + "&d=" + DateTime.Now.ToString());
            oRequest1.Method = "GET";
            oRequest1.Timeout = int.MaxValue;
            try
            {
                HttpWebResponse oResponse1 = (HttpWebResponse)oRequest1.GetResponse();
                Stream oStream1 = oResponse1.GetResponseStream();
                StreamReader reader1 = new StreamReader(oStream1);
                oStream1.Close();
                oResponse1.Close();
            }
            catch { }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (th != null)
            {
                th.Abort();
                if (th1 != null && th2 != null)
                {
                    th1.Abort();
                    th2.Abort();
                }
            }
            progressBar1.Value = 0;
            tientrinh1 = 0;
            tientrinh2 = 0;
            timer1.Stop();
            button1.Enabled = false;
            btnChon.Enabled = true;
            txtFileName.ResetText();
            db = null;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Value = tientrinh1 + tientrinh2;
            if (!th1.IsAlive && !th2.IsAlive && !th.IsAlive)
            {
                tientrinh1 = 0;
                tientrinh2 = 0;
                timer1.Stop();
                MessageBox.Show("Quá trình đồng bộ hoàn tât!");
                progressBar1.Value = 0;
                btnChon.Enabled = true;
                button1.Enabled = false;
                txtFileName.ResetText();
                db = null;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (th != null)
            {
                th.Abort();
                if (th1 != null && th2 != null)
                {
                    th1.Abort();
                    th2.Abort();
                }
            }
            Application.Exit();
        }
    }
}
