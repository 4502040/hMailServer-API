using ClosedXML.Excel;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace hMailImport
{
    public partial class Form1 : Form
    {
        DataTable dt;

        MySqlConnection connection;        

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            MySqlConnection conn = DbUtil.GetDBConnection();

            string sql = "select * from contacts limit 10;";

            MySqlCommand cmd = new MySqlCommand(sql, conn);

            try
            {                

                conn.Open();

                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {

                        while (reader.Read())
                        {
                            // Индекс (index) столбца Emp_ID в команде SQL.
                            int empIdIndex = reader.GetOrdinal("contact_id"); // 0
                            int contact_id = reader.GetInt32(empIdIndex);                                                       
                            
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string fileName = Application.StartupPath+ "\\import.xlsx";
            var workbook = new XLWorkbook(fileName);
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed(); // Skip header row

            dt = new DataTable();

            dt.Columns.Add("Username", System.Type.GetType("System.String"));
            dt.Columns.Add("Password", System.Type.GetType("System.String"));
            dt.Columns.Add("Email", System.Type.GetType("System.String"));
            dt.Columns.Add("Address", System.Type.GetType("System.String"));
            dt.Columns.Add("HomePhone", System.Type.GetType("System.String"));
            dt.Columns.Add("MobilePhone", System.Type.GetType("System.String"));
            dt.Columns.Add("WorkPhone", System.Type.GetType("System.String"));
            dt.Columns.Add("Otdel", System.Type.GetType("System.String"));
            dt.Columns.Add("Doljnost", System.Type.GetType("System.String"));
            dt.Columns.Add("Familia", System.Type.GetType("System.String"));
            dt.Columns.Add("Name", System.Type.GetType("System.String"));
            dt.Columns.Add("Otchestvo", System.Type.GetType("System.String"));
            dt.Columns.Add("OrgName", System.Type.GetType("System.String"));
            dt.Columns.Add("GroupName", System.Type.GetType("System.String"));

            foreach (var row in rows)
            {
                DataRow newRow = dt.NewRow();

                string Username = Convert.ToString(row.Cell(1).Value);
                string Password = Convert.ToString(row.Cell(2).Value);
                string Email = Convert.ToString(row.Cell(3).Value);
                string Address = Convert.ToString(row.Cell(4).Value);
                string HomePhone = Convert.ToString(row.Cell(5).Value);
                string MobilePhone = Convert.ToString(row.Cell(6).Value);
                string WorkPhone = Convert.ToString(row.Cell(7).Value);
                string Otdel = Convert.ToString(row.Cell(8).Value);
                string Doljnost = Convert.ToString(row.Cell(9).Value);
                string Familia = Convert.ToString(row.Cell(10).Value);
                string Name = Convert.ToString(row.Cell(11).Value);
                string Otchestvo = Convert.ToString(row.Cell(12).Value);
                string OrgName = Convert.ToString(row.Cell(13).Value);
                string GroupName = Convert.ToString(row.Cell(14).Value);

                newRow["Username"] = Username;
                newRow["Password"] = Password;
                newRow["Email"] = Email;
                newRow["Address"] = Address;
                newRow["HomePhone"] = HomePhone;
                newRow["MobilePhone"] = MobilePhone;
                newRow["WorkPhone"] = WorkPhone;
                newRow["Otdel"] = Otdel;
                newRow["Doljnost"] = Doljnost;
                newRow["Familia"] = Familia;
                newRow["Name"] = Name;
                newRow["Otchestvo"] = Otchestvo;
                newRow["OrgName"] = OrgName;
                newRow["GroupName"] = GroupName;
                //var rowNumber = row.RowNumber();

                dt.Rows.Add(newRow);
                
            }

            dataGridView1.DataSource = dt;

            button4.Enabled = false; 

            if (dt.Rows.Count > 0)
            {
                button4.Enabled = true;
            }

            txtLog.AppendText("Импортировано строк: "+dt.Rows.Count.ToString()+Environment.NewLine);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Left = button3.Width + button3.Left + 30;
            dataGridView1.Width = Width - dataGridView1.Left - 30;

            txtLog.Width = dataGridView1.Width;
            txtLog.Height = 200;
            txtLog.Left = dataGridView1.Left;

            txtLog.Top = dataGridView1.Height + dataGridView1.Top + 20;

            connection = DbUtil.GetDBConnection();
            connection.Open();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            dataGridView1.Left = button3.Width + button3.Left + 30;
            dataGridView1.Width = Width - dataGridView1.Left - 30;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            groupBox1.Enabled = false;

            hMailServer.Application hApp = new hMailServer.Application();

            if (IsAuthHMail(hApp) == false)
            {
                groupBox1.Enabled = true;
                return;
            }

            foreach (DataGridViewRow viewRow in dataGridView1.Rows)
            {
                if (viewRow.Cells["Username"].Value != null)
                {
                    string Username = viewRow.Cells["Username"].Value.ToString();
                    string Password = viewRow.Cells["Password"].Value.ToString();
                    string Email = viewRow.Cells["Email"].Value.ToString();
                    string GroupName = viewRow.Cells["GroupName"].Value.ToString();

                    int? groupId = GetGroupId(GroupName);

                    if (groupId == null)
                    {
                        InsertNewGroup(GroupName);
                    }

                    groupId = GetGroupId(GroupName);

                    if (groupId != null)
                    {
                        InsertNewContact(viewRow, groupId);

                        CreateHMailAccount(hApp,Email,Password);
                    }                    
                }                

            }

            groupBox1.Enabled = true;
        }

        private int? GetGroupId(string GroupName)
        {
            string groupSql = "select c.contactgroup_id from contactgroups c WHERE c.name = @gname AND c.del=0 ORDER BY c.contactgroup_id DESC LIMIT 1;";
                        
            MySqlCommand cmd = new MySqlCommand(groupSql, connection);

            cmd.Parameters.AddWithValue("@gname", GroupName.Trim());

            var groupId = cmd.ExecuteScalar();

            if (groupId != null)
            {
                return Convert.ToInt32(groupId);
            }

            return null;
        }

        private bool InsertNewGroup(string GroupName)
        {
            string sql = "insert into contactgroups(user_id,del,name) values(@user_id,@del,@name);";

            MySqlCommand cmd = new MySqlCommand(sql, connection);

            cmd.Parameters.AddWithValue("@user_id", 18);
            cmd.Parameters.AddWithValue("@del", 0);
            cmd.Parameters.AddWithValue("@name", GroupName.Trim());

            int insert = cmd.ExecuteNonQuery();

            if (Convert.ToInt32(insert)>0)
            {
                txtLog.AppendText("Добавлена новая группа в контакты: "+GroupName+Environment.NewLine);
                return true;
            }

            return false;
        }

        private void InsertNewContact(DataGridViewRow viewRow, int? groupId)
        {
            string Username = viewRow.Cells["Username"].Value.ToString();
            string Password = viewRow.Cells["Password"].Value.ToString();
            string Email = viewRow.Cells["Email"].Value.ToString();
            string Address = viewRow.Cells["Address"].Value.ToString();
            string HomePhone = viewRow.Cells["HomePhone"].Value.ToString();
            string MobilePhone = viewRow.Cells["MobilePhone"].Value.ToString();
            string WorkPhone = viewRow.Cells["WorkPhone"].Value.ToString();
            string Otdel = viewRow.Cells["Otdel"].Value.ToString();
            string Doljnost = viewRow.Cells["Doljnost"].Value.ToString();
            string Familia = viewRow.Cells["Familia"].Value.ToString();
            string Name = viewRow.Cells["Name"].Value.ToString();
            string Otchestvo = viewRow.Cells["Otchestvo"].Value.ToString();
            string OrgName = viewRow.Cells["OrgName"].Value.ToString();

            string sql = @"INSERT INTO contacts
                            (del, name, email, firstname, surname, vcard, words, user_id, changed) 
                            VALUES 
                            (@del,@name,@email,@firstname, @surname, @vcard, @words, @user_id, @changed)";

            MySqlCommand cmd = new MySqlCommand(sql, connection);

            string fio = string.Concat(Name," ",Familia," ",Otchestvo);
            string words = string.Concat(fio," ",OrgName," ",MobilePhone," ", Address," ",Otdel," ", Doljnost," ", HomePhone);
            string vcard = 
@"BEGIN:VCARD
VERSION:3.0
N:@familia;@name;@surname;;
FN:@fio
EMAIL;TYPE=INTERNET;TYPE=WORK:@email
TITLE:@doljnost
ORG:@orgName
X-DEPARTMENT:@dep
TEL;TYPE=CELL:@mobile
ADR;TYPE=home:;;@address;;;;
END:VCARD";
            vcard = vcard.Replace("@familia",Familia);
            vcard = vcard.Replace("@name", Name);
            vcard = vcard.Replace("@surname", Otchestvo);
            vcard = vcard.Replace("@fio", fio);
            vcard = vcard.Replace("@email", Email);
            vcard = vcard.Replace("@doljnost", Doljnost);
            vcard = vcard.Replace("@orgName", OrgName);
            vcard = vcard.Replace("@dep", Otdel);
            vcard = vcard.Replace("@mobile", MobilePhone);
            vcard = vcard.Replace("@address", Address);

            cmd.Parameters.AddWithValue("@del", 0);
            cmd.Parameters.AddWithValue("@name", fio);            
            cmd.Parameters.AddWithValue("@email", Email);            
            cmd.Parameters.AddWithValue("@firstname", Name);            
            cmd.Parameters.AddWithValue("@surname", Otchestvo);            
            cmd.Parameters.AddWithValue("@vcard", vcard);            
            cmd.Parameters.AddWithValue("@words", words);            
            cmd.Parameters.AddWithValue("@user_id", 18);            
            cmd.Parameters.AddWithValue("@changed", DateTime.Now);            

            int insert = cmd.ExecuteNonQuery();

            string select = @"select contact_id from contacts where del=0 and email=@email and user_id=@user_id";
            cmd = new MySqlCommand(select, connection);
            cmd.Parameters.AddWithValue("@email", Email);
            cmd.Parameters.AddWithValue("@user_id", 18);

            var contact_id = cmd.ExecuteScalar();

            if (contact_id != null)
            {
                string insertIntoGroup = @"insert into contactgroupmembers(contactgroup_id,contact_id,created) values(@contactgroup_id,@contact_id,@created);";
                cmd = new MySqlCommand(insertIntoGroup, connection);
                cmd.Parameters.AddWithValue("@contactgroup_id", groupId);
                cmd.Parameters.AddWithValue("@contact_id", Convert.ToInt32(contact_id));
                cmd.Parameters.AddWithValue("@created", DateTime.Now);

                int insertIntoGroupResult = cmd.ExecuteNonQuery();
            }

            txtLog.AppendText("Добавлено в контакты: "+Email+"-"+fio+Environment.NewLine);


        }

        private void CreateHMailAccount(hMailServer.Application hApp, string Email, string Password)
        {
            if (IsAuthHMail(hApp))
            {
                hMailServer.Domain domain = hApp.Domains[0];

                hMailServer.Account hAccount ;

                string log;

                try {

                    hAccount = domain.Accounts.ItemByAddress[Email];
                    log = "Изменено:";
                }
                catch (Exception ex) {

                    hAccount = domain.Accounts.Add();
                    log = "Создано:";
                }                

                hAccount.Active = true;
                hAccount.Address = Email;
                hAccount.Password = Password;
                hAccount.MaxSize = Convert.ToInt32(numericUpDown1.Value);

                hAccount.Save();

                log += " "+Email+Environment.NewLine;
                txtLog.AppendText(log);
            }
        }

        private bool IsAuthHMail(hMailServer.Application hApp)
        {
            string hAdministrator = txtHAdmin.Text;
            string hPassword = txtHPassword.Text;

            hMailServer.Account authenticated = hApp.Authenticate(hAdministrator, hPassword);

            if (authenticated != null)
            {
                return true;
            }
            
            txtLog.AppendText("hAdministrator или hPassword неправильный!"+Environment.NewLine);

            return false;
        }
        
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (connection.State == ConnectionState.Open)
            {
                connection.Close();

                connection.Dispose();
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
