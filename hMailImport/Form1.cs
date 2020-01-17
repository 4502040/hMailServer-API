using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace hMailImport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            hMailServer.Application newApp = new hMailServer.Application();
            hMailServer.Account authenticated = newApp.Authenticate("Administrator", "password");

            if (authenticated != null)
            {
                hMailServer.Domain domain = newApp.Domains[0];

                hMailServer.Account newAccount = domain.Accounts.Add();

                newAccount.Active = true;
                newAccount.Address = "test@osc.kz";
                newAccount.Password = "test123456";
                newAccount.MaxSize = 20;

                newAccount.Save();
            }
        }
    }
}
