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

        hMailServer.Application hApp ;

        public Form1()
        {
            InitializeComponent();

            hApp = new hMailServer.Application();
        }        
        

        private bool IsAuthHMail()
        {
            string hAdministrator = "Administrator";
            string hPassword = "gedeon4j";

            hMailServer.Account authenticated = hApp.Authenticate(hAdministrator, hPassword);

            if (authenticated != null)
            {
                return true;
            }
            
            Console.WriteLine("hAdministrator или hPassword неправильный!");

            return false;
        }
        
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }

           

        private void buttonUpdateList_Click(object sender, EventArgs e)
        {
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (IsAuthHMail())
            {
                var domain = hApp.Domains[0];

                var openFile = new OpenFileDialog();

                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    if (openFile.FileName != null)
                    {
                        Console.WriteLine($"Выбрали файл: {openFile.FileName}" + Environment.NewLine);

                        var workbook = new XLWorkbook(openFile.FileName);


                        //foreach (var xlSheet in workbook.Worksheets.Where(q=>q.Name.Equals("Форма табеля №1")).ToList())
                        foreach (var xlSheet in workbook.Worksheets.Where(q => q.Name.Contains("табеля")).ToList())
                        {
                            Console.WriteLine(xlSheet.Name);

                            var rows = xlSheet.RangeUsed().RowsUsed(); // Skip header row

                            int i = 0;

                            foreach (var row in rows)
                            {
                                i++;

                                string v = Convert.ToString(row.Cell(5).Value);

                                if (v == "") continue;

                                if (v.Contains("@"))
                                {
                                    if (i < rows.Count() && rows.ElementAt(i) != null)
                                    {
                                        string p = Convert.ToString(rows.ElementAt(i).Cell(5).Value);

                                        try
                                        {
                                            Random rnd = new Random();
                                            int part1 = rnd.Next(100, 999);
                                            int part2 = rnd.Next(100, 999);

                                            string newPass = $"{p.Substring(0, 1)}{part1.ToString()}{p.Substring(4, 1)}{part2}";

                                            var hAccount = domain.Accounts.ItemByAddress[v];

                                            hAccount.Password = newPass;

                                            hAccount.Save();

                                            Console.WriteLine($" Email: {v}, OldPassword: {p}, NewPassword; {newPass}");

                                            rows.ElementAt(i).Cell(5).Value = newPass;
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Email {v} not found on mail.osc.kz domain");
                                        }
                                        
                                    }


                                }
                                else continue;




                            }
                        }

                        workbook.Save();

                        Console.WriteLine("Файл обновлен успешно (записали новые пароли)");
                        Console.WriteLine("Можете закрывать программку");
                    }
                };
            }

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

        }
    }
}
