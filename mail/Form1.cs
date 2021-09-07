using System;
using System.Web;
using System.Net;
using System.Windows.Forms;
using xNet;
using HttpRequest = xNet.HttpRequest;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;

namespace mail
{
    public partial class mail : Form
    {
        public mail()
        {
            InitializeComponent();
        }



        void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")

            { MessageBox.Show("Введи ID"); }

            else
            {

                var date = DateTime.Now.ToString("dd.MM.yyyy");
                var orderid = textBox1.Text;

                HttpRequest http = new HttpRequest();

                http.Cookies = new CookieDictionary();
                http.KeepAlive = true;
                http.UserAgent = Http.ChromeUserAgent();
                http.AddHeader("Origin", "http://crm.iiko.ru");
                http.AddHeader("Referer", "http://crm.iiko.ru/index.php");
                http.AddHeader("Upgrade-Insecure-Requests", "1");


                RequestParams reqParams = new RequestParams();
                reqParams["module"] = "Users";
                reqParams["action"] = "Authenticate";
                reqParams["return_module"] = "Users";
                reqParams["return_action"] = "Login";
                reqParams["user_name"] = "aamelchenko";
                reqParams["user_password"] = "Gtht0306";
                reqParams["login_theme"] = "softed";
                reqParams["login_language"] = "ru_ru";
                reqParams["Login.x"] = "75";
                reqParams["Login.y"] = "14";


                string response = http.Post("http://crm.iiko.ru/index.php", reqParams, true).ToString();

                RequestParams reqParams2 = new RequestParams();
                reqParams["module"] = "Users";
                reqParams["action"] = "Authenticate";
                reqParams["return_module"] = "Users";
                reqParams["return_action"] = "Login";
                reqParams["user_name"] = "aamelchenko";
                reqParams["user_password"] = "Gtht0306";
                reqParams["login_theme"] = "softed";
                reqParams["login_language"] = "ru_ru";
                reqParams["Login.x"] = "75";
                reqParams["Login.y"] = "14";
                reqParams["action"] = "DetailView";
                reqParams["module"] = "Invoice";
                reqParams["record"] = "1901701";
                reqParams["parenttab"] = "Marketing";

                string url = $"http://crm.iiko.ru/index.php?action=DetailView&module=Invoice&record={orderid}&parenttab=Sales";



                string response2 = http.Post(url, reqParams2, true).ToString();





                var txtHTML = response2;
                var txtPrefix = @"<span id =""dtlview_Номер счета"">";
                var txtSuffix = @"</span>";
                var txtPrefixPosition = txtHTML.IndexOf(txtPrefix, StringComparison.OrdinalIgnoreCase);
                var txtSuffixPosition = txtHTML.IndexOf(txtSuffix, txtPrefixPosition + txtPrefix.Length, StringComparison.OrdinalIgnoreCase);

                string schet = txtHTML.Substring(
                    txtPrefixPosition + txtPrefix.Length,
                    txtSuffixPosition - txtPrefixPosition - txtPrefix.Length
                );

                var txtHTML2 = response2;
                var txtPrefix2 = @"title='Firm'>";
                var txtSuffix2 = @"</a></td>";
                var txtPrefixPosition2 = txtHTML2.IndexOf(txtPrefix2, StringComparison.OrdinalIgnoreCase);
                var txtSuffixPosition2 = txtHTML2.IndexOf(txtSuffix2, txtPrefixPosition2 + txtPrefix2.Length, StringComparison.OrdinalIgnoreCase);

                string urlic = txtHTML2.Substring(
                    txtPrefixPosition2 + txtPrefix2.Length,
                    txtSuffixPosition2 - txtPrefixPosition2 - txtPrefix2.Length
                );

                var txtHTML3 = response2;
                var txtPrefix3 = @"<td width=25% class=""dvtCellInfo"" align=""left"">&nbsp;<a href=""index.php?module=Accounts&action=DetailView&record=";
                var txtSuffix3 = @""">";
                var txtPrefixPosition3 = txtHTML3.IndexOf(txtPrefix3, StringComparison.OrdinalIgnoreCase);
                var txtSuffixPosition3 = txtHTML3.IndexOf(txtSuffix3, txtPrefixPosition3 + txtPrefix3.Length, StringComparison.OrdinalIgnoreCase);

                string crmid = txtHTML3.Substring(
                    txtPrefixPosition3 + txtPrefix3.Length,
                    txtSuffixPosition3 - txtPrefixPosition3 - txtPrefix3.Length
                );


                var txtHTML4 = response2;
                var txtPrefix4 = @"<td width=25% class=""dvtCellInfo"" align=""left"">&nbsp;<a href=""index.php?module=Accounts&action=DetailView&record=" + crmid + @""">";
                var txtSuffix4 = @"</a>";
                var txtPrefixPosition4 = txtHTML3.IndexOf(txtPrefix4, StringComparison.OrdinalIgnoreCase);
                var txtSuffixPosition4 = txtHTML3.IndexOf(txtSuffix4, txtPrefixPosition4 + txtPrefix4.Length, StringComparison.OrdinalIgnoreCase);

                string kontragent = txtHTML4.Substring(
                    txtPrefixPosition4 + txtPrefix4.Length,
                    txtSuffixPosition4 - txtPrefixPosition4 - txtPrefix4.Length
                );

                var txtHTML5 = response2;
                var txtPrefix5 = @"nbsp;<a href='index.php?module=Firm&action=DetailView&record=";
                var txtSuffix5 = @"' title='Firm'>";
                var txtPrefixPosition5 = txtHTML5.IndexOf(txtPrefix5, StringComparison.OrdinalIgnoreCase);
                var txtSuffixPosition5 = txtHTML5.IndexOf(txtSuffix5, txtPrefixPosition5 + txtPrefix5.Length, StringComparison.OrdinalIgnoreCase);

                string urlicid = txtHTML5.Substring(
                    txtPrefixPosition5 + txtPrefix5.Length,
                    txtSuffixPosition5 - txtPrefixPosition5 - txtPrefix5.Length
                );

                string learn = $"Полная отгрузка";
                if (response2.Contains("Director экспресс") || response2.Contains("Курс iikoChain") || response2.Contains("Консультационные услуги"))
                    learn = $"Частичная отгрузка, без обучения";


                



                string urlurlic = $"http%3A%2F%2Fcrm.iiko.ru%2Findex.php%3Fmodule%3DFirm%26action%3DDetailView%26record%3D{urlicid}";
                string urlkontr = $"http%3A%2F%2Fcrm.iiko.ru%2Findex.php%3Fmodule%3DAccounts%26action%3DDetailView%26record%3D{crmid}";
                string urlschet = $"http%3A%2F%2Fcrm.iiko.ru%2Findex.php%3Faction%3DDetailView%26module%3DInvoice%26record%3D{orderid}";




                Process.Start($"mailto:buh_kaz@iiko.ru?subject=Закрывающие%20документы%20{urlic}&body=Добрый%20день.%20Прошу%20подготовить%20закрывающие%20документы%0D%0A%0D%0AСчет:%20{schet}%20-%20{urlschet}%20%0D%0AЮр.лицо:%20{urlic}%20-%20{urlurlic}%20%0D%0AКонтрагент:%20{kontragent}%20-%20{urlkontr}%0D%0A%0D%0A{learn}%0D%0A%0D%0AДата%20отгрузки%20{date}%0D%0AСпасибо.");
            }
        }

        private void mail_Load(object sender, EventArgs e)
        {
            notifyIcon1.BalloonTipTitle = "Генерирователь писем все еще работает";
            notifyIcon1.BalloonTipText = "Теперь он в трее";
            notifyIcon1.Text = "Генерирователь писем";
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
            notifyIcon1.Visible = false;
            WindowState = FormWindowState.Normal;
        }

        private void mail_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                notifyIcon1.Visible = true;
                notifyIcon1.ShowBalloonTip(1000);
            }
            else if (FormWindowState.Normal == this.WindowState)
            { notifyIcon1.Visible = false; }
        }
    }
}

