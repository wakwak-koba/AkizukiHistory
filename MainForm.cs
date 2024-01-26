using System.Linq;

namespace AkizukiHistory
{
    public partial class MainForm : System.Windows.Forms.Form
    {
        [System.Runtime.InteropServices.DllImport("wininet.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
        private static extern bool InternetGetCookieEx(string pchURL, string pchCookieName, System.Text.StringBuilder pchCookieData, ref uint pcchCookieData, int dwFlags, System.IntPtr lpReserved);

        const int chunk = 50;
        const string host = @"https://akizukidenshi.com";
        System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog() { Description = "購入履歴詳細の保存先を選択してください", ShowNewFolderButton = true };
        System.Text.Encoding encoding = System.Text.Encoding.UTF8;
        System.Collections.Generic.SortedDictionary<string, System.IO.FileInfo> files = new System.Collections.Generic.SortedDictionary<string, System.IO.FileInfo>();

        public System.IO.DirectoryInfo folder = null;
        public string userID = null;
        public string password = null;

        public MainForm()
        {
            InitializeComponent();

            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            this.Load += (sender, e) => webBrowser.Navigate(host + @"/catalog/customer/menu.aspx");
            this.Shown += (sender, e) => {
                if(folder != null)
                    dialog.SelectedPath = folder.FullName;
                else if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    this.Close();
                    return;
                }
            };
            webBrowser.DocumentCompleted += (sender, e) =>
            {
                var web = sender as System.Windows.Forms.WebBrowser;
                if (web.Document.Body.InnerText.Contains(@"メールアドレスとパスワードを入力してログインしてください。"))
                {
                    if(!string.IsNullOrEmpty(userID) && !string.IsNullOrEmpty(password))
                    {
                        web.Document.GetElementById("login_uid").SetAttribute("value", userID); 
                        web.Document.GetElementById("login_pwd").SetAttribute("value", password);
                        var button = web.Document.Forms.OfType<System.Windows.Forms.HtmlElement>().Where(f => f.InnerHtml.Contains(@"ログインする")).First().All.OfType<System.Windows.Forms.HtmlElement>().Where(elm => "order".Equals(elm.Name)).First();
                        button.InvokeMember("click");
                    }
                    return;
                }

                if (web.Document.Body.InnerText.Contains(@"お客様のログイン情報を解除いたしました。"))
                {
                    this.Close();
                    return;
                }

                progressBar.Visible = true;
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                for (int page = 1; ; page++)
                {
                    htmlDoc.LoadHtml(GetResponse(new System.Uri(host  + @"/catalog/customer/history.aspx?ps=" + chunk.ToString() + "&p=" + page.ToString())));
                     
                    if (int.TryParse(htmlDoc.DocumentNode.SelectSingleNode(@"//span[@class=""pager-count""]/span")?.InnerText, out var maxPage))
                        progressBar.Maximum = maxPage;
                    foreach (var uri in htmlDoc.DocumentNode.SelectNodes(@"//ul[@class=""block-purchase-history--order-detail-list""]/a").Select(n => new System.Uri(host + n.Attributes["href"].Value)).ToArray())
                    {
                        var response = GetResponse(uri);
                        var orderid = System.Web.HttpUtility.ParseQueryString(uri.Query)["order_id"];
                        var file = new System.IO.FileInfo(System.IO.Path.Combine(dialog.SelectedPath, orderid) + @".html");
                        System.IO.File.WriteAllText(file.FullName, response, encoding);
                        files.Add(orderid, file);
                        progressBar.Value++;
                    }
                    if (maxPage == 0 || page * chunk >= progressBar.Maximum) break;
                }

                // csv
                var csv = new System.IO.FileInfo(System.IO.Path.Combine(dialog.SelectedPath, @"購入履歴詳細.tsv"));
                var csvAppend = csv.Exists;
                using (var writer = new System.IO.StreamWriter(new System.IO.FileStream(csv.FullName, System.IO.FileMode.Create, System.IO.FileAccess.Write), System.Text.Encoding.UTF8 ))
                {
                    if(!csvAppend)
                        writer.WriteLine(string.Join("\t", new[] { @"オーダーＩＤ", @"注文日", @"出荷日", @"販売コード", @"商品名", @"数量", @"単位", @"金額", @"href", @"image" }));
                    htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    foreach (var file in files.Values)
                    {
                        htmlDoc.Load(file.FullName, encoding);
                        System.DateTime.TryParse(System.Web.HttpUtility.HtmlDecode(htmlDoc.DocumentNode.SelectSingleNode(@"//td[@class=""block-purchase-history-detail--order-dt""]").InnerText).Trim(), out var 受注日);
                        System.DateTime.TryParse(System.Web.HttpUtility.HtmlDecode(htmlDoc.DocumentNode.SelectSingleNode(@"//td[@class=""block-purchase-history-detail--ship-dt""]").InnerText), out var 出荷日);

                        var meisais = htmlDoc.DocumentNode.SelectNodes(@"//div[@class=""block-purchase-history-detail--order-body-left""]/table/tbody/tr")
                            .Select(n1 => n1.SelectNodes("td").Select(n2 => System.Net.WebUtility.HtmlDecode(n2.InnerText).Trim().Replace(System.Environment.NewLine, "\n")).ToArray())
                            .Select(n => string.Join("\t", new[] { System.IO.Path.GetFileNameWithoutExtension(file.Name), 受注日.ToShortDateString(), 出荷日.ToShortDateString(), n[0], n[1].Split('\n').First(), n[2].Replace(@"￥", "").Replace(@",", ""), Split単位(n[1].Split('\n').Last()), n[3].Replace(@"￥", "").Replace(@",", ""), host + @"/catalog/g/g" + n[0] + @"/", host + @"/img/goods/M/" + n[0] + @".jpg" }));
                        foreach (var meisai in meisais)
                            writer.WriteLine(meisai);
                    }
                    writer.Close();
                }
                if(string.IsNullOrEmpty(userID) && string.IsNullOrEmpty(password))
                    System.Diagnostics.Process.Start(dialog.SelectedPath);

                webBrowser.Navigate(host + @"/catalog/customer/logout.aspx");
            };
        }

        private string GetResponse(System.Uri uri)
        {
            string cookie = webBrowser.Document.Cookie;
            uint length = 1000;
            var cookieData = new System.Text.StringBuilder((int)length);
            if (InternetGetCookieEx(uri.AbsoluteUri, null, cookieData, ref length, 0x00002000,  System.IntPtr.Zero) && cookieData.Length > 0)
                cookie = cookieData.ToString();

            var webRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(uri);
            webRequest.Method = @"GET";
            webRequest.Headers.Add(@"Cookie", cookie);
            webRequest.Timeout = 5000;
            using (var response = webRequest.GetResponse())
            using (var streamreader = new System.IO.StreamReader(response.GetResponseStream(), encoding))
                return streamreader.ReadToEnd();
        }

        private string Split単位(string value)
        {
            if (!value.StartsWith("・"))
                return null;
            value = value.Substring(1);

            var res = new System.Collections.ObjectModel.Collection<char>();
            for(int i = 0; i < value.Length;i ++)
            {
                if (i > 0 && isNumeric(value[i - 1]) != isNumeric(value[i]))
                    res.Add('\t');
                res.Add(value[i]);
            }
            var str = new string(res.ToArray()).Split('\t');
            return str[1];
        }

        private bool isNumeric(char c) => (c >= '0' && c <= '9');
    }
}