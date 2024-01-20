using System.Linq;

namespace AkizukiHistory
{
    public partial class MainForm : System.Windows.Forms.Form
    {
        const int chunk = 50;
        const string host = @"https://akizukidenshi.com";
        System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog() { Description = "購入履歴詳細の保存先を選択してください", ShowNewFolderButton = true };
        System.Text.Encoding encoding = System.Text.Encoding.GetEncoding("sjis");
        System.Collections.Generic.SortedDictionary<string, System.IO.FileInfo> files = new System.Collections.Generic.SortedDictionary<string, System.IO.FileInfo>();
        public MainForm()
        {
            InitializeComponent();

            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            this.Load += (sender, e) => webBrowser.Navigate(host + @"/catalog/customer/menu.aspx");
            this.Shown += (sender, e) => {
                if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    this.Close();
                    return;
                }
            };
            webBrowser.DocumentCompleted += (sender, e) =>
            {
                var web = sender as System.Windows.Forms.WebBrowser;
                if (web.Document.Body.InnerText.Contains(@"メールアドレスとパスワードを入力してログインしてください。"))
                    return;

                progressBar.Visible = true;
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                for (int page = 1; ; page++)
                {
                    htmlDoc.LoadHtml(GetResponse(new System.Uri(host  + @"/catalog/customer/history.aspx?ps=" + chunk.ToString() + "&p=" + page.ToString())));

                    if (int.TryParse(htmlDoc.DocumentNode.SelectSingleNode(@"//div[@class=""navipage_""]/b").InnerText, out var maxPage))
                        progressBar.Maximum = maxPage;
                    foreach (var uri in htmlDoc.DocumentNode.SelectNodes(@"//td[@class=""order_id_ order_detail_""]/a").Select(n => new System.Uri(host + n.Attributes["href"].Value)).ToArray())
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
                using (var writer = new System.IO.StreamWriter(new System.IO.FileStream(csv.FullName, System.IO.FileMode.Create, System.IO.FileAccess.Write), System.Text.Encoding.UTF8 ))
                {
                    writer.WriteLine(string.Join("\t", new[] { @"オーダーＩＤ", @"注文日", @"出荷日", @"通販コード", @"商品名", @"数量", @"単位", @"金額", @"href", @"image" }));
                    htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    foreach (var file in files.Values)
                    {
                        htmlDoc.Load(file.FullName, encoding);
                        System.DateTime.TryParse(System.Web.HttpUtility.HtmlDecode(htmlDoc.DocumentNode.SelectSingleNode(@"//tr[@class=""table_top_""]/td").ChildNodes[1].InnerText).Trim(), out var 受注日);
                        System.DateTime.TryParse(System.Web.HttpUtility.HtmlDecode(htmlDoc.DocumentNode.SelectNodes(@"//table[@class=""history_order_ top_""]/tr")[1].ChildNodes[1].ChildNodes[1].InnerText), out var 出荷日);

                        foreach (var meisai in htmlDoc.DocumentNode.SelectNodes(@"//table[@class=""history_loop_""]/tr").Skip(1).Select(n => string.Join("\t", new[] { System.IO.Path.GetFileNameWithoutExtension(file.Name), 受注日.ToShortDateString(), 出荷日.ToShortDateString(), n.ChildNodes[1].ChildNodes[0].InnerText, n.ChildNodes[3].ChildNodes[0].InnerText, n.ChildNodes[5].InnerText.Substring(0, n.ChildNodes[5].InnerText.Length - Split単位(n.ChildNodes[5].InnerText).Length).Replace(@",", ""), Split単位(n.ChildNodes[5].InnerText), n.ChildNodes[7].InnerText.Replace(@"￥", "").Replace(@",", ""), host + n.ChildNodes[1].ChildNodes[0].Attributes[0].Value, host + n.ChildNodes[3].ChildNodes[0].ChildNodes[0].Attributes[0].Value })))
                            writer.WriteLine(meisai);
                    }
                    writer.Close();
                }
                System.Diagnostics.Process.Start(dialog.SelectedPath);
                this.Close();
            };
        }
        private string GetResponse(System.Uri uri)
        {
            var webRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(uri);
            webRequest.Method = @"GET";
            webRequest.Headers.Add(@"Cookie", webBrowser.Document.Cookie);
            webRequest.Timeout = 5000;
            using (var response = webRequest.GetResponse())
            using (var streamreader = new System.IO.StreamReader(response.GetResponseStream(), encoding))
                return streamreader.ReadToEnd();
        }

        private string Split単位(string value) => new string(value.Reverse().TakeWhile(c => !(c >= '0' && c <= '9')).Reverse().ToArray());
    }
}