using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using System.Net;
using Facebook;
using System.IO;

namespace ExcelAddIn1 {
    public partial class ThisAddIn {
        private void ThisAddIn_Startup(object sender, System.EventArgs e) {

            Excel.Workbook template = Application.Workbooks.Open(@"C:\Users\ksawery.jasienski\Documents\Template.xlsx");
            Excel.Worksheet sheet = template.ActiveSheet;
            Excel.Workbook book = Application.Workbooks.Add();
            sheet.Copy(Application.ActiveWorkbook.Sheets.Add());
            template.Close();
            sheet = book.ActiveSheet;

            string json = "{ }";
            FacebookClient client = new FacebookClient("EAACEdEose0cBAKXs12QMh54IdpETXfDTFeb8KYiGwbcq54uj4t9KUZBR78gFI0AUxNVkQZCB3jXVAB2pAZCR9fzSr0AOvviuZA3pKb15RkVDFYotZAgfGShoZAWqCsOi1uLF7P1Doovn0iLXLRTfnhTZAKmA69n5mvZCNL78A98KqwZDZD");

            string pageName = "accenture";
            int postsToLoad = 0;

            var pnd = new PageNameDialog();
            pnd.ShowDialog();
            pageName = pnd.GetPageName();
            postsToLoad = pnd.GetPostCount();
            if (pnd.DialogResult != DialogResult.OK || pageName.Length == 0) {
                book.Close(0);
                Application.Quit();
            }

            try {
                json = client.Get(pageName + "?fields=name,posts.limit(" + postsToLoad + "){id,message,likes.summary(true),comments.summary(true),shares, created_time}").ToString();

            } catch (FacebookOAuthException) {

                MessageBox.Show("Your OAuth key expired, get a new one!");
                Application.Quit();

            } catch (FacebookApiException) {
                MessageBox.Show("Najprawdopodobniej nie ma strony na Facebooku o takim id!");
                pnd.ShowDialog();
                pageName = pnd.GetPageName();
                postsToLoad = pnd.GetPostCount();
                if (pnd.DialogResult != DialogResult.OK || pageName.Length == 0) {
                    book.Close(0);
                    Application.Quit();
                }
                json = client.Get(pageName + "?fields=name,posts.limit(" + postsToLoad + "){id,message,likes.summary(true),comments.summary(true),shares, created_time}").ToString();
            }

            if (json.Length < 5) Application.Quit();

            var obj = JObject.Parse(json);

            JArray posts = JArray.Parse(obj["posts"]["data"].ToString());

            postsToLoad = Math.Min(postsToLoad, posts.Count);
            sheet.Cells[4, 4] = obj["name"];

            for (int i = 0; i < postsToLoad; i++) {
                var post = posts[i];
                var likes = post["likes"];
                var shares = post["shares"];
                var comments = post["comments"];

                sheet.Cells[8 + i, 2] = i + 1;
                sheet.Cells[8 + i, 3] = post["created_time"].ToString().Substring(0, 11);
                sheet.Cells[8 + i, 4] = post["message"];
                sheet.Cells[8 + i, 5] = likes != null ? likes["summary"]["total_count"] : 0;
                sheet.Cells[8 + i, 6] = comments != null ? comments["summary"]["total_count"] : 0;
                sheet.Cells[8 + i, 7] = shares != null ? shares["count"] : 0;
            }

            Excel.Range chartRange;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = xlCharts.Add(627, 10, 360, 340);
            Excel.Chart chartPage = myChart.Chart;
            chartPage.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;

            chartRange = sheet.get_Range("e7", "g" + (postsToLoad + 7));
            chartPage.SetSourceData(chartRange, System.Reflection.Missing.Value);
            chartPage.ChartType = Excel.XlChartType.xlColumnStacked;


        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

    public class MyWebRequest {
        private WebRequest request;
        private Stream dataStream;

        private string status;

        public String Status {
            get {
                return status;
            }
            set {
                status = value;
            }
        }

        public MyWebRequest(string url) {
            // Create a request using a URL that can receive a post.

            request = WebRequest.Create(url);
        }

        public MyWebRequest(string url, string method)
            : this(url) {

            if (method.Equals("GET") || method.Equals("POST")) {
                // Set the Method property of the request to POST.
                request.Method = method;
            } else {
                throw new Exception("Invalid Method Type");
            }
        }

        public MyWebRequest(string url, string method, string data)
            : this(url, method) {

            // Create POST data and convert it to a byte array.
            string postData = data;
            byte[] byteArray = Encoding.UTF8.GetBytes(postData);

            // Set the ContentType property of the WebRequest.
            request.ContentType = "application/x-www-form-urlencoded";

            // Set the ContentLength property of the WebRequest.
            request.ContentLength = byteArray.Length;

            // Get the request stream.
            dataStream = request.GetRequestStream();

            // Write the data to the request stream.
            dataStream.Write(byteArray, 0, byteArray.Length);

            // Close the Stream object.
            dataStream.Close();

        }

        public string GetResponse() {
            // Get the original response.
            WebResponse response = request.GetResponse();

            this.Status = ((HttpWebResponse)response).StatusDescription;

            // Get the stream containing all content returned by the requested server.
            dataStream = response.GetResponseStream();

            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);

            // Read the content fully up to the end.
            string responseFromServer = reader.ReadToEnd();

            // Clean up the streams.
            reader.Close();
            dataStream.Close();
            response.Close();

            return responseFromServer;
        }

    }
}