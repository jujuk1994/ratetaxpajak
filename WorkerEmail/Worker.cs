using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WDSE;
using WDSE.Decorators;
using WDSE.ScreenshotMaker;
using System.Drawing;
using System.Net.Http;
using System.IO;
using Newtonsoft.Json;
using RestSharp;
using System.Net;
using System.Text.RegularExpressions;
using Telegram.Bot;
using Telegram.Bot.Types.Enums;

namespace WorkerEmail
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        static TelegramBotClient Bot = new TelegramBotClient("5478187618:AAENfPcaia3OMwc3alj57qil0uN7JrPFPP4");

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }
        public override Task StartAsync(CancellationToken cancellationToken)
        {
            return base.StartAsync(cancellationToken);
        }
        public override Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Service stopped");
            return base.StopAsync(cancellationToken);
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    if (DateTime.Now.Hour == 15)
                    {
                        IWebDriver ChromeDriver2 = new ChromeDriver();
                        ChromeDriver2.Manage().Window.Maximize();
                        ChromeDriver2.Navigate().GoToUrl("https://fiskal.kemenkeu.go.id/informasi-publik/kurs-pajak?date=" + DateTime.Now.ToString("yyyy-MM-dd"));

                        IWebElement table2 = ChromeDriver2.FindElement(By.XPath("//table"));

                        var html_text = table2.GetAttribute("innerHTML");
                        string allText = html_text;//File.ReadAllText(pathHTML);
                        var htmlremove = Regex.Replace(allText, "<.*?>", String.Empty);
                        var data = htmlremove.Split(new[] { "\r\n      \r\n\t\t\t\r\n\t\t\t\t\r\n\t\t\t\t" }, StringSplitOptions.None);
                        string[] tax = data[1].Split(new[] { "\r\n\t\t\t\r\n\t\t\t" }, StringSplitOptions.None);

                        ChromeDriver2.Quit();
                        var dataset = new DataSet1TableAdapters.RateTaxTableAdapter();
                        dataset.Insert(DateTime.Now.Date, Convert.ToDecimal(tax[0].Replace(".", "")));
                        sendTelegram("-778112650", "Kurs pajak hari ini Rp. " + tax[0].ToString());

                        _logger.LogInformation("Sukses");
                        monitoringServices("DOP_GetRateTax", "Insert rate pajak", "10.10.10.99", "Live");

                    }

                }
                catch (Exception ex)
                {
                    _logger.LogInformation("Worker Eror : " + ex.Message + " {time}", DateTimeOffset.Now);
                    sendTelegram("-778112650", "Eror kurs pajak hari ini" + ex.Message);
                    monitoringServices("DOP_GetRateTax", "Insert rate pajak eror "+ex.Message, "10.10.10.99", "Eror");
                }

                await Task.Delay(3600000, stoppingToken);
            }
        }
        private static void sendFileTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;

            var client = new RestClient("https://api.telegram.org/bot2144239635:AAFjcfn_GdHP4OkzzZomaZt4XbwpHDGyR-U/sendDocument");
            RestRequest requestWa = new RestRequest("https://api.telegram.org/bot2144239635:AAFjcfn_GdHP4OkzzZomaZt4XbwpHDGyR-U/sendDocument", Method.Post);


            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "multipart/form-data");
            requestWa.AddParameter("chat_id", chatId);
            requestWa.AddFile("document", body);
            var responseWa = client.ExecutePostAsync(requestWa);
            Console.WriteLine(responseWa.Result.Content);
        }
        private static string SendFile(string chatId, string data, string caption)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/sendFile?token=jkdjtwjkwq2gfkac");


            RestRequest requestWa = new RestRequest("https://api.chat-api.com/instance127354/sendFile?token=jkdjtwjkwq2gfkac", Method.Post);

            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("chatId", chatId);
            requestWa.AddParameter("filename", "1.png");
            requestWa.AddParameter("body", data);
            requestWa.AddParameter("caption", caption);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }
        public class MessageChat
        {
            public string id { get; set; }
            public string body { get; set; }
            public string fromMe { get; set; }
            public string self { get; set; }
            public string isForwarded { get; set; }
            public string author { get; set; }
            public double time { get; set; }
            public string chatId { get; set; }
            public int messageNumber { get; set; }
            public string type { get; set; }
            public string senderName { get; set; }
            public string caption { get; set; }
            public string quotedMsgBody { get; set; }
            public string quotedMsgId { get; set; }
            public string quotedMsgType { get; set; }
            public string chatName { get; set; }
        }
        public class ResponseChat
        {
            public IEnumerable<MessageChat> messages { get; set; }
            public int lastMessageNumber { get; set; }
        }
        private static string SendMessage(string chatId, string body)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/sendMessage?token=jkdjtwjkwq2gfkac");

            RestRequest requestWa = new RestRequest("https://api.chat-api.com/instance127354/sendMessage?token=jkdjtwjkwq2gfkac", Method.Post);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("phone", chatId);
            requestWa.AddParameter("body", body);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }
        private static string GetMessageList(string chatId, int lastMessageNumber)
        {
            string url = "https://api.chat-api.com/instance127354/messages?token=jkdjtwjkwq2gfkac&lastMessageNumber=" + lastMessageNumber + "&chatId=" + chatId;
            var client = new RestClient(url);
            var requestWa = new RestRequest(url, Method.Get);
            var responseWa = client.ExecuteGetAsync(requestWa);
            return responseWa.Result.Content;

        }
        private static void sendTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;

            var client = new RestClient("https://api.telegram.org/bot5478187618:AAENfPcaia3OMwc3alj57qil0uN7JrPFPP4/sendMessage?chat_id=" + chatId + "&text=" + body);
            RestRequest requestWa = new RestRequest("https://api.telegram.org/bot5478187618:AAENfPcaia3OMwc3alj57qil0uN7JrPFPP4/sendMessage?chat_id=" + chatId + "&text=" + body, Method.Get);
            requestWa.Timeout = -1;
            var responseWa = client.ExecutePostAsync(requestWa);
            Console.WriteLine(responseWa.Result.Content);
        }
        private static string monitoringServices(string servicename, string servicedescription, string servicelocation, string appstatus)
        {
            string jsonString = "{" +
                                "\"service_name\" : \"" + servicename + "\"," +
                                "\"service_description\": \"" + servicedescription + "\"," +
                                "\"service_location\":\"" + servicelocation + "\"," +
                                "\"app_status\":\"" + appstatus + "\"," +
                                "}";
            var client = new RestClient("http://10.10.10.99:84/api/ServiceStatus");

            RestRequest requestWa = new RestRequest("http://10.10.10.99:84/api/ServiceStatus", Method.Post);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("data", jsonString);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }

    }
}
