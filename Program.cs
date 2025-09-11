using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;
using System.Text.RegularExpressions;
using DotNetEnv;
using Microsoft.Data.SqlClient;
using System;
using ReadMail.Models;
using MongoDB.Driver;
using MongoDB.Bson;
class Program
{
    static GraphServiceClient? graphClient;
    static IPublicClientApplication? pca;
    static string[]? scopes;

    static async Task Main(string[] args)
    {
        Env.Load();
        var clientId = Env.GetString("CLIENT_ID");
        var redirectUri = Env.GetString("REDIRECT_URI");
        scopes = Env.GetString("SCOPES").Split(',');

        pca = PublicClientApplicationBuilder
            .Create(clientId)
            .WithRedirectUri(redirectUri)
            .WithAuthority("https://login.microsoftonline.com/consumers")
            .WithCacheOptions(CacheOptions.EnableSharedCacheOptions) // เปิด session cache
            .Build();

        var accessToken = await GetTokenAsync();

        graphClient = new GraphServiceClient(new BaseBearerTokenAuthenticationProvider(
            new SimpleAccessTokenProvider(accessToken)));

        var timer = new System.Timers.Timer(60000); // 1 นาที
        timer.Elapsed += async (sender, e) =>
        {
            var token = await GetTokenAsync();
            graphClient = new GraphServiceClient(new BaseBearerTokenAuthenticationProvider(
                new SimpleAccessTokenProvider(token)));
            await FetchMail();
        };
        timer.AutoReset = true;
        timer.Enabled = true;

        Console.WriteLine("เริ่มดึงอีเมลทุก 1 นาที กด Enter เพื่อออก...");
        TestLocalDbConnection();
        await FetchMail(); // ดึงครั้งแรกทันที
        Console.ReadLine();
    }

    static async Task<string> GetTokenAsync()
    {
        AuthenticationResult result = null;
        if (pca == null || scopes == null) throw new InvalidOperationException("PCA หรือ Scopes ยังไม่ถูกตั้งค่า");
        var accounts = await pca.GetAccountsAsync();
        
        try
        {
            // พยายามดึง token จาก cache ก่อน
            result = await pca.AcquireTokenSilent(scopes, accounts.FirstOrDefault()).ExecuteAsync();
        }
        catch (MsalUiRequiredException)
        {
            // ถ้าไม่มีใน cache หรือหมดอายุ ให้ login ใหม่
            result = await pca.AcquireTokenInteractive(scopes).ExecuteAsync();
        }
        return result.AccessToken;
    }

    static async Task FetchMail()
    {
        try
        {
            if (graphClient == null)
            {
                Console.WriteLine("Graph client ยัง Login ไม่สำเร็จ");
                return;
            }
            var messages = await graphClient.Me.Messages.GetAsync(config =>
            {
                config.QueryParameters.Filter = "isRead eq false";
                config.QueryParameters.Orderby = new[] { "receivedDateTime desc" };
                config.QueryParameters.Top = 10;
                config.QueryParameters.Select = new[] { "subject", "receivedDateTime", "bodyPreview", "from", "body", "internetMessageId", "id" };
            });

            if (messages == null || messages.Value == null || messages.Value.Count == 0)
            {
                Console.WriteLine("ไม่พบอีเมล");
                Console.WriteLine("--------------------------------------------------------------------------------------------------------");
                return;
            }

            var filters = messages?.Value?.Where(x => x.Subject != null && x.Subject.Contains("คุณได้รับเงินผ่านรายการพร้อมเพย์")).ToList();
            if (filters == null || filters.Count == 0)
            {
                Console.WriteLine("ไม่พบอีเมลที่ตรงกับเงื่อนไข");
                Console.WriteLine("--------------------------------------------------------------------------------------------------------");
                return;
            }
            
            foreach (var message in filters)
            {
                //var plainText = StripHtml(message?.Body?.Content);

                Console.WriteLine($"Subject: {message.Subject}");
                Console.WriteLine($"Received: {message.ReceivedDateTime}");
                Console.WriteLine($"Form: {message?.From?.EmailAddress?.Name}");
                Console.WriteLine($"MessageId: {message?.InternetMessageId}");
                Console.WriteLine($"Id: {message?.Id}");
                //Console.WriteLine($"plainText : {plainText}");

                var database = TestMongoDbConnection();
                var scb = database.GetCollection<BsonDocument>("scb");
                var existing = await scb.Find(new BsonDocument { { "emailId", message?.Id } }).FirstOrDefaultAsync();

                if (existing != null)
                {

                    Console.WriteLine("อีเมลนี้ถูกบันทึกแล้ว ข้ามการประมวลผล");
                    Console.WriteLine("--------------------------------------------------------------------------------------------------------");
                    //return; // ข้ามอีเมลนี้
                    continue; 
                }   
                
                Console.WriteLine("อีเมลนี้ยังไม่ถูกบันทึก ดำเนินการประมวลผลต่อ");

                var contentType = message?.Body?.ContentType;
                var rawBody = message?.Body?.Content;
                var input = "";

                if (contentType == BodyType.Text)
                {
                    //Console.WriteLine("Body (Text):");
                    //Console.WriteLine(rawBody);
                    input = rawBody;
                }
                else if (contentType == BodyType.Html)
                {
                    input = StripHtml(rawBody);
                    //Console.WriteLine(plainText);
                }

                //Console.WriteLine("plainText : ", input);

                var notification = ParseNotification(input);

                Console.WriteLine($"Recipient: {notification.Recipient}");
                Console.WriteLine($"Amount: {notification.Transaction.AmountBaht} บาท");
                Console.WriteLine($"From Bank: {notification.Transaction.From.Bank}");
                Console.WriteLine($"From Account: {notification.Transaction.From.Account}");
                Console.WriteLine($"To Account: {notification.Transaction.ToAccount}");
                Console.WriteLine($"DateTime: {notification.Transaction.DateTime}");

                var doc = new BsonDocument
                {
                    { "recipient", notification.Recipient },
                    { "amount", notification.Transaction.AmountBaht },
                    { "fromBank", notification.Transaction.From.Bank },
                    { "fromAccount", notification.Transaction.From.Account },
                    { "toAccount", notification.Transaction.ToAccount },
                    { "datetime", notification.Transaction.DateTime },
                    { "messageId", message?.InternetMessageId },
                    { "emailId", message?.Id },
                    { "createdAt", DateTime.UtcNow }
                };
                scb.InsertOne(doc);
                //var collection = database.GetCollection<BsonDocument>("scb");
                await graphClient.Me.Messages[message.Id].PatchAsync(new Message
                {
                    IsRead = true
                });
                Console.WriteLine("--------------------------------------------------------------------------------------------------------");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
    public class SimpleAccessTokenProvider : IAccessTokenProvider
    {
        private readonly string _accessToken;
        public SimpleAccessTokenProvider(string accessToken)
        {
            _accessToken = accessToken;
        }
        public Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object>? context = null, CancellationToken cancellationToken = default)
        {
            return Task.FromResult(_accessToken);
        }
        public AllowedHostsValidator AllowedHostsValidator => new AllowedHostsValidator();
    }
    public static string StripHtml(string? html)
    {
        if (string.IsNullOrEmpty(html)) return string.Empty;
        return Regex.Replace(html, "<.*?>", string.Empty);
    }
    static string GetLocalDbConnectionString()
    {
        // ตัวอย่าง connection string สำหรับ localDB
        return @"Server=(localdb)\MSSQLLocalDB;Database=TestDBA;Trusted_Connection=True;";
    }
    static void TestLocalDbConnection()
    {
        var connStr = GetLocalDbConnectionString();
        using (var conn = new SqlConnection(connStr))
        {
            conn.Open();
            Console.WriteLine("เชื่อมต่อ localDB สำเร็จ!");
            // ตัวอย่าง query
            using (var cmd = new SqlCommand("SELECT GETDATE()", conn))
            {
                var result = cmd.ExecuteScalar();
                Console.WriteLine($"เวลาปัจจุบันในฐานข้อมูล: {result}");
            }
        }
    }
    static IMongoDatabase TestMongoDbConnection()
    {
        var connStr = Env.GetString("MONGODB_CONNECTION_STRING");
        var client = new MongoClient(connStr);
        var database = client.GetDatabase("paymentdb");
        //var collection = database.GetCollection<BsonDocument>("scb");
        Console.WriteLine("เชื่อมต่อ MongoDB สำเร็จ!");
        return database;
        // ตัวอย่าง insert document
        //var doc = new BsonDocument { { "test", DateTime.Now } };
        //collection.InsertOne(doc);

        // ตัวอย่าง query
        //var count = collection.CountDocuments(new BsonDocument());
        //Console.WriteLine($"จำนวนเอกสารใน TestCollection: {count}");
    }
    public static PromptPayNotification ParseNotification(string? text)
    {
        var recipient = Regex.Match(text, @"เรียน\s+(.*?)\s+คุณได้รับ").Groups[1].Value.Trim();
        var bank = Regex.Match(text, @"จาก:\s*(\w+)").Groups[1].Value.Trim();
        var fromAccount = Regex.Match(text, @"/\s*(\w+)จำนวน").Groups[1].Value.Trim();
        var amount = decimal.Parse(Regex.Match(text, @"จำนวน \(บาท\):\s*(.*?)เข้าบัญชี").Groups[1].Value);
        var toAccount = Regex.Match(text, @"เข้าบัญชี:\s*(\w+)").Groups[1].Value.Trim().Replace("วัน", "");
        var datetime = Regex.Match(text, @"วัน/เวลา:\s*(.*?)ขอแสดงความนับถือธนาคารไทยพาณิชย์").Groups[1].Value.Trim();

        return new PromptPayNotification
        {
            Recipient = recipient,
            Transaction = new TransactionDetail
            {
                Method = "พร้อมเพย์",
                From = new BankAccount
                {
                    Bank = bank,
                    Account = fromAccount
                },
                AmountBaht = amount,
                ToAccount = toAccount,
                DateTime = datetime
            }
        };
    }

}