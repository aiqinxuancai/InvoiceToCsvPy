namespace InvoiceToCsvSharp
{
    using CsvHelper;
    using CsvHelper.Configuration;
    using Newtonsoft.Json;
    using System;
    using System.Collections.Generic;
    using System.Formats.Asn1;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;

    // 发票数据模型
    public class InvoiceData
    {
        [CsvHelper.Configuration.Attributes.Name("发票代码")]
        [JsonProperty("发票代码")]
        public string InvoiceCode { get; set; }

        [CsvHelper.Configuration.Attributes.Name("发票号码")]
        [JsonProperty("发票号码")]
        public string InvoiceNumber { get; set; }

        [CsvHelper.Configuration.Attributes.Name("销方识别号")]
        [JsonProperty("销方识别号")]
        public string SellerTaxId { get; set; }

        [CsvHelper.Configuration.Attributes.Name("销方名称")]
        [JsonProperty("销方名称")]
        public string SellerName { get; set; }

        [CsvHelper.Configuration.Attributes.Name("购方识别号")]
        [JsonProperty("购方识别号")]
        public string BuyerTaxId { get; set; }

        [CsvHelper.Configuration.Attributes.Name("购买方名称")]
        [JsonProperty("购买方名称")]
        public string BuyerName { get; set; }

        [CsvHelper.Configuration.Attributes.Name("开票日期")]
        [JsonProperty("开票日期")]
        public string IssueDate { get; set; }

        [CsvHelper.Configuration.Attributes.Name("项目名称")]
        [JsonProperty("项目名称")]
        public string ItemName { get; set; }

        [CsvHelper.Configuration.Attributes.Name("数量")]
        [JsonProperty("数量")]
        public string Quantity { get; set; }

        [CsvHelper.Configuration.Attributes.Name("金额")]
        [JsonProperty("金额")]
        public string Amount { get; set; }

        [CsvHelper.Configuration.Attributes.Name("税率")]
        [JsonProperty("税率")]
        public string TaxRate { get; set; }

        [CsvHelper.Configuration.Attributes.Name("税额")]
        [JsonProperty("税额")]
        public string TaxAmount { get; set; }

        [CsvHelper.Configuration.Attributes.Name("价税合计")]
        [JsonProperty("价税合计")]
        public string TotalAmount { get; set; }

        [CsvHelper.Configuration.Attributes.Name("发票票种")]
        [JsonProperty("发票票种")]
        public string InvoiceType { get; set; }

        [CsvHelper.Configuration.Attributes.Name("类别")]
        [JsonProperty("类别")]
        public string Category { get; set; }
    }

    // Moonshot API 客户端（与之前的实现相同）
    public class MoonshotAIClient
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;
        private const string BASE_URL = "https://api.moonshot.cn/v1";

        public MoonshotAIClient(string apiKey)
        {
            _apiKey = apiKey;
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", apiKey);
        }

        public class FileUploadResponse
        {
            public string id { get; set; }
            public string @object { get; set; }
            public long bytes { get; set; }
            public long created_at { get; set; }
            public string filename { get; set; }
            public string purpose { get; set; }
        }

        public class ChatMessage
        {
            public string role { get; set; }
            public string content { get; set; }
        }

        public class ChatCompletionRequest
        {
            public string model { get; set; }
            public List<ChatMessage> messages { get; set; }
            public double temperature { get; set; }
            public ResponseFormat response_format { get; set; }
        }

        public class ResponseFormat
        {
            public string type { get; set; }
        }

        public class ChatCompletionResponse
        {
            public List<Choice> choices { get; set; }
        }

        public class Choice
        {
            public ChatMessage message { get; set; }
            public int index { get; set; }
            public string finish_reason { get; set; }
        }

        public async Task<FileUploadResponse> UploadFileAsync(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"文件未找到: {filePath}");
            }

            using var form = new MultipartFormDataContent();

            var fileContent = await File.ReadAllBytesAsync(filePath);
            var fileStreamContent = new ByteArrayContent(fileContent);
            fileStreamContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/octet-stream");

            form.Add(fileStreamContent, "file", Path.GetFileName(filePath));
            form.Add(new StringContent("file-extract"), "purpose");

            var response = await _httpClient.PostAsync($"{BASE_URL}/files", form);
            response.EnsureSuccessStatusCode();

            var responseContent = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<FileUploadResponse>(responseContent);
        }

        public async Task<string> GetFileContentAsync(string fileId)
        {
            var response = await _httpClient.GetAsync($"{BASE_URL}/files/{fileId}/content");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        public async Task DeleteFileAsync(string fileId)
        {
            var response = await _httpClient.DeleteAsync($"{BASE_URL}/files/{fileId}");
            response.EnsureSuccessStatusCode();
        }

        public async Task<ChatCompletionResponse> CreateChatCompletionAsync(ChatCompletionRequest request)
        {
            var json = JsonConvert.SerializeObject(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync($"{BASE_URL}/chat/completions", content);
            response.EnsureSuccessStatusCode();

            var responseContent = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<ChatCompletionResponse>(responseContent);
        }
    }

    // 主程序类
    public class InvoiceProcessor
    {
        // --- 配置区域 ---
        private const string API_KEY_FILE = "moonshot.txt";
        private const string PDF_FOLDER_PATH = "./invoices";
        private const string CSV_OUTPUT_PATH = "invoices_data.csv";
        private const int MAX_RETRIES = 3;
        private const int RETRY_DELAY = 2; // 秒
        private const string BUYER_FILE = "buyer.txt";


        // --- 配置区域结束 ---

        private MoonshotAIClient _client;
        private string _buyer;

        /// <summary>
        /// 从文件中读取API密钥
        /// </summary>
        private string GetApiKey(string filename)
        {
            try
            {
                if (!File.Exists(filename))
                {
                    throw new FileNotFoundException();
                }

                var key = File.ReadAllText(filename, Encoding.UTF8).Trim();

                if (string.IsNullOrWhiteSpace(key))
                {
                    Console.WriteLine($"错误：API密钥文件 '{filename}' 是空的。");
                    Console.WriteLine("请在文件中填入您的Kimi API密钥后重试。按 Enter 键退出。");
                    Console.ReadLine();
                    Environment.Exit(1);
                }

                return key;
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine($"错误：未找到API密钥文件 '{filename}'。");
                Console.WriteLine($"请在程序目录下创建一个名为 '{filename}' 的文本文件，并将您的Kimi API密钥粘贴进去。");
                Console.WriteLine("按 Enter 键退出。");
                Console.ReadLine();
                Environment.Exit(1);
                return null;
            }
        }

        /// <summary>
        /// 移除文件名中的非法字符
        /// </summary>
        private string SanitizeFilename(string name)
        {
            return Regex.Replace(name, @"[\\/*?:""<>|]", "");
        }

        /// <summary>
        /// 创建一个所有字段均为'N/A'的记录
        /// </summary>
        private InvoiceData CreateNARecord()
        {
            return new InvoiceData
            {
                InvoiceCode = "N/A",
                InvoiceNumber = "N/A",
                SellerTaxId = "N/A",
                SellerName = "N/A",
                BuyerTaxId = "N/A",
                BuyerName = "N/A",
                IssueDate = "N/A",
                ItemName = "N/A",
                Quantity = "N/A",
                Amount = "N/A",
                TaxRate = "N/A",
                TaxAmount = "N/A",
                TotalAmount = "N/A",
                InvoiceType = "N/A",
                Category = "N/A"
            };
        }

        /// <summary>
        /// 根据提取的数据重命名PDF文件
        /// </summary>
        private void RenameSuccessfulPdf(string originalPath, InvoiceData data)
        {
            try
            {
                var invoiceNum = data.InvoiceNumber ?? "N_A";
                var issueDate = data.IssueDate ?? "N_A";
                var category = data.Category ?? "N_A";
                var totalAmountStr = data.TotalAmount ?? "0";

                // 将总金额转换为不带小数点的整数
                int totalAmountInt = 0;
                if (double.TryParse(totalAmountStr, out double totalAmount))
                {
                    totalAmountInt = (int)totalAmount;
                }

                // 构建新文件名并清理非法字符
                var newBaseName = $"{invoiceNum}-{issueDate}-{category}-{totalAmountInt}.pdf";
                var sanitizedName = SanitizeFilename(newBaseName);

                var dirName = Path.GetDirectoryName(originalPath);
                var newPath = Path.Combine(dirName, sanitizedName);

                // 检查新文件名是否已存在，避免覆盖
                if (File.Exists(newPath))
                {
                    Console.WriteLine($"警告：重命名失败，文件 '{newPath}' 已存在。");
                    return;
                }

                File.Move(originalPath, newPath);
                Console.WriteLine($"文件已成功重命名为: {sanitizedName}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"警告：文件 '{Path.GetFileName(originalPath)}' 重命名失败: {e.Message}");
            }
        }

        /// <summary>
        /// 上传PDF，提取信息，包含重试和重命名逻辑
        /// </summary>
        private async Task<InvoiceData> ExtractInvoiceInfoFromPdfAsync(string pdfPath)
        {
            if (_client == null)
            {
                Console.WriteLine("错误：API客户端未初始化。");
                return CreateNARecord();
            }

            Console.WriteLine($"--- 开始处理文件: {Path.GetFileName(pdfPath)} ---");

            // --- API请求与重试逻辑 ---
            MoonshotAIClient.ChatCompletionResponse completion = null;
            MoonshotAIClient.FileUploadResponse fileObject = null;

            for (int attempt = 0; attempt < MAX_RETRIES; attempt++)
            {
                try
                {
                    // 仅在第一次尝试时上传文件
                    if (fileObject == null)
                    {
                        fileObject = await _client.UploadFileAsync(pdfPath);
                    }

                    var fileContent = await _client.GetFileContentAsync(fileObject.id);

                    // 构建完整的提取信息Prompt
                    var prompt = $@"
            你是一个发票信息提取助手。请从下面的文本内容中提取结构化的发票信息，其中发票票种为""增值税电子普通发票""或""电子发票（普通发票）""。

            文件内容:
            ---
            {fileContent}
            ---

            请严格按照以下JSON格式返回提取的信息，如果某个字段在文件中不存在，请用 ""N/A"" 表示。
            {{
                ""发票代码"": ""invoice_code"",
                ""发票号码"": ""invoice_number"",
                ""销方识别号"": ""seller_tax_id"",
                ""销方名称"": ""seller_name"",
                ""购方识别号"": ""buyer_tax_id"",
                ""购买方名称"": ""buyer_name"",
                ""开票日期"": ""issue_date"",
                ""项目名称"": ""item_name"",
                ""数量"": ""quantity"",
                ""金额"": ""amount"",
                ""税率"": ""tax_rate"",
                ""税额"": ""tax_amount"",
                ""价税合计"": ""total_amount"",
                ""发票票种"": ""invoice_type"",
                ""类别"": ""category""
            }}

            提取的例子：
            ```csv
            发票代码,发票号码,销方识别号,销方名称,购方识别号,购买方名称,开票日期,项目名称,数量,金额,税率,税额,价税合计,发票票种,类别
            ,24442000000657111111,91440812MAD2B8BJX7,湛江市西海岸西厨餐饮管理有限公司,91111111MA9W511111,广州咖喱网络科技有限公司,2024年12月30日,*餐饮服务*餐饮服务,1,109.9,1%,1.1,111,电子发票（普通发票）,餐饮服务
            044002301111,45311111,9144000061740323XQ,百胜餐饮（广东）有限公司,91111111MA9W511111,广州咖喱网络科技有限公司,2024年12月30日,*餐饮服务*餐饮服务,1,177.28,6%,10.64,187.92,增值税电子普通发票,餐饮服务
            ```
            
            {_buyer}
            
            类别参考：餐饮服务,住宿服务,交通运输服务,居民日常服务,办公用品,电子设备,咨询服务,技术服务,租赁服务,建筑服务,医疗服务,教育服务,商品零售 等

            金额相关数值请不要带有货币符号，例如¥等。

            注意不要搞反销售方和购买方信息。 
            ";

                    var messages = new List<MoonshotAIClient.ChatMessage>
                {
                    new MoonshotAIClient.ChatMessage
                    {
                        role = "system",
                        content = "你是一个专业的发票数据提取机器人，请根据用户提供的发票文本内容，准确地抽取出关键信息并以JSON格式返回。"
                    },
                    new MoonshotAIClient.ChatMessage
                    {
                        role = "user",
                        content = prompt
                    }
                };

                    completion = await _client.CreateChatCompletionAsync(new MoonshotAIClient.ChatCompletionRequest
                    {
                        model = "moonshot-v1-32k",
                        messages = messages,
                        temperature = 0.0,
                        response_format = new MoonshotAIClient.ResponseFormat { type = "json_object" }
                    });

                    Console.WriteLine(completion.choices[0].message.content = completion.choices[0].message.content.Trim());

                    Console.WriteLine($"API请求成功 (尝试 {attempt + 1}/{MAX_RETRIES})");
                    break; // 成功则跳出重试循环
                }
                catch (Exception e)
                {
                    Console.WriteLine($"API请求失败 (尝试 {attempt + 1}/{MAX_RETRIES}): {e.Message}");
                    if (attempt < MAX_RETRIES - 1)
                    {
                        Console.WriteLine($"将在 {RETRY_DELAY} 秒后重试...");
                        Thread.Sleep(RETRY_DELAY * 1000);
                    }
                    else
                    {
                        Console.WriteLine("所有重试均失败。");
                    }
                }
            }

            // 清理已上传的文件
            if (fileObject != null)
            {
                try
                {
                    await _client.DeleteFileAsync(fileObject.id);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"警告：清理云端文件 {fileObject.id} 失败: {e.Message}");
                }
            }

            // --- 处理API结果 ---
            if (completion == null)
            {
                Console.WriteLine($"处理失败，为文件 {Path.GetFileName(pdfPath)} 插入N/A记录。");
                return CreateNARecord();
            }

            try
            {
                var extractedData = JsonConvert.DeserializeObject<InvoiceData>(completion.choices[0].message.content);
                Console.WriteLine("API响应解析成功。");

                // 成功提取数据后，重命名文件
#if !DEBUG
                RenameSuccessfulPdf(pdfPath, extractedData);
#endif
                return extractedData;
            }
            catch (Exception e)
            {
                Console.WriteLine($"解析API响应或重命名文件时出错: {e.Message}");
                return CreateNARecord();
            }
        }

        /// <summary>
        /// 主函数，初始化并执行所有流程
        /// </summary>
        public async Task Main()
        {
            var apiKey = GetApiKey(API_KEY_FILE);
            _client = new MoonshotAIClient(apiKey);

            if (File.Exists(BUYER_FILE))
            {
                _buyer = "购买方名称及代码：\n" + File.ReadAllText(BUYER_FILE);
            }

            


            if (!Directory.Exists(PDF_FOLDER_PATH))
            {
                Console.WriteLine($"错误：文件夹 '{PDF_FOLDER_PATH}' 不存在。");
                Console.WriteLine("按 Enter 键退出。");
                Console.ReadLine();
                return;
            }

            var pdfFiles = Directory.GetFiles(PDF_FOLDER_PATH, "*.pdf", SearchOption.TopDirectoryOnly)
                                   .ToList();

            if (!pdfFiles.Any())
            {
                Console.WriteLine($"在文件夹 '{PDF_FOLDER_PATH}' 中没有找到PDF文件。");
                Console.WriteLine("按 Enter 键退出。");
                Console.ReadLine();
                return;
            }

            Console.WriteLine($"找到 {pdfFiles.Count} 个PDF文件。开始处理...");

            var allData = new List<InvoiceData>();

            foreach (var pdfPath in pdfFiles)
            {
                // 检查文件是否还存在，因为它可能已被成功处理并重命名
                if (File.Exists(pdfPath))
                {
                    var invoiceData = await ExtractInvoiceInfoFromPdfAsync(pdfPath);
                    allData.Add(invoiceData);
                }
                else
                {
                    Console.WriteLine($"跳过文件 {Path.GetFileName(pdfPath)}，因为它可能已被重命名。");
                }
            }

            // 使用CsvHelper写入CSV文件
            using (var writer = new StreamWriter(CSV_OUTPUT_PATH, false, new UTF8Encoding(true))) // UTF-8 with BOM
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(allData);
            }

            Console.WriteLine($"\n--- 处理完成！---\n所有发票信息已保存到文件: {CSV_OUTPUT_PATH}");
            Console.WriteLine("按 Enter 键退出。");
            Console.ReadLine();
        }
    }

    // 程序入口点
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var processor = new InvoiceProcessor();
            await processor.Main();
        }
    }
}
