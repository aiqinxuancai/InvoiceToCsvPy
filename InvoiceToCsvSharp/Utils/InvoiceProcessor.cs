using CsvHelper;
using InvoiceToCsvSharp.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace InvoiceToCsvSharp.Utils
{
    public class InvoiceProcessor
    {
        // --- 配置区域 --- 
        private const string API_KEY_FILE = "moonshot.txt";
        private const string PDF_FOLDER_PATH = "./invoices";
        private const string CSV_OUTPUT_PATH = "invoices_data.csv";
        private const int MAX_RETRIES = 3;
        private const int RETRY_DELAY = 2; // 秒
        private const string BUYER_FILE = "buyer.txt";

        // 多线程配置
        private const int MAX_CONCURRENT_TASKS = 5; // 最大并发数
        private const int API_RATE_LIMIT_DELAY_MS = 500; // API调用间隔（毫秒） 

        // --- 配置区域结束 --- 

        private MoonshotAIClient _client;
        private string _buyer;

        // 多线程相关
        private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(MAX_CONCURRENT_TASKS);
        private readonly SemaphoreSlim _apiRateLimiter = new SemaphoreSlim(1, 1);
        private readonly object _consoleLock = new object();
        private readonly object _fileLock = new object();
        private int _processedCount = 0;
        private int _totalCount = 0;

        /// <summary>
        /// 线程安全的控制台输出
        /// </summary>
        private void SafeConsoleWriteLine(string message, ConsoleColor? color = null)
        {
            lock (_consoleLock)
            {
                if (color.HasValue)
                {
                    var originalColor = Console.ForegroundColor;
                    Console.ForegroundColor = color.Value;
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");
                    Console.ForegroundColor = originalColor;
                }
                else
                {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");
                }
            }
        }

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
                    SafeConsoleWriteLine($"错误：API密钥文件 '{filename}' 是空的。", ConsoleColor.Red);
                    SafeConsoleWriteLine("请在文件中填入您的Kimi API密钥后重试。按 Enter 键退出。");
                    Console.ReadLine();
                    Environment.Exit(1);
                }

                return key;
            }
            catch (FileNotFoundException)
            {
                SafeConsoleWriteLine($"错误：未找到API密钥文件 '{filename}'。", ConsoleColor.Red);
                SafeConsoleWriteLine($"请在程序目录下创建一个名为 '{filename}' 的文本文件，并将您的Kimi API密钥粘贴进去。");
                SafeConsoleWriteLine("按 Enter 键退出。");
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
        /// 根据提取的数据重命名PDF文件（线程安全） 
        /// </summary>
        private void RenameSuccessfulPdf(string originalPath, InvoiceData data)
        {
            lock (_fileLock)
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
                        SafeConsoleWriteLine($"警告：重命名失败，文件 '{newPath}' 已存在。", ConsoleColor.Yellow);
                        return;
                    }

                    File.Move(originalPath, newPath);
                    SafeConsoleWriteLine($"文件已成功重命名为: {sanitizedName}", ConsoleColor.Green);
                }
                catch (Exception e)
                {
                    SafeConsoleWriteLine($"警告：文件 '{Path.GetFileName(originalPath)}' 重命名失败: {e.Message}", ConsoleColor.Yellow);
                }
            }
        }

        /// <summary>
        /// 上传PDF，提取信息，包含重试和重命名逻辑（支持并发） 
        /// </summary>
        private async Task<InvoiceData> ExtractInvoiceInfoFromPdfAsync(string pdfPath, int index)
        {
            await _semaphore.WaitAsync();
            try
            {
                if (_client == null)
                {
                    SafeConsoleWriteLine("错误：API客户端未初始化。", ConsoleColor.Red);
                    return CreateNARecord();
                }

                SafeConsoleWriteLine($"[{index + 1}/{_totalCount}] 开始处理文件: {Path.GetFileName(pdfPath)}", ConsoleColor.Cyan);

                // --- API请求与重试逻辑 --- 
                MoonshotAIClient.ChatCompletionResponse completion = null;
                MoonshotAIClient.FileUploadResponse fileObject = null;

                for (int attempt = 0; attempt < MAX_RETRIES; attempt++)
                {
                    try
                    {
                        // 使用速率限制器控制API调用频率
                        await _apiRateLimiter.WaitAsync();
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
                                model = "kimi-k2-0905-preview", //"moonshot-v1-32k", 
                                messages = messages,
                                temperature = 0.0,
                                response_format = new MoonshotAIClient.ResponseFormat { type = "json_object" }
                            });

                            completion.choices[0].message.content = completion.choices[0].message.content.Trim();

                            SafeConsoleWriteLine($"[{index + 1}] API响应: {completion.choices[0].message.content}", ConsoleColor.Gray);
                            SafeConsoleWriteLine($"[{index + 1}] API请求成功 (尝试 {attempt + 1}/{MAX_RETRIES})", ConsoleColor.Green);
                            break; // 成功则跳出重试循环
                        }
                        finally
                        {
                            // 延迟释放，实现速率限制
                            _ = Task.Delay(API_RATE_LIMIT_DELAY_MS).ContinueWith(_ => _apiRateLimiter.Release());
                        }
                    }
                    catch (Exception e)
                    {
                        SafeConsoleWriteLine($"[{index + 1}] API请求失败 (尝试 {attempt + 1}/{MAX_RETRIES}): {e.Message}", ConsoleColor.Yellow);
                        if (attempt < MAX_RETRIES - 1)
                        {
                            SafeConsoleWriteLine($"[{index + 1}] 将在 {RETRY_DELAY} 秒后重试...", ConsoleColor.Yellow);
                            await Task.Delay(RETRY_DELAY * 1000);
                        }
                        else
                        {
                            SafeConsoleWriteLine($"[{index + 1}] 所有重试均失败。", ConsoleColor.Red);
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
                        SafeConsoleWriteLine($"[{index + 1}] 警告：清理云端文件 {fileObject.id} 失败: {e.Message}", ConsoleColor.Yellow);
                    }
                }

                // --- 处理API结果 --- 
                if (completion == null)
                {
                    SafeConsoleWriteLine($"[{index + 1}] 处理失败，为文件 {Path.GetFileName(pdfPath)} 插入N/A记录。", ConsoleColor.Red);
                    return CreateNARecord();
                }

                try
                {
                    var extractedData = JsonConvert.DeserializeObject<InvoiceData>(completion.choices[0].message.content);
                    SafeConsoleWriteLine($"[{index + 1}] API响应解析成功。", ConsoleColor.Green);

                    // 成功提取数据后，重命名文件
#if !DEBUG
RenameSuccessfulPdf(pdfPath, extractedData); 
#endif

                    // 更新进度
                    var processed = Interlocked.Increment(ref _processedCount);
                    SafeConsoleWriteLine($"进度: {processed}/{_totalCount} ({processed * 100 / _totalCount}%)", ConsoleColor.Magenta);

                    return extractedData;
                }
                catch (Exception e)
                {
                    SafeConsoleWriteLine($"[{index + 1}] 解析API响应或重命名文件时出错: {e.Message}", ConsoleColor.Red);
                    return CreateNARecord();
                }
            }
            finally
            {
                _semaphore.Release();
            }
        }

        /// <summary>
        /// 处理一批文件
        /// </summary>
        private async Task<List<InvoiceData>> ProcessBatchAsync(List<(string path, int index)> batch)
        {
            var tasks = batch.Select(item => ExtractInvoiceInfoFromPdfAsync(item.path, item.index));
            var results = await Task.WhenAll(tasks);
            return results.ToList();
        }

        /// <summary>
        /// 主函数，初始化并执行所有流程
        /// </summary>
        public async Task Main()
        {
            var startTime = DateTime.Now;

            var apiKey = GetApiKey(API_KEY_FILE);
            _client = new MoonshotAIClient(apiKey);

            if (File.Exists(BUYER_FILE))
            {
                _buyer = "购买方名称及代码：\n" + File.ReadAllText(BUYER_FILE);
            }

            if (!Directory.Exists(PDF_FOLDER_PATH))
            {
                SafeConsoleWriteLine($"错误：文件夹 '{PDF_FOLDER_PATH}' 不存在。", ConsoleColor.Red);
                SafeConsoleWriteLine("按 Enter 键退出。");
                Console.ReadLine();
                return;
            }

            var pdfFiles = Directory.GetFiles(PDF_FOLDER_PATH, "*.pdf", SearchOption.TopDirectoryOnly)
                                   .ToList();

            if (!pdfFiles.Any())
            {
                SafeConsoleWriteLine($"在文件夹 '{PDF_FOLDER_PATH}' 中没有找到PDF文件。", ConsoleColor.Red);
                SafeConsoleWriteLine("按 Enter 键退出。");
                Console.ReadLine();
                return;
            }

            _totalCount = pdfFiles.Count;
            _processedCount = 0;

            SafeConsoleWriteLine($"找到 {pdfFiles.Count} 个PDF文件。", ConsoleColor.Green);
            SafeConsoleWriteLine($"使用 {MAX_CONCURRENT_TASKS} 个并发任务开始处理...", ConsoleColor.Green);
            SafeConsoleWriteLine(new string('-', 50));

            // 使用线程安全的集合来收集结果
            var allData = new ConcurrentBag<InvoiceData>();
            var indexedFiles = pdfFiles.Select((path, index) => (path, index)).ToList();

            // 分批处理文件
            var batches = new List<List<(string path, int index)>>();
            for (int i = 0; i < indexedFiles.Count; i += MAX_CONCURRENT_TASKS)
            {
                var batch = indexedFiles.Skip(i).Take(MAX_CONCURRENT_TASKS).ToList();
                batches.Add(batch);
            }

            // 处理所有批次
            foreach (var batch in batches)
            {
                var batchData = await ProcessBatchAsync(batch);
                foreach (var data in batchData)
                {
                    allData.Add(data);
                }
            }

            // 按原始顺序排序（如果需要） 
            var sortedData = allData.OrderBy(x => x.InvoiceNumber).ToList();

            // 使用CsvHelper写入CSV文件
            using (var writer = new StreamWriter(CSV_OUTPUT_PATH, false, new UTF8Encoding(true))) // UTF-8 with BOM
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(sortedData);
            }

            var endTime = DateTime.Now;
            var duration = endTime - startTime;

            SafeConsoleWriteLine(new string('-', 50));
            SafeConsoleWriteLine($"处理完成！", ConsoleColor.Green);
            SafeConsoleWriteLine($"总耗时: {duration.TotalSeconds:F2} 秒", ConsoleColor.Cyan);
            SafeConsoleWriteLine($"平均每个文件: {duration.TotalSeconds / _totalCount:F2} 秒", ConsoleColor.Cyan);
            SafeConsoleWriteLine($"所有发票信息已保存到文件: {CSV_OUTPUT_PATH}", ConsoleColor.Green);
            SafeConsoleWriteLine("按 Enter 键退出。");
            Console.ReadLine();
        }
    }
}