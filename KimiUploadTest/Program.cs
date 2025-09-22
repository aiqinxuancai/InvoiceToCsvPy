namespace KimiUploadTest
{
    using System;
    using System.IO;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Newtonsoft.Json;
    using System.Collections.Generic;

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

        // 文件上传响应模型
        public class FileUploadResponse
        {
            public string id { get; set; }
            public string @object { get; set; }
            public long bytes { get; set; }
            public long created_at { get; set; }
            public string filename { get; set; }
            public string purpose { get; set; }
        }

        // 文件内容响应模型
        public class FileContentResponse
        {
            public string content { get; set; }
        }

        // 聊天消息模型
        public class ChatMessage
        {
            public string role { get; set; }
            public string content { get; set; }
        }

        // 聊天完成请求模型
        public class ChatCompletionRequest
        {
            public string model { get; set; }
            public List<ChatMessage> messages { get; set; }
            public double temperature { get; set; }
        }

        // 聊天完成响应模型
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

        /// <summary>
        /// 上传文件到Moonshot API
        /// 注意：purpose 参数必须设置为 "file-extract"，这是目前唯一支持的值
        /// </summary>
        public async Task<FileUploadResponse> UploadFileAsync(string filePath)
        {
            // 验证文件是否存在
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"文件未找到: {filePath}");
            }

            using var form = new MultipartFormDataContent();

            // 读取文件内容
            var fileContent = await File.ReadAllBytesAsync(filePath);
            var fileStreamContent = new ByteArrayContent(fileContent);
            fileStreamContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/octet-stream");

            // 添加文件到表单，字段名必须为 "file"
            form.Add(fileStreamContent, "file", Path.GetFileName(filePath));

            // 重要：purpose 参数必须设置为 "file-extract"
            // 这是目前文件上传接口唯一支持的 purpose 值
            form.Add(new StringContent("file-extract"), "purpose");

            // 发送请求
            var response = await _httpClient.PostAsync($"{BASE_URL}/files", form);
            response.EnsureSuccessStatusCode();

            var responseContent = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<FileUploadResponse>(responseContent);
        }

        /// <summary>
        /// 获取文件内容
        /// 注意：使用新版API endpoint /files/{file_id}/content
        /// </summary>
        public async Task<string> GetFileContentAsync(string fileId)
        {
            var response = await _httpClient.GetAsync($"{BASE_URL}/files/{fileId}/content");
            response.EnsureSuccessStatusCode();

            // 直接返回文本内容
            return await response.Content.ReadAsStringAsync();
        }

        /// <summary>
        /// 调用聊天完成API
        /// </summary>
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

    // 使用示例
    public class Program
    {
        public static async Task Main(string[] args)
        {
            // 替换为你的 API Key
            string apiKey = "";
            var client = new MoonshotAIClient(apiKey);

            try
            {
                // 1. 上传文件
                // moonshot.pdf 是示例文件，支持文本文件和图片文件（具有OCR能力）
                // purpose 必须设置为 "file-extract"
                Console.WriteLine("正在上传文件...");
                var fileObject = await client.UploadFileAsync("22222-22222-餐饮-597.pdf");
                Console.WriteLine($"文件上传成功，文件ID: {fileObject.id}");

                // 2. 获取文件内容
                // 使用新版 API，旧版的 retrieve_content 已标记为 warning
                Console.WriteLine("正在获取文件内容...");
                var fileContent = await client.GetFileContentAsync(fileObject.id);
                Console.WriteLine("文件内容获取成功");

                // 3. 构建消息列表
                // 重要：将文件内容通过系统提示词(system prompt)放入请求中
                var messages = new List<MoonshotAIClient.ChatMessage>
            {
                new MoonshotAIClient.ChatMessage
                {
                    role = "system",
                    content = "你是 Kimi，由 Moonshot AI 提供的人工智能助手，你更擅长中文和英文的对话。你会为用户提供安全，有帮助，准确的回答。同时，你会拒绝一切涉及恐怖主义，种族歧视，黄色暴力等问题的回答。Moonshot AI 为专有名词，不可翻译成其他语言。"
                },
                new MoonshotAIClient.ChatMessage
                {
                    role = "system",
                    content = fileContent  // <-- 重要：这里将抽取后的文件内容（注意是文件内容，不是文件ID）放置在请求中
                },
                new MoonshotAIClient.ChatMessage
                {
                    role = "user",
                    content = "请简单介绍 moonshot.pdf 的具体内容"
                }
            };

                // 4. 调用聊天完成API，获取 Kimi 的回答
                Console.WriteLine("正在生成回答...");
                var completion = await client.CreateChatCompletionAsync(new MoonshotAIClient.ChatCompletionRequest
                {
                    model = "kimi-k2-0905-preview",
                    messages = messages,
                    temperature = 0.6
                });

                // 5. 输出结果
                Console.WriteLine("\nKimi的回答：");
                Console.WriteLine(completion.choices[0].message.content);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误: {ex.Message}");
            }
        }
    }
}
