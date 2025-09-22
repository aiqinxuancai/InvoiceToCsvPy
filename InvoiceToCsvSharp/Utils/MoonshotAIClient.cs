using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceToCsvSharp.Utils
{
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
}
