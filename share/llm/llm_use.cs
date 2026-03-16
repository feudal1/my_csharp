using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace tools
{
    /// <summary>
    /// 通义千问 (DashScope) 流式调用服务
    /// 兼容 OpenAI 格式接口
    /// </summary>
    public class VlmService
    {
        // 配置项
        private const string DefaultModel = "qwen-vl-max-latest"; // 推荐：支持高清长图，理解力强
        private const string ApiUrl = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions";
        
        private readonly HttpClient _httpClient;

        public VlmService()
        {
            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromMinutes(5) // 给大图分析足够的时间
            };
        }

        /// <summary>
        /// 执行流式调用
        /// </summary>
        /// <param name="imagePath">本地图片路径</param>
        /// <param name="prompt">提示词</param>
        /// <param name="apiKey">DashScope API Key</param>
        /// <param name="model">模型名称 (可选，默认 qwen-vl-max-latest)</param>
        /// <returns>完整的回复文本</returns>
        public async Task<string> CallStreamingAsync(string imagePath, string prompt, string? model = null)
        {
             string apiKey = Environment.GetEnvironmentVariable("DASHSCOPE_API_KEY") ?? "";
        
        if (string.IsNullOrEmpty(apiKey))
        {
            Console.WriteLine("\n[警告] 未找到 DASHSCOPE_API_KEY 环境变量。");
            Console.Write("请输入临时 API Key: ");
            apiKey = Console.ReadLine()?.Trim() ?? "";
        }

  
            if (string.IsNullOrEmpty(apiKey)) throw new ArgumentException("API Key 不能为空");
            if (!File.Exists(imagePath)) throw new FileNotFoundException($"图片文件未找到: {imagePath}");

            model ??= DefaultModel;
            var fullResponse = new StringBuilder();
            bool hasOutputStarted = false;

            try
            {
                // 1. 准备图片数据 (自动识别 MIME 类型)
                string ext = Path.GetExtension(imagePath).ToLower();
                string mimeType = ext switch
                {
                    ".png" => "image/png",
                    ".jpg" => "image/jpeg",
                    ".jpeg" => "image/jpeg",
                    ".webp" => "image/webp",
                    _ => "image/png" // 默认 fallback
                };

                // .NET Framework 4.8 不支持 File.ReadAllBytesAsync，使用 FileStream 替代
                byte[] imageBytes;
                using (var fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
                {
                    var ms = new MemoryStream();
                    await fs.CopyToAsync(ms);
                    imageBytes = ms.ToArray();
                }
                string base64Image = Convert.ToBase64String(imageBytes);
                string dataUrl = $"data:{mimeType};base64,{base64Image}";

                // 2. 构建请求体 (OpenAI 兼容格式)
                var requestBody = new
                {
                    model = model,
                    stream = true, // 开启流式
                    messages = new[]
                    {
                        new
                        {
                            role = "user",
                            content = new object[]
                            {
                                new { type = "text", text = prompt },
                                new { type = "image_url", image_url = new { url = dataUrl } }
                            }
                        }
                    }
                };

                string jsonBody = JsonSerializer.Serialize(requestBody);
                using var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
                
                // 设置 Header
                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

                // 3. 发送请求 (使用 SendAsync 以支持 HttpCompletionOption)
                var request = new HttpRequestMessage(HttpMethod.Post, ApiUrl)
                {
                    Content = content
                };
                using var response = await _httpClient.SendAsync(
                    request,
                    HttpCompletionOption.ResponseHeadersRead
                );
                
                // 4. 处理非 200 状态码 (如 401, 404, 500)
                if (!response.IsSuccessStatusCode)
                {
                    string errorBody = await response.Content.ReadAsStringAsync();
                    throw new HttpRequestException($"API 请求失败 ({response.StatusCode}): {errorBody}");
                }

                // 5. 【重构】读取流式响应 - 改用字节流缓冲读取以实现真正的实时输出
                using var stream = await response.Content.ReadAsStreamAsync();
                var buffer = new byte[4096];
                var sbLine = new StringBuilder(); // 用于暂存当前行的数据
                int bytesRead;

                Console.WriteLine(); // 确保在开始接收前换行

                while ((bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                {
                    string textChunk = Encoding.UTF8.GetString(buffer, 0, bytesRead);
                    
                    // 逐字符处理，寻找换行符
                    foreach (char c in textChunk)
                    {
                        if (c == '\n')
                        {
                            // 遇到换行符，处理当前累积的一行
                            string line = sbLine.ToString().Trim();
                            sbLine.Clear();

                            if (string.IsNullOrEmpty(line)) continue;

                            // 检查 SSE 前缀
                            if (line.StartsWith("data: "))
                            {
                                string dataJson = line.Substring(6).Trim();

                                if (dataJson == "[DONE]")
                                {
                                    goto EndStream; // 跳出循环
                                }

                                try
                                {
                                    using var doc = JsonDocument.Parse(dataJson);
                                    var root = doc.RootElement;

                                    // 错误检查
                                    if (root.TryGetProperty("error", out var errorElem))
                                    {
                                        string errMsg = errorElem.GetProperty("message").GetString() ?? "未知错误";
                                        throw new Exception($"流式传输中发生错误：{errMsg}");
                                    }

                                    // 提取内容
                                    if (root.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0)
                                    {
                                        var choice = choices[0];
                                        if (choice.TryGetProperty("delta", out var delta) && 
                                            delta.TryGetProperty("content", out var contentElem))
                                        {
                                            string? chunk = contentElem.GetString();
                                            if (!string.IsNullOrEmpty(chunk))
                                            {
                                                fullResponse.Append(chunk);
                                                Console.Write(chunk); // 实时输出
                                                Console.Out.Flush();  // 强制刷新控制台缓冲区
                                                hasOutputStarted = true;
                                            }
                                        }
                                    }
                                }
                                catch (JsonException)
                                {
                                    // 忽略解析错误，可能是空行或不完整数据
                                }
                            }
                            else if (line.StartsWith("{")) 
                            {
                                // 兼容某些非标准实现直接返回 JSON 而没有 data: 前缀的情况（极少见，但作为兜底）
                                // 这里可以根据需要添加逻辑，通常 SSE 必须有 data:
                            }
                        }
                        else
                        {
                            // 累积字符
                            sbLine.Append(c);
                        }
                    }
                }

EndStream:
                if (!hasOutputStarted)
                {
                    Console.WriteLine("\n[提示] 模型未返回任何文本内容。");
                }

                return fullResponse.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n\n[严重错误] 调用失败: {ex.Message}");
                throw; // 向上抛出，让调用者决定如何处理
            }
        }
    }
}