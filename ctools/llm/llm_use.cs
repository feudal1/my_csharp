using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace tools
{
    /// <summary>
    /// 通义千问 (DashScope) 服务类
    /// 支持普通 LLM 对话和 VLM 图像分析
    /// </summary>
    public class LlmService
    {
        // 配置项
        private const string DefaultModel = "qwen3.5-flash";
        
        private const string ApiUrl = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions";
        
        private readonly HttpClient _httpClient;
        private readonly string _shotMemoryFile;
        private readonly string _memoryDir;
        private readonly string _worksKnowledgeFile;
        private readonly string _commandsDescriptionFile;

        public LlmService()
        {
            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromMinutes(5)
            };

            // 初始化文件路径
            string llmDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "llm");
            if (!Directory.Exists(llmDir))
            {
                Directory.CreateDirectory(llmDir);
            }
            
            _shotMemoryFile = Path.Combine(llmDir, "shot_memory.json");
            _memoryDir = Path.Combine(llmDir, "memory_data");
            _worksKnowledgeFile = Path.Combine(llmDir, "works_knowledge.txt");
            _commandsDescriptionFile = Path.Combine(llmDir, "commands_description.txt");
            
            if (!Directory.Exists(_memoryDir))
            {
                Directory.CreateDirectory(_memoryDir);
            }
        }

        /// <summary>
        /// 从本地文件加载消息历史（不包含 system）
        /// </summary>
        private List<ChatMessage> LoadMessagesFromDisk()
        {
            try
            {
                if (File.Exists(_shotMemoryFile))
                {
                    var json = File.ReadAllText(_shotMemoryFile, Encoding.UTF8);
                    var messages = JsonSerializer.Deserialize<List<ChatMessage>>(json) ?? new List<ChatMessage>();
                    return messages.Where(m => m.Role != "system").ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载消息历史失败：{ex.Message}");
            }
            
            return new List<ChatMessage>();
        }

        /// <summary>
        /// 保存消息历史到本地 JSON 文件
        /// </summary>
        private void SaveMessagesToDisk(List<ChatMessage> messages)
        {
            try
            {
                var filteredMessages = messages.Where(m => m.Role != "system").ToList();
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                var json = JsonSerializer.Serialize(filteredMessages, options);
                File.WriteAllText(_shotMemoryFile, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存消息历史失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 读取 works_knowledge.txt 文件内容
        /// </summary>
        private string ReadWorksKnowledge()
        {
            try
            {
                if (File.Exists(_worksKnowledgeFile))
                {
                    return File.ReadAllText(_worksKnowledgeFile, Encoding.UTF8).Trim();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"读取 works_knowledge.txt 失败：{ex.Message}");
            }
            
            return "";
        }

        /// <summary>
        /// 执行 search 搜索并获取相关命令（使用 main.cs 中的现成方法）
        /// </summary>
        private string SearchCommands(string keyword, double threshold = 0.3, int? topK = null)
        {
            try
            {
                // 通过反射调用 Program.SearchCommands 方法
                var programType = typeof(Program);
                var searchMethod = programType.GetMethod("SearchCommands", 
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
                
                if (searchMethod == null)
                {
                    Console.WriteLine("未找到 SearchCommands 方法");
                    return "";
                }
                
                // 捕获控制台输出
                var sbOutput = new StringBuilder();
                var originalOut = Console.Out;
                using (var writer = new StringWriter(sbOutput))
                {
                    Console.SetOut(writer);
                    
                    // 调用静态方法
                    searchMethod.Invoke(null, [keyword, threshold, topK]);
                    
                    Console.SetOut(originalOut);
                }
                
                // 返回捕获的输出
                return sbOutput.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"搜索命令失败：{ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// 构建 System Prompt（仅包含基础信息，不包含全部命令）
        /// </summary>
        private string BuildSystemPrompt()
        {
            var sysPrompt = new StringBuilder();

            sysPrompt.AppendLine("你是一个 SolidWorks 自动化助手，可以帮助用户执行各种 SolidWorks 操作。");
            sysPrompt.AppendLine("当用户需要执行 SolidWorks 操作时，请使用以下严格格式调用命令：");
            sysPrompt.AppendLine("do_【命令名】参数 1 参数 2 ...");
            sysPrompt.AppendLine("例如：do_【export_dwg】C:\\path\\to\\file.sldprt");
            sysPrompt.AppendLine("重要：");
            sysPrompt.AppendLine("1. 必须使用 do_前缀和中文方括号【】来标识命令");
            sysPrompt.AppendLine("2. 执行命令前必须等待用户确认，用户输入 'y' 才会执行，'n' 跳过，'auto' 切换到自动模式。");
            sysPrompt.AppendLine("3. 不要使用其他格式（如单独的【】或英文括号）来调用命令。");
            
            // 添加 works_knowledge.txt 的内容
            var worksKnowledge = ReadWorksKnowledge();
            if (!string.IsNullOrEmpty(worksKnowledge))
            {
                sysPrompt.AppendLine("\n***工作知识与规范:***");
                sysPrompt.AppendLine(worksKnowledge);
            }

            sysPrompt.AppendLine("\n***命令格式示例:***");
            sysPrompt.AppendLine("正确：do_【export_dwg】C:\\work\\part.sldprt");
            sysPrompt.AppendLine("错误：【export_dwg】C:\\work\\part.sldprt (缺少 do_前缀)");
            sysPrompt.AppendLine("错误：do_[export_dwg] C:\\work\\part.sldprt (使用了英文括号)");

            return sysPrompt.ToString();
        }

        /// <summary>
        /// 获取 API Key（从环境变量或用户输入）
        /// </summary>
        private async Task<string> GetApiKeyAsync()
        {
            string apiKey = Environment.GetEnvironmentVariable("DASHSCOPE_API_KEY") ?? "";
            
            if (string.IsNullOrEmpty(apiKey))
            {
                Console.WriteLine("\n[警告] 未找到 DASHSCOPE_API_KEY 环境变量。");
                Console.Write("请输入临时 API Key: ");
                
                // 使用异步方式读取控制台输入
                apiKey = await Task.Run(() => Console.ReadLine()?.Trim() ?? "");
            }

            if (string.IsNullOrEmpty(apiKey)) 
                throw new ArgumentException("API Key 不能为空");
            
            return apiKey;
        }

        /// <summary>
        /// 统一的对话接口（支持文本和图像）
        /// </summary>
        public async Task<string> ChatAsync(string userPrompt, string? imagePath = null)
        {
            string apiKey = await GetApiKeyAsync();

            // 先对用户输入进行 search 搜索，只导入相关的命令到上下文
            string searchResult = SearchCommands(userPrompt, threshold: 0.3, topK: 5);
            Debug.Write($"searchResult:{searchResult}");
            // 加载历史消息（不包含 system）
            var messages = LoadMessagesFromDisk();
            
            // 构建基础的 system prompt（不包含具体命令）
            var sysPrompt = BuildSystemPrompt();
            
            // 将搜索结果添加到 system prompt 中
            if (!string.IsNullOrEmpty(searchResult))
            {
                sysPrompt += "\n***相关命令:***\n";
                sysPrompt += searchResult;
            }
            
            // 构建完整的消息列表
            var messagesWithSystem = new List<ChatMessage>
            {
                new ChatMessage { Role = "system", Content = sysPrompt.ToString() }
            };
            messagesWithSystem.AddRange(messages);
            
     
            
            var startTime = DateTime.Now;
            var fullResponse = !string.IsNullOrEmpty(imagePath) 
                ? await CallStreamingWithImageAsync(messagesWithSystem, userPrompt, imagePath, apiKey, DefaultModel)
                : await CallStreamingAsync(messagesWithSystem, apiKey, DefaultModel);
            
            var elapsedMs = (DateTime.Now - startTime).TotalMilliseconds;
            Console.WriteLine($"\n{(imagePath != null ? "VLM" : "LLM")} 调用耗时：{elapsedMs:F0}毫秒");
            
            // 保存助手回复到消息历史
            messages.Add(new ChatMessage { Role = "assistant", Content = fullResponse });
            SaveMessagesToDisk(messages);
            
            // 保存长期记忆
            var memoryContent = imagePath != null ? $"[图像分析]{imagePath}\n{fullResponse}" : fullResponse;
            SaveLongTermMemoryLog(memoryContent);
            
            return fullResponse;
        }

        /// <summary>
        /// 执行流式调用（纯文本 LLM）
        /// </summary>
        private async Task<string> CallStreamingCoreAsync(object requestBodyObj, string apiKey)
        {
            if (string.IsNullOrEmpty(apiKey)) 
                throw new ArgumentException("API Key 不能为空");

            var fullResponse = new StringBuilder();
            bool hasOutputStarted = false;

            try
            {
                string jsonBody = JsonSerializer.Serialize(requestBodyObj);
                using var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
                
                // 设置 Header
                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

                // 发送请求
                var request = new HttpRequestMessage(HttpMethod.Post, ApiUrl)
                {
                    Content = content
                };
                using var response = await _httpClient.SendAsync(
                    request,
                    HttpCompletionOption.ResponseHeadersRead
                );
                
                // 处理非 200 状态码
                if (!response.IsSuccessStatusCode)
                {
                    string errorBody = await response.Content.ReadAsStringAsync();
                    throw new HttpRequestException($"API 请求失败 ({response.StatusCode}): {errorBody}");
                }

                // 读取流式响应
                using var stream = await response.Content.ReadAsStreamAsync();
                var buffer = new byte[4096];
                var sbLine = new StringBuilder();
                int bytesRead;

                Console.WriteLine();

                while ((bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                {
                    string textChunk = Encoding.UTF8.GetString(buffer, 0, bytesRead);
                    
                    foreach (char c in textChunk)
                    {
                        if (c == '\n')
                        {
                            string line = sbLine.ToString().Trim();
                            sbLine.Clear();

                            if (string.IsNullOrEmpty(line)) continue;

                            if (line.StartsWith("data: "))
                            {
                                string dataJson = line.Substring(6).Trim();

                                if (dataJson == "[DONE]")
                                {
                                    goto EndStream;
                                }

                                try
                                {
                                    using var doc = JsonDocument.Parse(dataJson);
                                    var root = doc.RootElement;

                                    if (root.TryGetProperty("error", out var errorElem))
                                    {
                                        string errMsg = errorElem.GetProperty("message").GetString() ?? "未知错误";
                                        throw new Exception($"流式传输中发生错误：{errMsg}");
                                    }

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
                                                Console.Write(chunk);
                                                Console.Out.Flush();
                                                hasOutputStarted = true;
                                            }
                                        }
                                    }
                                }
                                catch (JsonException)
                                {
                                    // 忽略解析错误
                                }
                            }
                        }
                        else
                        {
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
                Console.WriteLine($"\n\n[严重错误] 调用失败：{ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 执行流式调用（纯文本 LLM）
        /// </summary>
        public async Task<string> CallStreamingAsync(List<ChatMessage> messages, string apiKey, string model)
        {
            var requestBody = new
            {
                model = model,
                stream = true,
                messages = messages.Select(m => new
                {
                    role = m.Role,
                    content = m.Content
                })
            };
            
            return await CallStreamingCoreAsync(requestBody, apiKey);
        }

        /// <summary>
        /// 执行流式调用（带图像 VLM）
        /// </summary>
        public async Task<string> CallStreamingWithImageAsync(List<ChatMessage> messages, string userPrompt, string imagePath, string apiKey, string? model = null)
        {
            model ??= DefaultModel;
            
            // 准备图片数据
            string ext = Path.GetExtension(imagePath).ToLower();
            string mimeType = ext switch
            {
                ".png" => "image/png",
                ".jpg" => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".webp" => "image/webp",
                _ => "image/png"
            };

            byte[] imageBytes;
            using (var fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
            {
                var ms = new MemoryStream();
                await fs.CopyToAsync(ms);
                imageBytes = ms.ToArray();
            }
            string base64Image = Convert.ToBase64String(imageBytes);
            string dataUrl = $"data:{mimeType};base64,{base64Image}";

            // 构建请求体（OpenAI 兼容格式，带图像）
            var messagesList = new List<object>();
            foreach (var m in messages)
            {
                messagesList.Add(new { role = m.Role, content = m.Content });
            }
            messagesList.Add(new
            {
                role = "user",
                content = new object[]
                {
                    new { type = "text", text = userPrompt },
                    new { type = "image_url", image_url = new { url = dataUrl } }
                }
            });

            var requestBody = new
            {
                model = model,
                stream = true,
                messages = messagesList.ToArray()
            };
            
            return await CallStreamingCoreAsync(requestBody, apiKey);
        }

        /// <summary>
        /// 保存长期记忆日志
        /// </summary>
        private void SaveLongTermMemoryLog(string content)
        {
            try
            {
                var logFilePath = Path.Combine(_memoryDir, "longterm_memory.txt");
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                File.AppendAllText(logFilePath, $"[{timestamp}] {content}\n", Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存日志失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 清空对话历史
        /// </summary>
        public void ClearHistory()
        {
            try
            {
                if (File.Exists(_shotMemoryFile))
                {
                    File.Delete(_shotMemoryFile);
                    Console.WriteLine("对话历史已清空");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清空历史失败：{ex.Message}");
            }
        }
    }

    /// <summary>
    /// 聊天消息类
    /// </summary>
    public class ChatMessage
    {
        public string Role { get; set; } = "";
        public string Content { get; set; } = "";
    }
}
