using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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
        private readonly string _logFilePath;
        private readonly string _worksKnowledgeFile;
        private readonly Func<string>? _getCommandsDescriptionFunc;

        public LlmService(Func<string>? getCommandsDescriptionFunc = null)
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
      _logFilePath = Path.Combine(llmDir, "longterm_memory.txt");
            _worksKnowledgeFile = Path.Combine(llmDir, "works_knowledge.txt");
            _getCommandsDescriptionFunc = getCommandsDescriptionFunc;
            
        
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
                    var messages = JsonConvert.DeserializeObject<List<ChatMessage>>(json) ?? new List<ChatMessage>();
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
        /// 保存消息历史到本地 JSON 文件（超过 10 条自动截断）
        /// </summary>
        private void SaveMessagesToDisk(List<ChatMessage> messages)
        {
            try
            {
                var filteredMessages = messages.Where(m => m.Role != "system").ToList();
                
                // 如果消息超过 10 条，保留最近的 10 条（5 轮对话）
                if (filteredMessages.Count > 10)
                {
                    filteredMessages = filteredMessages.Skip(filteredMessages.Count - 10).ToList();
                    Console.WriteLine($"\n[调试] 消息数量超过 10 条，已截断保留最近 10 条");
                }
                
                var json = JsonConvert.SerializeObject(filteredMessages, Formatting.Indented);
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
        /// 执行 search 搜索并获取相关命令（基于相似度匹配）
        /// </summary>
        private string SearchCommands(string keyword, double threshold = 0.3, int? topK = null)
        {
            // 使用委托获取实时命令描述
            if (_getCommandsDescriptionFunc != null)
            {
                try
                {
                    string allContent = _getCommandsDescriptionFunc();
                    return SearchInContent(allContent, keyword, threshold, topK);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"调用命令描述获取函数失败：{ex.Message}");
                }
            }
            
            return $"未找到与 '{keyword}' 相关的命令";
        }

        /// <summary>
        /// 在内容中搜索相关命令（简化版：只使用模糊搜索）
        /// </summary>
        private string SearchInContent(string content, string keyword, double threshold, int? topK)
        {
            if (string.IsNullOrWhiteSpace(keyword))
            {
                return content;
            }

            var lines = content.Split('\n');
            var resultBlocks = new List<List<string>>();
            var currentBlock = new List<string>();
            bool inBlock = false;
            
            foreach (var line in lines)
            {
                // 检测新命令开始（格式：【分组】命令名）
                if (line.StartsWith("【") && line.Contains("】"))
                {
                    // 保存上一个匹配的块
                    if (inBlock && currentBlock.Count > 0)
                    {
                        resultBlocks.Add(new List<string>(currentBlock));
                    }

                    currentBlock.Clear();
                    currentBlock.Add(line);
                    inBlock = false;

                    // 检查命令名是否包含关键词
                    if (line.ToLower().Contains(keyword.ToLower()))
                    {
                        inBlock = true;
                    }
                }
                else if (inBlock)
                {
                    // 如果在匹配的块中，继续收集后续行（说明、参数等）
                    if (line.StartsWith("    ") || string.IsNullOrWhiteSpace(line))
                    {
                        currentBlock.Add(line);
                    }
                    else
                    {
                        // 遇到非缩进的内容，块结束
                        resultBlocks.Add(new List<string>(currentBlock));
                        inBlock = false;
                    }
                }
                else
                {
                    // 模糊搜索：使用相似度匹配而不是简单包含
                    double score = CalculateLineSimilarity(keyword, line);
                    
                    if (score >= threshold)
                    {
                        // 回溯找到这个命令块的开头
                        int currentIndex = Array.IndexOf(lines, line);
                        for (int i = currentIndex - 1; i >= 0; i--)
                        {
                            if (lines[i].StartsWith("【") && lines[i].Contains("】"))
                            {
                                // 从命令头开始重新收集
                                for (int j = i; j <= currentIndex; j++)
                                {
                                    currentBlock.Add(lines[j]);
                                }
                                inBlock = true;
                                break;
                            }
                        }
                        
                        // 如果没找到命令头，直接添加当前行（兼容其他格式）
                        if (!inBlock)
                        {
                            currentBlock.Add(line);
                            inBlock = true;
                        }
                    }
                }
            }

            // 处理最后一个块
            if (inBlock && currentBlock.Count > 0)
            {
                resultBlocks.Add(currentBlock);
            }

            if (resultBlocks.Count == 0)
            {
                return $"未找到与 '{keyword}' 相关的命令";
            }

            // 合并所有匹配的块
            var result = resultBlocks.SelectMany(b => b);
            if (topK.HasValue)
            {
                result = result.Take(topK.Value);
            }

            return string.Join("\n", result);
        }

        /// <summary>
        /// 计算单行的相似度
        /// </summary>
        private double CalculateLineSimilarity(string keyword, string line)
        {
            if (string.IsNullOrEmpty(keyword) || string.IsNullOrEmpty(line))
            {
                return 0.0;
            }

            string keywordLower = keyword.ToLower();
            string lineLower = line.ToLower();

            // 直接包含关键词则返回高分
            if (lineLower.Contains(keywordLower))
            {
                return 0.8;
            }

            // 否则计算编辑距离相似度
            return CalculateLevenshteinRatioSimple(keywordLower, lineLower);
        }

        /// <summary>
        /// 简化的 Levenshtein 距离比率计算
        /// </summary>
        private double CalculateLevenshteinRatioSimple(string s1, string s2)
        {
            if (string.IsNullOrEmpty(s1)) return string.IsNullOrEmpty(s2) ? 1.0 : 0.0;
            if (string.IsNullOrEmpty(s2)) return 0.0;

            int[,] matrix = new int[s1.Length + 1, s2.Length + 1];

            for (int i = 0; i <= s1.Length; i++)
                matrix[i, 0] = i;

            for (int j = 0; j <= s2.Length; j++)
                matrix[0, j] = j;

            for (int i = 1; i <= s1.Length; i++)
            {
                for (int j = 1; j <= s2.Length; j++)
                {
                    int cost = (s1[i - 1] == s2[j - 1]) ? 0 : 1;
                    matrix[i, j] = Math.Min(
                        Math.Min(
                            matrix[i - 1, j] + 1,
                            matrix[i, j - 1] + 1
                        ),
                        matrix[i - 1, j - 1] + cost
                    );
                }
            }

            int distance = matrix[s1.Length, s2.Length];
            int maxLen = Math.Max(s1.Length, s2.Length);
            return maxLen > 0 ? 1.0 - (distance / (double)maxLen) : 1.0;
        }

        /// <summary>
        /// 构建 System Prompt（精简版，不包含命令格式说明）
        /// </summary>
        private string BuildSystemPrompt()
        {
            var sysPrompt = new StringBuilder();

            sysPrompt.AppendLine("你是一个 SolidWorks 自动化助手，可以帮助用户执行各种 SolidWorks 操作。");
            
            // 添加 works_knowledge.txt 的内容
            var worksKnowledge = ReadWorksKnowledge();
            if (!string.IsNullOrEmpty(worksKnowledge))
            {
                sysPrompt.AppendLine("\n***工作知识与规范:***");
                sysPrompt.AppendLine(worksKnowledge);
            }

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
        /// 统一的对话接口（支持文本和图像，不支持 Tool）
        /// </summary>
        public async Task<string> ChatAsync(string userPrompt, string? imagePath = null)
        {
            string apiKey = await GetApiKeyAsync();

            // 先对用户输入进行 search 搜索，只导入相关的命令到上下文
            string searchResult = SearchCommands(userPrompt, threshold: 0.3, topK: 5);
            Console.Write($"searchResult:{searchResult}");
            // 加载历史消息（不包含 system）
            var messages = LoadMessagesFromDisk();
            
            // 构建基础的 system prompt
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
            
            // 保存用户输入到消息历史
            messages.Add(new ChatMessage { Role = "user", Content = userPrompt });
            
            // 保存助手回复到消息历史
            messages.Add(new ChatMessage { Role = "assistant", Content = fullResponse });
            SaveMessagesToDisk(messages);
            
            // 保存长期记忆
            var memoryContent = imagePath != null ? $"[图像分析]{imagePath}\n{fullResponse}" : fullResponse;
            SaveLongTermMemoryLog(memoryContent);
            
            return fullResponse;
        }

        /// <summary>
        /// 带 Tool 调用的对话接口（优化版：只传递搜索后的工具）
        /// </summary>
        public async Task<(string Response, List<ToolCall>? ToolCalls)> ChatWithToolsAsync(
            string userPrompt, 
            List<ToolDefinition> allTools)
        {
            string apiKey = await GetApiKeyAsync();
        
            // 先对用户输入进行 search 搜索，获取相关命令名称
            string searchResult = SearchCommands(userPrompt, threshold: 0.3, topK: 5);
            Console.Write($"searchResult:{searchResult}");
                    
            // 从搜索结果中提取匹配的命令名，过滤工具列表
            var filteredTools = FilterToolsBySearchResult(allTools, searchResult);
                    
            // 加载历史消息
            var messages = LoadMessagesFromDisk();
                    
            // 构建 system prompt（不包含命令列表）
            var sysPrompt = BuildSystemPrompt();
                    
            // 构建消息列表
            var messagesWithSystem = new List<ChatMessage>
            {
                new ChatMessage { Role = "system", Content = sysPrompt.ToString() }
            };
            messagesWithSystem.AddRange(messages);
                    
            var startTime = DateTime.Now;
            var (response, toolCalls) = await CallStreamingWithToolsAsync(
                messagesWithSystem, 
                userPrompt, 
                filteredTools,  // 只传递筛选后的工具
                apiKey, 
                DefaultModel
            );
                    
            var elapsedMs = (DateTime.Now - startTime).TotalMilliseconds;
            Console.WriteLine($"\nLLM Tool 调用耗时：{elapsedMs:F0}毫秒");
                    
            // 保存用户输入到消息历史
            messages.Add(new ChatMessage { Role = "user", Content = userPrompt });
                    
            // 保存助手回复
            if (!string.IsNullOrEmpty(response))
            {
                messages.Add(new ChatMessage { Role = "assistant", Content = response });
            }
            else if (toolCalls != null && toolCalls.Count > 0)
            {
                // 保存 Tool 调用到消息历史
                foreach (var toolCall in toolCalls)
                {
                    messages.Add(new ChatMessage 
                    { 
                        Role = "assistant", 
                        Content = "",
                        ToolCallId = toolCall.Id,
                        ToolCalls = toolCalls
                    });
                }
            }
                    
            SaveMessagesToDisk(messages);
            SaveLongTermMemoryLog(response ?? $"ToolCall: {JsonConvert.SerializeObject(toolCalls)}");
                    
            return (response, toolCalls);
        }
        
        /// <summary>
        /// 根据搜索结果过滤工具列表
        /// </summary>
        private List<ToolDefinition> FilterToolsBySearchResult(List<ToolDefinition> allTools, string searchResult)
        {
            // 如果搜索未找到结果，返回全部工具，并给出提示
            if (searchResult.StartsWith("未找到"))
            {
                Console.WriteLine($"\n[调试] 未找到匹配的命令，返回全部 {allTools.Count} 个工具");
                        
                // 添加一个特殊的工具来提示 AI
                var hintTool = new ToolDefinition
                {
                    Type = "function",
                    Function = new FunctionDefinition
                    {
                        Name = "no_matching_command_found",
                        Description = "提示：未在命令库中找到与用户问题直接相关的命令。请根据用户的实际需求，从可用工具中选择最接近的功能，或者询问用户更具体的需求。",
                        Parameters = new FunctionParameters()
                    }
                };
                        
                // 将提示工具添加到全部工具前面
                var result = new List<ToolDefinition> { hintTool };
                result.AddRange(allTools);
                return result;
            }
                    
            // 解析搜索结果中的命令名
            var matchedCommandNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var lines = searchResult.Split('\n');
                    
            foreach (var line in lines)
            {
                // 匹配格式：【分组】命令名
                if (line.StartsWith("【") && line.Contains("】"))
                {
                    int closeBracketIndex = line.IndexOf('】');
                    if (closeBracketIndex > 0 && closeBracketIndex < line.Length - 1)
                    {
                        string commandName = line.Substring(closeBracketIndex + 1).Trim();
                        // 去掉异步标记
                        int asyncIndex = commandName.IndexOf(" (异步)");
                        if (asyncIndex >= 0)
                        {
                            commandName = commandName.Substring(0, asyncIndex).Trim();
                        }
                                
                        if (!string.IsNullOrEmpty(commandName))
                        {
                            matchedCommandNames.Add(commandName.ToLower());
                        }
                    }
                }
            }
                    
            Console.WriteLine($"\n[调试] 从搜索结果中提取到 {matchedCommandNames.Count} 个命令名：{string.Join(", ", matchedCommandNames)}");
                    
            // 过滤工具列表，只保留匹配的命令
            var filteredTools = allTools
                .Where(t => 
                {
                    string funcName = t.Function.Name;
                    // 去掉 execute_前缀进行比较
                    if (funcName.StartsWith("execute_"))
                    {
                        funcName = funcName.Substring(8);
                    }
                    return matchedCommandNames.Contains(funcName.ToLower());
                })
                .ToList();
                    
            Console.WriteLine($"\n[调试] 过滤后剩余 {filteredTools.Count} 个工具");
            return filteredTools;
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
                string jsonBody = JsonConvert.SerializeObject(requestBodyObj);
                using var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
                
                Console.WriteLine($"\n[调试] 请求体 JSON 长度：{jsonBody.Length} 字符");
                Console.WriteLine($"[调试] HttpClient 状态：{(_httpClient != null ? "已初始化" : "未初始化")}");
                Console.WriteLine($"[调试] HttpClient Timeout: {_httpClient?.Timeout.TotalSeconds} 秒");
                
                // 设置 Header
                _httpClient!.DefaultRequestHeaders.Clear();
                _httpClient!.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
                
                Console.WriteLine($"[调试] Authorization Header 已设置");

                // 发送请求
                var request = new HttpRequestMessage
                {
                    Method = HttpMethod.Post,
                    RequestUri = new Uri(ApiUrl),
                    Content = content
                };
                
                Console.WriteLine($"\n[调试] 正在发送请求到：{ApiUrl}");
                Console.WriteLine($"[调试] 请求方法：{request.Method}");
                Console.WriteLine($"[调试] API Key 前缀：{apiKey.Substring(0, Math.Min(10, apiKey.Length))}...");
                
                HttpResponseMessage? response = null;
                try
                {
                    Console.WriteLine($"[调试] 即将调用 SendAsync...");
                    response = await _httpClient.SendAsync(
                        request,
                        HttpCompletionOption.ResponseHeadersRead
                    );
                    Console.WriteLine($"[调试] SendAsync 返回，response 为 {(response == null ? "null" : "非 null")}");
                    
                    if (response != null)
                    {
                        Console.WriteLine($"[调试] 响应状态码：{response.StatusCode}");
                        Console.WriteLine($"[调试] 响应是否成功：{response.IsSuccessStatusCode}");
                    }
                }
                catch (HttpRequestException httpEx)
                {
                    Console.WriteLine($"\n[严重错误] HTTP 请求异常：{httpEx.Message}");
                    Console.WriteLine($"[调试] 异常类型：{httpEx.GetType().FullName}");
                    if (httpEx.InnerException != null)
                    {
                        Console.WriteLine($"[调试] 内部异常：{httpEx.InnerException.Message}");
                        Console.WriteLine($"[调试] 内部异常类型：{httpEx.InnerException.GetType().FullName}");
                    }
                    throw;
                }
                catch (TaskCanceledException cancelEx)
                {
                    Console.WriteLine($"\n[严重错误] 请求被取消或超时：{cancelEx.Message}");
                    Console.WriteLine($"[调试] 是否因超时取消：{cancelEx.InnerException is TimeoutException}");
                    throw;
                }
                catch (Exception sendEx)
                {
                    Console.WriteLine($"\n[严重错误] 发送请求时出错：{sendEx.Message}");
                    Console.WriteLine($"[调试] 异常类型：{sendEx.GetType().FullName}");
                    if (sendEx.InnerException != null)
                    {
                        Console.WriteLine($"[调试] 内部异常：{sendEx.InnerException.Message}");
                        Console.WriteLine($"[调试] 内部异常类型：{sendEx.InnerException.GetType().FullName}");
                    }
                    throw;
                }
                
                // 处理非 200 状态码
                if (response == null)
                {
                    throw new HttpRequestException("API 请求失败：响应为 null");
                }
                
                if (!response.IsSuccessStatusCode)
                {
                    string errorBody;
                    try
                    {
                        errorBody = await response.Content.ReadAsStringAsync();
                    }
                    catch (Exception readEx)
                    {
                        errorBody = $"无法读取错误内容：{readEx.Message}";
                    }
                    Console.WriteLine($"\n[严重错误] API 请求失败 ({response.StatusCode}): {errorBody}");
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
                                    var jObject = JObject.Parse(dataJson);
                                    var root = jObject;

                                    var errorToken = root["error"];
                                    if (errorToken != null && errorToken["message"] != null)
                                    {
                                        string errMsg = errorToken["message"]!.ToString();
                                        throw new Exception($"流式传输中发生错误：{errMsg}");
                                    }

                                    if (root["choices"] is JArray choices && choices.Count > 0)
                                    {
                                        var choiceToken = choices[0];
                                        var contentToken = choiceToken?["delta"]?["content"];
                                        if (contentToken != null)
                                        {
                                            string chunk = contentToken.ToString();
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
        public async Task<string> CallStreamingWithImageAsync(List<ChatMessage> messages, string userPrompt, string? imagePath, string apiKey, string? model = null)
        {
            model ??= DefaultModel;
            
            // 如果 imagePath 为 null 或空，直接返回
            if (string.IsNullOrEmpty(imagePath))
            {
                throw new ArgumentException("imagePath 不能为空", nameof(imagePath));
            }
            
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
        /// 执行流式调用（带 Tool）
        /// </summary>
        public async Task<(string Response, List<ToolCall>? ToolCalls)> CallStreamingWithToolsAsync(
            List<ChatMessage> messages, 
            string userPrompt, 
            List<ToolDefinition> tools, 
            string apiKey, 
            string model)
        {
            // 添加用户消息
            var messagesWithUser = new List<object>();
            foreach (var m in messages)
            {
                messagesWithUser.Add(new { role = m.Role, content = m.Content });
            }
            messagesWithUser.Add(new { role = "user", content = userPrompt });

            var requestBody = new
            {
                model = model,
                stream = false, // Tool 调用不使用流式
                messages = messagesWithUser.ToArray(),
                tools = tools.Select(t => new
                {
                    type = t.Type,
                    function = new
                    {
                        name = t.Function.Name,
                        description = t.Function.Description,
                        parameters = t.Function.Parameters
                    }
                }).ToArray()
            };
            
            string jsonBody = JsonConvert.SerializeObject(requestBody);
            using var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
            
            _httpClient!.DefaultRequestHeaders.Clear();
            _httpClient!.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
            
            var request = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(ApiUrl),
                Content = content
            };
            
            Console.WriteLine($"\n[调试] 发送 Tool 调用请求到：{ApiUrl}");
            
            var response = await _httpClient.SendAsync(request);
            response.EnsureSuccessStatusCode();
            
            var responseJson = await response.Content.ReadAsStringAsync();
            var result = JObject.Parse(responseJson);
            
            var choices = result["choices"] as JArray;
            if (choices == null || choices.Count == 0)
            {
                return ("", null);
            }
            
            var message = choices[0]["message"];
            var toolCalls = message["tool_calls"] as JArray;
            
            if (toolCalls != null && toolCalls.Count > 0)
            {
                var calls = toolCalls.ToObject<List<ToolCall>>()!;
                Console.WriteLine($"\n[调试] 检测到 {calls.Count} 个 Tool 调用");
                return ("", calls);
            }
            
            var contentToken = message["content"];
            return (contentToken?.ToString() ?? "", null);
        }

        /// <summary>
        /// 保存长期记忆日志
        /// </summary>
        private void SaveLongTermMemoryLog(string content)
        {
            try
            {
                
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                File.AppendAllText(_logFilePath, $"[{timestamp}] {content}\n", Encoding.UTF8);
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
    /// Tool 定义（用于 Function Calling）
    /// </summary>
    public class ToolDefinition
    {
        [JsonProperty("type")]
        public string Type { get; set; } = "function";
        
        [JsonProperty("function")]
        public FunctionDefinition Function { get; set; }
    }
    
    /// <summary>
    /// 函数定义
    /// </summary>
    public class FunctionDefinition
    {
        [JsonProperty("name")]
        public string Name { get; set; }
        
        [JsonProperty("description")]
        public string Description { get; set; }
        
        [JsonProperty("parameters")]
        public object Parameters { get; set; }
    }
    
    /// <summary>
    /// 函数参数定义
    /// </summary>
    public class FunctionParameters
    {
        [JsonProperty("type")]
        public string Type { get; set; } = "object";
        
        [JsonProperty("properties")]
        public Dictionary<string, PropertyDefinition> Properties { get; set; } = new Dictionary<string, PropertyDefinition>();
        
        [JsonProperty("required")]
        public List<string> Required { get; set; } = new List<string>();
    }
    
    /// <summary>
    /// 属性定义
    /// </summary>
    public class PropertyDefinition
    {
        [JsonProperty("type")]
        public string Type { get; set; } = "string";
        
        [JsonProperty("description")]
        public string Description { get; set; }
    }
    
    /// <summary>
    /// Tool 调用响应
    /// </summary>
    public class ToolCall
    {
        [JsonProperty("id")]
        public string Id { get; set; }
        
        [JsonProperty("type")]
        public string Type { get; set; } = "function";
        
        [JsonProperty("function")]
        public FunctionCall Function { get; set; }
    }
    
    /// <summary>
    /// 函数调用
    /// </summary>
    public class FunctionCall
    {
        [JsonProperty("name")]
        public string Name { get; set; }
        
        [JsonProperty("arguments")]
        public string Arguments { get; set; }
    }
    
    /// <summary>
    /// 聊天消息类
    /// </summary>
    public class ChatMessage
    {
        [JsonProperty("role")]
        public string Role { get; set; } = "";
        
        [JsonProperty("content")]
        public string Content { get; set; } = "";
        
        // Tool 调用相关字段
        [JsonProperty("tool_call_id", NullValueHandling = NullValueHandling.Ignore)]
        public string? ToolCallId { get; set; }
        
        [JsonProperty("tool_calls", NullValueHandling = NullValueHandling.Ignore)]
        public List<ToolCall>? ToolCalls { get; set; }
    }
}
