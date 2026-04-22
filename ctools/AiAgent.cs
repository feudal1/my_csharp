using System;
using System.Threading.Tasks;

namespace ctool
{
    /// <summary>
    /// 最小 AI Agent：仅提供对话能力，供 sw_plugin 直接引用调用。
    /// </summary>
    public class AiAgent
    {
        private readonly tools.LlmService _llmService;

        public AiAgent(Func<string>? getCommandsDescriptionFunc = null)
        {
            _llmService = new tools.LlmService(getCommandsDescriptionFunc);
        }

        public Task<string> ChatAsync(string prompt)
        {
            if (string.IsNullOrWhiteSpace(prompt))
            {
                return Task.FromResult(string.Empty);
            }

            return _llmService.ChatAsync(prompt);
        }

        public void ClearHistory()
        {
            _llmService.ClearHistory();
        }
    }
}
