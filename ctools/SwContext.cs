using SolidWorks.Interop.sldworks;

namespace tools
{
    /// <summary>
    /// SolidWorks 上下文管理器（单例模式）
    /// 提供全局可访问的 SolidWorks 应用程序和文档实例
    /// </summary>
    public class SwContext
    {
        private static readonly Lazy<SwContext> _instance = new Lazy<SwContext>(() => new SwContext());
        
        public static SwContext Instance => _instance.Value;
        
        private SldWorks? _swApp;
        private ModelDoc2? _swModel;
        private object _lock = new object();
        
        /// <summary>
        /// 私有构造函数（单例模式）
        /// </summary>
        private SwContext()
        {
        }
        
        /// <summary>
        /// SolidWorks 应用程序实例
        /// </summary>
        public SldWorks? SwApp
        {
            get
            {
                lock (_lock)
                {
                    return _swApp;
                }
            }
            set
            {
                lock (_lock)
                {
                    _swApp = value;
                }
            }
        }
        
        /// <summary>
        /// 当前激活的文档
        /// </summary>
        public ModelDoc2? SwModel
        {
            get
            {
                lock (_lock)
                {
                    return _swModel;
                }
            }
            set
            {
                lock (_lock)
                {
                    _swModel = value;
                }
            }
        }
        
        /// <summary>
        /// 初始化 SolidWorks 上下文
        /// </summary>
        public void Initialize(SldWorks? swApp, ModelDoc2? swModel)
        {
            SwApp = swApp;
            SwModel = swModel;
        }
        
        /// <summary>
        /// 清除 SolidWorks 上下文
        /// </summary>
        public void Clear()
        {
            SwApp = null;
            SwModel = null;
        }
    }
}
