# my_ai - SolidWorks 智能助手

<div align="center">

🤖 基于 AI 的 SolidWorks 自动化工具集 | 🎯 智能对话控制 | ⚡ 高效批量处理

[主要功能](#-主要功能) • [快速开始](#-快速开始) • [使用文档](#-使用文档) • [项目结构](#-项目结构)

</div>

---

## 📖 项目简介

**my_ai** 是一个功能强大的 SolidWorks 二次开发项目，集成了 AI 智能对话、命令行工具和插件界面，为机械工程师提供自动化设计解决方案。

### ✨ 核心特性

- **🧠 AI 智能对话** - 通过自然语言描述需求，AI 自动识别并执行相应操作
- **⌨️ 命令行工具** - 支持直接输入命令快速执行 SolidWorks 任务
- **🖱️ 右键菜单集成** - 在 SolidWorks 中无缝集成自定义功能菜单
- **📦 丰富功能库** - 涵盖零件、装配体、工程图、CAD 文件处理等多种场景
- **🔍 模糊搜索** - 支持命令相似度匹配，快速定位所需功能
- **📊 性能监控** - 内置性能分析工具，实时追踪命令执行效率

---

## 🚀 主要功能

### ctools 命令行工具

```bash
# 启动交互式对话模式
ctool.exe

# 示例命令
> exportdxf                    # 导出 DXF 文件
> 导出当前零件                 # AI 识别意图并调用相应命令
> search export                # 搜索相关命令
> history                      # 查看历史对话
> clear                        # 清空历史记录
> mode                         # 切换工作模式
> llm                          # 进入纯对话模式
> quit/exit                    # 退出程序
```

### plugin SolidWorks 插件

- ✅ 自动加载到 SolidWorks 菜单栏
- ✅ 提供欢迎界面和版本信息
- ✅ 集成控制台输出窗口
- ✅ 支持右键菜单快捷操作
- ✅ 命令管理器集成

### share 功能库

#### 🔧 零件操作 (part/)
- `exportdwg` - 导出 DWG 工程图
- `get_thickness` - 获取钣金厚度
- `unsuppress` - 解除压缩特征
- `new_drw` - 新建工程图
- `select_part_byname` - 按名称选择零件

#### 🔩 装配体操作 (asm/)
- `asm2bom` - 生成 BOM 表
- `asm2do` - 装配体转工程图
- `asm2step` - 导出 STEP 格式
- `get_all_part_name` - 获取所有零件名称

#### 📄 工程图操作 (drw/)
- `drw2dwg` - 工程图转 DWG
- `drw2dxf` - 工程图转 DXF
- `drw2png` - 导出 PNG 图片
- `get_all_visable_edge` - 获取可见边线

#### 📐 CAD 文件处理 (cad/)
- `dwg2dxf` - DWG 转 DXF
- `dxf2dwg` - DXF 转 DWG
- `merge_dwg` - 合并 DWG 文件
- `draw_divider` - 绘制分隔线

#### 🧮 高级功能 (train/)
- `topology_labeler` - 拓扑结构标注
- `similarity_calculator` - 模型相似度计算
- `face_graph_builder` - 面图构建
- `WL 图核算法` - Weisfeiler-Lehman 图核算法

---

## 📦 安装与配置

### 环境要求

- **SolidWorks** - 已安装并正确配置
- **.NET SDK** - .NET 9.0 或更高版本
- **Visual Studio** - 用于编译项目（可选）

### 编译项目

```bash
# 进入项目目录
cd e:\cqh\code\my_c#

# 使用 Visual Studio 打开 my_ai.sln
# 或使用命令行编译
dotnet build -c Release
```

### 注册 SolidWorks 插件

#### 方法一：使用注册脚本（推荐）

1. **右键点击** `plugin\register_addin.bat`
2. **选择"以管理员身份运行"**
3. 等待注册成功的提示
4. 按照提示在 SolidWorks 中启用插件

#### 方法二：手动注册

```bash
# 以管理员身份打开命令行，进入项目目录
cd e:\cqh\code\my_c#

# 执行注册命令
%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe "plugin\bin\Release\net48\plugin.dll" /codebase /tlb
```

---

## 🔧 在 SolidWorks 中启用插件

1. **打开 SolidWorks**
2. 点击顶部菜单栏的 **"工具"** (Tools)
3. 选择 **"插件"** (Add-Ins)
4. 在弹出的对话框中找到 **"SolidWorksAddinStudy"**
5. **勾选** 该插件名称旁边的复选框
   - ✓ 勾选当前会话：仅本次有效
   - ✓ 勾选"启动时加载"：每次启动自动加载
6. 点击 **"确定"**

---

## 💡 使用指南

### ctools 使用示例

#### 1. 直接命令模式

```bash
> exportdxf
> get_thickness
> asm2bom
```

#### 2. 自然语言模式

```bash
> 导出当前零件的 DXF 文件
> 获取这个钣金的厚度
> 生成装配体的材料明细表
```

#### 3. 命令搜索

```bash
> search export              # 搜索包含"export"的命令
> find dwg                   # 查找与 DWG 相关的功能
```

#### 4. 查看帮助

```bash
> help                       # 显示所有可用命令
> exportdxf -h              # 查看特定命令的帮助
```

### plugin 插件使用

#### 访问插件功能

1. **菜单栏** - SolidWorks 顶部菜单会出现自定义工具栏
2. **右键菜单** - 在图形区域右键可看到快捷菜单
3. **控制台** - 按提示打开控制台查看输出信息

#### 常用操作

- **显示控制台** - 点击插件菜单中的"显示控制台"
- **查看版本** - 启动时自动显示欢迎界面和版本信息
- **清空缓存** - 欢迎界面倒计时结束后可清空工程文件

---

## 🗂️ 项目结构

```
my_ai/
├── ctools/                     # 命令行工具主程序
│   ├── main.cs                 # 程序入口和命令调度
│   ├── llm_service.cs          # AI 大模型服务
│   ├── llm_loop_caller.cs      # AI 对话循环控制器
│   ├── command_executor.cs     # 命令执行器
│   ├── CommandAttribute.cs     # 命令特性定义
│   └── connect.cs              # SolidWorks 连接模块
│
├── plugin/                     # SolidWorks 插件
│   ├── addin.cs                # 插件主程序
│   ├── ConsoleOutputForm.cs    # 控制台输出窗口
│   ├── body_context_menu.cs    # 实体右键菜单
│   ├── function_adder.cs       # 功能添加器
│   └── welcome.png             # 欢迎界面图片
│
├── share/                      # 共享功能库
│   ├── part/                   # 零件相关功能
│   │   ├── exportdwg.cs        # 导出 DWG
│   │   ├── get_thickness.cs    # 获取厚度
│   │   └── ...
│   ├── asm/                    # 装配体相关功能
│   │   ├── asm2bom.cs          # 生成 BOM
│   │   ├── asm2do.cs           # 转工程图
│   │   └── ...
│   ├── drw/                    # 工程图相关功能
│   │   ├── drw2dwg.cs          # 转 DWG
│   │   ├── drw2dxf.cs          # 转 DXF
│   │   └── ...
│   ├── cad/                    # CAD 文件处理
│   │   ├── dwg2dxf.cs          # DWG 转 DXF
│   │   ├── merge_dwg.cs        # 合并 DWG
│   │   └── ...
│   ├── train/                  # AI 训练和算法
│   │   ├── topology_labeler.cs # 拓扑标注
│   │   ├── similarity_calculator.cs
│   │   └── wl_graph_kernel.cs  # WL 图核
│   └── nomal/                  # 通用工具
│       ├── comhelp.cs          # COM 帮助类
│       └── get_folder_file.cs  # 文件夹遍历
│
├── reference/                  # 引用 DLL 文件
│   ├── Autodesk.AutoCAD.Interop.*.dll
│   ├── SolidWorks.Interop.*.dll
│   └── solidworkstools.dll
│
├── design_notes/               # 设计笔记
│   ├── draw_knowledge.txt      # 制图知识
│   ├── works_knowledge.txt     # 工作知识
│   └── *.png                   # 示意图
│
├── my_ai.sln                   # Visual Studio 解决方案
└── README.md                   # 本说明文档
```

---

## 🔧 技术细节

### 命令系统架构

- **CommandAttribute** - 命令特性标记，定义命令名称、描述、参数
- **CommandRegistry** - 全局命令注册中心，管理所有命令
- **CommandExecutor** - 命令执行器，解析和执行命令
- **LlmLoopCaller** - AI 对话循环控制器，集成命令解析

### AI 集成

- **服务提供商** - 阿里云通义千问 (DashScope)
- **支持模型** - qwen3.5-flash（默认）、VLM 图像分析
- **功能特性** - 
  - 动态构建 System Prompt
  - 长短期记忆管理
  - 工作知识库集成
  - 命令描述动态生成

### SolidWorks API

- **接口实现** - ISwAddin 标准接口
- **COM 注册** - Guid: `D9C5D3A4-3B9F-4ACF-BC19-6D77D39C47CD`
- **命令管理器** - ICommandManager 集成
- **文档类型** - 零件、装配体、工程图全支持

---

## ⚠️ 常见问题

### Q1: 插件注册失败

**解决方案：**
- ✓ 确保以**管理员身份**运行注册脚本
- ✓ 检查 DLL 文件是否存在于 `plugin\bin\Release\net48\` 目录
- ✓ 确认 SolidWorks 版本兼容性

### Q2: 在 SolidWorks 中找不到插件

**解决方案：**
- ✓ 重新运行 `register_addin.bat`
- ✓ 重启 SolidWorks
- ✓ 检查注册表项：`HKEY_CURRENT_USER\Software\SolidWorks\AddInsStartup`

### Q3: ctools 无法连接 SolidWorks

**解决方案：**
- ✓ 先启动 SolidWorks 应用程序
- ✓ 确保有激活的文档
- ✓ 以管理员身份运行 ctool.exe

### Q4: 命令执行无响应

**解决方案：**
- ✓ 查看控制台输出信息
- ✓ 检查 SolidWorks 是否弹出错误提示
- ✓ 确认当前文档类型是否符合命令要求

### Q5: AI 对话无法识别命令

**解决方案：**
- ✓ 使用更明确的命令描述
- ✓ 使用 `search` 命令查看可用命令列表
- ✓ 切换到直接命令模式

---

## 🗑️ 卸载插件

### 方法一：使用卸载脚本

1. **右键点击** `plugin\unregister_addin.bat`
2. **选择"以管理员身份运行"**

### 方法二：手动卸载

```bash
# 以管理员身份打开命令行
%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe "plugin\bin\Release\net48\plugin.dll" /unregister
```

### 方法三：在 SolidWorks 中移除

1. 打开 SolidWorks
2. 点击 **"工具"** > **"插件"**
3. 取消勾选 **"SolidWorksAddinStudy"**
4. 点击 **"确定"**

---

## 📝 开发规范

### 命令命名规范

- ✅ 使用小写字母和下划线：`export_dwg`, `get_thickness`
- ✅ 动词 + 名词结构：`export*`, `get*`, `create_*`
- ❌ 避免使用大写：`ExportDXF` (错误)
- ❌ 避免使用中文：`导出 dxf` (错误)

### 命令特性定义

```csharp
[Command("exportdxf")]
[Description("导出当前零件的 DXF 文件")]
[Parameters("无")]
[Group("part")]
public static void ExportDXF(string[] args)
{
    // 命令实现
}
```

### 示例格式规范

- **无参数命令**：直接使用命令名（如 `rename`）
- **有参数命令**：`命令名 [参数值]` 格式
- **禁止使用**：过时的 `do_【命令名】` 或 `do_【命令名】<参数值>` 格式

---

## 🔗 相关链接

- [SolidWorks API 文档](https://help.solidworks.com/2025/english/api/sldworksapi/)
- [通义千问 API 文档](https://help.aliyun.com/zh/dashscope/)
- [.NET 官方文档](https://learn.microsoft.com/zh-cn/dotnet/)

---

## 👥 技术支持

### 交流群

点击链接加入群聊【solidworks 神经网络自动标注小白群】：  
https://qm.qq.com/q/n5HGmImlCC

### 问题反馈

如遇到问题，请通过以下方式反馈：
1. GitHub Issues（如有）
2. QQ 群内提问
3. 邮件联系开发者

---

## 📄 许可证

本项目仅供学习和研究使用。

---

## 🙏 致谢

感谢以下开源项目和工具：
- SolidWorks API
- 阿里云通义千问
- Newtonsoft.Json
- AutoCAD Interop

---

<div align="center">

**Made with ❤️ by my_ai Team**  
*最后更新：2026 年 3 月 27 日*

</div>
