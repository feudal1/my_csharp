# SolidWorks 集成

<cite>
**本文引用的文件**
- [sw_plugin/addin.cs](file://sw_plugin/addin.cs)
- [sw_plugin/function.cs](file://sw_plugin/function.cs)
- [sw_plugin/body_context_menu.cs](file://sw_plugin/body_context_menu.cs)
- [sw_plugin/register_addin.bat](file://sw_plugin/register_addin.bat)
- [sw_plugin/unregister_addin.bat](file://sw_plugin/unregister_addin.bat)
- [sw_plugin/CommandAttribute.cs](file://sw_plugin/CommandAttribute.cs)
- [ctools/SwContext.cs](file://ctools/SwContext.cs)
- [ctools/connect.cs](file://ctools/connect.cs)
- [ctools/CommandRegistry.cs](file://ctools/CommandRegistry.cs)
- [ctools/CommandAttribute.cs](file://ctools/CommandAttribute.cs)
- [ctools/command_executor.cs](file://ctools/command_executor.cs)
- [ctools/main.cs](file://ctools/main.cs)
- [ctools/solidworks_commands/part_commands.cs](file://ctools/solidworks_commands/part_commands.cs)
- [ctools/solidworks_commands/asm_commands.cs](file://ctools/solidworks_commands/asm_commands.cs)
- [ctools/solidworks_commands/drw_commands.cs](file://ctools/solidworks_commands/drw_commands.cs)
- [ctools/llm_loop_caller.cs](file://ctools/llm_loop_caller.cs)
- [share/drw/benddim.cs](file://share/drw/benddim.cs)
- [cad_plugin/cad_addin.cs](file://cad_plugin/cad_addin.cs)
- [cad_plugin/register.ps1](file://cad_plugin/register.ps1)
- [cad_plugin/unregister.ps1](file://cad_plugin/unregister.ps1)
</cite>

## 目录
1. [简介](#简介)
2. [项目结构](#项目结构)
3. [核心组件](#核心组件)
4. [架构总览](#架构总览)
5. [详细组件分析](#详细组件分析)
6. [依赖关系分析](#依赖关系分析)
7. [性能考虑](#性能考虑)
8. [故障排查指南](#故障排查指南)
9. [结论](#结论)
10. [附录](#附录)

## 简介
本技术文档面向 SolidWorks 插件开发者，系统阐述基于 COM 组件的插件架构与实现要点，涵盖以下主题：
- ISwAddin 接口的标准实现与 Guid 配置
- 插件注册与生命周期管理（自动加载与手动注册）
- 命令管理器集成与实体右键菜单实现
- SwContext 上下文管理器的作用与使用方式
- SolidWorks API 最佳实践与性能优化建议
- 常见集成问题与调试技巧
- **新增**：弯曲标注功能集成（DimensionBends 命令）

目标是帮助开发者快速理解 COM 组件编程与 SolidWorks API 的使用方法，并在实际项目中稳定落地。

## 项目结构
该项目由两部分组成：
- sw_plugin：SolidWorks 插件主程序，负责 COM 插件生命周期、命令注册、UI 交互与实体右键菜单集成
- ctools：命令与工具集合，提供命令注册中心、命令执行器、LLM 对话循环调用器以及大量 SolidWorks 命令实现

```mermaid
graph TB
subgraph "SolidWorks 插件"
A["addin.cs<br/>ISwAddin 实现<br/>COM 注册/反注册"]
B["function.cs<br/>命令入口Command 特性"]
C["body_context_menu.cs<br/>实体右键菜单"]
D["register_addin.bat / unregister_addin.bat<br/>注册/卸载脚本"]
E["CommandAttribute.cs<br/>命令特性定义"]
end
subgraph "命令与工具集合"
F["SwContext.cs<br/>上下文管理器"]
G["connect.cs<br/>连接 SolidWorks"]
H["CommandRegistry.cs<br/>命令注册中心"]
I["CommandAttribute.cs<br/>命令特性"]
J["command_executor.cs<br/>命令执行器"]
K["main.cs<br/>命令注册与交互入口"]
L["part_commands.cs / asm_commands.cs<br/>命令实现"]
M["drw_commands.cs<br/>工程图命令实现"]
N["llm_loop_caller.cs<br/>LLM 对话循环调用器"]
O["benddim.cs<br/>弯曲标注核心逻辑"]
end
A --> B
A --> C
A --> E
B --> G
H --> J
J --> K
K --> L
K --> M
M --> O
K --> N
```

**图表来源**
- [sw_plugin/addin.cs:18-339](file://sw_plugin/addin.cs#L18-L339)
- [sw_plugin/function.cs:29-698](file://sw_plugin/function.cs#L29-L698)
- [sw_plugin/body_context_menu.cs:141-166](file://sw_plugin/body_context_menu.cs#L141-L166)
- [sw_plugin/CommandAttribute.cs:1-27](file://sw_plugin/CommandAttribute.cs#L1-L27)
- [ctools/SwContext.cs:9-85](file://ctools/SwContext.cs#L9-L85)
- [ctools/connect.cs:9-51](file://ctools/connect.cs#L9-L51)
- [ctools/CommandRegistry.cs:12-242](file://ctools/CommandRegistry.cs#L12-L242)
- [ctools/CommandAttribute.cs:5-18](file://ctools/CommandAttribute.cs#L5-L18)
- [ctools/command_executor.cs:12-116](file://ctools/command_executor.cs#L12-L116)
- [ctools/main.cs:34-377](file://ctools/main.cs#L34-L377)
- [ctools/solidworks_commands/part_commands.cs:11-149](file://ctools/solidworks_commands/part_commands.cs#L11-L149)
- [ctools/solidworks_commands/asm_commands.cs:11-158](file://ctools/solidworks_commands/asm_commands.cs#L11-L158)
- [ctools/solidworks_commands/drw_commands.cs:156-162](file://ctools/solidworks_commands/drw_commands.cs#L156-L162)
- [ctools/llm_loop_caller.cs:19-800](file://ctools/llm_loop_caller.cs#L19-L800)
- [share/drw/benddim.cs:11-331](file://share/drw/benddim.cs#L11-331)

**章节来源**
- [sw_plugin/addin.cs:18-339](file://sw_plugin/addin.cs#L18-L339)
- [ctools/main.cs:54-109](file://ctools/main.cs#L54-L109)

## 核心组件
- COM 插件宿主与生命周期
  - 通过 Guid 与 SwAddin 特性声明插件元数据，实现 ISwAddin 接口，完成连接与断开 SolidWorks 的回调
  - 提供 ComRegisterFunction/ComUnregisterFunction 完成注册表写入与清理
- 命令系统
  - 基于 Command 特性的命令注册中心，支持同步与异步命令
  - 命令执行器统一解析命令、解析参数、连接 SolidWorks、更新上下文并执行
  - **新增**：支持声明式命令注册，通过 CommandAttribute 特性定义命令属性
- 上下文管理
  - SwContext 单例提供全局可访问的 SldWorks 与 ModelDoc2 实例，线程安全更新
- 实体右键菜单
  - 通过 AddMenuPopupItem2 为特征管理器树中的实体添加右键菜单项
- **新增**：弯曲标注功能
  - 通过 benddim 类实现复杂的折弯特征识别与标注逻辑
  - 支持自动检测折弯特征、计算标注位置、生成尺寸标注
- LLM 集成
  - LlmLoopCaller 支持 Tool 调用模式，将命令封装为工具，实现自然语言到命令的映射与执行

**章节来源**
- [sw_plugin/addin.cs:18-339](file://sw_plugin/addin.cs#L18-L339)
- [sw_plugin/CommandAttribute.cs:1-27](file://sw_plugin/CommandAttribute.cs#L1-L27)
- [ctools/CommandRegistry.cs:12-242](file://ctools/CommandRegistry.cs#L12-L242)
- [ctools/command_executor.cs:12-116](file://ctools/command_executor.cs#L12-L116)
- [ctools/SwContext.cs:9-85](file://ctools/SwContext.cs#L9-L85)
- [sw_plugin/body_context_menu.cs:141-166](file://sw_plugin/body_context_menu.cs#L141-L166)
- [share/drw/benddim.cs:11-331](file://share/drw/benddim.cs#L11-331)
- [ctools/llm_loop_caller.cs:19-800](file://ctools/llm_loop_caller.cs#L19-L800)

## 架构总览
下图展示 SolidWorks 插件的总体架构与数据流：

```mermaid
graph TB
subgraph "SolidWorks"
SW["SldWorks 应用实例"]
DOC["ModelDoc2 文档实例"]
END["工程图视图"]
END --> SW
END --> DOC
END --> BEND["折弯特征"]
BEND --> SW
BEND --> DOC
END --> BEND
END --> DIM["尺寸标注"]
DIM --> SW
END --> DIM
end
subgraph "插件宿主"
ADDIN["AddinStudy<br/>ISwAddin 实现"]
CMGR["ICommandManager<br/>命令管理器"]
POP["右键菜单<br/>AddMenuPopupItem2"]
CMDATTR["CommandAttribute<br/>命令特性"]
END
subgraph "命令层"
REG["CommandRegistry<br/>命令注册中心"]
EXE["CommandExecutor<br/>命令执行器"]
CMD["命令实现<br/>part_commands.cs / asm_commands.cs / drw_commands.cs"]
END
subgraph "上下文与工具"
CTX["SwContext<br/>全局上下文"]
CONN["Connect<br/>连接器"]
LLM["LlmLoopCaller<br/>对话循环"]
BENDDIM["benddim<br/>弯曲标注核心"]
END
subgraph "弯曲标注流程"
SEL["选择视图"] --> FIND["查找折弯特征"]
FIND --> DETECT["检测圆柱面"]
DETECT --> PARALLEL["寻找平行边"]
PARALLEL --> LEVEL["多级面分析"]
LEVEL --> SELECT["选择标注面"]
SELECT --> PLACE["确定标注位置"]
PLACE --> DIM["生成尺寸标注"]
END
SW --> ADDIN
ADDIN --> CMGR
ADDIN --> POP
ADDIN --> CMDATTR
ADDIN --> CTX
CTX --> CONN
REG --> EXE
EXE --> CMD
CMD --> BENDDIM
BENDDIM --> SEL
BENDDIM --> FIND
BENDDIM --> DETECT
BENDDIM --> PARALLEL
BENDDIM --> LEVEL
BENDDIM --> SELECT
BENDDIM --> PLACE
BENDDIM --> DIM
EXE --> SW
LLM --> REG
LLM --> EXE
```

**图表来源**
- [sw_plugin/addin.cs:96-120](file://sw_plugin/addin.cs#L96-L120)
- [sw_plugin/body_context_menu.cs:141-166](file://sw_plugin/body_context_menu.cs#L141-L166)
- [sw_plugin/CommandAttribute.cs:18-25](file://sw_plugin/CommandAttribute.cs#L18-L25)
- [ctools/CommandRegistry.cs:12-242](file://ctools/CommandRegistry.cs#L12-L242)
- [ctools/command_executor.cs:12-116](file://ctools/command_executor.cs#L12-L116)
- [ctools/SwContext.cs:9-85](file://ctools/SwContext.cs#L9-L85)
- [ctools/connect.cs:9-51](file://ctools/connect.cs#L9-L51)
- [ctools/llm_loop_caller.cs:44-67](file://ctools/llm_loop_caller.cs#L44-L67)
- [ctools/solidworks_commands/drw_commands.cs:156-162](file://ctools/solidworks_commands/drw_commands.cs#L156-L162)
- [share/drw/benddim.cs:13-62](file://share/drw/benddim.cs#L13-62)

## 详细组件分析

### COM 插件宿主与生命周期（ISwAddin 实现）
- Guid 与 SwAddin 特性
  - 通过 Guid 与 SwAddin 特性声明插件标识、标题与描述，并指定启动时加载
- 生命周期回调
  - ConnectToSW：保存 SldWorks 与 Cookie，获取 ICommandManager，初始化命令注册中心与右键菜单
  - DisconnectFromSW：插件卸载时的清理工作
- 注册与反注册
  - ComRegisterFunction/ComUnregisterFunction：向注册表写入/删除插件信息，支持启动时自动加载

```mermaid
sequenceDiagram
participant SW as "SolidWorks"
participant Host as "AddinStudy(ISwAddin)"
participant Reg as "注册表"
participant Cmd as "命令注册中心"
SW->>Host : 调用 ConnectToSW(ThisSW, Cookie)
Host->>Host : 保存 SldWorks 与 Cookie
Host->>SW : GetCommandManager(Cookie)
Host->>Cmd : InitializeCommandRegistry()
Host->>SW : AddMenuPopupItem2(...)
SW-->>Host : 返回连接状态
Note over Host : 插件加载完成
SW->>Host : 调用 DisconnectFromSW()
Host-->>Reg : 清理注册表项
```

**图表来源**
- [sw_plugin/addin.cs:96-120](file://sw_plugin/addin.cs#L96-L120)
- [sw_plugin/addin.cs:262-333](file://sw_plugin/addin.cs#L262-L333)

**章节来源**
- [sw_plugin/addin.cs:18-339](file://sw_plugin/addin.cs#L18-L339)

### 命令管理器集成与命令注册
- 命令注册中心
  - 单例模式，支持从程序集与实例类型批量注册命令，维护命令名到 CommandInfo 的映射
  - 支持命令别名注册，统一解析与执行
- 命令执行器
  - 解析命令文本，提取命令名与参数，解析当前激活文档，调用 CommandInfo.AsyncAction
  - 统一异常处理与日志输出
- **新增**：声明式命令注册
  - 通过 CommandAttribute 特性定义命令属性，包括命令 ID、名称、工具提示、本地化名称和文档类型
  - 支持 UI 命令和命令行命令的统一管理
- 命令实现
  - 通过 Command 特性标注命令，支持同步与异步方法；在插件中通过 Command 特性标注命令入口

```mermaid
classDiagram
class CommandRegistry {
+RegisterCommand(commandInfo)
+RegisterAssembly(assembly)
+RegisterType(instance, type)
+GetCommand(name) CommandInfo?
+GetAllCommands() Dictionary
+Clear()
}
class CommandExecutor {
+ExecuteCommandAsync(commandText) Task~string~
}
class CommandAttribute {
+int Id
+string Name
+string Tooltip
+string LocalizedName
+int[] DocumentTypes
+bool ShowOutputWindow
}
class CommandInfo {
+string Name
+string Description
+string Parameters
+string Group
+string[] Aliases
+CommandType CommandType
+AsyncAction(args) Task
}
CommandRegistry --> CommandInfo : "管理"
CommandExecutor --> CommandRegistry : "解析命令"
CommandExecutor --> CommandInfo : "调用 AsyncAction"
CommandAttribute <.. CommandInfo : "特性驱动"
```

**图表来源**
- [ctools/CommandRegistry.cs:12-242](file://ctools/CommandRegistry.cs#L12-L242)
- [ctools/command_executor.cs:12-116](file://ctools/command_executor.cs#L12-L116)
- [ctools/CommandAttribute.cs:5-18](file://ctools/CommandAttribute.cs#L5-L18)
- [sw_plugin/CommandAttribute.cs:8-26](file://sw_plugin/CommandAttribute.cs#L8-L26)

**章节来源**
- [ctools/CommandRegistry.cs:12-242](file://ctools/CommandRegistry.cs#L12-L242)
- [ctools/command_executor.cs:12-116](file://ctools/command_executor.cs#L12-L116)
- [ctools/CommandAttribute.cs:5-18](file://ctools/CommandAttribute.cs#L5-L18)
- [sw_plugin/CommandAttribute.cs:1-27](file://sw_plugin/CommandAttribute.cs#L1-L27)
- [sw_plugin/function.cs:627-657](file://sw_plugin/function.cs#L627-L657)

### 实体右键菜单实现
- 初始化右键菜单
  - 通过 AddMenuPopupItem2 为特定文档类型与选择类型添加右键菜单项
  - 将菜单项与具体方法绑定，实现"从实体创建工程图""导出 STEP"等功能
- 选择实体处理
  - 在方法中获取当前文档与选中实体，必要时切换到组件模型文档

```mermaid
flowchart TD
Start(["初始化右键菜单"]) --> CheckSW["检查 SldWorks 与 ICommandManager"]
CheckSW --> AddItems["为 PART/ASSEMBLY + FACE 选择类型添加菜单项"]
AddItems --> BindActions["绑定菜单项到具体方法"]
BindActions --> End(["完成"])
subgraph "菜单点击处理"
Click(["用户点击菜单"]) --> GetDoc["获取当前文档"]
GetDoc --> GetSel["获取选中实体"]
GetSel --> HasSel{"是否选中实体?"}
HasSel --> |否| Warn["提示用户先选中实体"]
HasSel --> |是| Exec["执行对应命令"]
Exec --> Done(["完成"])
Warn --> Done
end
```

**图表来源**
- [sw_plugin/body_context_menu.cs:141-166](file://sw_plugin/body_context_menu.cs#L141-L166)
- [sw_plugin/body_context_menu.cs:19-133](file://sw_plugin/body_context_menu.cs#L19-L133)

**章节来源**
- [sw_plugin/body_context_menu.cs:141-166](file://sw_plugin/body_context_menu.cs#L141-L166)
- [sw_plugin/body_context_menu.cs:19-133](file://sw_plugin/body_context_menu.cs#L19-L133)

### SwContext 上下文管理器
- 单例模式，提供全局可访问的 SldWorks 与 ModelDoc2 实例
- 线程安全：通过锁保护读写，避免并发访问导致的状态不一致
- 初始化与清理：在插件连接与断开时分别设置与清空上下文

```mermaid
classDiagram
class SwContext {
-static Lazy~SwContext~ _instance
+static Instance : SwContext
-SldWorks _swApp
-ModelDoc2 _swModel
-object _lock
+SwApp : SldWorks?
+SwModel : ModelDoc2?
+Initialize(swApp, swModel)
+Clear()
}
```

**图表来源**
- [ctools/SwContext.cs:9-85](file://ctools/SwContext.cs#L9-L85)

**章节来源**
- [ctools/SwContext.cs:9-85](file://ctools/SwContext.cs#L9-L85)
- [sw_plugin/addin.cs:105-111](file://sw_plugin/addin.cs#L105-L111)

### **新增**：弯曲标注功能集成
- **benddim 类核心功能**
  - 自动识别折弯特征：遍历模型体的特征，查找 OneBend 类型的折弯特征
  - 几何分析：检测圆柱面、平行边、多级面结构，确定标注面
  - 标注生成：根据视图变换和几何关系，自动放置尺寸标注
- **DimensionBends 命令实现**
  - UI 命令：通过 CommandAttribute 特性定义，支持工程图文档类型
  - 命令行命令：通过 [Command] 特性注册，支持 dimension_bends 命令
  - 集成调用：调用 benddim.标折弯尺寸(swApp) 执行核心逻辑
- **弯曲标注算法流程**
  1. 选择视图验证：确保选择了工程图视图
  2. 特征扫描：遍历折弯特征，识别最大圆柱面
  3. 边分析：寻找与圆柱轴线平行的直线边
  4. 面级联：通过相邻面关系构建多级面分析
  5. 标注面选择：寻找互相平行的标注面
  6. 位置计算：根据视图变换矩阵确定标注位置
  7. 标注生成：创建尺寸标注并更新偏移量

```mermaid
flowchart TD
Start(["DimensionBends 命令"]) --> CheckDoc["检查活动文档"]
CheckDoc --> SelView["验证视图选择"]
SelView --> GetFeat["扫描折弯特征"]
GetFeat --> FindCyl["识别圆柱面"]
FindCyl --> FindEdges["寻找平行边"]
FindEdges --> MultiLevel["多级面分析"]
MultiLevel --> SelectFaces["选择标注面"]
SelectFaces --> CalcPos["计算标注位置"]
CalcPos --> CreateDim["创建尺寸标注"]
CreateDim --> UpdateOffset["更新偏移量"]
UpdateOffset --> NextBend{"还有折弯?"}
NextBend --> |是| GetFeat
NextBend --> |否| Complete["标注完成"]
Complete --> End(["结束"])
```

**图表来源**
- [sw_plugin/function.cs:627-657](file://sw_plugin/function.cs#L627-L657)
- [ctools/solidworks_commands/drw_commands.cs:156-162](file://ctools/solidworks_commands/drw_commands.cs#L156-L162)
- [share/drw/benddim.cs:13-62](file://share/drw/benddim.cs#L13-62)
- [share/drw/benddim.cs:64-266](file://share/drw/benddim.cs#L64-266)

**章节来源**
- [sw_plugin/function.cs:627-657](file://sw_plugin/function.cs#L627-L657)
- [ctools/solidworks_commands/drw_commands.cs:156-162](file://ctools/solidworks_commands/drw_commands.cs#L156-L162)
- [share/drw/benddim.cs:11-331](file://share/drw/benddim.cs#L11-331)

### 插件注册与生命周期管理
- 自动加载机制
  - 通过 SwAddin 特性中的 LoadAtStartup=true，使插件在 SolidWorks 启动时自动加载
- 手动注册流程
  - 使用批处理脚本调用 regasm.exe 完成 COM 注册与反注册
  - 注册时写入 HKLM 与 HKCU 键值，实现开机自动加载与用户级启动项

```mermaid
flowchart TD
Start(["开始注册"]) --> CheckAdmin["检查管理员权限"]
CheckAdmin --> RunRegasm["调用 regasm.exe 注册 DLL"]
RunRegasm --> WriteHKLM["写入 HKLM\\...\\Addins\\{GUID}"]
WriteHKLM --> WriteHKCU["写入 HKCU\\...\\AddInsStartup\\{GUID}"]
WriteHKCU --> Done(["注册完成"])
Unreg(["开始卸载"]) --> DelHKLM["删除 HKLM 注册项"]
DelHKLM --> DelHKCU["删除 HKCU 注册项"]
DelHKCU --> UnregDone(["卸载完成"])
```

**图表来源**
- [sw_plugin/register_addin.bat:7-7](file://sw_plugin/register_addin.bat#L7-L7)
- [sw_plugin/unregister_addin.bat:7-7](file://sw_plugin/unregister_addin.bat#L7-L7)
- [sw_plugin/addin.cs:262-333](file://sw_plugin/addin.cs#L262-L333)

**章节来源**
- [sw_plugin/register_addin.bat:1-10](file://sw_plugin/register_addin.bat#L1-L10)
- [sw_plugin/unregister_addin.bat:1-11](file://sw_plugin/unregister_addin.bat#L1-L11)
- [sw_plugin/addin.cs:262-333](file://sw_plugin/addin.cs#L262-L333)

### SolidWorks API 最佳实践与性能优化
- 文档与模型获取
  - 优先使用 ActiveDoc；若为空，尝试 IActiveDoc2 获取，避免空引用
- 异步执行
  - 对耗时操作使用异步命令，减少 UI 阻塞
- 参数解析与校验
  - 在命令执行器中统一解析参数，提前校验文档状态与必填条件
- **新增**：弯曲标注性能优化
  - 使用几何缓存避免重复计算
  - 批量处理多个折弯特征时使用增量偏移
  - 通过异常处理确保单个特征失败不影响整体执行
- 性能监控
  - 使用 Profiled 特性与计时器包装命令，定位瓶颈
- 资源释放
  - 在命令结束时及时关闭临时打开的文档，避免句柄泄漏

**章节来源**
- [ctools/command_executor.cs:60-94](file://ctools/command_executor.cs#L60-L94)
- [ctools/solidworks_commands/asm_commands.cs:63-78](file://ctools/solidworks_commands/asm_commands.cs#L63-L78)
- [share/drw/benddim.cs:34-56](file://share/drw/benddim.cs#L34-56)

### LLM 集成与自然语言到命令映射
- 命令工具定义
  - 将命令注册中心中的命令动态转换为 Tool 定义，支持别名与参数说明
- 对话循环
  - 用户输入经模糊匹配与 LLM 确认后，转化为命令执行，支持确认/自动两种模式
- 输出捕获
  - 拦截命令执行过程中的 Console 输出，统一反馈给用户

```mermaid
sequenceDiagram
participant User as "用户"
participant LLM as "LlmLoopCaller"
participant Reg as "CommandRegistry"
participant Exec as "CommandExecutor"
participant SW as "SolidWorks"
User->>LLM : 输入自然语言
LLM->>Reg : 获取命令定义
LLM->>LLM : 模糊匹配/完全匹配
LLM->>Exec : ExecuteCommandAsync(命令+参数)
Exec->>SW : 获取 ActiveDoc/IActiveDoc2
Exec->>Exec : 调用 AsyncAction(args)
Exec-->>LLM : 返回执行结果与捕获输出
LLM-->>User : 结果与输出
```

**图表来源**
- [ctools/llm_loop_caller.cs:117-172](file://ctools/llm_loop_caller.cs#L117-L172)
- [ctools/llm_loop_caller.cs:493-726](file://ctools/llm_loop_caller.cs#L493-L726)
- [ctools/command_executor.cs:32-113](file://ctools/command_executor.cs#L32-L113)

**章节来源**
- [ctools/llm_loop_caller.cs:19-800](file://ctools/llm_loop_caller.cs#L19-L800)
- [ctools/command_executor.cs:12-116](file://ctools/command_executor.cs#L12-L116)

## 依赖关系分析
- 组件耦合
  - 插件宿主依赖 SwContext 与命令注册中心；命令执行器依赖注册中心与连接器
  - 命令实现位于独立模块，通过特性与注册中心解耦
  - **新增**：DimensionBends 命令依赖 benddim 类进行几何计算
- 外部依赖
  - SolidWorks Interop 类库、Windows 注册表 API、System.Windows.Forms 控件
- 潜在风险
  - COM 注册失败、注册表权限不足、命令参数解析错误、UI 线程阻塞
  - **新增**：弯曲标注算法复杂度高，可能影响性能

```mermaid
graph LR
ADDIN["AddinStudy"] --> CTX["SwContext"]
ADDIN --> REG["CommandRegistry"]
ADDIN --> POP["右键菜单"]
ADDIN --> CMDATTR["CommandAttribute"]
REG --> EXE["CommandExecutor"]
EXE --> CONN["Connect"]
EXE --> SW["SldWorks"]
CMDP["part_commands.cs"] --> REG
CMDA["asm_commands.cs"] --> REG
DRWCMD["drw_commands.cs"] --> REG
DRWCMD --> BENDDIM["benddim.cs"]
BENDDIM --> SW
LLM["LlmLoopCaller"] --> REG
LLM --> EXE
```

**图表来源**
- [sw_plugin/addin.cs:96-120](file://sw_plugin/addin.cs#L96-L120)
- [sw_plugin/CommandAttribute.cs:18-25](file://sw_plugin/CommandAttribute.cs#L18-L25)
- [ctools/CommandRegistry.cs:12-242](file://ctools/CommandRegistry.cs#L12-L242)
- [ctools/command_executor.cs:12-116](file://ctools/command_executor.cs#L12-L116)
- [ctools/connect.cs:9-51](file://ctools/connect.cs#L9-L51)
- [ctools/solidworks_commands/part_commands.cs:11-149](file://ctools/solidworks_commands/part_commands.cs#L11-L149)
- [ctools/solidworks_commands/asm_commands.cs:11-158](file://ctools/solidworks_commands/asm_commands.cs#L11-L158)
- [ctools/solidworks_commands/drw_commands.cs:156-162](file://ctools/solidworks_commands/drw_commands.cs#L156-L162)
- [ctools/llm_loop_caller.cs:44-67](file://ctools/llm_loop_caller.cs#L44-L67)
- [share/drw/benddim.cs:11-331](file://share/drw/benddim.cs#L11-331)

**章节来源**
- [sw_plugin/addin.cs:96-120](file://sw_plugin/addin.cs#L96-L120)
- [ctools/CommandRegistry.cs:12-242](file://ctools/CommandRegistry.cs#L12-L242)
- [ctools/command_executor.cs:12-116](file://ctools/command_executor.cs#L12-L116)

## 性能考虑
- 异步命令
  - 对耗时操作（如批量导出、打开文档）采用异步命令，避免阻塞 UI
- **新增**：弯曲标注性能优化
  - 使用几何缓存避免重复计算圆柱面和边的平行关系
  - 批量处理多个折弯特征时使用增量偏移，减少重复计算
  - 通过异常处理确保单个特征失败不影响整体执行
- 文档状态检查
  - 在执行前检查文档是否保存、是否选中实体，减少无效调用
- 日志与诊断
  - 使用 Console 输出与 Debug.WriteLine 记录关键路径，便于定位性能瓶颈
- 批处理优化
  - 对装配体批量处理时，尽量减少重复打开/关闭文档的次数

## 故障排查指南
- 插件未加载
  - 检查注册表项是否正确写入，确认管理员权限与 regasm 路径
  - 确认 SwAddin 特性中的 LoadAtStartup 与 Guid 配置
- 命令不可用或报错
  - 检查命令注册中心是否正确注册，确认命令名与别名拼写
  - 在命令执行器中查看参数解析与文档状态判断
  - **新增**：检查 benddim 类的几何计算是否成功，确认折弯特征存在
- UI 无响应
  - 将耗时操作改为异步命令，避免在 UI 线程执行长任务
  - **新增**：弯曲标注可能需要较长时间，建议在后台执行
- 右键菜单不显示
  - 确认 AddMenuPopupItem2 的文档类型与选择类型参数正确
  - 检查 ConnectToSW 中的初始化顺序与异常日志
- **新增**：弯曲标注失败
  - 确认选择了正确的工程图视图
  - 检查折弯特征是否存在于对应的零件文档中
  - 验证几何计算过程中是否有异常抛出

**章节来源**
- [sw_plugin/register_addin.bat:7-7](file://sw_plugin/register_addin.bat#L7-L7)
- [sw_plugin/addin.cs:262-333](file://sw_plugin/addin.cs#L262-L333)
- [ctools/command_executor.cs:32-113](file://ctools/command_executor.cs#L32-L113)
- [sw_plugin/body_context_menu.cs:141-166](file://sw_plugin/body_context_menu.cs#L141-L166)
- [share/drw/benddim.cs:15-61](file://share/drw/benddim.cs#L15-61)

## 结论
本项目通过标准的 COM 插件架构与命令系统，实现了 SolidWorks 的自动化扩展能力。借助 SwContext 上下文管理器、CommandRegistry 命令注册中心与 CommandExecutor 命令执行器，开发者可以快速扩展命令、集成 LLM 对话，并通过右键菜单提升用户体验。

**新增的弯曲标注功能**进一步增强了插件的能力，通过 benddim 类实现了复杂的折弯特征识别与标注逻辑，为钣金件设计提供了强大的自动化工具。该功能展示了如何将复杂的几何分析算法与 SolidWorks API 结合，为用户提供直观的尺寸标注解决方案。

遵循本文的最佳实践与性能优化建议，可在保证稳定性的同时显著提升开发效率与运行性能。

## 附录
- 相关文件清单
  - 插件宿主与注册：[sw_plugin/addin.cs:18-339](file://sw_plugin/addin.cs#L18-L339)、[sw_plugin/register_addin.bat:1-10](file://sw_plugin/register_addin.bat#L1-L10)、[sw_plugin/unregister_addin.bat:1-11](file://sw_plugin/unregister_addin.bat#L1-L11)
  - 命令系统：[ctools/CommandRegistry.cs:12-242](file://ctools/CommandRegistry.cs#L12-L242)、[ctools/CommandAttribute.cs:5-18](file://ctools/CommandAttribute.cs#L5-L18)、[ctools/command_executor.cs:12-116](file://ctools/command_executor.cs#L12-L116)、[ctools/main.cs:54-109](file://ctools/main.cs#L54-L109)、[sw_plugin/CommandAttribute.cs:1-27](file://sw_plugin/CommandAttribute.cs#L1-L27)
  - 命令实现：[ctools/solidworks_commands/part_commands.cs:11-149](file://ctools/solidworks_commands/part_commands.cs#L11-L149)、[ctools/solidworks_commands/asm_commands.cs:11-158](file://ctools/solidworks_commands/asm_commands.cs#L11-L158)、[ctools/solidworks_commands/drw_commands.cs:156-162](file://ctools/solidworks_commands/drw_commands.cs#L156-L162)、[sw_plugin/function.cs:627-657](file://sw_plugin/function.cs#L627-L657)
  - 上下文与连接：[ctools/SwContext.cs:9-85](file://ctools/SwContext.cs#L9-L85)、[ctools/connect.cs:9-51](file://ctools/connect.cs#L9-L51)
  - 右键菜单：[sw_plugin/body_context_menu.cs:141-166](file://sw_plugin/body_context_menu.cs#L141-L166)
  - LLM 集成：[ctools/llm_loop_caller.cs:19-800](file://ctools/llm_loop_caller.cs#L19-L800)
  - **新增**：弯曲标注核心：[share/drw/benddim.cs:11-331](file://share/drw/benddim.cs#L11-331)
  - AutoCAD 插件（对比参考）：[cad_plugin/cad_addin.cs:16-80](file://cad_plugin/cad_addin.cs#L16-L80)、[cad_plugin/register.ps1:1-93](file://cad_plugin/register.ps1#L1-L93)、[cad_plugin/unregister.ps1:1-92](file://cad_plugin/unregister.ps1#L1-L92)