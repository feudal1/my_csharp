# SolidWorks 插件连接指南

## 📌 前提条件

1. 已安装 SolidWorks
2. 已安装 .NET SDK（本项目使用 .NET 10.0）
3. 项目已编译成功

## 🚀 快速开始

### 方法一：使用注册脚本（推荐）

**步骤：**

1. **右键点击** `register_addin.bat` 文件
2. **选择"以管理员身份运行"**
3. 等待注册成功的提示
4. 按照提示在 SolidWorks 中启用插件

### 方法二：手动注册

如果脚本无法运行，可以手动执行以下命令：

```bash
# 以管理员身份打开命令行，进入项目目录
cd e:\cqh\code\csharp3

# 执行注册命令
%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe "csharp3\bin\Release\net10.0\csharp3.dll" /codebase /tlb
```

## 🔧 在 SolidWorks 中启用插件

1. **打开 SolidWorks**
2. 点击顶部菜单栏的 **"工具"** (Tools)
3. 选择 **"插件"** (Add-Ins)
4. 在弹出的对话框中找到 **"SolidWorksAddinStudy"**
5. **勾选** 该插件名称旁边的复选框
   - 勾选当前会话：仅本次有效
   - 勾选"启动时加载"：每次启动 SolidWorks 自动加载
6. 点击 **"确定"**

## ✅ 验证插件是否工作

如果插件成功加载，当你打开 SolidWorks 时会自动显示消息："fuck you"

## 🗑️ 卸载插件

### 方法一：使用卸载脚本

1. **右键点击** `unregister_addin.bat`
2. **选择"以管理员身份运行"**

### 方法二：手动卸载

```bash
# 以管理员身份打开命令行
%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe "csharp3\bin\Release\net10.0\csharp3.dll" /unregister
```

### 方法三：在 SolidWorks 中移除

1. 打开 SolidWorks
2. 点击 **"工具"** > **"插件"**
3. 取消勾选 **"SolidWorksAddinStudy"**
4. 点击 **"确定"**

## ⚠️ 常见问题

### 问题 1：注册失败
**解决方案：** 
- 确保以**管理员身份**运行注册脚本或命令
- 检查 DLL 文件是否存在于 `csharp3\bin\Release\net10.0\` 目录

### 问题 2：在 SolidWorks 中找不到插件
**解决方案：**
- 重新运行注册脚本
- 重启 SolidWorks
- 检查插件是否正确实现了 `ISwAddin` 接口

### 问题 3：插件加载但没有反应
**解决方案：**
- 检查代码中的 `ConnectToSW` 方法是否正确实现
- 查看 SolidWorks 是否有错误提示

## 📝 项目结构

```
csharp3/
├── csharp3/
│   ├── Class1.cs              # 插件主代码
│   └── csharp3.csproj         # 项目配置文件
├── register_addin.bat         # 注册脚本
├── unregister_addin.bat       # 卸载脚本
└── README.md                  # 本说明文档
```

## 🔍 技术细节

### COM 注册说明

- **Guid**: `E3112324-138D-4636-B3A5-B0AAC1438C4E`
- **类名**: `tools.AddinStudy`
- **接口**: `ISwAddin`
- **特性**: `[ComVisible(true)]`

### 必需的方法

1. `ConnectToSW(object ThisSW, int Cookie)` - 插件加载时调用
2. `DisconnectFromSW()` - 插件卸载时调用

## 📞 需要帮助？

如果遇到问题，请检查：
1. SolidWorks 版本是否兼容
2. .NET Framework 是否正确安装
3. 是否有足够的权限运行脚本
