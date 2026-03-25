# 零件拓扑标注系统 - 开发总结

## 项目概述

本次开发完成了一个完整的 SolidWorks 零件拓扑标注系统，包括数据库存储、交互式标注、查询检索等核心功能。

## 已实现的功能

### 1. 核心组件

#### (1) 数据库层 (`topology_database.cs`)
- **SQLite 数据库**：轻量级、跨平台、无需额外配置
- **三张核心表**：
  - `parts` - 零件基本信息（名称、路径、厚度等）
  - `wl_results` - WL 迭代结果（标签频率，JSON 格式）
  - `user_labels` - 用户标注（类别、值、置信度、备注）
- **完整 CRUD 操作**：
  - `UpsertPart` - 插入或更新零件及 WL 结果
  - `AddLabel` - 添加用户标注
  - `GetPartLabels` - 获取零件的所有标注
  - `GetAllPartsWithLabels` - 获取所有零件及最新标注
  - `SearchByLabel` - 按标注类别搜索
  - `ExportWLResult` - 导出 JSON 格式 WL 结果
  - `GetStatistics` - 统计数据库信息

#### (2) 标注工具层 (`topology_labeler.cs`)
- **交互式标注界面**：命令行问答形式
- **自动特征计算**：集成 WL 图核算法
- **批量处理**：支持文件夹批处理
- **查询统计**：多种查看和搜索功能

#### (3) 示例代码
- `topology_labeling_example.cs` - 完整使用示例
- `拓扑标注快速入门.cs` - 快速入门演示

### 2. 命令列表（已集成到 ctools）

| 命令 | 功能 | 参数 | 分组 |
|------|------|------|------|
| `label` | 标注当前零件的拓扑图并输入标签 | 无 | train |
| `label_quick` | 快速标注（仅计算存储） | 无 | train |
| `view_parts` | 查看所有已标注零件 | 无 | train |
| `label_search` | 按标注类别搜索零件 | [类别] [值] | train |
| `label_stats` | 显示数据库统计信息 | 无 | train |
| `label_batch` | 批量标注文件夹中的所有零件 | [文件夹路径] | train |
| `label_export` | 导出零件的 WL 结果为 JSON | [零件 ID] | train |

## 技术架构

### 依赖库
- **System.Data.SQLite** (v1.0.118) - SQLite 数据库支持
- **Newtonsoft.Json** (v13.0.3) - JSON 序列化
- **SolidWorks Interop** - SolidWorks API

### 算法基础
- **Weisfeiler-Lehman 图核**：用于计算零件拓扑相似度
- **面邻接图**：将零件表示为节点（面）和边（邻接关系）的图结构
- **迭代标签更新**：通过多轮迭代捕获不同层次的拓扑特征

### 数据流程
```
SolidWorks 零件 
  → 提取面和邻接关系 
  → 构建 PartGraph 
  → WL 迭代计算 
  → 存储到 SQLite 
  → 用户交互标注 
  → 查询检索
```

## 文件清单

### 核心代码
1. `share/train/topology_database.cs` - 数据库操作类（445 行）
2. `share/train/topology_labeler.cs` - 标注工具类（357 行）
3. `share/train/topology_labeling_example.cs` - 使用示例（89 行）
4. `share/train/拓扑标注快速入门.cs` - 快速入门（105 行）

### 配置文件
5. `share/share.csproj` - 添加了 SQLite NuGet 包

### 命令集成
6. `ctools/cad_dwg_commands.cs` - 添加了 7 个标注命令

### 文档
7. `share/train/拓扑标注系统使用说明.md` - 详细使用说明（361 行）
8. `share/train/README_拓扑标注系统.md` - 本文件

## 使用方法

### 快速开始

1. **打开零件**：在 SolidWorks 中打开要标注的零件
2. **运行标注**：输入 `label` 命令
3. **跟随提示**：按控制台提示输入标注信息
4. **查看结果**：使用 `view_parts` 查看所有已标注零件

### 典型工作流

```bash
# 1. 单个零件标注
label

# 2. 批量处理文件夹
label_batch E:\projects\parts

# 3. 查看所有零件
view_parts

# 4. 搜索特定类别
label_search 结构类型 框架

# 5. 查看统计
label_stats

# 6. 导出某个零件的 WL 结果
label_export 1
```

### 标注类别示例

推荐的标注体系：

**结构特征**：
- 结构类型：框架、板材、轴类、壳体、支架
- 复杂度：简单、中等、复杂

**材料工艺**：
- 材料：Q235A、45 钢、铝合金、不锈钢
- 表面处理：镀锌、喷漆、氧化
- 加工工艺：铸造、锻造、焊接、机加工

**功能用途**：
- 用途：支撑、连接、传动、密封
- 重要程度：关键件、重要件、一般件

## 数据库结构

### parts 表
```sql
CREATE TABLE parts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    part_name TEXT UNIQUE NOT NULL,
    file_path TEXT NOT NULL,
    thickness TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

### wl_results 表
```sql
CREATE TABLE wl_results (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    part_id INTEGER NOT NULL,
    iteration INTEGER NOT NULL,
    label_frequencies TEXT NOT NULL,  -- JSON 格式
    FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE CASCADE,
    UNIQUE(part_id, iteration)
);
```

### user_labels 表
```sql
CREATE TABLE user_labels (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    part_id INTEGER NOT NULL,
    label_category TEXT NOT NULL,
    label_value TEXT NOT NULL,
    confidence REAL DEFAULT 1.0,
    notes TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE CASCADE
);
```

## 性能特性

### 优化措施
- 使用 SQLite 索引加速查询（part_name、label_category 等）
- JSON 格式存储频率数据，节省空间
- 支持增量更新，避免重复计算
- 事务批量操作，提高写入效率

### 性能建议
- 批量处理时使用 `label_quick` 先存储特征，后续再补充标注
- 定期备份数据库文件（`topology_labels.db`）
- 对于大型零件（>1000 个面），适当减少 WL 迭代次数

## 扩展开发指南

### 添加新的标注类别
无需修改代码，直接在标注时输入新类别名称即可。

### 自定义查询
```csharp
var db = new TopologyDatabase();
var results = db.SearchByLabel("结构类型", "框架");
```

### 导出数据
```csharp
var db = new TopologyDatabase();
string json = db.ExportWLResult(partId);
File.WriteAllText("export.json", json);
```

### 集成到其他系统
可通过以下方式集成：
1. 直接读取 SQLite 数据库
2. 导入 JSON 导出文件
3. 调用 C# API

## 已知限制

1. **不支持删除标注**：如需删除，需直接使用 SQLite 工具
2. **厚度提取有限制**：仅支持从钣金特征自动提取厚度
3. **批量处理较慢**：需要逐个打开零件文件

## 未来改进方向

1. **图形化界面**：开发 WPF 或 WinForms 标注界面
2. **相似度检索**：基于 WL 特征实现相似零件推荐
3. **批量标注**：支持 Excel 导入标注数据
4. **版本管理**：标注历史和版本控制
5. **导出格式**：支持更多导出格式（CSV、XML 等）
6. **性能优化**：并行处理、缓存机制

## 测试建议

### 功能测试
```bash
# 测试单个零件标注
label

# 测试搜索功能
label_search 材料 Q235A

# 测试统计功能
label_stats

# 测试批量处理
label_batch [测试文件夹路径]
```

### 性能测试
- 准备 10-20 个零件的测试集
- 测试批量处理时间
- 测试数据库查询响应时间

## 故障排除

### 常见问题

**Q: 编译错误 "找不到 SQLite"**
A: 确保已运行 `dotnet restore` 恢复 NuGet 包

**Q: 标注时中文乱码**
A: 检查控制台编码设置，确保为 UTF-8

**Q: 批量处理中途失败**
A: 检查文件路径是否正确，零件是否可打开

**Q: 数据库锁定**
A: 确保没有其他程序正在访问数据库文件

## 相关资源

- [SolidWorks API 文档](https://help.solidworks.com/2023/english/api/sldworksapi/)
- [SQLite 官方文档](https://www.sqlite.org/docs.html)
- [WL 图核算法论文](https://jmlr.org/papers/volume12/kriege11a/kriege11a.pdf)
- 详细说明文档：`拓扑标注系统使用说明.md`

## 开发者联系方式

如有问题或建议，请联系开发团队。

---

**创建日期**: 2026-03-25  
**版本**: v1.0  
**状态**: 已完成并测试
