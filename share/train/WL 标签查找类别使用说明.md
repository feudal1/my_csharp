# 使用 WL 标签查找类别方法说明

## 概述

新增了基于 WL（Weisfeiler-Lehman）图核的零件类别查找功能，可以根据零件的拓扑特征自动推荐可能的类别。

## 核心方法

### 1. 数据库层方法（TopologyDatabase）

#### FindCategoriesByWLTags
```csharp
public List<(string Category, string PartName, double Similarity, double Confidence, string Notes)> 
    FindCategoriesByWLTags(
        List<Dictionary<string, int>> wlFrequencies, 
        int topK = 10, 
        double minSimilarity = 0.3)
```

**功能**：根据 WL 标签频率查找具有相似拓扑特征的零件类别

**参数**：
- `wlFrequencies`：待查询零件的 WL 标签频率列表（通过 WLGraphKernel.PerformWLIterations 获得）
- `topK`：返回最相似的 K 个结果（默认 10）
- `minSimilarity`：最小相似度阈值（默认 0.3）

**返回值**：列表包含以下字段：
- `Category`：类别名称
- `PartName`：匹配的零件名称
- `Similarity`：拓扑相似度（0-1）
- `Confidence`：标注置信度
- `Notes`：备注信息

#### FindTopCategoriesByWLTags
```csharp
public List<(string Category, int Count, double AvgSimilarity, double AvgConfidence)> 
    FindTopCategoriesByWLTags(
        List<Dictionary<string, int>> wlFrequencies, 
        int topK = 5, 
        double minSimilarity = 0.3)
```

**功能**：简化版，返回统计意义上的推荐类别

**返回值**：列表包含以下字段：
- `Category`：类别名称
- `Count`：该类别在相似零件中出现的次数
- `AvgSimilarity`：平均相似度
- `AvgConfidence`：平均置信度

### 2. 用户接口方法（TopologyLabeler）

#### FindCategoriesByWL
```csharp
public static void FindCategoriesByWL(ModelDoc2 swModel, int wlIterations = 1, int topK = 5)
```

**功能**：交互式查找当前零件的推荐类别

**参数**：
- `swModel`：当前打开的 SolidWorks 零件文档
- `wlIterations`：WL 迭代次数（默认 1）
- `topK`：显示前 K 个推荐类别（默认 5）

**使用流程**：
1. 构建零件拓扑图
2. 执行 WL 迭代计算特征
3. 在数据库中查找相似零件
4. 显示推荐类别及其统计信息
5. 可按 Enter 查看详细匹配结果

## 使用示例

### 示例 1：在命令行中使用

```csharp
// 假设当前已打开一个 SolidWorks 零件
SldWorks swApp = ...;  // SolidWorks 应用程序对象
ModelDoc2 model = swApp.ActiveDoc;

// 查找推荐类别（使用默认参数）
TopologyLabeler.FindCategoriesByWL(model);

// 或者指定参数：WL 迭代 2 次，显示前 10 个类别
TopologyLabeler.FindCategoriesByWL(model, wlIterations: 2, topK: 10);
```

### 示例 2：直接调用数据库方法

```csharp
// 初始化数据库
var database = new TopologyDatabase("topology_labels.db");

// 构建当前零件的 WL 特征
ModelDoc2 model = swApp.ActiveDoc;
var graph = FaceGraphBuilder.BuildGraph(model);
var wlFrequencies = WLGraphKernel.PerformWLIterations(graph, iterations: 1);

// 查找推荐类别
var recommendations = database.FindTopCategoriesByWLTags(
    wlFrequencies, 
    topK: 5, 
    minSimilarity: 0.5  // 提高相似度阈值
);

// 显示结果
Console.WriteLine("\n推荐类别:");
foreach (var (category, count, avgSim, avgConf) in recommendations)
{
    Console.WriteLine($"类别：{category}");
    Console.WriteLine($"  出现次数：{count}");
    Console.WriteLine($"  平均相似度：{avgSim:F3}");
    Console.WriteLine($"  平均置信度：{avgConf:F2}");
}
```

### 示例 3：获取详细匹配结果

```csharp
// 获取详细的匹配结果（包含具体零件信息）
var detailedResults = database.FindCategoriesByWLTags(
    wlFrequencies, 
    topK: 20, 
    minSimilarity: 0.4
);

foreach (var (category, partName, similarity, confidence, notes) in detailedResults)
{
    Console.WriteLine($"零件：{partName}");
    Console.WriteLine($"  类别：{category}");
    Console.WriteLine($"  相似度：{similarity * 100:F1}%");
    Console.WriteLine($"  置信度：{confidence:F2}");
    if (!string.IsNullOrEmpty(notes))
    {
        Console.WriteLine($"  备注：{notes}");
    }
    Console.WriteLine();
}
```

## 工作流程

```
1. 用户打开零件
   ↓
2. 构建拓扑图（面、邻接关系）
   ↓
3. 执行 WL 迭代 → 生成标签频率向量
   ↓
4. 在数据库中查找已标注零件
   ↓
5. 计算 WL 标签频率的余弦相似度
   ↓
6. 筛选相似度 ≥ 阈值的零件
   ↓
7. 统计这些零件的类别标注
   ↓
8. 按出现次数、相似度排序返回
```

## 相似度计算原理

使用**余弦相似度**计算两个零件的 WL 标签频率向量的相似性：

```
           Σ(count1_i × count2_i)
similarity = ─────────────────────────
             √(Σcount1_i²) × √(Σcount2_i²)
```

其中：
- `count1_i`：零件 1 中标签 i 的出现次数
- `count2_i`：零件 2 中标签 i 的出现次数

相似度范围：0（完全不同）到 1（完全相同）

## 应用场景

1. **零件分类辅助**：帮助工程师快速确定新零件的类别
2. **标准化检查**：检查零件是否符合已有分类体系
3. **知识复用**：找到相似零件，复用其设计和工艺知识
4. **质量提升**：通过类别识别确保标注的一致性

## 注意事项

1. **数据库要求**：需要先有一定数量的已标注零件才能提供有效推荐
2. **相似度阈值**：默认 0.3，可根据实际情况调整
   - 提高阈值 → 更严格匹配，结果更少但更精确
   - 降低阈值 → 更宽松匹配，结果更多但可能不够准确
3. **WL 迭代次数**：增加迭代次数可以捕获更高阶的拓扑特征，但会增加计算时间
4. **性能考虑**：对于大型数据库，建议定期清理和优化

## 输出示例

```
============================================================
TOP-5 推荐类别
============================================================
1. 类别：连接件
   出现次数：8 次
   平均相似度：0.856 (85.6%)
   平均置信度：0.95

2. 类别：支撑板
   出现次数：5 次
   平均相似度：0.723 (72.3%)
   平均置信度：0.88

3. 类别：安装座
   出现次数：3 次
   平均相似度：0.689 (68.9%)
   平均置信度：0.90

============================================================

按 Enter 查看详细匹配结果...
```

## 相关文件

- `topology_database.cs`：数据库实现，包含核心查找方法
- `topology_labeler.cs`：用户接口，提供交互式功能
- `wl_graph_kernel.cs`：WL 图核算法实现
- `similarity_calculator.cs`：相似度计算工具
