# WL 标签查找类别功能 - 快速开始指南

## 一、功能简介

这是一个基于**拓扑相似度**的零件类别推荐系统。它会分析你的零件几何特征（面的连接关系），然后在数据库中查找具有相似特征的已标注零件，最后告诉你这些相似零件都被标注为什么类别。

### 核心优势
- ✅ **智能识别**：根据几何拓扑特征自动推荐类别
- ✅ **无需建模知识**：新手也能快速分类零件
- ✅ **标准化辅助**：确保同类零件使用统一分类
- ✅ **知识复用**：找到相似零件，复用设计经验

## 二、快速上手（3 步即可）

### 步骤 1：准备数据库（首次使用）

在使用之前，你需要先标注一些零件作为"训练数据"：

```csharp
// 打开一个零件，然后运行标注命令
TopologyLabeler.LabelCurrentPart(model);
// 按提示输入类别，例如："连接件"、"支撑板"、"安装座"等
```

建议至少标注 10-20 个不同类别的零件。

### 步骤 2：打开要查找的零件

在 SolidWorks 中打开你想要分类的零件。

### 步骤 3：运行查找命令

```csharp
// 一行代码搞定！
TopologyLabeler.FindCategoriesByWL(model, wlIterations: 1, topK: 5);
```

输出示例：
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

...
```

## 三、应用场景

### 场景 1：新零件分类
你设计了一个新零件，不知道应该归类为什么：
```csharp
// 直接运行查找命令
TopologyLabeler.FindCategoriesByWL(swApp.ActiveDoc);
// 系统会告诉你：这个零件跟数据库中的"连接件"最相似
```

### 场景 2：批量检查
你有一批零件需要确认分类是否正确：
```csharp
// 批量处理整个文件夹
WLTagCategoryFinderExample.Example4_BatchFindCategories(swApp, "E:/parts");
```

### 场景 3：详细分析
你想了解具体是哪些零件被归为某类：
```csharp
// 获取详细匹配结果（包含零件名称）
var database = new TopologyDatabase();
var detailedResults = database.FindCategoriesByWLTags(wlFrequencies, topK: 20);

foreach (var (category, partName, similarity, confidence, notes) in detailedResults)
{
    Console.WriteLine($"零件 {partName} => 类别：{category}, 相似度：{similarity * 100:F1}%");
}
```

## 四、参数说明

### FindCategoriesByWL 方法参数

```csharp
TopologyLabeler.FindCategoriesByWL(
    ModelDoc2 swModel,      // 当前打开的零件文档
    int wlIterations = 1,   // WL 迭代次数（通常用 1 就够了）
    int topK = 5           // 显示前 K 个推荐类别
)
```

**参数建议**：
- `wlIterations`：保持默认值 1 即可
  - 增加迭代次数可以捕获更复杂的拓扑特征
  - 但计算时间会增加，通常没必要
  
- `topK`：根据需求调整
  - 快速浏览：设为 3-5
  - 详细分析：设为 10-20

### 高级 API 参数

```csharp
database.FindTopCategoriesByWLTags(
    wlFrequencies,          // WL 特征向量（由系统自动生成）
    topK: 5,               // 返回前 K 个类别
    minSimilarity: 0.3     // 最小相似度阈值
)
```

**相似度阈值调整**：
- `0.7+`：非常严格的匹配，结果很少但很精确
- `0.3-0.7`：常用范围，平衡准确性和覆盖度
- `<0.3`：宽松匹配，可能找到意想不到的关联

## 五、常见问题

### Q1：为什么显示"未找到相似的已标注零件"？
**A**：数据库中没有已标注的零件。你需要先运行一些标注：
```csharp
TopologyLabeler.LabelCurrentPart(model);
// 输入类别，例如："支架"
```

### Q2：相似度多少才算可靠？
**A**：建议参考标准：
- `>0.8`：非常可靠，几乎可以确定
- `0.5-0.8`：比较可靠，可以作为主要参考
- `0.3-0.5`：有一定参考价值
- `<0.3`：仅供参考，可能是不同类型的零件

### Q3：如何提高准确率？
**A**：三个方法：
1. **增加标注样本**：标注更多零件，特别是建立清晰的类别体系
2. **提高相似度阈值**：将 `minSimilarity` 从 0.3 提高到 0.5 或更高
3. **优化类别定义**：确保类别之间有明显区分，避免模糊分类

### Q4：能否自定义类别体系？
**A**：可以！类别名称完全由你决定，例如：
- 按功能：连接件、支撑件、传动件
- 按形状：板类、轴类、箱体类
- 按工艺：钣金件、机加件、铸件

### Q5：计算速度慢怎么办？
**A**：优化建议：
1. 减少 `topK` 值（如从 20 降到 5）
2. 降低 `wlIterations`（保持为 1）
3. 提高 `minSimilarity` 阈值（减少候选数量）
4. 分批处理大型数据库

## 六、最佳实践

### 1. 建立标准类别体系
```
推荐分类结构：
├── 结构件
│   ├── 连接件
│   ├── 支撑件
│   └── 固定件
├── 传动件
│   ├── 齿轮
│   └── 皮带轮
└── 覆盖件
    ├── 面板
    └── 护罩
```

### 2. 标注一致性原则
- 同一类零件使用相同的类别名称
- 避免过于细分（如"连接件_类型 A"、"连接件_类型 B"）
- 为特殊零件添加备注说明

### 3. 定期维护数据库
```csharp
// 查看已有标注
TopologyLabeler.ViewAllParts();

// 查看统计信息
TopologyLabeler.ShowStatistics();
```

### 4. 结合人工判断
系统推荐仅供参考，最终分类应结合：
- 零件的实际功能
- 企业的标准规范
- 工艺加工要求

## 七、技术原理（简化版）

```
1. 提取零件的面和邻接关系 → 构建拓扑图
2. 运行 WL 算法 → 生成拓扑特征向量
3. 与数据库中的零件对比 → 计算余弦相似度
4. 找出最相似的零件 → 统计它们的类别
5. 按出现次数排序 → 返回推荐类别
```

**相似度计算**：使用余弦相似度衡量两个零件的拓扑特征向量夹角
- 1.0：完全相同
- 0.0：完全不同

## 八、完整示例代码

```csharp
using SolidWorks.Interop.sldworks;
using tools;

// 假设这是你的 SolidWorks 应用程序对象
SldWorks swApp = ...;

// 【最简单用法】一行代码搞定
TopologyLabeler.FindCategoriesByWL(swApp.ActiveDoc);


// 【进阶用法】获取更多详细信息
ModelDoc2 model = swApp.ActiveDoc;

// 1. 初始化数据库
var database = new TopologyDatabase("topology_labels.db");

// 2. 构建零件拓扑图
var graph = FaceGraphBuilder.BuildGraph(model);

// 3. 执行 WL 迭代
var wlFrequencies = WLGraphKernel.PerformWLIterations(graph, 1);

// 4. 查找推荐类别
var recommendations = database.FindTopCategoriesByWLTags(
    wlFrequencies, 
    topK: 5, 
    minSimilarity: 0.3
);

// 5. 显示结果
Console.WriteLine("\n推荐类别:");
foreach (var (category, count, avgSim, avgConf) in recommendations)
{
    Console.WriteLine($"{category}: {count}次匹配 " +
        $"[相似度：{avgSim:F3}, 置信度：{avgConf:F2}]");
}
```

## 九、相关文件

- `topology_database.cs` - 数据库核心方法
- `topology_labeler.cs` - 用户交互接口
- `wl_graph_kernel.cs` - WL 图核算法
- `WL 标签查找类别使用说明.md` - 详细技术文档
- `wl_category_finder_examples.cs` - 完整示例代码

## 十、下一步

1. **试用**：打开一个零件，运行 `FindCategoriesByWL`
2. **标注**：建立你的第一批训练数据（10-20 个零件）
3. **验证**：用已知类别的零件测试准确率
4. **优化**：调整参数，完善类别体系
5. **推广**：应用到实际工作中，提高效率

---

**提示**：如有问题或建议，欢迎反馈！祝使用愉快！ 😊
