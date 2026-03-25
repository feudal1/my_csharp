using tools;
using cad_tools;
using System.Windows.Forms;
using System.Data.SQLite;

namespace test
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
      
            
          //   var swApp = Connect.run();
         //   var swModel = swApp!.IActiveDoc2;
        //  TopologyLabeler.LabelCurrentPart(swModel, wlIterations: 1);
         //   TopologyLabeler.FindCategoriesByWL(swModel, wlIterations: 1, topK: 5);
         draw_divider.process_subfolders_with_divider();
            
            CheckDatabase();
        }
        
        static void CheckDatabase()
        {
            // 检查两个可能的数据库文件
            string[] dbPaths = new string[]
            {
                "e:\\cqh\\code\\my_c#\\test\\topology_labels.db",
                "e:\\cqh\\code\\my_c#\\test\\bin\\Debug\\net9.0-windows\\topology_labels.db"
            };
            
            foreach (string dbPath in dbPaths)
            {
                if (!System.IO.File.Exists(dbPath))
                {
                    Console.WriteLine($"数据库文件不存在：{dbPath}\n");
                    continue;
                }
                
                Console.WriteLine($"=== 检查数据库：{dbPath} ===\n");
            
                using (var connection = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                connection.Open();
                
                // 1. 统计零件数
                using (var cmd = new SQLiteCommand("SELECT COUNT(*) FROM parts", connection))
                {
                    int partCount = Convert.ToInt32(cmd.ExecuteScalar());
                    Console.WriteLine($"parts 表中的零件数：{partCount}");
                }
                
                // 2. 统计标注数
                using (var cmd = new SQLiteCommand("SELECT COUNT(*) FROM user_labels", connection))
                {
                    int labelCount = Convert.ToInt32(cmd.ExecuteScalar());
                    Console.WriteLine($"user_labels 表中的标注数：{labelCount}");
                }
                
                // 3. 显示所有零件
                Console.WriteLine("\n=== 所有零件 ===");
                using (var cmd = new SQLiteCommand("SELECT id, part_name, file_path FROM parts", connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Console.WriteLine($"[{reader.GetInt32(0)}] {reader.GetString(1)}");
                    }
                }
                
                // 4. 显示所有标注
                Console.WriteLine("\n=== 所有标注 ===");
                using (var cmd = new SQLiteCommand(@"
                    SELECT ul.id, ul.part_id, p.part_name, ul.label_category, ul.label_value, ul.confidence, ul.notes 
                    FROM user_labels ul 
                    INNER JOIN parts p ON ul.part_id = p.id 
                    ORDER BY ul.id", connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Console.WriteLine($"标注 ID:{reader.GetInt32(0)}, 零件 ID:{reader.GetInt32(1)}, 零件名:{reader.GetString(2)}, 类别:{reader.GetString(3)}, 值:{reader.GetString(4)}, 置信度:{reader.GetDouble(5)}");
                    }
                }
                
                // 5. 按类别分组统计
                Console.WriteLine("\n=== 按类别统计 ===");
                using (var cmd = new SQLiteCommand(@"
                    SELECT label_category, COUNT(*) as count 
                    FROM user_labels 
                    GROUP BY label_category", connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Console.WriteLine($"{reader.GetString(0)}: {reader.GetInt32(1)} 个标注");
                    }
                }
                
                Console.WriteLine();
                }
            }
        }
        
        static void DeleteExample()
        {
            // 示例：删除标注 ID 为 1 的记录
            // TopologyLabeler.DeleteLabel(1);
            
            // 示例：删除零件 ID 为 2 的所有标注
            // TopologyLabeler.DeletePartLabels(2);
            
            Console.WriteLine("=== 删除标注示例 ===");
            Console.WriteLine("使用以下命令删除标注:");
            Console.WriteLine("  TopologyLabeler.DeleteLabel(标注 ID);     // 删除单条标注");
            Console.WriteLine("  TopologyLabeler.DeletePartLabels(零件 ID); // 删除零件的所有标注");
            Console.WriteLine();
        }
    }
}