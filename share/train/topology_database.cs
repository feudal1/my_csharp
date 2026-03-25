using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace tools
{
    /// <summary>
    /// 零件拓扑图数据库 - 使用 SQLite 存储 WL 图核结果
    /// </summary>
    public class TopologyDatabase
    {
        private readonly string _connectionString;
        private readonly string _dbPath;

        public TopologyDatabase(string dbPath = "topology_labels.db")
        {
            _dbPath = dbPath;
            _connectionString = $"Data Source={dbPath};Version=3;";
            InitializeDatabase();
        }

        /// <summary>
        /// 初始化数据库表结构
        /// </summary>
        private void InitializeDatabase()
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();

                // 创建零件表
                string createPartsTable = @"
                    CREATE TABLE IF NOT EXISTS parts (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        part_name TEXT UNIQUE NOT NULL,
                        file_path TEXT NOT NULL,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    );";

                // 创建 WL 迭代结果表
                string createWLResultsTable = @"
                    CREATE TABLE IF NOT EXISTS wl_results (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        part_id INTEGER NOT NULL,
                        iteration INTEGER NOT NULL,
                        label_frequencies TEXT NOT NULL,
                        FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE CASCADE,
                        UNIQUE(part_id, iteration)
                    );";

                // 创建用户标注表
                string createLabelsTable = @"
                    CREATE TABLE IF NOT EXISTS user_labels (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        part_id INTEGER NOT NULL,
                        label_category TEXT NOT NULL,
                        label_value TEXT NOT NULL,
                        confidence REAL DEFAULT 1.0,
                        notes TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE CASCADE
                    );";

                // 创建索引
                string createIndexes = @"
                    CREATE INDEX IF NOT EXISTS idx_part_name ON parts(part_name);
                    CREATE INDEX IF NOT EXISTS idx_wl_part ON wl_results(part_id);
                    CREATE INDEX IF NOT EXISTS idx_label_category ON user_labels(label_category);
                    CREATE INDEX IF NOT EXISTS idx_label_value ON user_labels(label_value);
                ";

                using (var command = new SQLiteCommand(connection))
                {
                    command.CommandText = createPartsTable;
                    command.ExecuteNonQuery();

                    command.CommandText = createWLResultsTable;
                    command.ExecuteNonQuery();

                    command.CommandText = createLabelsTable;
                    command.ExecuteNonQuery();

                    command.CommandText = createIndexes;
                    command.ExecuteNonQuery();
                }

                Console.WriteLine($"数据库初始化完成：{_dbPath}");
            }
        }

        /// <summary>
        /// 添加或更新零件及其 WL 结果
        /// </summary>
        public int UpsertPart(string partName, string filePath, List<Dictionary<string, int>> wlFrequencies)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        // 插入或更新零件信息
                        int partId = GetOrCreatePart(connection, partName, filePath);

                        // 删除旧的 WL 结果
                        string deleteOld = "DELETE FROM wl_results WHERE part_id = @part_id";
                        using (var cmd = new SQLiteCommand(deleteOld, connection, transaction))
                        {
                            cmd.Parameters.AddWithValue("@part_id", partId);
                            cmd.ExecuteNonQuery();
                        }

                        // 插入新的 WL 结果
                        string insertWL = @"
                            INSERT INTO wl_results (part_id, iteration, label_frequencies) 
                            VALUES (@part_id, @iteration, @frequencies)";

                        using (var cmd = new SQLiteCommand(insertWL, connection, transaction))
                        {
                            for (int i = 0; i < wlFrequencies.Count; i++)
                            {
                                cmd.Parameters.Clear();
                                cmd.Parameters.AddWithValue("@part_id", partId);
                                cmd.Parameters.AddWithValue("@iteration", i);
                                cmd.Parameters.AddWithValue("@frequencies", JsonConvert.SerializeObject(wlFrequencies[i]));
                                cmd.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                        Console.WriteLine($"✓ 零件 '{partName}' 已存入数据库 ({wlFrequencies.Count} 次迭代)");
                        return partId;
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        Console.WriteLine($"× 存储零件失败：{ex.Message}");
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// 获取或创建零件记录
        /// </summary>
        private int GetOrCreatePart(SQLiteConnection connection, string partName, string filePath)
        {
            // 先尝试查找
            string select = "SELECT id FROM parts WHERE part_name = @part_name";
            using (var cmd = new SQLiteCommand(select, connection))
            {
                cmd.Parameters.AddWithValue("@part_name", partName);
                var result = cmd.ExecuteScalar();
                if (result != null)
                {
                    return Convert.ToInt32(result);
                }
            }

            // 不存在则插入
            string insert = @"
                INSERT INTO parts (part_name, file_path) 
                VALUES (@part_name, @file_path)";
            
            using (var cmd = new SQLiteCommand(insert, connection))
            {
                cmd.Parameters.AddWithValue("@part_name", partName);
                cmd.Parameters.AddWithValue("@file_path", filePath);
                cmd.ExecuteNonQuery();
                
                // 返回新插入的 ID
                cmd.CommandText = "SELECT last_insert_rowid()";
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        /// <summary>
        /// 添加用户标注
        /// </summary>
        public void AddLabel(int partId, string category, string value, double confidence = 1.0, string? notes = null)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                
                string insert = @"
                    INSERT INTO user_labels (part_id, label_category, label_value, confidence, notes)
                    VALUES (@part_id, @category, @value, @confidence, @notes)";

                using (var cmd = new SQLiteCommand(insert, connection))
                {
                    cmd.Parameters.AddWithValue("@part_id", partId);
                    cmd.Parameters.AddWithValue("@category", category);
                    cmd.Parameters.AddWithValue("@value", value);
                    cmd.Parameters.AddWithValue("@confidence", confidence);
                    cmd.Parameters.AddWithValue("@notes", notes ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                }

                Console.WriteLine($"✓ 已标注：{category} = {value} (置信度：{confidence})");
            }
        }

        /// <summary>
        /// 获取零件的所有标注
        /// </summary>
        public Dictionary<string, List<(string Value, double Confidence, string Notes)>> GetPartLabels(int partId)
        {
            var result = new Dictionary<string, List<(string, double, string)>>();

            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                string select = @"
                    SELECT label_category, label_value, confidence, notes 
                    FROM user_labels 
                    WHERE part_id = @part_id
                    ORDER BY created_at DESC";

                using (var cmd = new SQLiteCommand(select, connection))
                {
                    cmd.Parameters.AddWithValue("@part_id", partId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string category = reader.GetString(0);
                            string value = reader.GetString(1);
                            double confidence = reader.GetDouble(2);
                            string notes = reader.IsDBNull(3) ? "" : reader.GetString(3);

                            if (!result.ContainsKey(category))
                                result[category] = new List<(string, double, string)>();

                            result[category].Add((value, confidence, notes));
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 获取所有零件及其最新标注
        /// </summary>
        public List<(int PartId, string PartName, Dictionary<string, string> Labels)> GetAllPartsWithLabels()
        {
            var result = new List<(int, string, Dictionary<string, string>)>();

            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                
                // 获取所有零件
                string selectParts = "SELECT id, part_name FROM parts ORDER BY part_name";
                using (var cmd = new SQLiteCommand(selectParts, connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        int partId = reader.GetInt32(0);
                        string partName = reader.GetString(1);

                        // 获取该零件的最新标注（每个类别取最新的）
                        var labels = GetLatestLabels(connection, partId);
                        
                        result.Add((partId, partName, labels));
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 获取零件的最新标注（每个类别一条）
        /// </summary>
        private Dictionary<string, string> GetLatestLabels(SQLiteConnection connection, int partId)
        {
            var labels = new Dictionary<string, string>();

            string select = @"
                SELECT label_category, label_value 
                FROM (
                    SELECT label_category, label_value, 
                           ROW_NUMBER() OVER (PARTITION BY label_category ORDER BY created_at DESC) as rn
                    FROM user_labels
                    WHERE part_id = @part_id
                )
                WHERE rn = 1";

            using (var cmd = new SQLiteCommand(select, connection))
            {
                cmd.Parameters.AddWithValue("@part_id", partId);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string category = reader.GetString(0);
                        string value = reader.GetString(1);
                        labels[category] = value;
                    }
                }
            }

            return labels;
        }

        /// <summary>
        /// 根据标注类别查询零件
        /// </summary>
        public List<(int PartId, string PartName, string LabelValue)> SearchByLabel(string category, string? value = null)
        {
            var result = new List<(int, string, string)>();

            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                
                string query = @"
                    SELECT p.id, p.part_name, ul.label_value
                    FROM parts p
                    INNER JOIN user_labels ul ON p.id = ul.part_id
                    WHERE ul.label_category = @category";

                if (!string.IsNullOrEmpty(value))
                {
                    query += " AND ul.label_value = @value";
                }

                query += " ORDER BY p.part_name";

                using (var cmd = new SQLiteCommand(query, connection))
                {
                    cmd.Parameters.AddWithValue("@category", category);
                    if (!string.IsNullOrEmpty(value))
                    {
                        cmd.Parameters.AddWithValue("@value", value);
                    }

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add((
                                reader.GetInt32(0),
                                reader.GetString(1),
                                reader.GetString(2)
                            ));
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 导出 WL 结果为 JSON
        /// </summary>
        public string ExportWLResult(int partId)
        {
            var result = new
            {
                PartId = partId,
                Iterations = new List<object>()
            };

            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                string select = "SELECT iteration, label_frequencies FROM wl_results WHERE part_id = @part_id ORDER BY iteration";
                
                using (var cmd = new SQLiteCommand(select, connection))
                {
                    cmd.Parameters.AddWithValue("@part_id", partId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Iterations.Add(new
                            {
                                Iteration = reader.GetInt32(0),
                                Frequencies = JsonConvert.DeserializeObject<Dictionary<string, int>>(reader.GetString(1))
                            });
                        }
                    }
                }
            }

            return JsonConvert.SerializeObject(result, Formatting.Indented);
        }

        /// <summary>
        /// 统计数据库信息
        /// </summary>
        public (int PartCount, int LabelCount, Dictionary<string, int> CategoryStats) GetStatistics()
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                
                // 统计零件数
                int partCount = 0;
                using (var cmd = new SQLiteCommand("SELECT COUNT(*) FROM parts", connection))
                {
                    partCount = Convert.ToInt32(cmd.ExecuteScalar());
                }

                // 统计标注数
                int labelCount = 0;
                using (var cmd = new SQLiteCommand("SELECT COUNT(*) FROM user_labels", connection))
                {
                    labelCount = Convert.ToInt32(cmd.ExecuteScalar());
                }

                // 按类别统计
                var categoryStats = new Dictionary<string, int>();
                using (var cmd = new SQLiteCommand(
                    "SELECT label_category, COUNT(*) FROM user_labels GROUP BY label_category", 
                    connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        categoryStats[reader.GetString(0)] = reader.GetInt32(1);
                    }
                }

                return (partCount, labelCount, categoryStats);
            }
        }
    }
}
