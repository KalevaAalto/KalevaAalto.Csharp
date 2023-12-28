using KalevaAalto;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using static KalevaAalto.Main;




namespace KalevaAalto.Models
{
    /// <summary>
    /// 利用SqlSugar整合的个人使用的Mysql类
    /// </summary>
    public class Mysql
    {
        public SqlSugarClient db { private set; get; }

        /// <summary>
        /// Mysql常规构造函数
        /// </summary>
        /// <param name="localhost">主机名</param>
        /// <param name="port">端口</param>
        /// <param name="username">用户名</param>
        /// <param name="password">密码</param>
        /// <param name="databasename">数据库名称</param>
        public Mysql(string localhost, int port, string username, string password, string databasename)
        {
            db = new SqlSugarClient(new ConnectionConfig()
            {
                ConnectionString = $"server={localhost};port={port};user={username};password={password}; database={databasename};",
                DbType = SqlSugar.DbType.MySql,
                IsAutoCloseConnection = true, // 自动释放数据务，如果存在事务，在事务结束后释放  
                InitKeyType = InitKeyType.Attribute, // 从实体特性中获取主键自增列信息  
            });


        }


        /// <summary>
        /// Mysql常规析造函数
        /// </summary>
        ~Mysql()
        {
            db.Close();
            db.Dispose();
        }


        public async Task<System.Data.DataTable> Query(string sql)
        {
            try
            {
                return await db.Ado.GetDataTableAsync(sql);

            }
            catch (Exception error)
            {
                throw new Exception(error.Message + '|' + sql);
            }

        }


        public async Task Run(string sql)
        {
            try
            {
                await db.Ado.ExecuteCommandAsync(sql);
            }
            catch (Exception error)
            {
                throw new Exception(error.Message + '|' + sql);
            }
        }



        public async Task UploadDataTable(System.Data.DataTable dataTable)
        {
            if(dataTable.Columns.Count == 0)
            {
                throw new Exception($"表单“{dataTable.TableName}”没有任何字段；");
            }
            if (dataTable.Rows.Count == 0)
            {
                return;
            }

            StringBuilder columnString = new StringBuilder();
            StringBuilder dataRowString = new StringBuilder();

            await Task.WhenAll(
                Task.Run(() =>
                {
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        columnString.Append('`');
                        columnString.Append(column.ColumnName);
                        columnString.Append('`');
                        columnString.Append(',');
                    }
                    columnString.Remove(columnString.Length - 1, 1);
                }),
                Task.Run(() =>
                {
                    foreach (DataRow dataRow in dataTable.Rows)
                    {
                        dataRowString.Append('(');
                        foreach (object? cell in dataRow.ItemArray)
                        {
                            string cellStr = cell?.ToString() ?? string.Empty;
                            if (cellStr.Length == 0)
                            {
                                dataRowString.Append(@"NULL");
                            }
                            else
                            {
                                dataRowString.Append('\'');
                                dataRowString.Append(cellStr);
                                dataRowString.Append('\'');
                            }
                            dataRowString.Append(',');
                        }
                        dataRowString.Remove(dataRowString.Length - 1, 1);
                        dataRowString.Append(')');
                        dataRowString.Append(',');
                    }
                    dataRowString.Remove(dataRowString.Length - 1, 1);
                })
                );
            await db.Ado.ExecuteCommandAsync($"insert into `{dataTable.TableName}`({columnString}) value{dataRowString};");
        }




        public async Task ClearTable(string tableName, string[]? conditions = null)
        {
            if (conditions is null || conditions.Length == 0)
            {
                await db.Ado.ExecuteCommandAsync($"delete from `{tableName}`;");
            }
            else
            {
                StringBuilder conditionsString = new StringBuilder();
                string partString = @" and ";

                foreach (string condition in conditions)
                {

                    conditionsString.Append('(');
                    conditionsString.Append(condition);
                    conditionsString.Append(')');

                    conditionsString.Append(partString);
                }
                conditionsString.Remove(conditionsString.Length - partString.Length, partString.Length);

                await db.Ado.ExecuteCommandAsync($"delete from `{tableName}` where {conditionsString.ToString()};");
            }




        }



        public async Task ClearTable(System.Data.DataTable table, string[]? conditions = null)
        {
            await ClearTable(table.TableName, conditions);
            await UploadDataTable(table);
        }

        public async Task<string[]> GetColumnNames(string tableName)
        {
            System.Data.DataTable dataTable = await Query($"DESC {tableName};");
            return dataTable.Rows.Cast<DataRow>().Select(it => (string)it[@"Field"]).ToArray();
        }


        public async Task Sync(string from_table_name, string to_table_name, string[]? conditions = null)
        {
            await ClearTable(to_table_name, conditions);


            string[] columns = await GetColumnNames(from_table_name);
            StringBuilder columns_str = new StringBuilder();
            foreach (string column in columns)
            {
                columns_str.Append("`");
                columns_str.Append(column);
                columns_str.Append("`");
                columns_str.Append(",");
            }
            columns_str.Remove(columns_str.Length - 1, 1);

            StringBuilder sql = new StringBuilder($"insert into `{to_table_name}`({columns_str.ToString()}) select {columns_str.ToString()} from `{from_table_name}`");
            if (conditions != null)
            {
                sql.Append(" where ");
                foreach (string condition in conditions)
                {

                    sql.Append("(");
                    sql.Append(condition);
                    sql.Append(")");

                    sql.Append(" and ");
                }
                sql.Remove(sql.Length - 5, 5);

            }
            sql.Append(";");

            await db.Ado.ExecuteCommandAsync(sql.ToString());

        }





        public async Task CleanTableRepetitiveContent<T>() where T : class, new()
        {
            T[] values = (await db.Queryable<T>().ToArrayAsync()).AsParallel().ToHashSet().ToArray();
            await db.Deleteable<T>().ExecuteCommandAsync();
            await db.Insertable(values).ExecuteCommandAsync();
        }






        public async void Reconnect()
        {
            // 如果MySQL连接已关闭，则尝试重新连接
            await Query(@"select 1;");// 向MySQL服务器发送一个简单的心跳查询

        }


        public bool Status()
        {
            try
            {
                // 打开连接以检查是否连接有效
                db.Ado.Open();
                // 在这里可以执行其他数据库操作
            }
            catch (Exception ex)
            {
                Console.WriteLine("Connection error: " + ex.Message);
                // 在此处处理连接错误，可以记录日志或采取其他适当的措施
                db.Ado.Close();
                return false;
            }
            // 如果MySQL连接已关闭，则尝试重新连接
            return true;
        }

        /// <summary>
        /// 获取数据表的列信息
        /// </summary>
        /// <param name="tableName">表名称</param>
        /// <returns>返回数据表的列信息</returns>
        public DbColumnInfo[] GetColumnInfosByTableName(string tableName)
        {
            List<DbColumnInfo> result = db.DbMaintenance.GetColumnInfosByTableName(tableName);
            return result.ToArray();
        }



        private readonly static Dictionary<string, string> mysqlTypeStringToCsharpTypeString = new Dictionary<string, string>
        {
            {@"varchar",@"string" },
            {@"text",@"string" },
            {@"decimal",@"decimal" },
            {@"int",@"int" },
            {@"bigint",@"long" },
            {@"bigint unsigned",@"long" },
            {@"date",@"DateTime" },
            {@"datetime",@"DateTime" },
            {@"time",@"DateTime" },
            {@"tinyint",@"bool" },
        };

        public string GetClassBuilding(string tableName)
        {
            DbColumnInfo[] dbColumnInfos = GetColumnInfosByTableName(tableName);

            StringBuilder result = new StringBuilder();

            result.AppendLine($"/// <summary>");
            result.AppendLine($"/// 数据表：{tableName}");
            result.AppendLine($"/// <summary>");
            result.AppendLine($"[SugarTable(@\"{tableName}\")]");
            result.AppendLine($"public class {tableName}");
            result.AppendLine($"{{");
            foreach (var item in dbColumnInfos)
            {
                result.AppendLine($"\t/// <summary>");
                result.AppendLine($"\t/// 字段名称：{item.DbColumnName}");
                result.AppendLine($"\t/// <summary>");

                if (item.IsPrimarykey)
                {
                    result.AppendLine($"\t[SugarColumn(ColumnName = @\"{item.DbColumnName}\", IsPrimaryKey = true)]");
                }
                else
                {
                    result.AppendLine($"\t[SugarColumn(ColumnName = @\"{item.DbColumnName}\")]");
                }

                if (item.IsNullable)
                {
                    result.AppendLine($"\tpublic {mysqlTypeStringToCsharpTypeString[item.DataType]}? {item.DbColumnName} {{ get; set; }}");
                }
                else
                {
                    result.AppendLine($"\tpublic {mysqlTypeStringToCsharpTypeString[item.DataType]} {item.DbColumnName} {{ get; set; }}");
                }



                result.AppendLine();
            }


            result.AppendLine("\tpublic readonly static string[] columnNames = new string[]");
            result.AppendLine($"\t{{");
            foreach (var item in dbColumnInfos)
            {
                result.AppendLine($"\t\t@\"{item.DbColumnName}\",");
            }

            result.AppendLine($"\t}}");
            result.AppendLine();


            result.AppendLine($"}};");



            return result.ToString();
        }



    }





}
