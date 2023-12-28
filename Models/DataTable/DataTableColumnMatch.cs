using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Models.DataTable
{
    public class DataTableColumnMatch
    {
        public readonly string[] SourceColumnNames;
        public readonly string ColumnName;
        public readonly Type Type = typeof(object);

        public DataTableColumnMatch(string[] sourceColumnNames,string columnName)
        {
            
            this.SourceColumnNames = sourceColumnNames;
            this.ColumnName = columnName;
            if (sourceColumnNames.Length <= 0)
            {
                this.SourceColumnNames = new string[] { columnName };
            }
        }

        public DataTableColumnMatch(string[] sourceColumnNames, string columnName,Type type)
        {
            this.SourceColumnNames = sourceColumnNames;
            this.ColumnName = columnName;
            this.Type = type;
            if (sourceColumnNames.Length <= 0)
            {
                this.SourceColumnNames = new string[] { columnName };
            }
        }
    }
}


namespace KalevaAalto.Static
{
    public static partial class Main
    {
        public static System.Data.DataTable GetEmptyDataTable(this Models.DataTable.DataTableColumnMatch[] dataTableColumnMatches,string tableName = emptyString)
        {
            System.Data.DataTable result = new System.Data.DataTable(tableName);
            foreach (Models.DataTable.DataTableColumnMatch dataTableColumnMatch in dataTableColumnMatches)
            {
                result.Columns.Add(dataTableColumnMatch.ColumnName, dataTableColumnMatch.Type);
            }
            return result;
        }


        public static System.Data.DataTable TranslateDataTable(this System.Data.DataTable dataTable, Models.DataTable.DataTableColumnMatch[] dataTableColumnMatches)
        {
            #region 运行前检查
            if (dataTableColumnMatches.Length <= 0)
            {
                throw new Exception($"表单“{dataTable.TableName}”可匹配的字段数应大于零；");
            }
            Dictionary<string,string> columnMatchs = new Dictionary<string,string>();
            HashSet<string> sourceColumnNamesCollection = dataTable.Columns.Cast<DataColumn>().Select(it => it.ColumnName).ToHashSet();
            HashSet<string> unMatchs = new HashSet<string>();
            Parallel.ForEach(dataTableColumnMatches, dataTableColumnMatch =>
            {
                foreach(string matchColumnName in dataTableColumnMatch.SourceColumnNames)
                {
                    if (sourceColumnNamesCollection.Contains(matchColumnName))
                    {
                        columnMatchs.Add(dataTableColumnMatch.ColumnName,matchColumnName);
                        sourceColumnNamesCollection.Remove(matchColumnName);
                        break;
                    }
                }
                if (!columnMatchs.ContainsKey(dataTableColumnMatch.ColumnName)) { unMatchs.Add(dataTableColumnMatch.ColumnName); }
            });
            if (unMatchs.Count > 0)
            {
                StringBuilder errorMessage = new StringBuilder();
                foreach(string columnName in unMatchs)
                {
                    errorMessage.Append($"“{columnName}”");
                }
                throw new Exception($"找不到字段{errorMessage}；");
            }
            #endregion


            //添加字段
            System.Data.DataTable result = dataTableColumnMatches.GetEmptyDataTable(dataTable.TableName);


            //添加数据
            foreach(DataRow dataRowSource in dataTable.Rows)
            {
                DataRow dataRow = result.Rows.Add();
                foreach(KeyValuePair<string,string> item in columnMatchs)
                {
                    object? valueSource = dataRowSource[item.Value];
                    object? value = result.Columns[item.Key]!.DataType.GetValue(valueSource);
                    if(value == null)
                    {
                        dataRow[item.Key] = DBNull.Value;
                    }
                    else
                    {
                        dataRow[item.Key] = value;
                    }
                }
            }



            return result;
        }
    }
}
