using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace KalevaAalto.Models.Excel
{
    public class CellPos
    {
        public const int XlsxMaxRow = 1048576;
        public const int XlsxMaxColumn = 16384;
        private const int baseChar = 'A' - 1;
        private readonly static Regex regexAddress = new Regex(@"(?<column>[a-zA-Z]+)(?<row>\d+)", RegexOptions.Compiled);
        public int Row { get; set; }
        public int Column { get; set; }


        public CellPos(int row, int column)
        {
            if (row <= 0 || column <= 0)
            {
                throw new Exception(@"单元格的行号和列号都应大于零；");
            }
            if (row > XlsxMaxRow || column > XlsxMaxColumn)
            {
                throw new Exception(@"行号或列号超出取值范围");
            }
            Row = row;
            Column = column;
        }


        public static int LetterToInt(string letter)
        {
            int result = 0;
            for (int i = 0; i < letter.Length; i++)
            {
                result *= 26;
                result += letter[i] - baseChar;
            }

            return result;
        }

        public CellPos(string address)
        {

            Match match = regexAddress.Match(address);
            if (match.Success)
            {
                int row = Convert.ToInt32(match.Groups[@"row"].Value);
                int column = LetterToInt(match.Groups[@"column"].Value.ToUpper());
                if (row <= 0 || column <= 0)
                {
                    throw new Exception(@"单元格的行号和列号都应大于零；");
                }
                if (row > XlsxMaxRow || column > XlsxMaxColumn)
                {
                    throw new Exception(@"行号或列号超出取值范围");
                }
                Row = row;
                Column = column;

            }
            else
            {
                throw new Exception($"“{address}”不是合法的单元格地址；");
            }
        }


        public static CellPos DefaultStartPos { get; } = new CellPos(1, 1);






        public string Address
        {
            get
            {
                string columnName = "";
                // 如果列号小于等于26，直接转换为对应字母
                if (Column <= 26)
                {
                    columnName = ((char)(baseChar + Column)).ToString();
                }
                else
                {
                    // 如果列号大于26，将列号拆分为多个字母
                    int dividend = Column;


                    while (dividend > 0)
                    {
                        int modulo = (dividend - 1) % 26;
                        columnName = ((char)(baseChar + modulo + 1)).ToString() + columnName;
                        dividend = (dividend - modulo - 1) / 26;
                    }
                }


                return columnName + Row.ToString();
            }
        }


        public override string ToString()
        {
            return Address;
        }



    }
}
