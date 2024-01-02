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
        private const int s_baseChar = 'A' - 1;
        private readonly static Regex s_regexAddress = new Regex(@"(?<column>[a-zA-Z]+)(?<row>\d+)", RegexOptions.Compiled);

        public static CellPos DefaultStartPos => new CellPos(1, 1);

        public static int LetterToInt(string letter)
        {
            int result = 0;
            for (int i = 0; i < letter.Length; i++)
            {
                result *= 26;
                result += letter[i] - s_baseChar;
            }
            return result;
        }

        private int _row;
        private int _column;
        public CellPos(int row, int column)
        {
            if (row <= 0 || column <= 0)throw new Exception(@"单元格的行号和列号都应大于零；");
            if (row > XlsxMaxRow || column > XlsxMaxColumn)throw new Exception(@"行号或列号超出取值范围");
            _row = row;
            _column = column;
        }
        public CellPos(string address)
        {

            Match match = s_regexAddress.Match(address);
            if (match.Success)
            {
                int row = Convert.ToInt32(match.Groups[@"row"].Value);
                int column = LetterToInt(match.Groups[@"column"].Value.ToUpper());
                if (row <= 0 || column <= 0) throw new Exception(@"单元格的行号和列号都应大于零；");
                if (row > XlsxMaxRow || column > XlsxMaxColumn) throw new Exception(@"行号或列号超出取值范围");
                _row = row;
                _column = column;
            }
            else throw new Exception($"“{address}”不是合法的单元格地址；");
            
        }



        public int Row { get=> _row; set=> _row=value; }
        public int Column { get=> _column; set=> _column= value; }


        

        

        public string Address
        {
            get
            {
                string columnName = string.Empty ;
                if (Column <= 26) columnName = ((char)(s_baseChar + Column)).ToString();
                else
                {
                    int dividend = Column;
                    while (dividend > 0)
                    {
                        int modulo = (dividend - 1) % 26;
                        columnName = ((char)(s_baseChar + modulo + 1)).ToString() + columnName;
                        dividend = (dividend - modulo - 1) / 26;
                    }
                }
                return columnName + Row.ToString();
            }
        }


        public override string ToString() => Address;




    }
}
