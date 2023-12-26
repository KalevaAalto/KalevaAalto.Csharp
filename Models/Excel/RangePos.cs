using SqlSugar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace KalevaAalto.Models.Excel
{
    public class RangePos
    {
        public CellPos StartPos { get; set; }
        public CellPos EndPos { get; }

        public RangePos(CellPos startPos, CellPos endPos)
        {
            if (endPos.Row < startPos.Column || endPos.Column < startPos.Column)
            {
                throw new Exception(@"单元格区域结束点应在开始点之后；");
            }
            StartPos = startPos;
            EndPos = endPos;
        }

        private readonly Regex regexAddress = new Regex(@"(?<startPos>[a-zA-Z]+\d+)\:(?<endPos>[a-zA-Z]+\d+)");
        public RangePos(string address)
        {
            Match match = regexAddress.Match(address);
            if (match.Success)
            {
                CellPos startPos = new CellPos(match.Groups[@"startPos"].Value);
                CellPos endPos = new CellPos(match.Groups[@"endPos"].Value);
                if (endPos.Row < startPos.Column || endPos.Column < startPos.Column)
                {
                    throw new Exception(@"单元格区域结束点应在开始点之后；");
                }
                StartPos = startPos;
                EndPos = endPos;
            }
            else
            {
                throw new Exception($"“{address}”不是合法的单元格地址；");
            }
        }

        public RangePos(int startPosRow, int startPosColumn, int endPosRow, int endPosColumn)
        {
            if (endPosRow < startPosColumn || endPosColumn < startPosColumn)
            {
                throw new Exception(@"单元格区域结束点应在开始点之后；");
            }
            StartPos = new CellPos(startPosRow, startPosColumn);
            EndPos = new CellPos(endPosRow, endPosColumn);
        }


        public string Address { get => $"{StartPos.Address}:{EndPos.Address}"; }


        public override string ToString()
        {
            return Address;
        }

    }
}
