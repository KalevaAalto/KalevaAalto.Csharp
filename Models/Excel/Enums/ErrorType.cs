using System;
using System.Collections.Generic;
using System.Diagnostics.Tracing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Models.Excel.Enums
{
    public enum ErrorType : byte
    {
        None,
        //
        // 摘要:
        //     Division by zero
        Div0,
        //
        // 摘要:
        //     Not applicable
        NA,
        //
        // 摘要:
        //     Name error
        Name,
        //
        // 摘要:
        //     Null error
        Null,
        //
        // 摘要:
        //     Num error
        Num,
        //
        // 摘要:
        //     Reference error
        Ref,
        //
        // 摘要:
        //     Value error
        Value,
        //
        // 摘要:
        //     Calc error
        Calc,
        //
        // 摘要:
        //     Spill error from a dynamic array formula.
        Spill
    }
}
