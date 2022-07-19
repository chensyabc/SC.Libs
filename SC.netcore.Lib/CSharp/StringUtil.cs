using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Util
{
    public static class StringUtil
    {
        public static string MsgCombine(string resultMsg, string detailMsg, bool isDetailMsgTechinicalDetail)
        {
            return string.Format("{0}" + "\n" + (isDetailMsgTechinicalDetail ? "Technical Details" : "") + ": {1}", resultMsg, detailMsg);
        }
    }
}
