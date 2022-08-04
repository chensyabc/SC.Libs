using System;
using System.Data;

namespace SC.dotnet.Lib.CSharp
{
    /// <summary>
    /// Get UTC time or local time from database 
    /// </summary>
    public class UTCUtil
    {
        public static DateTime GetUTCTime()
        {
            //DataSet ds = AdoUtil.Query("select GETUTCDATE() as databaseTime");
            DataSet ds = null;
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                return Convert.ToDateTime(ds.Tables[0].Rows[0]["databaseTime"]);
            }
            else
            {
                return DateTime.UtcNow;
            }
        }
    }
}
