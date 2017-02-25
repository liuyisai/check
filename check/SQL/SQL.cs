using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using check;
using System.Data.SqlClient;


namespace check.SQL
{
    class SQL
    {
        public static DataTable getMeet(string time)//根据时间获取会议信息
        {
            try
            {
                string sqlStr = "select * from MeMeetInfo where mDate='"+time+"'";

                DataSet dt = SqlHelper.ExecuteDataset(SqlHelper.GetConnSting(), CommandType.Text, sqlStr);

                if (dt.Tables[0].Rows.Count > 0)
                {
                    return dt.Tables[0];
                }
                return null;
            }
            catch (SyntaxErrorException e)
            {
                ErrorHandle.showError(e);

                return null;
            }

        }

        public static DataTable getMeeter(string  id)//根据会议获取人员信息
        {
            try
            {
                string sqlStr = "select MePerAttend.*,MeDelegation.* from MePerAttend,MeDelegation where MePerAttend.meetingId=" + id + " and MeDelegation.id=MePerAttend.delegationId";

                DataSet dt = SqlHelper.ExecuteDataset(SqlHelper.GetConnSting(), CommandType.Text, sqlStr);

                if (dt.Tables[0].Rows.Count > 0)
                {
                    return dt.Tables[0];
                }
                return null;
            }
            catch (SyntaxErrorException e)
            {
                ErrorHandle.showError(e);

                return null;
            }

        }




















    }
}
