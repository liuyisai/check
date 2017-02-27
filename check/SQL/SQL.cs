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
                string sqlStr = "select MePerAttend.*,MeDelegation.*,MeUserInfo.* from MePerAttend,MeDelegation,MeUserInfo where MePerAttend.meetingId=" + id + " and MeDelegation.id=MePerAttend.delegationId and MeUserInfo.id=MePerAttend.uId order by MeDelegation.id";

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

        public static DataTable getMeeterInfo(string code,string meetid)//根据二维码获取人员信息
        {
            try
            {
                string sqlStr = "select MePerAttend.*,MeDelegation.*,MeUserInfo.* from MePerAttend,MeDelegation,MeUserInfo where MePerAttend.QRcode='" + code 
                    + "' and MeDelegation.id=MePerAttend.delegationId and MeUserInfo.id=MePerAttend.uId and MePerAttend.meetingId="+meetid;

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
        public static int   setMeeterInfo(string QRcode,string checkTime)//根据二维码设置人员信息
        {
            try
            {
                string sqlStr = "update MePerAttend set attendState=1,attendTime='"+checkTime+"' where QRcode='"+QRcode+"'";

                int  i = SqlHelper.ExecuteNonQuery(SqlHelper.GetConnSting(), CommandType.Text, sqlStr);

                if (i > 0)
                {
                    return i;
                }
                return -1;
            }
            catch (SyntaxErrorException e)
            {
                ErrorHandle.showError(e);

                return -1;
            }

        }

        public static DataTable getIsMeeting(string meetingId)//
        {
            try
            {
                string sqlStr = "select * from MeNumber where meetingId=" + meetingId;

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
        public static int insertNumber(string meetingId)//根据二维码设置人员信息
        {
            //string totalNum, string arriveNum,string noarriveNum,
            try
            {
                string sqlStr = "insert MeNumber(meetingId) values("+meetingId +")";
                    //"insert MeNumber(nTotal,nReal,nNotArrive,meetingId) values('" + totalNum + "','" + arriveNum + "','" + noarriveNum + "'," + meetingId + ")";

                int i = SqlHelper.ExecuteNonQuery(SqlHelper.GetConnSting(), CommandType.Text, sqlStr);

                if (i > 0)
                {
                    return i;
                }
                return -1;
            }
            catch (SyntaxErrorException e)
            {
                ErrorHandle.showError(e);

                return -1;
            }

        }

        public static int updateNumber(string totalNum, string arriveNum, string noarriveNum, string meetingId)//根据二维码设置人员信息
        {
            try
            {
                string sqlStr = "update MeNumber set nTotal='" + totalNum + "',nReal='" + arriveNum + "',nNotArrive='" + noarriveNum + "' where meetingId=" + meetingId;

                int i = SqlHelper.ExecuteNonQuery(SqlHelper.GetConnSting(), CommandType.Text, sqlStr);

                if (i > 0)
                {
                    return i;
                }
                return -1;
            }
            catch (SyntaxErrorException e)
            {
                ErrorHandle.showError(e);

                return -1;
            }

        }

        public static DataTable getCheck(string QRcode)//根据会议获取人员信息
        {
            try
            {
                string sqlStr = "select * from MePerAttend where QRcode='"+QRcode+"'";

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
