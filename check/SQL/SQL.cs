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
        public static DataTable getMeet(string time,string departID)//根据时间获取会议信息
        {
            try
            {
                string sqlStr = "select * from MeMeetInfo where mDate='"+time+"' and mDeptCode='"+departID+"'";

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
                string sqlStr = "select MePerAttend.*,MeDelegation.* from MePerAttend,MeDelegation where MePerAttend.meetingId=" + id + " and MeDelegation.id=MePerAttend.delegationId order by MePerAttend.attendTime DESC";

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
                string sqlStr = "select MePerAttend.*,MeDelegation.* from MePerAttend,MeDelegation,MeUserInfo where MePerAttend.QRcode='" + code
                    + "' and MeDelegation.id=MePerAttend.delegationId and MePerAttend.meetingId=" + meetid + " order by MePerAttend.timeStamp desc";

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
        public static int   setMeeterInfo(string QRcode,string checkTime,string mID)//根据二维码设置人员信息
        {
            try
            {
                string sqlStr = "update MePerAttend set attendState=1,attendTime='"+checkTime+"',manageId='"+mID+"' where QRcode='"+QRcode+"'";

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
        public static int setloginState(string userId, string loginState)//根据二维码设置人员信息
        {
            try
            {
                string sqlStr = "update MeUserInfo set loginState="+loginState+" where uId="+userId ;

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

        public static int updatePassword(string userId, string pssword,string psd)//根据二维码设置人员信息
        {
            try
            {
                string sqlStr = "select * from MeUserInfo where uId=" + userId + " and uPassword='" + pssword + "'";
                DataSet dt = SqlHelper.ExecuteDataset(SqlHelper.GetConnSting(), CommandType.Text, sqlStr);
                if (dt.Tables[0].Rows.Count > 0)
                {
                    string sqlStr1 = "update MeUserInfo set uPassword='" + psd + "' where uId="+userId;
                    int i = SqlHelper.ExecuteNonQuery(SqlHelper.GetConnSting(), CommandType.Text, sqlStr1);
                    if (i > 0)
                    {
                        return i;
                    }
                    return -1;
                }
                else
                {
                    return -1;
                }
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
                string sqlStr = "insert MeNumber(meetingId,state) values("+meetingId +",0)";
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
                string sqlStr = "update MeNumber set  nTotal='" + totalNum + "',nReal='" + arriveNum + "',nNotArrive='" + noarriveNum + "' where meetingId=" + meetingId;

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
        public static DataRow Login(string userName, string userPwd)
        {

            try
            {
                string sq = "select * from MeUserInfo where uId='" + userName + "' and uPassword='" + userPwd + "'";
                DataSet obj = SqlHelper.ExecuteDataset(SqlHelper.GetConnSting(), CommandType.Text,sq );

                if (obj.Tables[0].Rows.Count > 0)
                {

                    if (obj.Tables[0].Rows[0]["uId"].ToString() == userName && obj.Tables[0].Rows[0]["uPassword"].ToString() == userPwd)
                    {
                        return obj.Tables[0].Rows[0];

                    }
                }
                return null;
            }
            catch (SystemException e)
            {
                ErrorHandle.showError(e);
                return null;
            }

        }













    }
}
