using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.ViewModels.NewModels;

namespace WCals.Controllers.NSPMIS {
    public class TDE0110Controller : Controller {
        int pageSize = 10;
        // GET: TDE0110
        public ActionResult Index(string SYR, int page = 1) {
            TDE0110Model model = new TDE0110Model();

            string userid = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            string userdp = Helpers.WebSiteHelper.getUserDept();

            // 1. 準備年度下拉選單
            string yearSql = "SELECT DISTINCT SYR FROM CMETMS.dbo.CCDMCNT ORDER BY SYR DESC";
            DataTable dtYear = BASE_DAO.SqlDao.CMBASEQuery4List(yearSql);
            ViewBag.SYR = SYR; // 保存當前選擇
            ViewBag.SYR_List = tools.ToSelectList(dtYear, "SYR", "SYR", SYR);

            string sql = "select CITYLIST from CMBASE.dbo.OGCDEPT where DEPT_NO='"+userdp+"'";
            SqlCommand cmd = new SqlCommand(sql);
            
            //cmd.Parameters.Add(new SqlParameter("SYR", filter("SYR")));
            //cmd.Parameters.Add(new SqlParameter("END_SYR", "END_SYR"));
            //cmd.Parameters.Add(new SqlParameter("PID", "1"));
            DataTable ps = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);


            if(ps.Rows.Count == 0) {
                return View("TDE0110Index", model); // 修正 View 名稱
            }


            StringBuilder sb=new StringBuilder("select case isnull(IS_LOCK,'N') when 'N' then 'MOD' else 'VIEW' end as MODE,SYR,SMN,CITY,PC_CNT,AC_CNT,BC_CNT,CC_CNT,P_PC_CNT,P_AC_CNT,P_BC_CNT,P_CC_CNT,B_PC_CNT,B_AC_CNT,B_BC_CNT,B_CC_CNT,CHANGE_NOTE,UNUSUAL_NOTE,isnull(IS_LOCK,'N') as IS_LOCK from CMETMS.dbo.CCDMCNT where CITY In (");
            for(int i=0;i<ps.Rows.Count;i++)
            {
                
                    
                sb.Append("'" + ps.Rows[i]["CITYLIST"]+"',"); 
            }
            sb.Length--;

            sb.Append(") ");

            // 2. 加入年度篩選
            if (!string.IsNullOrEmpty(SYR)) {
                sb.Append(" AND SYR = '" + SYR.Replace("'", "''") + "' "); // 簡單防注入
            }

            sb.Append(" order by SYR DESC, SMN DESC"); // 修正排序

            ps = BASE_DAO.SqlDao.CMBASEQuery4List(sb.ToString());
            List<DataRow> list=ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist= new PagedList<DataRow>(list, page, pageSize); 
            
            
            model.List1 = plist; // 指派資料給 Model

            return View("TDE0110Index", model); // 修正 View 名稱
        }

        // GET: TDE0110/Create
        public ActionResult Create() {
            // 準備下拉選單
            ViewBag.CityList = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);
            return View("Create");
        }

        // POST: TDE0110/Create
        [HttpPost]
        public ActionResult Create(string SYR, string SMN, string CITY, decimal? PC_CNT, decimal? AC_CNT, decimal? BC_CNT, decimal? CC_CNT, string CHANGE_NOTE, string UNUSUAL_NOTE) {
            try {
                if (string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN) || string.IsNullOrEmpty(CITY)) {
                    TempData["Error"] = "年度、月份、縣市為必填欄位";
                    return Create(); // 返回建立頁面
                }

                // 檢查是否已存在
                string checkSql = "SELECT COUNT(*) FROM CMETMS.dbo.CCDMCNT WHERE SYR=@SYR AND SMN=@SMN AND CITY=@CITY";
                SqlCommand cmdCheck = new SqlCommand(checkSql);
                cmdCheck.Parameters.Add(new SqlParameter("SYR", SYR));
                cmdCheck.Parameters.Add(new SqlParameter("SMN", SMN));
                cmdCheck.Parameters.Add(new SqlParameter("CITY", CITY));
                DataTable dtCheck = BASE_DAO.SqlDao.CMETMSQuery4List(cmdCheck);
                
                if (dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0][0]) > 0) {
                    TempData["Error"] = "該年月縣市資料已存在";
                    return Create();
                }

                // 插入資料 (注意：這裡假設只需要插入基本資料，P_ (累計) 欄位可能需要另外處理或讓使用者在編輯時輸入，此處簡化為基本輸入)
                // 這裡也需要插入 IS_LOCK 預設值，參考 TDE0070 可能為 'N'
                // 如果需要插入 P_ columns，需增加參數。
                // 根據 TDE0110Index 顯示，有 P_PC_CNT 等欄位。為簡化 Create，先只填基本，詳細可 Edit。
                // 實際 SQL 應包含所有欄位
                string insertSql = @"INSERT INTO CMETMS.dbo.CCDMCNT 
                    (SYR, SMN, CITY, PC_CNT, AC_CNT, BC_CNT, CC_CNT, CHANGE_NOTE, UNUSUAL_NOTE, IS_LOCK, ADD_DT, UPD_DT, UPD_UID)
                    VALUES (@SYR, @SMN, @CITY, @PC_CNT, @AC_CNT, @BC_CNT, @CC_CNT, @CHANGE_NOTE, @UNUSUAL_NOTE, 'N', GETDATE(), GETDATE(), @UPD_UID)";

                SqlCommand cmd = new SqlCommand(insertSql);
                cmd.Parameters.Add(new SqlParameter("SYR", SYR));
                cmd.Parameters.Add(new SqlParameter("SMN", SMN));
                cmd.Parameters.Add(new SqlParameter("CITY", CITY));
                cmd.Parameters.Add(new SqlParameter("PC_CNT", (object)PC_CNT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("AC_CNT", (object)AC_CNT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("BC_CNT", (object)BC_CNT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("CC_CNT", (object)CC_CNT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("CHANGE_NOTE", (object)CHANGE_NOTE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("UNUSUAL_NOTE", (object)UNUSUAL_NOTE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", Helpers.WebSiteHelper.getUserId()));

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                TempData["Success"] = "新增成功";
                return RedirectToAction("Index");
            } catch (Exception ex) {
                TempData["Error"] = "新增失敗：" + ex.Message;
                return Create();
            }
        }

        // GET: TDE0110/Edit
        public ActionResult Edit(string SYR, string SMN, string CITY) {
            if (string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN) || string.IsNullOrEmpty(CITY)) {
                return RedirectToAction("Index");
            }

            string sql = @"SELECT * FROM CMETMS.dbo.CCDMCNT WHERE SYR=@SYR AND SMN=@SMN AND CITY=@CITY";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("SYR", SYR));
            cmd.Parameters.Add(new SqlParameter("SMN", SMN));
            cmd.Parameters.Add(new SqlParameter("CITY", CITY));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            if (dt.Rows.Count == 0) {
                TempData["Error"] = "找不到資料";
                return RedirectToAction("Index");
            }

            // 檢查是否鎖定
            // 如果是 View 模式 (按鈕連結進來的檢視)，不需要檢查鎖定，或者如果鎖定只是禁止編輯，那這裡應該要能顯示資料
            // 但現在是因為 IS_LOCK = 'Y' 會跳轉回 Index，所以 "檢視" 按鈕也會失效 (因為共用 Edit Action)
            // 修改邏輯：如果是檢視 (View)，就不應該被 IS_LOCK 擋住。
            // 但 Edit Action 無法區分 View 或 Edit，除非 View 頁面把欄位設為唯讀。
            // 解決方案：允許進入 Edit 頁面，但在 View 裡面根據 IS_LOCK 決定是否將表單設為唯讀。
            // 現在先註解掉這裡的跳轉，讓 "檢視" 可以進去。真正的編輯權限控制應該在 POST Action 或在 View 中隱藏儲存按鈕。
            
            // if (dt.Rows[0]["IS_LOCK"].ToString() == "Y") {
            //    TempData["Error"] = "資料已鎖定，無法編輯";
            //    return RedirectToAction("Index");
            // }

            var row = dt.Rows[0];
            // 傳遞資料到 View，可以使用 Model 或 ViewBag
            // 這裡簡單使用 ViewBag 傳遞 DataRow 內容
            ViewBag.Data = row;
            return View("Edit");
        }

        // POST: TDE0110/Edit
        [HttpPost]
        public ActionResult Edit(string SYR, string SMN, string CITY, decimal? PC_CNT, decimal? AC_CNT, decimal? BC_CNT, decimal? CC_CNT, 
                                decimal? P_PC_CNT, decimal? P_AC_CNT, decimal? P_BC_CNT, decimal? P_CC_CNT,
                                string CHANGE_NOTE, string UNUSUAL_NOTE) {
            try {
                // 檢查是否鎖定
                string checkSql = "SELECT IS_LOCK FROM CMETMS.dbo.CCDMCNT WHERE SYR=@SYR AND SMN=@SMN AND CITY=@CITY";
                SqlCommand checkCmd = new SqlCommand(checkSql);
                checkCmd.Parameters.Add(new SqlParameter("SYR", SYR));
                checkCmd.Parameters.Add(new SqlParameter("SMN", SMN));
                checkCmd.Parameters.Add(new SqlParameter("CITY", CITY));
                DataTable dtCheck = BASE_DAO.SqlDao.CMETMSQuery4List(checkCmd);
                
                if (dtCheck.Rows.Count > 0 && dtCheck.Rows[0]["IS_LOCK"].ToString() == "Y") {
                    TempData["Error"] = "資料已鎖定，無法編輯";
                    return RedirectToAction("Index");
                }

                // 更新資料 (包含本月與累計)
                string updateSql = @"UPDATE CMETMS.dbo.CCDMCNT SET 
                    PC_CNT=@PC_CNT, AC_CNT=@AC_CNT, BC_CNT=@BC_CNT, CC_CNT=@CC_CNT,
                    P_PC_CNT=@P_PC_CNT, P_AC_CNT=@P_AC_CNT, P_BC_CNT=@P_BC_CNT, P_CC_CNT=@P_CC_CNT,
                    CHANGE_NOTE=@CHANGE_NOTE, UNUSUAL_NOTE=@UNUSUAL_NOTE,
                    UPD_DT=GETDATE(), UPD_UID=@UPD_UID
                    WHERE SYR=@SYR AND SMN=@SMN AND CITY=@CITY";

                SqlCommand cmd = new SqlCommand(updateSql);
                cmd.Parameters.Add(new SqlParameter("SYR", SYR));
                cmd.Parameters.Add(new SqlParameter("SMN", SMN));
                cmd.Parameters.Add(new SqlParameter("CITY", CITY));
                
                cmd.Parameters.Add(new SqlParameter("PC_CNT", (object)PC_CNT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("AC_CNT", (object)AC_CNT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("BC_CNT", (object)BC_CNT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("CC_CNT", (object)CC_CNT ?? DBNull.Value));
                
                cmd.Parameters.Add(new SqlParameter("P_PC_CNT", (object)P_PC_CNT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("P_AC_CNT", (object)P_AC_CNT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("P_BC_CNT", (object)P_BC_CNT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("P_CC_CNT", (object)P_CC_CNT ?? DBNull.Value));

                cmd.Parameters.Add(new SqlParameter("CHANGE_NOTE", (object)CHANGE_NOTE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("UNUSUAL_NOTE", (object)UNUSUAL_NOTE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", Helpers.WebSiteHelper.getUserId()));

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                TempData["Success"] = "更新成功";
                return RedirectToAction("Index", new { SYR = SYR }); // 返回列表並維持年度篩選
            } catch (Exception ex) {
                TempData["Error"] = "更新失敗：" + ex.Message;
                return RedirectToAction("Index");
            }
        }
    }
}