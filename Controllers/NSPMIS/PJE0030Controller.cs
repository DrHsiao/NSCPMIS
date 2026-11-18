using Aspose.Cells;
using Newtonsoft.Json.Linq;
using PagedList;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.Helpers;
using WCals.Models.CMEPMSModels;
using WCals.ViewModels.NewModels;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.NSPMIS
{
    public class PJE0030Controller : Controller
    {
        private readonly Models.CMEPMSModels.CMEPMSEntities db_pms = new Models.CMEPMSModels.CMEPMSEntities();
        private readonly Models.CMETMSModels.CMETMSEntities db_tms = new Models.CMETMSModels.CMETMSEntities();
        private readonly Models.CMPSGSModels.CMPSGSEntities db_sgs = new Models.CMPSGSModels.CMPSGSEntities();

        int pageSize = 10;

        private IEnumerable GetAllPJDFDCR()
        {
            string userid = Helpers.WebSiteHelper.getUserId();
            var result = from pf in db_pms.PJDFDCR
                         join pj in db_pms.PJMFORM on pf.PID equals pj.PID
                         where pj.ADV_UID == userid
                         select pj.PID;
            return result.ToList();
        }

        private void AppendIndexSearchFilter(StringBuilder sql, SqlCommand cmd, PJE0030Model model)
        {
            if (string.IsNullOrWhiteSpace(model.keyword))
            {
                return;
            }

            string stype = string.IsNullOrEmpty(model.stype) ? "PJ_NAME" : model.stype;
            string columnName;

            switch (stype)
            {
                case "SYR":
                    columnName = "W.SYR";
                    break;
                case "PJ_CODE":
                    columnName = "B.PJ_CODE";
                    break;
                default:
                    model.stype = "PJ_NAME";
                    columnName = "B.PJ_NAME";
                    break;
            }

            sql.Append(" AND " + columnName + " LIKE @keyword");
            cmd.Parameters.Add(new SqlParameter("keyword", "%" + model.keyword.Trim() + "%"));
        }

        private DataTable GetIndexList(PJE0030Model model)
        {
            string userid = Helpers.WebSiteHelper.getUserId();
            StringBuilder sql = new StringBuilder("select B.PJ_CODE, B.PID, W.SYR, W.PJ_NAME as WPJ_NAME, B.PJ_NAME as BPJ_NAME, B.PJ_FUND," +
                "case (select count(1) from PJDFDCR where PID = B.PID)" +
                "when 0 then '原預算'else '第' + convert(nvarchar(3), " +
                "(select max(SEQ_NO) from PJDFDCR where PID = B.PID)) +'期修正'" +
                "end as SEQ_NAME from PJMFORM B left join PJMFORM W on B.AN_PID = W.PID and W.PJ_TYPE = 'W' " +
                "where B.PJ_TYPE = 'B' and B.ADV_UID=@userid");
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add(new SqlParameter("userid", userid));
            AppendIndexSearchFilter(sql, cmd, model);
            sql.Append(" order by SYR desc");
            cmd.CommandText = sql.ToString();

            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

            sql = new StringBuilder("select B.PJ_CODE, B.PID, W.SYR, W.PJ_NAME as WPJ_NAME, B.PJ_NAME as BPJ_NAME, B.PJ_FUND," +
                "case (select count(1) from PJDFDCR where PID = B.PID)" +
                "when 0 then '原預算'else '第' + convert(nvarchar(3), " +
                "(select max(SEQ_NO) from PJDFDCR where PID = B.PID)) +'期修正'" +
                "end as SEQ_NAME from (PJMFORM B left join PJMFORM W on B.AN_PID = W.PID and W.PJ_TYPE = 'W')left join PJDAUTH a ON B.PID=A.PID " +
                "where B.PJ_TYPE = 'B' and A.EXE_UID=@userid");
            cmd = new SqlCommand();
            cmd.Parameters.Add(new SqlParameter("userid", userid));
            AppendIndexSearchFilter(sql, cmd, model);
            sql.Append(" order by SYR desc");
            cmd.CommandText = sql.ToString();
            ps.Merge(BASE_DAO.SqlDao.CMEPMSQuery4List(cmd));

            DataView view = new DataView(ps);
            view.Sort = "SYR DESC";
            return view.ToTable();
        }

        [HttpPost]
        public ActionResult EzSv(int pid, int tid, string id, string val)
        {
            PJMFORM_CMEPMS p = db_pms.PJMFORM.Where(w => w.PID == pid).FirstOrDefault();
            string sql;
            DateTime now = DateTime.Now;
            CTDITFD c;
            if (id.Equals("M_00"))
            {
                //中央款比例
                //檢查存在
                c = db_pms.CTDITFD.Where(w => w.TID == tid && w.SYR == p.SYR && w.FIT_CODE == "00").FirstOrDefault();
                if (c == null)
                {
                    c = new CTDITFD();
                    c.FIT_CODE = "00";
                    c.SYR = p.SYR;
                    c.TID = tid;
                    c.FUND = tools.safe2Decimal(val);
                    c.UPD_UID = WebSiteHelper.getUserId();
                    c.UPD_DT = now;
                    c.ADD_DT = now;
                    db_pms.CTDITFD.Add(c);
                    db_pms.SaveChanges();
                    return Json("已新增", JsonRequestBehavior.AllowGet);
                }
                else
                {
                    c.FUND = tools.safe2Decimal(val);
                    c.UPD_UID = WebSiteHelper.getUserId();
                    c.UPD_DT = now;
                    c.ADD_DT = now;
                    db_pms.SaveChanges();
                    return Json("已儲存", JsonRequestBehavior.AllowGet);
                }
            }

            if (id.Equals("M_"))
            {
                //地方款比例
                //檢查存在
                c = db_pms.CTDITFD.Where(w => w.TID == tid && w.SYR == p.SYR && w.FIT_CODE == "11").FirstOrDefault();
                if (c == null)
                {
                    c = new CTDITFD();
                    c.FIT_CODE = "11";
                    c.SYR = p.SYR;
                    c.TID = tid;
                    c.FUND = tools.safe2Decimal(val);
                    c.UPD_UID = WebSiteHelper.getUserId();
                    c.UPD_DT = now;
                    c.ADD_DT = now;
                    db_pms.CTDITFD.Add(c);
                    db_pms.SaveChanges();
                    return Json("已新增", JsonRequestBehavior.AllowGet);
                }
                else
                {
                    c.FUND = tools.safe2Decimal(val);
                    c.UPD_UID = WebSiteHelper.getUserId();
                    c.UPD_DT = now;
                    c.ADD_DT = now;
                    db_pms.SaveChanges();
                    return Json("已儲存", JsonRequestBehavior.AllowGet);
                }
            }




            //CTDMNDS c = db_pms.CTDMNDS.Where(w => w.TID == tid).FirstOrDefault();
            sql = "UPDATE　CTDMNDS SET UPD_DT='" + now.ToString("yyyy-MM-dd HH:mm") + "', UPD_UID='" + WebSiteHelper.getUserId() + "', " + id + "='" + val + "'WHERE TID=" + tid;
            SqlDao.CMEPMSExecuteSql(sql);
            return Json("已儲存", JsonRequestBehavior.AllowGet);

        }

        public ActionResult Index(PJE0030Model model)
        {

            int page = (model.prePage == 0) ? 1 : model.prePage;
            DataTable sortedTable = GetIndexList(model);


            List<DataRow> list = sortedTable.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.prePage = page;
            model.MainList = plist;
            return View("PJE0030Index", model);
        }

        /// <summary>
        /// 查詢
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SearchIndex")]
        public ActionResult SearchIndex(PJE0030Model model)
        {

            model.prePage = 1;
            return Index(model);
        }

        /// <summary>
        /// 列印
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Print")]
        public ActionResult Print(PJE0030Model model)
        {
            /// 載入模板
            Workbook workbook = new Workbook(Server.MapPath("~/Templates/PJE0030.xlsx"));
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            Worksheet sheet = workbook.Worksheets[0];

            /// 做預存程序
            DataTable dt;

            using (Models.CMEPMSModels.CMEPMSEntities cmepms = new Models.CMEPMSModels.CMEPMSEntities())
            using (SqlConnection conn = new SqlConnection(cmepms.Database.Connection.ConnectionString))
            using (SqlCommand cmd = new SqlCommand("SP_RCD0010_UDF003_3", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                /// 輸入參數
                cmd.Parameters.AddWithValue("@PID", model.PID);
                cmd.Parameters.AddWithValue("@SEQ_NO", model.SEQ_NO);

                dt = new DataTable();
                conn.Open();

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    dt.Load(reader);  /// 將結果讀入 DataTable
                }
            }

            /// Excel 第 3 列是資料樣板列，Aspose row index 為 2。
            int templateRow = 2;
            int startRow = templateRow;

            if (dt.Rows.Count > 1)
            {
                sheet.Cells.InsertRows(startRow + 1, dt.Rows.Count - 1, true);
            }

            for (int row = startRow + 1; row < startRow + dt.Rows.Count; row++)
            {
                sheet.Cells.CopyRow(sheet.Cells, templateRow, row);
            }

            foreach (DataRow d in dt.Rows)
            {

                sheet.Cells[startRow, 0].PutValue(d["TD_CODE"]);                                        /// 工程編號
                sheet.Cells[startRow, 1].PutValue(d["SYR"].ToString());                                 /// 年度
                sheet.Cells[startRow, 2].PutValue(d["W_CODE"].ToString());                              /// 計畫
                sheet.Cells[startRow, 3].PutValue(d["PJ_CODE"].ToString());                             /// 分支
                sheet.Cells[startRow, 4].PutValue(d["IPJ_CODE"].ToString());                            /// 分項
                sheet.Cells[startRow, 5].PutValue(d["CITY_CODE"].ToString());                           /// 縣市鄉鎮
                sheet.Cells[startRow, 6].PutValue(d["TL_CODE"].ToString());                             /// 工程
                sheet.Cells[startRow, 7].PutValue(d["TOD_CODE"].ToString());                            /// 序
                sheet.Cells[startRow, 8].PutValue(d["TD_NAME"].ToString());                             /// 工程名稱
                sheet.Cells[startRow, 9].PutValue(Convert.ToDouble(d["YR_FUND"].ToString()) / 1000);    /// 總經費
                sheet.Cells[startRow, 10].PutValue(Convert.ToDouble(d["CEN_FUND"].ToString()) / 1000);  /// 本署經費
                sheet.Cells[startRow, 11].PutValue(Convert.ToDouble(d["LOC_FUND"].ToString()) / 1000);  /// 中央直撥

                //DataTable cts = BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT * FROM CTDITFD c WHERE c.TID=" + d["TID"] + " ORDER BY c.FID DESC");
                //DataRow[] dr = cts.Select("FIT_CODE='11'");
                sheet.Cells[startRow, 12].PutValue(Convert.ToDouble(d["FUND_11"].ToString()) / 1000);   /// 地方自籌
                //sheet.Cells[startRow, 12].PutValue(dr.Length > 0 ? Convert.ToDouble(dr[0]["FUND"].ToString()) :0);   /// 地方自籌
                sheet.Cells[startRow, 13].PutValue(Convert.ToDouble(d["FUND_20"].ToString()) / 1000);   /// 特別預算
                //dr = cts.Select("FIT_CODE='03'");
                //sheet.Cells[startRow, 13].PutValue(dr.Length > 0 ? Convert.ToDouble(dr[0]["FUND"].ToString()) : 0);   /// 地方自籌
                //sheet.Cells[startRow, 13].PutValue(0);   /// 特別預算
                sheet.Cells[startRow, 14].PutValue(d["FIT_DESC"]);                                      /// 辦理方式
                sheet.Cells[startRow, 15].PutValue(d["TIT_DESC"]);                                      /// 費用別

                startRow++;
            }

            /// sheet.AutoFitColumns();
            sheet.Cells.SetColumnWidth(0, 25);   /// A 項次
            sheet.Cells.SetColumnWidth(1, 6);   /// B 縣市別
            sheet.Cells.SetColumnWidth(2, 6);  /// C 
            sheet.Cells.SetColumnWidth(3, 6);  /// D
            sheet.Cells.SetColumnWidth(4, 6);  /// E
            sheet.Cells.SetColumnWidth(5, 6);   /// F
            sheet.Cells.SetColumnWidth(6, 6);  /// G 
            sheet.Cells.SetColumnWidth(7, 6);  /// H
            sheet.Cells.SetColumnWidth(8, 55);  /// I
            sheet.Cells.SetColumnWidth(9, 16);  /// J
            sheet.Cells.SetColumnWidth(10, 16); /// K
            sheet.Cells.SetColumnWidth(11, 16); /// L
            sheet.Cells.SetColumnWidth(12, 16); /// M
            sheet.Cells.SetColumnWidth(13, 16); /// N
            sheet.Cells.SetColumnWidth(14, 16); /// O
            sheet.Cells.SetColumnWidth(15, 16); /// P
            sheet.Cells.SetColumnWidth(15, 16); /// P
            sheet.PageSetup.Zoom = 70;

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"PJE0030_{timestamp}.xlsx";

            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                ms.Position = 0;
                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        /// <summary>
        /// 維護
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Maintain")]
        public ActionResult Maintain(PJE0030Model model)
        {
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)

            int page = (model.page == 0) ? 1 : model.page;
            SqlCommand cmd = new SqlCommand("SELECT*FROM PJDFDCR p WHERE p.PID=@PID ORDER BY SEQ_NO DESC");
            //cmd.Parameters.Add(new SqlParameter("PID", tools.SqlFilterInt(model.PID)));
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.List1 = plist;
            model.PJDFDCR = new PJDFDCR();
            model.PJMFORM = db_pms.PJMFORM.Where(w => w.PID == model.PID).FirstOrDefault();
            return View("PJE0030Maintain", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "AddPJDFDCR")]
        public ActionResult AddPJDFDCR(PJE0030Model model)
        {
            //做預存程序
            int returnValue = -1;
            Models.CMEPMSModels.CMEPMSEntities cmepms = new Models.CMEPMSModels.CMEPMSEntities();
            using (SqlConnection conn = new SqlConnection(cmepms.Database.Connection.ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand("SP_PJP0050_PJNFCR", conn))
                {

                    cmd.CommandType = CommandType.StoredProcedure;

                    // 輸入參數
                    cmd.Parameters.AddWithValue("@PID", model.PID);
                    cmd.Parameters.AddWithValue("@UPD_UID", WebSiteHelper.getUserId());

                    // 回傳值參數
                    SqlParameter returnParam = new SqlParameter();
                    returnParam.Direction = ParameterDirection.ReturnValue;
                    returnParam.SqlDbType = SqlDbType.Int;
                    cmd.Parameters.Add(returnParam);

                    conn.Open();
                    try
                    {
                        cmd.ExecuteNonQuery(); // 注意：這裡使用 ExecuteNonQuery（因為不取結果集）
                        returnValue = (int)returnParam.Value;
                    }
                    catch (Exception e)
                    {
                        ViewBag.msg = "資料處理錯誤: " + e.Message;
                        return Maintain(model);
                    }


                    // 取得儲存過程的 RETURN 值

                    switch (returnValue)
                    {
                        case 0:

                            ViewBag.msg = "新增一期 成功";
                            break;
                        default:
                            ViewBag.msg = "資料處理錯誤，錯誤代碼：";
                            break;
                    }
                }
            }

            return Maintain(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SavePJDFDCR")]
        public ActionResult SavePJDFDCR(PJE0030Model model)
        {
            DateTime now = DateTime.Now;

            PJDFDCR p;
            p = (from pj in db_pms.PJDFDCR where pj.PID == model.PID && pj.SEQ_NO == model.PJDFDCR.SEQ_NO select pj).FirstOrDefault();
            p.APROV_FN = model.PJDFDCR.APROV_FN;
            p.APROV_DT = model.PJDFDCR.APROV_DT;
            p.AP_DT = model.PJDFDCR.AP_DT;
            p.AP_FN = model.PJDFDCR.AP_FN;
            p.AP_OUT_FUND = model.PJDFDCR.AP_OUT_FUND;
            p.FDCR_RESN = model.PJDFDCR.FDCR_RESN;
            p.UPD_UID = Helpers.WebSiteHelper.getUserId();
            p.UPD_DT = now;

            if (!string.IsNullOrEmpty(model.addkey))
            {
                db_pms.PJDFDCR.Add(p);
            }
            db_pms.SaveChanges();

            PJMFORM_CMEPMS pjmform = db_pms.PJMFORM.Where(w => w.PID == model.PID).FirstOrDefault();
            if (model.FCAL_FG.Equals("Y"))
            {

                pjmform.FLOCK_FG = true;
                db_pms.SaveChanges();
            }
            else
            {
                pjmform.FLOCK_FG = false;
                db_pms.SaveChanges();
            }




            //做預存程序
            /*
            int returnValue = -1;
            Models.CMEPMSModels.CMEPMSEntities cmepms = new Models.CMEPMSModels.CMEPMSEntities();
            using(SqlConnection conn = new SqlConnection(cmepms.Database.Connection.ConnectionString)) {
                using(SqlCommand cmd = new SqlCommand("SP_PJP0050_PJNFCR", conn)) {

                    cmd.CommandType = CommandType.StoredProcedure;

                    // 輸入參數
                    cmd.Parameters.AddWithValue("@PID", model.PID);
                    cmd.Parameters.AddWithValue("@UPD_UID", WebSiteHelper.getUserId());

                    // 回傳值參數
                    SqlParameter returnParam = new SqlParameter();
                    returnParam.Direction = ParameterDirection.ReturnValue;
                    returnParam.SqlDbType = SqlDbType.Int;
                    cmd.Parameters.Add(returnParam);

                    conn.Open();
                    try {
                        cmd.ExecuteNonQuery(); // 注意：這裡使用 ExecuteNonQuery（因為不取結果集）
                        returnValue = (int)returnParam.Value;
                    } catch(Exception e) {
                        ViewBag.msg = "資料處理錯誤: "+e.Message;
                        return Maintain(model);
                    }
                    

                    // 取得儲存過程的 RETURN 值
                    
                    switch(returnValue) {
                        case 0:

                            ViewBag.msg = "新增一期 成功";
                            break;
                        default:
                            ViewBag.msg = "資料處理錯誤，錯誤代碼：";
                            break;
                    }
                }
            }
            */




            return Maintain(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "DelPJDFDCR")]
        public ActionResult DelPJDFDCR(PJE0030Model model)
        {

            SqlCommand cmd = new SqlCommand("DELETE FROM PJDFDCR WHERE PID=@PID AND SEQ_NO=@SEQ_NO");
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            cmd.Parameters.Add(new SqlParameter("SEQ_NO", model.PJDFDCR.SEQ_NO));
            SqlDao.CMEPMSExecuteSql(cmd);
            ViewBag.msg = "刪除完成";
            //ViewBag.desc = "請確認資料";
            return Maintain(model);
        }

        [HttpPost]
        public JsonResult GetPJDFDCR(string PID, string SEQ_NO, string edit)
        {

            SqlCommand cmd = new SqlCommand("SELECT m.FLOCK_FG, p.* FROM PJDFDCR p, PJMFORM m WHERE m.PID=p.PID AND p.PID=@PID AND p.SEQ_NO=@SEQ_NO");
            cmd.Parameters.Add(new SqlParameter("PID", PID));
            cmd.Parameters.Add(new SqlParameter("SEQ_NO", SEQ_NO));
            DataTable dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 取得單一縣市的工程資料
        /// </summary>
        /// <param name="city"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetContinuingProjects(string city)
        {
            string sql = "SELECT TD_CODE, TD_NAME as AP_NAME, TD_CITY as CITY, TD_TOWN as TOWN FROM TDMFORM WHERE TD_CITY = @city AND TD_CODE IS NOT NULL ORDER BY TD_CODE";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("city", city));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 管制
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Manage")]
        public ActionResult Manage(PJE0030Model model)
        {
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)

            int page = (model.page == 0) ? 1 : model.page;


            int SEQ_MAX = util.GetSEQ_NO(model.PID);
            int SEQ_NO;
            if (model.SEQ_NO == 0)
            {
                //0為初次進入
                SEQ_NO = util.GetSEQ_NO(model.PID);
                model.SEQ_NO = SEQ_NO;
                //GetPJDFDCR_Max(model.PID, model);
                model = GetPJDFDCR_Max(model.PID, model, page, 1);
            }
            else
            {
                //已選取期數
                SEQ_NO = model.SEQ_NO;
                if (SEQ_NO == SEQ_MAX)
                {
                    //取最新
                    //GetPJDFDCR_Max(model.PID, model);
                    model = GetPJDFDCR_Max(model.PID, model, page, 1);
                }
                else
                {
                    //取某期
                    model = GetPJDFDCR_By(model.PID, model, page, 1);
                }
            }

            Session["fund"] =
                SqlDao.CMEPMSQuery4List("select sum(YR_FUND)as YR_FUND, " +
                "sum(CEN_FUND)as CEN_FUND, sum(LOC_FUND)as LOC_FUND from " +
                "TDV_TDDCFTS where PID =" + model.PID);
            GetSelectList(model.PID);
            return View("PJE0030Manage", model);
        }

        public ActionResult ManageInfo(int PID, int SEQ, string SYR)
        {

            int page = 1;
            PJE0030Model model = new PJE0030Model();
            model.PID = PID;
            model.SEQ_NO = SEQ;
            model.SYR = SYR;

            int SEQ_MAX = util.GetSEQ_NO(model.PID);
            int SEQ_NO;
            if (model.SEQ_NO == 0)
            {
                //0為初次進入
                SEQ_NO = util.GetSEQ_NO(model.PID);
                model.SEQ_NO = SEQ_NO;
                //GetPJDFDCR_Max(model.PID, model);
                model = GetPJDFDCR_Max(model.PID, model, page, 1);
            }
            else
            {
                //已選取期數
                SEQ_NO = model.SEQ_NO;
                if (SEQ_NO == SEQ_MAX)
                {
                    //取最新
                    //GetPJDFDCR_Max(model.PID, model);
                    model = GetPJDFDCR_Max(model.PID, model, page, 1);
                }
                else
                {
                    //取某期
                    model = GetPJDFDCR_By(model.PID, model, page, 1);
                }
            }



            //Session["fund"] = SqlDao.CMEPMSQuery4List("select sum(YR_FUND)as YR_FUND, sum(CEN_FUND)as CEN_FUND, sum(LOC_FUND)as LOC_FUND from TDV_TDDCFTS where PID =" + model.PID);
            GetSelectList(model.PID);
            return View("PJE0030ManageInfo", model);
        }

        public ActionResult ManageList1(string PJ_CATY_CODE, string TD_NAME, string CITY, string TOWN, string TD_CATY, string EXE_TYPE, string IPJ_CODE, string N_IPJ_CODE, string EMG_CATY, string ACC_ID, int PID, int page, int SEQ_NO)
        {
            PJE0030Model model = new PJE0030Model();
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.ACC_ID = ACC_ID;
            model.PID = PID;
            model.SEQ_NO = SEQ_NO;
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.EMG_CATY = EMG_CATY;
            model.ACC_ID = ACC_ID;

            page = (page == 0) ? 1 : page;
            model.page = page;
            int SEQ_MAX = util.GetSEQ_NO(model.PID);
            //int SEQ_NO;
            if (SEQ_NO == 0)
            {
                //0為初次進入
                SEQ_NO = util.GetSEQ_NO(model.PID);
                model = GetPJDFDCR_Max(model.PID, model, model.page, 1);
            }
            else
            {
                //已選取期數                
                if (SEQ_NO == SEQ_MAX)
                {
                    //取最新
                    model = GetPJDFDCR_Max(model.PID, model, model.page, 1);
                }
                else
                {
                    //取某期
                    model = GetPJDFDCR_By(model.PID, model, model.page, 1);
                }
            }
            //model.SEQ_NO = SEQ_NO;
            GetSelectList(model.PID);
            return View("PJE0030ManageInfo", model);
        }

        public ActionResult Search(int SEQ_NO, int PID, string PJ_CATY_CODE, string TD_NAME, string CITY, string TOWN, string TD_CATY, string EXE_TYPE, string IPJ_CODE, string N_IPJ_CODE, string EMG_CATY, string ACC_ID, int zero)
        {
            PJE0030Model model = new PJE0030Model();
            int max = util.GetSEQ_NO(PID);
            model.PID = PID;
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.EMG_CATY = EMG_CATY;
            model.ACC_ID = ACC_ID;
            model.page = 1;
            model.SEQ_NO = SEQ_NO;

            if (model.SEQ_NO == max)
            {
                model = GetPJDFDCR_Max(PID, model, 1, 1);//2:經費為限
            }
            else
            {
                model = GetPJDFDCR_By(PID, model, 1, 1);
            }
            GetSelectList(model.PID);
            return View("PJE0030ManageInfo", model);
        }

        /// <summary>
        /// 經費為0頁
        /// </summary>
        /// <param name="PJ_CATY_CODE"></param>
        /// <param name="TD_NAME"></param>
        /// <param name="CITY"></param>
        /// <param name="TOWN"></param>
        /// <param name="TD_CATY"></param>
        /// <param name="EXE_TYPE"></param>
        /// <param name="IPJ_CODE"></param>
        /// <param name="N_IPJ_CODE"></param>
        /// <param name="EMG_CATY"></param>
        /// <param name="ACC_ID"></param>
        /// <param name="PID"></param>
        /// <param name="page"></param>
        /// <param name="SEQ_NO"></param>
        /// <returns></returns>
        public ActionResult ManageList2(string PJ_CATY_CODE, string TD_NAME, string CITY, string TOWN, string TD_CATY, string EXE_TYPE, string IPJ_CODE, string N_IPJ_CODE, string EMG_CATY, string ACC_ID, int PID, int page, int SEQ_NO)
        {
            PJE0030Model model = new PJE0030Model();
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.ACC_ID = ACC_ID;
            model.PID = PID;
            model.SEQ_NO = SEQ_NO;
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.EMG_CATY = EMG_CATY;
            model.ACC_ID = ACC_ID;


            page = (page == 0) ? 1 : page;
            model.page = page;
            int SEQ_MAX = util.GetSEQ_NO(model.PID);

            if (SEQ_NO == 0)
            {
                //0為初次進入
                SEQ_NO = util.GetSEQ_NO(model.PID);
                model = GetPJDFDCR_Max(model.PID, model, model.page, 2);
            }
            else
            {
                //已選取期數
                //SEQ_NO = model.SEQ_NO;
                if (SEQ_NO == SEQ_MAX)
                {
                    //取最新
                    model = GetPJDFDCR_Max(model.PID, model, model.page, 2);
                }
                else
                {
                    //取某期
                    model = GetPJDFDCR_By(model.PID, model, model.page, 2);
                }
            }
            //model.SEQ_NO = SEQ_NO;
            GetSelectList(model.PID);
            return View("PJE0030ManageInfo", model);
        }

        /// <summary>
        /// 保留款頁
        /// </summary>
        /// <param name="tid"></param>
        /// <returns></returns>
        public ActionResult EditCRDMNDS(decimal tid, decimal pid)
        {
            PJE0030Model model = new PJE0030Model();

            model.PID = pid;

            try
            {
                model.CRDMNDS = db_pms.CRDMNDS.Where(w => w.TID == tid).FirstOrDefault();
            }
            catch
            {
                model.CRDMNDS = new CRDMNDS();
            }


            return View("PJE0030EditCRDMNDS", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SaveCRDMNDS")]
        public ActionResult SaveCRDMNDS(PJE0030Model model)
        {
            decimal total = 0;
            total += model.CRDMNDS.M_01.Value;
            total += model.CRDMNDS.M_02.Value;
            total += model.CRDMNDS.M_03.Value;
            total += model.CRDMNDS.M_04.Value;
            total += model.CRDMNDS.M_05.Value;
            total += model.CRDMNDS.M_06.Value;
            total += model.CRDMNDS.M_07.Value;
            total += model.CRDMNDS.M_08.Value;
            total += model.CRDMNDS.M_09.Value;
            total += model.CRDMNDS.M_10.Value;
            total += model.CRDMNDS.M_11.Value;
            total += model.CRDMNDS.M_12.Value;

            if (total != model.CRDMNDS.REV_FUND)
            {
                ViewBag.msg = "月分配數(每月合計總數 應等於 保留金額)";
                return EditCRDMNDS(model.TID, model.PID);
            }
            DateTime now = DateTime.Now;
            CRDMNDS c;
            PJMFORM_CMEPMS p = db_pms.PJMFORM.Where(w => w.PID == model.PID).FirstOrDefault();
            CTDMNDS t = db_pms.CTDMNDS.Where(w => w.TID == model.TID).FirstOrDefault();



            if (model.CRDMNDS.TID == 0)
            {
                // TODO: 如果會有新增的情況
                c = new CRDMNDS();
                c.PID = p.PID;
                c.SYR = p.SYR;
                c.TD_CODE = model.CRDMNDS.TD_CODE;
                c.TD_CODE1 = model.CRDMNDS.TD_CODE.Replace("-", "");
                c.TD_NAME = t.TD_NAME;
                //c.CEN_SP_RATIO
                //c.TD_CITY
                //c.ACC_ID
                //c.CEN_SP_RATIO
                //c.FD_CODE
                //c.FIT_CODE
                c.CATY_CODE = t.CATY_CODE;
                c.TID = t.TID;
                c.ADD_DT = now;
            }
            else
            {
                c = db_pms.CRDMNDS.Where(w => w.TID == model.CRDMNDS.TID).FirstOrDefault();
            }
            // 保留款無法修改"保留金額"、"發生權責數"、"特別分配數"、"保留原因"欄位
            //c.REV_FUND = model.CRDMNDS.REV_FUND;
            //c.OCR_FUND = model.CRDMNDS.OCR_FUND;
            //c.SPD_FUND = model.CRDMNDS.SPD_FUND;
            //c.REV_RESN = model.CRDMNDS.REV_RESN;
            c.M_01 = model.CRDMNDS.M_01;
            c.M_02 = model.CRDMNDS.M_02;
            c.M_03 = model.CRDMNDS.M_03;
            c.M_04 = model.CRDMNDS.M_04;
            c.M_05 = model.CRDMNDS.M_05;
            c.M_06 = model.CRDMNDS.M_06;
            c.M_07 = model.CRDMNDS.M_07;
            c.M_08 = model.CRDMNDS.M_08;
            c.M_09 = model.CRDMNDS.M_09;
            c.M_10 = model.CRDMNDS.M_10;
            c.M_11 = model.CRDMNDS.M_11;
            c.M_12 = model.CRDMNDS.M_12;

            c.UPD_DT = now;
            c.UPD_UID = WebSiteHelper.getUserId();

            if (model.CRDMNDS.TID == 0)
            {
                db_pms.CRDMNDS.Add(c);
                db_pms.SaveChanges();
            }
            else
            {
                db_pms.SaveChanges();
            }

            ViewBag.msg = "已儲存";
            return EditCRDMNDS(model.TID, model.PID);
        }

        public ActionResult ManageList3(string PJ_CATY_CODE, string TD_NAME, string CITY, string TOWN, string TD_CATY, string EXE_TYPE, string IPJ_CODE, string N_IPJ_CODE, string EMG_CATY, string ACC_ID, int PID, int page, int SEQ_NO)
        {
            PJE0030Model model = new PJE0030Model();

            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.ACC_ID = ACC_ID;
            model.PID = PID;
            model.SEQ_NO = SEQ_NO;
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.EMG_CATY = EMG_CATY;
            model.ACC_ID = ACC_ID;


            page = (page == 0) ? 1 : page;
            model.page = page;
            model.PJMFORM = db_pms.PJMFORM.Where(w => w.PID == model.PID).FirstOrDefault();
            if (model.PJMFORM != null)
            {
                model.SYR = model.PJMFORM.SYR;
            }

            int SEQ_MAX = util.GetSEQ_NO(model.PID);

            /// int SEQ_NO;
            if (SEQ_NO == 0)
            {
                /// 0　為初次進入
                SEQ_NO = util.GetSEQ_NO(model.PID);
                model = GetPJDFDCR_Max(model.PID, model, model.page, 3);
            }
            else
            {
                /// 已選取期數                
                if (SEQ_NO == SEQ_MAX)
                {
                    /// 取最新
                    model = GetPJDFDCR_Max(model.PID, model, model.page, 3);
                }
                else
                {
                    /// 取某期
                    model = GetPJDFDCR_By(model.PID, model, model.page, 3);
                }
            }

            GetSelectList(model.PID);

            return View("PJE0030ManageInfo", model);
        }

        /// <summary>
        /// 各項下拉選單
        /// </summary>
        /// <param name="PID"></param>
        private void GetSelectList(object PID)
        {
            SqlCommand cmd = new SqlCommand("select SEQ_NO from PJDFDCR where PID = @PID ORDER BY SEQ_NO DESC");
            cmd.Parameters.Add(new SqlParameter("PID", PID));
            ViewBag.sels = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            DataTable dt;
            ViewBag.cities = tools.ToSelectList(util.GetCities(), "CITY", "CITY", null);

            ViewBag.towns = tools.ToSelectList(GetTowns(null), "TOWN", "TOWN", null);

            //ViewBag.towns = tools.ToSelectList(util.GetTownNew(null), "TOWN", "TOWN", null);
            //類型/工作類別
            //ViewBag.TD_CATY = tools.ToSelectList(util.GetCMEPMSCODES("FD", "00"), "id", "name", null);

            //執行方式
            dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("name");
            DataRow dr = dt.NewRow();
            dr["id"] = "01";
            dr["name"] = "01-非BOT";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["id"] = "02";
            dr["name"] = "02-BOT";
            dt.Rows.Add(dr);
            ViewBag.EXE_TYPE = tools.ToSelectList(dt, "id", "name", null);

            //經費層級
            ViewBag.IPJ_CODE = tools.ToSelectList(util.GetCMEPMSCODES("SM", "02"), "id", "name", null);


            //辦理層級
            ViewBag.N_IPJ_CODE = tools.ToSelectList(util.GetCMETMSCODES("SM", "02"), "id", "name", null);

            //分類(雨水用)
            ViewBag.EMG_CATY = tools.ToSelectList(util.GetCMEPMSCODES("TT", "05"), "id", "name", null);

            //會計科目
            ViewBag.ACC_ID = tools.ToSelectList(util.GetCMEPMSCODES("FD", "02"), "id", "name", null);
            //計畫分項
            ViewBag.PJ_CATY_CODE = tools.ToSelectList(BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT CATY_CODE as id, CATY_CODE+CATY_NAME as name FROM PJDCATY WHERE PID=" + PID), "id", "name", null);
            //工程類別
            ViewBag.YR_WORK = tools.ToSelectList(util.GetCMETMSCODES("TT", "04"), "id", "name", null);

            //工作類別
            //ViewBag.TD_CATY = tools.ToSelectList(util.GetCMETMSCODES("SM", "09"), "id", "name", null);
            ViewBag.TD_CATY = tools.ToSelectList(util.GetCMETMSCODES("SM", "09"), "id", "name", null);

            //單位
            string sql = "SELECT o.DEPT_NO,\r\n    MAX(o.DEPT_NAME) AS DEPT_NAME, \r\n COUNT(r.USR_ID) AS USER_COUNT \r\nFROM OGCDPUR r\r\nJOIN OGCDEPT o ON o.DEPT_NO = r.DEPT_NO\r\nJOIN OGCUSER u ON u.USR_ID = r.USR_ID\r\nGROUP BY o.DEPT_NO;";
            ViewBag.depts = tools.ToSelectList(SqlDao.CMBASEQuery4List(sql), "DEPT_NO", "DEPT_NAME", "請選擇單位");

            //案件類型
            ViewBag.CATY_COD = tools.ToSelectList(util.GetCMETMSCODES("TT", "03"), "id", "name", null);

            //經費來源
            ViewBag.FIT_CODE = tools.ToSelectList(util.GetCMEPMSCODES("FD", "05"), "id", "name", null);
            //人員
            ViewBag.users = tools.ToSelectList(SqlDao.CMBASEQuery4List("SELECT USR_ID, USR_NAME FROM OGCUSER"), "USR_ID", "USR_NAME", "請選擇");

        }

        private DataTable GetRetainageFundSummary(object PID, bool useRetainageView)
        {
            string retainageSource = useRetainageView ? "PJV_CRDMNDS" : "TDV_TDDRFLS";


            //SqlCommand cmd = new SqlCommand(
            //    "SELECT ISNULL(SUM(T.YR_FUND), 0) AS YR_FUND, " +
            //    "ISNULL(SUM(T.CEN_FUND), 0) AS CEN_FUND, " +
            //    "ISNULL(SUM(T.LOC_FUND), 0) AS LOC_FUND " +
            //    "FROM TDV_TDDCFTS T " +
            //    "WHERE EXISTS (SELECT 1 FROM " + retainageSource + " R WHERE R.PID = @PID AND R.TID = T.TID)");

            /// 20260514 版本修改 => 修改 "保留款" 頁籤取的方式
            string sql = $@"
                    SELECT 
                        ISNULL(SUM(ISNULL(T.CEN_FUND, 0) + ISNULL(T.LOC_FUND, 0)), 0) AS YR_FUND,
                        ISNULL(SUM(T.CEN_FUND), 0) AS CEN_FUND,
                        ISNULL(SUM(T.LOC_FUND), 0) AS LOC_FUND
                    FROM TDV_TDDCFTS T
                    INNER JOIN {retainageSource} R
                        ON R.TID = T.TID
                       AND R.PID = T.PID
                    WHERE T.PID = @PID ";

            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("PID", PID));

            Debug.WriteLine($"SQL:{cmd}");

            return SqlDao.CMEPMSQuery4List(cmd);
        }

        /// <summary>
        /// 保留款: 取最新期 CTDMNDS
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="m"></param>
        /// <param name="page"></param>
        /// <param name="zero"></param>
        /// <returns></returns>
        private PJE0030Model GetPJDFDCR_Max(object PID, PJE0030Model m, int page, int zero)
        {
            /// 取最新期 PID
            StringBuilder sb = new StringBuilder("SELECT * FROM TDV_TDDCFTS WHERE PID = @PID ");

            string sql;

            /// 分項
            if (!string.IsNullOrEmpty(m.PJ_CATY_CODE))
            {
                sb.Append("AND PJ_CATY_CODE='" + m.PJ_CATY_CODE + "'");
            }

            /// 名稱
            if (!string.IsNullOrEmpty(m.TD_NAME))
            {
                sb.Append("AND TD_NAME LIKE'%" + m.TD_NAME + "%'");
            }

            /// 縣市
            if (!string.IsNullOrEmpty(m.CITY))
            {
                sb.Append("AND CITY='" + m.CITY + "'");
            }

            /// 鄉鎮
            if (!string.IsNullOrEmpty(m.TOWN) && m.TOWN.Trim() != "不分區")
            {
                sb.Append("AND TOWN='" + m.TOWN + "'");
            }

            /// 工作類型
            if (!string.IsNullOrEmpty(m.TD_CATY))
            {
                sb.Append("AND TD_CATY='" + m.TD_CATY + "'");
            }

            /// 執行方式
            if (!string.IsNullOrEmpty(m.EXE_TYPE))
            {
                sb.Append("AND EXE_TYPE='" + m.EXE_TYPE + "'");
            }

            /// 經費層級
            if (!string.IsNullOrEmpty(m.IPJ_CODE))
            {
                sb.Append("AND IPJ_CODE='" + m.IPJ_CODE + "'");
            }

            /// 辦理層級
            if (!string.IsNullOrEmpty(m.N_IPJ_CODE))
            {
                sb.Append("AND N_IPJ_CODE='" + m.N_IPJ_CODE + "'");
            }

            /// 分類
            if (!string.IsNullOrEmpty(m.EMG_CATY))
            {
                sb.Append("AND EMG_CATY='" + m.EMG_CATY + "'");
            }

            /// 會計科目
            if (!string.IsNullOrEmpty(m.ACC_ID))
            {
                sb.Append("AND ACC_ID='" + m.ACC_ID + "'");
            }

            sql = sb.ToString();

            SqlCommand cmd;

            List<DataRow> list;

            DataTable ps;

            IPagedList<DataRow> plist;

            if (zero == 1)
            {
                sql += "AND(YR_FUND>0 OR YR_FUND IS NULL) ORDER BY TID DESC";

                cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("PID", PID));

                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List1 = plist;



                /// 20260518 版本修改 => 不管如何切換頁數維持顯示"分支計畫總經費"

                #region 舊版邏輯 "執行中": 總經費/中央款/地方款
                //sb = new StringBuilder("select sum(YR_FUND)as YR_FUND, " +
                //"sum(CEN_FUND)as CEN_FUND, sum(LOC_FUND)as LOC_FUND from " +
                //"TDV_TDDCFTS where TID IN(");

                //foreach (DataRow d in plist)
                //{
                //    sb.Append("'" + d["TID"] + "',");
                //}

                //sb.Length--;
                //sb.Append(")");

                //if (plist.Count > 0)
                //{
                //    ViewBag.fund1 = SqlDao.CMEPMSQuery4List(sb.ToString());

                //    DataTable fund1 = SqlDao.CMEPMSQuery4List(sb.ToString());

                //    if (!string.IsNullOrEmpty(Convert.ToString(fund1.Rows[0][0])))
                //    {
                //        ViewBag.fund1 = fund1;
                //    }
                //    else
                //    {
                //        ViewBag.fund1 = null;
                //    }
                //}
                #endregion

                ViewBag.fund1 = GetProjectFundSummary(PID, m.SYR, m.SEQ_NO, m.CITY, m.PJ_CATY_CODE);
            }

            if (zero == 2)
            {
                //經費為0
                sql = sb.ToString();
                sql += "AND YR_FUND=0 ORDER BY TID DESC";
                cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("PID", PID));
                //ViewBag.List2 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List2 = plist;
                sb.Length--;
                sb = new StringBuilder("select sum(YR_FUND)as YR_FUND, " +
                "sum(CEN_FUND)as CEN_FUND, sum(LOC_FUND)as LOC_FUND from " +
                "TDV_TDDCFTS where TID IN(");
                foreach (DataRow d in plist)
                {
                    sb.Append("'" + d["TID"] + "',");
                }
                sb.Length--;
                sb.Append(")");

                if (plist.Count > 0)
                {
                    ViewBag.fund2 = SqlDao.CMEPMSQuery4List(sb.ToString());

                    DataTable fund2 = SqlDao.CMEPMSQuery4List(sb.ToString());

                    if (!string.IsNullOrEmpty(Convert.ToString(fund2.Rows[0][0])))
                    {
                        ViewBag.fund2 = fund2;
                    }
                    else
                    {
                        ViewBag.fund2 = null;
                    }
                }

            }

            /// 20260514 版本修改 => 修改 "保留款" 頁籤取的方式
            if (zero == 3)
            {
                /// 保留款：清單主資料改從 TDV_TDDCFTS 取得
                /// PJV_CRDMNDS 只用來判斷哪些 TID 屬於保留款 
                string cmetmsDbName = db_tms.Database.Connection.Database;

                string retainageSql = $@"
                                        SELECT
                                            A.PID,
                                            A.TID,
                                            C.TD_CITY,
                                            ISNULL(B.SYR, '') AS SYR,
                                            ISNULL(B.SMN, '') AS SMN,
                                            A.FD_CODE,
                                            A.TD_CODE,
                                            A.TD_NAME,
                                            ISNULL(B.BAM, 0) AS BAM,
                                            ISNULL(B.AE_CTM, 0) + ISNULL(B.PP_CTM, 0) AS REV_FUND,
                                            ISNULL(D.REV_RESN, '') AS REV_RESN
                                        FROM CTDMNDS A
                                        LEFT JOIN TDDCTET B ON B.FD_CODE = A.TD_CODE
                                        LEFT JOIN [{cmetmsDbName}].dbo.TDMFORM C ON C.TD_CODE = A.TD_CODE
                                        LEFT JOIN CRDMNDS D ON D.TID = A.TID
                                        WHERE A.PID = @PID ; ";

                cmd = new SqlCommand(retainageSql);
                cmd.Parameters.Add(new SqlParameter("PID", PID));

                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List3 = plist;

                ViewBag.fund3 = GetProjectFundSummary(PID, m.SYR, m.SEQ_NO);
            }

            return m;
        }

        /// <summary>
        /// 取某期CTHMNDS
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="m"></param>
        /// <param name="page"></param>
        /// <param name="zero"></param>
        /// <returns></returns>
        private PJE0030Model GetPJDFDCR_By(object PID, PJE0030Model m, int page, int zero)
        {
            StringBuilder sb = new StringBuilder("SELECT * FROM TDV_TDHCFTS WHERE PID = @PID ");

            string sql;

            /// 分項
            if (!string.IsNullOrEmpty(m.PJ_CATY_CODE))
            {
                sb.Append("AND PJ_CATY_CODE='" + m.PJ_CATY_CODE + "'");
            }

            /// 名稱
            if (!string.IsNullOrEmpty(m.TD_NAME))
            {
                sb.Append("AND TD_NAME LIKE'%" + m.TD_NAME + "%'");
            }

            /// 縣市
            if (!string.IsNullOrEmpty(m.CITY))
            {
                sb.Append("AND CITY='" + m.CITY + "'");
            }

            /// 鄉鎮
            if (!string.IsNullOrEmpty(m.TOWN))
            {
                sb.Append("AND TOWN='" + m.TOWN + "'");
            }

            /// 工作類型
            if (!string.IsNullOrEmpty(m.TD_CATY))
            {
                sb.Append("AND TD_CATY='" + m.TD_CATY + "'");
            }

            /// 執行方式
            if (!string.IsNullOrEmpty(m.EXE_TYPE))
            {
                sb.Append("AND EXE_TYPE='" + m.EXE_TYPE + "'");
            }

            /// 經費層級
            if (!string.IsNullOrEmpty(m.IPJ_CODE))
            {
                sb.Append("AND IPJ_CODE='" + m.IPJ_CODE + "'");
            }

            /// 辦理層級
            if (!string.IsNullOrEmpty(m.N_IPJ_CODE))
            {
                sb.Append("AND N_IPJ_CODE='" + m.N_IPJ_CODE + "'");
            }

            /// 分類
            if (!string.IsNullOrEmpty(m.EMG_CATY))
            {
                sb.Append("AND EMG_CATY='" + m.EMG_CATY + "'");
            }

            /// 會計科目
            if (!string.IsNullOrEmpty(m.ACC_ID))
            {
                sb.Append("AND ACC_ID='" + m.ACC_ID + "'");
            }

            /// SEQ_NO
            sb.Append("AND SEQ_NO='" + m.SEQ_NO + "'");

            sql = sb.ToString();
            SqlCommand cmd;
            List<DataRow> list;
            DataTable ps;
            IPagedList<DataRow> plist;

            if (zero == 1)
            {
                sql += "AND(YR_FUND>0 OR YR_FUND IS NULL) ORDER BY TID DESC";
                cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("PID", PID));
                //ViewBag.List1 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List1 = plist;

            }

            if (zero == 2)
            {
                //經費為0
                sql = sb.ToString();
                sql += "AND YR_FUND=0 ORDER BY TID DESC";
                cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("PID", PID));
                //ViewBag.List2 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List2 = plist;
            }

            /// 20260514 版本修改 => 修改 "保留款" 頁籤取得方式
            if (zero == 3)
            {
                /// 保留款: 歷史期清單主資料也改從 TDV_TDDCFTS 取得
                /// TDV_TDDRFLS 只用來判斷哪些 TID 屬於保留款
                /// cmd = new SqlCommand("select * from TDV_TDDRFLS where PID = @PID ORDER BY PID_SYR DESC, PID DESC");
                /// 
                string retainageSql = $@"
                            SELECT 
                                T.*,
                                R.REV_FUND,
                                R.OCR_FUND,
                                R.SPD_FUND,
                                R.REV_RESN
                            FROM TDV_TDDCFTS T
                            INNER JOIN TDV_TDDRFLS R
                                ON R.TID = T.TID
                               AND R.PID = T.PID
                            WHERE T.PID = @PID
                            ORDER BY T.PID_SYR DESC, T.PID DESC, T.TID DESC
                        ";

                cmd = new SqlCommand(retainageSql);
                cmd.Parameters.Add(new SqlParameter("PID", PID));

                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List3 = plist;

                ViewBag.fund3 = GetProjectFundSummary(PID, m.SYR, m.SEQ_NO);
            }
            return m;
        }

        /// <summary>
        /// 管制 - 查詢
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Search")]
        public ActionResult Search(PJE0030Model model)
        {
            int max = util.GetSEQ_NO(model.PID);
            model.page = (model.page == 0) ? 1 : model.page;

            if (model.SEQ_NO == max)
            {
                model = GetPJDFDCR_Max(model.PID, model, 1, 1);
            }
            else
            {
                model = GetPJDFDCR_By(model.PID, model, 1, 1);
            }
            GetSelectList(model.PID);
            return View("PJE0030ManageInfo", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "delCTDMNDS")]
        public ActionResult delCTDMNDS(PJE0030Model model)
        {
            try
            {
                //1先檢查是否有其他年度經費使用相同工程編號
                int s1 = Convert.ToInt16(SqlDao.CMEPMSQuery4Object("select count(*)as cnt from CMEPMS.dbo.CTDMNDS where TD_CODE='" + model.TID + "'")["cnt"]);
                //2再檢查是否有延續性提案工程使用相同工程編號
                int s2 = Convert.ToInt16(SqlDao.CMPSGSQuery4Object("select count(*)as cnt from APMFORM where TD_CODE='" + model.TID + "'")["cnt"]);
                //符合條件「1」=1, 「2」=0 me  則刪除額外資料動作 
                if (s1 == 1 && s2 == 0)
                {
                    SqlDao.CMEPMSExecuteSql("delete from CMETMS.dbo.TDMFORM where TD_CODE='" + model.TID + "'");
                }
            }
            catch
            {
                //查詢出現問題

            }

            //刪除動作
            SqlDao.CMEPMSExecuteSql("DELETE FROM CTDMNDS WHERE TID=" + model.TID);
            SqlDao.CMEPMSExecuteSql("DELETE FROM CTDITFD WHERE TID=" + model.TID);
            ViewBag.msg = "已刪除";
            //return Search(model);
            return Redirect(Request.UrlReferrer.ToString());
        }

        /// <summary>
        /// 管制-新增介面, model->失誤返回
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="TID"></param>
        /// <param name="SEQ_NO"></param>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult AddMnd(decimal PID, decimal TID, int SEQ_NO, PJE0030Model model)
        {
            PJV_PJMFORM p = db_pms.PJV_PJMFORM.Where(x => x.PID == PID).FirstOrDefault();

            model.TID = TID;
            model.PID = PID;
            model.SEQ_NO = SEQ_NO;

            if (model.TDV_TDHCFTS == null)
            {
                model.TDV_TDHCFTS = new TDV_TDHCFTS();
                model.TDV_TDHCFTS.ADV_UID = WebSiteHelper.getUserId();
                model.TDV_TDHCFTS.ADV_DP = WebSiteHelper.getUserDept();
            }

            // ==========================================================
            // 【修改重點】：取得計畫名稱並擷取年度
            // ==========================================================
            string projName = WebSiteHelper.GetCurrentProjectName(0, model.PID);
            string projYear = (DateTime.Now.Year - 1911).ToString(); // 預設值防呆

            if (!string.IsNullOrEmpty(projName))
            {
                // 尋找 "[" 和 "]" 的位置
                int startIndex = projName.IndexOf('[');
                int endIndex = projName.IndexOf(']');

                // 如果有找到括號，且格式正確
                if (startIndex != -1 && endIndex != -1 && endIndex > startIndex)
                {
                    // 擷取括號中間的字串
                    projYear = projName.Substring(startIndex + 1, endIndex - startIndex - 1);
                }
            }
            // 將擷取出來的年度放入 ViewBag
            ViewBag.ProjYear = projYear;

            //選單            
            GetSelectList(PID);
            //ViewBag.allap = SqlDao.CMPSGSQuery4List("SELECT AID, AP_NAME FROM APMFORM");
            ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY_ID", "CITY", null);
            ViewBag.NO1 = p.SYR;
            ViewBag.NO2 = p.SIP_W_CODE;
            ViewBag.NO3 = p.PJ_CODE;
            //ViewBag.NO4 = p.PJ_CODE;
            //ViewBag.NO5 = p.PJ_CODE;
            //ViewBag.NO6 = p.PJ_CODE;


            ViewBag.FIT_CODES = BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT CATY_COD as id, CODE_DESC as name  FROM SMCCODE c WHERE c.SYS_ID='FD' AND CATY_ID='05'");

            return View("PJE0030ManageAdd", model);
        }

        /// <summary>
        /// 引用列表查詢介面 
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult RefSelectFrame(PJE0030Model model)
        {
            ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);
            ViewBag.towns = tools.ToSelectList(util.GetTowns(null), "TOWN", "TOWN", null);
            return View("RefSelectFrame", model);
        }

        /// <summary>
        /// 轉入核定提案
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult RefBatchFrame(PJE0030Model model)
        {
            //string sql= "Select PPID, PJ_NAME from PJMFORM Where ADV_UID ='" + WebSiteHelper.getUserId() + "'";
            string sql = "Select PPID, PJ_NAME from PJMFORM";
            DataTable dt = BASE_DAO.SqlDao.CMPSGSQuery4List(sql);
            ViewBag.PPIDS = tools.ToSelectList(dt, "PPID", "PJ_NAME", null);
            sql = "Select YRC_ID, PSG_NAME from PJDSYRC Where PPID = " + model.PPID + " Order by SYR desc";
            dt = SqlDao.CMPSGSQuery4List(sql);
            ViewBag.PJDSYRC = tools.ToSelectList(dt, "YRC_ID", "PSG_NAME", null);

            PJV_PJMFORM p = db_pms.PJV_PJMFORM.Where(x => x.PID == model.PID).FirstOrDefault();
            model.NO1 = p.SYR;
            model.NO2 = p.SIP_W_CODE;
            model.NO3 = p.PJ_CODE;

            RefBatchWork(model);
            /*
            sql = @"select * 
                    from APV_TDMFORM 
                    where PPID=@PPID 
                    and YRC_ID=@YRC_ID
                    And AP_STAT='ZZ' 
                    and TD_CODE not in (select TD_CODE from CTDMNDS where PID=@PID)";


            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("PPID", model.PPID));
            cmd.Parameters.Add(new SqlParameter("YRC_ID", model.PJDSYRC));
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            dt = SqlDao.CMPSGSQuery4List(cmd);            
            ViewBag.apmforms = dt;
            */
            return View("RefBatchFrame", model);
        }

        private void RefBatchWork(PJE0030Model model)
        {
            GetSelectList(model.PID);
            DataTable citys = util.GetCities(); //util.GetCitiesNew();
            DataTable towns = GetTowns(model.CITY);
            ViewBag.cities = tools.ToSelectList(citys, "CITY", "CITY", null);//查詢要用名稱
            ViewBag.towns = tools.ToSelectList(towns, "TOWN", "TOWN", null);

            ViewBag.citiesTable = citys;
            ViewBag.townsTable = towns;


            ViewBag.pjcatycode = BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT CATY_CODE as id, ACC_ID+CATY_NAME as name FROM PJDCATY WHERE PID=" + model.PID);//分項計畫
            ViewBag.catycod = util.GetCMETMSCODES("TT", "03");//案件類型            
            ViewBag.ipjcode = util.GetCMETMSCODES("SM", "02");//經費層級
            ViewBag.tdcaty = util.GetCMETMSCODES("SM", "09");//工程類別            
            ViewBag.yrworks = util.GetCMETMSCODES("TT", "04");//工作類別
        }

        /// <summary>
        /// 取鄉鎮列表 - 新
        /// </summary>
        /// <param name="city"></param>
        /// <returns></returns>
        public static DataTable GetTowns(string city)
        {
            StringBuilder sql = new StringBuilder("SELECT * FROM SMCPOSTN WHERE TOWN<>''");
            SqlCommand cmd = new SqlCommand();

            if (!string.IsNullOrEmpty(city))
            {
                sql.Append(" AND CITY=@city");
                cmd.Parameters.Add(new SqlParameter("@city", city));
            }
            sql.Append(" ORDER BY CITY, CPA_TOWN_ID");

            cmd.CommandText = sql.ToString();
            DataTable dt = BASE_DAO.SqlDao.CMBASEQuery4List(cmd);

            return dt;
        }

        /// <summary>
        /// 轉入核定提案查詢
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "RefBatchSearch")]
        public ActionResult RefBatchSearch(PJE0030Model model)
        {

            string sql = "Select PPID, PJ_NAME from PJMFORM";
            DataTable dt = BASE_DAO.SqlDao.CMPSGSQuery4List(sql);
            ViewBag.PPIDS = tools.ToSelectList(dt, "PPID", "PJ_NAME", null);


            SqlCommand cmd = new SqlCommand("Select YRC_ID, PSG_NAME from PJDSYRC Where PPID = @PPID Order by SYR desc");
            cmd.Parameters.Add(new SqlParameter("PPID", model.PPID));
            dt = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            //dt = SqlDao.CMPSGSQuery4List("Select YRC_ID, PSG_NAME from PJDSYRC Where PPID = " + model.PPID + " Order by SYR desc");
            ViewBag.PJDSYRC = tools.ToSelectList(dt, "YRC_ID", "PSG_NAME", null);
            string search_ppid;
            if (model.PPID == 0)
            {
                search_ppid = "";
            }
            else
            {
                search_ppid = model.PPID.ToString();
            }

            int limit = Convert.ToInt32(model.Limit);
            /*
            sql = $@"select * from APV_TDMFORM 
                where PPID LIKE @PPID 
                and YRC_ID LIKE @YRC_ID
                and CITY LIKE @CITY
                and TOWN LIKE @TOWN
                And AP_STAT='90' 
                -- 將 NOT IN 替換為 NOT EXISTS
                and not exists (select 1 from CTDMNDS where PID=@PID and TD_CODE = APV_TDMFORM.TD_CODE)";
            */
            sql = @" select * 
                    from APV_TDMFORM 
                    where PPID=@PPID 
                     and YRC_ID LIKE @YRC_ID 
                        and CITY LIKE @CITY 
                        @strTOWN 
                        And AP_STAT='90'
                        and (TD_CODE IS NULL OR TD_CODE NOT IN (select TD_CODE from CTDMNDS where PID=@PID))
                        and convert(nvarchar(20), AID) not in (select isnull(DA_CODE,'') from  CTDMNDS where PID=@PID) 
                        order by CITY,TOWN ASC, TD_CODE DESC";

            if (!string.IsNullOrEmpty(model.TOWN))
                sql = sql.Replace("@strTOWN", "and TOWN LIKE @TOWN");
            else
                sql = sql.Replace("@strTOWN", "");
            cmd = new SqlCommand(sql);
            if (string.IsNullOrEmpty(model.CITY))
            {
                model.CITY = "";
            }
            cmd.Parameters.Add(new SqlParameter("PPID", search_ppid));
            cmd.Parameters.Add(new SqlParameter("YRC_ID", model.PJDSYRC + "%"));
            cmd.Parameters.Add(new SqlParameter("CITY", model.CITY + "%"));
            cmd.Parameters.Add(new SqlParameter("TOWN", model.TOWN + "%"));
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            dt = SqlDao.CMEPMSQuery4List(cmd);

            // 實現分頁邏輯
            int page = (model.prePage == 0) ? 1 : model.prePage;
            int pageSize = 10; // 每頁10筆

            List<DataRow> list = dt.AsEnumerable().ToList();
            IPagedList<DataRow> pagedList = new PagedList<DataRow>(list, page, pageSize);

            ViewBag.apmforms = pagedList;

            RefBatchWork(model);
            return View("RefBatchFrame", model);
        }

        /// <summary>
        /// 轉入核定提案查詢 - GET方法用於分頁
        /// </summary>
        /// <param name="prePage"></param>
        /// <param name="PPID"></param>
        /// <param name="PJDSYRC"></param>
        /// <param name="CITY"></param>
        /// <param name="TOWN"></param>
        /// <param name="Limit"></param>
        /// <param name="PID"></param>
        /// <param name="SYR"></param>
        /// <param name="NO1"></param>
        /// <param name="NO2"></param>
        /// <param name="NO3"></param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult RefBatchSearchPage(int? prePage, int? PPID, string PJDSYRC, string CITY, string TOWN, decimal? Limit, decimal? PID, string SYR, string NO1, string NO2, string NO3)
        {
            PJE0030Model model = new PJE0030Model();
            model.prePage = prePage ?? 0;
            model.PPID = PPID ?? 0;
            model.PJDSYRC = PJDSYRC ?? "";
            model.CITY = CITY ?? "";
            model.TOWN = TOWN ?? "";
            //model.Limit = Limit ?? 10;
            model.PID = PID ?? 0;
            model.SYR = SYR ?? "";
            model.NO1 = NO1;
            model.NO2 = NO2;
            model.NO3 = NO3;

            string sql = "Select PPID, PJ_NAME from PJMFORM";
            DataTable dt = BASE_DAO.SqlDao.CMPSGSQuery4List(sql);
            ViewBag.PPIDS = tools.ToSelectList(dt, "PPID", "PJ_NAME", null);

            SqlCommand cmd = new SqlCommand("Select YRC_ID, PSG_NAME from PJDSYRC Where PPID = @PPID Order by SYR desc");
            cmd.Parameters.Add(new SqlParameter("PPID", model.PPID));
            dt = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            ViewBag.PJDSYRC = tools.ToSelectList(dt, "YRC_ID", "PSG_NAME", null);

            string search_ppid;
            if (model.PPID == 0)
            {
                search_ppid = "";
            }
            else
            {
                search_ppid = model.PPID.ToString();
            }



            sql = @"select * 
                    from APV_TDMFORM 
                    where PPID=@PPID 
                     and YRC_ID LIKE @YRC_ID 
                        and CITY LIKE @CITY 
                        and TOWN LIKE @TOWN 
                        And AP_STAT='90'                        
                        and(TD_CODE IS NULL OR TD_CODE NOT IN(select TD_CODE from CTDMNDS where PID = @PID))
                        and convert(nvarchar(20), AID) not in (select isnull(DA_CODE, '') from CTDMNDS where PID = @PID) 
                        order by CITY,TOWN ASC, TD_CODE DESC";




            /*sql = $@"select * from APV_TDMFORM
                where PPID LIKE @PPID
                and YRC_ID LIKE @YRC_ID
                and CITY LIKE @CITY
                and TOWN LIKE @TOWN
                And AP_STAT='90'
                and not exists (select 1 from CTDMNDS where PID=@PID and TD_CODE = APV_TDMFORM.TD_CODE)";*/

            cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("PPID", search_ppid));
            cmd.Parameters.Add(new SqlParameter("YRC_ID", model.PJDSYRC + "%"));
            cmd.Parameters.Add(new SqlParameter("CITY", model.CITY + "%"));
            cmd.Parameters.Add(new SqlParameter("TOWN", model.TOWN + "%"));
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            dt = SqlDao.CMEPMSQuery4List(cmd);

            // 實現分頁邏輯
            int page = (model.prePage == 0) ? 1 : model.prePage;
            int pageSize = 10; // 每頁10筆

            List<DataRow> list = dt.AsEnumerable().ToList();
            IPagedList<DataRow> pagedList = new PagedList<DataRow>(list, page, pageSize);

            ViewBag.apmforms = pagedList;

            RefBatchWork(model);
            return View("RefBatchFrame", model);
        }

        /// <summary>
        /// 批次
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult RefBatchAddOne(string data)
        {
            JObject json = JObject.Parse(data);

            //TODO apitoken規範
            string PID = Convert.ToString(json["PID"]);
            string TD_CODE = Convert.ToString(json["TD_CODE"]);
            string FD_CODE = Convert.ToString(json["FD_CODE"]);
            string EXT_FG = Convert.ToString(json["EXT_FG"]);
            string AID = Convert.ToString(json["AID"]);
            SqlCommand cmd = new SqlCommand("select 1 from CTDMNDS where FD_CODE =@FD_CODE");
            cmd.Parameters.Add(new SqlParameter("FD_CODE", FD_CODE));

            DataTable t = SqlDao.CMEPMSQuery4List(cmd);
            if (t.Rows.Count > 0)
            {
                var jsonObject = new
                {
                    //id = newid,
                    msg = "經費編號已存在, 請更換經費編號",
                };
                return Json(jsonObject, JsonRequestBehavior.AllowGet);
            }

            // 驗證工程編號
            cmd = new SqlCommand("select 1 from CTDMNDS where PID=@PID AND TD_CODE =@TD_CODE ");
            cmd.Parameters.Add(new SqlParameter("PID", PID));
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            t = SqlDao.CMEPMSQuery4List(cmd);
            if (t.Rows.Count > 0)
            {
                var jsonObject = new
                {
                    //id = newid,
                    msg = "工程編號已存在, 請變更工程編號",
                };
                return Json(jsonObject, JsonRequestBehavior.AllowGet);
            }

            PJE0030Model model = new PJE0030Model();
            model.PID = Convert.ToDecimal(json["PID"]);
            model.TID = Convert.ToDecimal(json["TID"]);
            model.TDV_TDHCFTS = new TDV_TDHCFTS();
            model.TDV_TDHCFTS.TD_CODE = TD_CODE;
            model.TDV_TDHCFTS.FD_CODE = FD_CODE;
            model.AID = AID;
            model.TDV_TDHCFTS.CATY_CODE = Convert.ToString(json["CATY_CODE"]);

            model.TDV_TDHCFTS.TD_NAME = Convert.ToString(json["TD_NAME"]); //工程名稱
            model.TDV_TDHCFTS.PJ_CATY_CODE = Convert.ToString(json["PJ_CATY_CODE"]); ;//分項 
            model.TDV_TDHCFTS.TD_CATY = Convert.ToString(json["TD_CATY"]); //工程類別
            model.TDV_TDHCFTS.IPJ_CODE = Convert.ToString(json["IPJ_CODE"]);//辦理層級

            model.TDV_TDHCFTS.CITY = Convert.ToString(json["CITY"]);
            model.TDV_TDHCFTS.TOWN = Convert.ToString(json["TOWN"]);
            model.TDV_TDHCFTS.YR_WORK = Convert.ToString(json["YR_WORK"]);

            model.TDV_TDHCFTS.ADV_DP = Convert.ToString(json["ADV_DP"]);
            model.TDV_TDHCFTS.ADV_UID = Convert.ToString(json["ADV_UID"]);

            //經費
            model.CTDITFD_SYR = new string[2];

            model.CTDITFD_FD_CODE = new string[2];
            model.CTDITFD_FUND = new decimal[2];

            model.CTDITFD_SYR = new string[2];
            model.CTDITFD_FD_CODE = new string[2];
            model.CTDITFD_FUND = new decimal[2];

            model.CITY = Convert.ToString(json["CITY"]);
            model.TOWN = Convert.ToString(json["TOWN"]);

            // 直接從資料庫撈取對應的年度，取代json["YEAR"]
            var YEAR = db_pms.PJMFORM.FirstOrDefault(s => s.PID == model.PID)?.SYR;
            if (string.IsNullOrEmpty(YEAR))
            {
                var jsonObject = new
                {
                    //id = newid,
                    msg = "無法取得對應資料",
                };
                return Json(jsonObject, JsonRequestBehavior.AllowGet);
            }

            model.CTDITFD_SYR[0] = YEAR;// Convert.ToString(json["YEAR"]);
            model.CTDITFD_FD_CODE[0] = "00";
            try
            {
                model.CTDITFD_FUND[0] = Convert.ToDecimal(json["F00"]);
            }
            catch
            {
                model.CTDITFD_FUND[0] = 0;
            }


            model.CTDITFD_SYR[1] = YEAR;// Convert.ToString(json["YEAR"]);
            model.CTDITFD_FD_CODE[1] = "11";
            try
            {
                model.CTDITFD_FUND[1] = Convert.ToDecimal(json["F11"]);
            }
            catch
            {
                model.CTDITFD_FUND[1] = 0;
            }



            decimal newid = addCTDMNDS(model);


            if (newid == 0)
            {
                var jsonObject = new
                {
                    id = newid,
                    msg = "建立失敗",
                };
                return Json(jsonObject, JsonRequestBehavior.AllowGet);
            }
            else
            {
                var jsonObject = new
                {
                    id = newid,
                    msg = "建立完成",
                };
                return Json(jsonObject, JsonRequestBehavior.AllowGet);
            }
        }

        /// 引用列表 分頁
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "RefSelectSearch")]
        public ActionResult RefSelectSearch(PJE0030Model model)
        {

            StringBuilder sql = new StringBuilder("Select SIP_TD_CODE,TD_CODE,TD_NAME,TD_CITY,TD_TOWN From CMETMS.dbo.TDV_TDMFORM WHERE TD_CODE IS NOT NULL ");
            if (!string.IsNullOrEmpty(model.CITY))
            {
                sql.Append("AND TD_CITY='" + model.CITY + "'");
            }
            if (!string.IsNullOrEmpty(model.TOWN))
            {
                sql.Append("AND TD_TOWN='" + model.TOWN + "'");
            }
            if (!string.IsNullOrEmpty(model.TD_NAME))
            {
                sql.Append("AND TD_NAME LIKE'%" + model.TD_NAME + "%'");
            }

            if (!string.IsNullOrEmpty(model.WK_CODE))
            {
                sql.Append("AND TD_CODE ='" + model.WK_CODE + "'");
            }

            DataTable dt = BASE_DAO.SqlDao.CMPSGSQuery4List(sql.ToString());
            List<DataRow> list = dt.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, 1, 20);
            model.MainList = plist;
            return RefSelectFrame(model);
            //return View("RefSelectFrame");
        }

        /// <summary>
        /// 引用列表 查詢
        /// </summary>
        /// <param name="page"></param>
        /// <param name="city"></param>
        /// <param name="town"></param>
        /// <param name="name"></param>
        /// <param name="no"></param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult RefSelectSearch(int page, string city, string town, string name, string no)
        {
            StringBuilder sql = new StringBuilder("Select SIP_TD_CODE,TD_CODE,TD_NAME,TD_CITY,TD_TOWN From CMETMS.dbo.TDV_TDMFORM WHERE TD_CODE IS NOT NULL ");
            if (!string.IsNullOrEmpty(city))
            {
                sql.Append("AND TD_CITY='" + city + "'");
            }
            if (!string.IsNullOrEmpty(town))
            {
                sql.Append("AND TD_TOWN='" + town + "'");
            }
            if (!string.IsNullOrEmpty(name))
            {
                sql.Append("AND TD_NAME LIKE'%" + name + "%'");
            }

            if (!string.IsNullOrEmpty(no))
            {
                sql.Append("AND TD_CODE ='" + no + "'");
            }

            DataTable dt = BASE_DAO.SqlDao.CMPSGSQuery4List(sql.ToString());
            List<DataRow> list = dt.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, 20);

            PJE0030Model model = new PJE0030Model();
            model.CITY = city;
            model.TOWN = town;
            model.WK_CODE = no;
            model.TD_NAME = name;
            model.MainList = plist;



            return RefSelectFrame(model);
        }

        /// <summary>
        /// 取引用
        /// </summary>
        /// <param name="tdcode"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult RefSelectOne(string tdcode)
        {
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List("SELECT*FROM TDV_TDMFORM WHERE TD_CODE='" + tdcode + "'");
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 引用列表 json查詢 
        /// </summary>
        /// <param name="aid"></param>
        /// <returns></returns>
        /*public ActionResult GetCMETMSModel(string city, string town, string name, string no) {
            StringBuilder sql= new StringBuilder("Select SIP_TD_CODE,TD_CODE,TD_NAME,TD_CITY,TD_TOWN From CMETMS.dbo.TDV_TDMFORM WHERE TD_CODE IS NOT NULL ");
            if(!string.IsNullOrEmpty(city)) {
                sql.Append("AND TD_CITY='"+city+"'");
            }
            if(!string.IsNullOrEmpty(town)) {
                sql.Append("AND TD_TOWN='" + town + "'");
            }
            if(!string.IsNullOrEmpty(name)) {
                sql.Append("AND TD_NAME LIKE'%" + name + "%'");
            }

            if(!string.IsNullOrEmpty(no)) {
                sql.Append("AND TD_CODE ='" + no + "'");
            }
            DataTable dt = BASE_DAO.SqlDao.CMPSGSQuery4List(sql.ToString());            
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }*/

        /// 管制[CTDMNDS 進入版本 Frame]
        public ActionResult EditMnd(decimal PID, decimal TID, string Lock)
        {
            // ==========================================================
            // 【修改重點】：取得經費年度
            // ==========================================================

            string projYear = (DateTime.Now.Year - 1911).ToString(); // 預設值防呆
            var pjm = db_pms.PJMFORM.FirstOrDefault(f => f.PID == PID);
            projYear = pjm != null ? pjm.SYR : projYear; // 如果找不到計畫，則使用預設值
            // 將擷取出來的年度放入 ViewBag
            ViewBag.ProjYear = projYear;

            if (!string.IsNullOrEmpty(Lock))
            {
                ViewBag.LOCK = Lock;
            }
            PJE0030Model model = new PJE0030Model();
            int max = util.GetSEQ_NO(PID);
            //最新版本
            CTDMNDS m = db_pms.CTDMNDS.Where(f => f.TID == TID).FirstOrDefault();
            TDV_TDDCFTS t = db_pms.TDV_TDDCFTS.Where(f => f.TID == TID).FirstOrDefault();
            model.TDV_TDHCFTS = new TDV_TDHCFTS();


            //選單            
            GetSelectList(PID);
            model.TID = TID;
            model.PID = PID;
            //經費
            DataTable cts = BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT *     FROM CTDITFD c WHERE c.TID=" + TID + " ORDER BY c.FID DESC");
            foreach (DataRow row in cts.Rows)
            {
                // 遍歷每一欄去除小數
                foreach (DataColumn col in cts.Columns)
                {
                    object val = row[col];

                    if (val != DBNull.Value)
                    {
                        Type type = val.GetType();

                        // 處理 decimal、double、float 類型
                        if (type == typeof(decimal))
                        {
                            row[col] = Decimal.ToInt64((decimal)val);  // 去除小數
                        }
                        else if (type == typeof(double))
                        {
                            row[col] = Convert.ToInt64((double)val);
                        }
                        else if (type == typeof(float))
                        {
                            row[col] = Convert.ToInt64((float)val);
                        }
                    }
                }
            }
            ViewBag.CTDITFD = cts;
            ViewBag.FIT_CODES = BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT CATY_COD as id, CODE_DESC as name  FROM SMCCODE c WHERE c.SYS_ID='FD' AND CATY_ID='05'");

            ViewBag.CTDUPRC = BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT*FROM CTDUPRC c WHERE c.TID=" + TID);
            if (t == null)
            {
                //return Content("<xml>資料對應有誤,請刪除並重新建立</xml>", "text/xml");
                return View("PJE0030ManageEdit", model);
            }

            try
            {
                model.TDV_TDHCFTS.CEN_SP_RATIO_S = t.CEN_SP_RATIO_S;
            }
            catch
            {
                model.TDV_TDHCFTS.CEN_SP_RATIO_S = 0;
            }

            try
            {
                model.TDV_TDHCFTS.CEN_SP_RATIO = t.CEN_SP_RATIO;
            }
            catch
            {
                model.TDV_TDHCFTS.CEN_SP_RATIO = 0;
            }

            model.TDV_TDHCFTS.TID = TID;
            model.TDV_TDHCFTS.PID = t.PID;
            model.TDV_TDHCFTS.SIP_TD_CODE = t.SIP_TD_CODE;
            model.TDV_TDHCFTS.TD_CODE = t.TD_CODE;
            model.TDV_TDHCFTS.FD_CODE = t.FD_CODE;
            model.TDV_TDHCFTS.TD_NAME = t.TD_NAME;
            model.TDV_TDHCFTS.TD_CATY = t.TD_CATY;
            model.TDV_TDHCFTS.IPJ_CODE_DESC = t.IPJ_CODE_DESC;
            model.TDV_TDHCFTS.YR_FUND = t.YR_FUND;
            model.TDV_TDHCFTS.ACC_DESC = t.ACC_DESC;
            model.TDV_TDHCFTS.CATY_CODE = t.CATY_CODE;
            model.TDV_TDHCFTS.CEN_SP_RATIO = t.CEN_SP_RATIO;
            model.TDV_TDHCFTS.CEN_SP_RATIO_S = t.CEN_SP_RATIO_S;
            //model.TDV_TDHCFTS.CEN_SP_RATIO_U = t.CEN_SP_RATIO_U;
            model.TDV_TDHCFTS.NOCTRL_FG = t.NOCTRL_FG;
            model.TDV_TDHCFTS.FCAL_FG = t.FCAL_FG;
            model.TDV_TDHCFTS.UCNT_FG = t.UCNT_FG;
            //model.TDV_TDHCFTS.WORKTYPE = t.WORKTYPE;
            model.TDV_TDHCFTS.IPJ_CODE = t.IPJ_CODE;
            model.TDV_TDHCFTS.EXE_TYPE = t.EXE_TYPE;
            model.TDV_TDHCFTS.EMG_CATY = t.EMG_CATY;
            model.TDV_TDHCFTS.ADV_DP = t.ADV_DP;
            if (!string.IsNullOrEmpty(model.TDV_TDHCFTS.ADV_UID))
                model.TDV_TDHCFTS.ADV_UID = t.ADV_UID;
            model.TDV_TDHCFTS.UPD_DT = t.UPD_DT;
            model.TDV_TDHCFTS.UPD_UID = t.UPD_UID;
            model.TDV_TDHCFTS.PD_FG = t.PD_FG;
            model.TDV_TDHCFTS.CITY = t.CITY;
            model.TDV_TDHCFTS.TOWN = t.TOWN;
            model.TDV_TDHCFTS.M_01 = t.M_01;
            model.TDV_TDHCFTS.M_02 = t.M_02;
            model.TDV_TDHCFTS.M_03 = t.M_03;
            model.TDV_TDHCFTS.M_04 = t.M_04;
            model.TDV_TDHCFTS.M_05 = t.M_05;
            model.TDV_TDHCFTS.M_06 = t.M_06;
            model.TDV_TDHCFTS.M_07 = t.M_07;
            model.TDV_TDHCFTS.M_08 = t.M_08;
            model.TDV_TDHCFTS.M_09 = t.M_09;
            model.TDV_TDHCFTS.M_10 = t.M_10;
            model.TDV_TDHCFTS.M_11 = t.M_11;
            model.TDV_TDHCFTS.M_12 = t.M_12;
            model.TDV_TDHCFTS.TD_GROUP = t.TD_GROUP;
            model.TDV_TDHCFTS.B_GROUP = t.B_GROUP;
            model.TDV_TDHCFTS.F_GROUP = t.F_GROUP;
            model.TDV_TDHCFTS.PJ_CATY_CODE = t.PJ_CATY_CODE;
            model.TDV_TDHCFTS.TD_CATY = t.TD_CATY;
            model.TDV_TDHCFTS.YR_WORK = t.YR_WORK;



            return View("PJE0030ManageEdit", model);
        }

        /// <summary>
        /// 管制 - 新增CTDMNDS
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private decimal addCTDMNDS(PJE0030Model model)
        {
            DateTime now = DateTime.Now;

            string tdcode = model.TDV_TDHCFTS.TD_CODE.Replace("-", ""); ;
            //SqlDao.CMEPMSExecuteSql("DELETE FROM CTDMNDS WHERE TD_CODE='" + tdcode + "'");
            //SqlDao.CMETMSExecuteSql(new SqlCommand("DELETE FROM TDMFORM WHERE TD_CODE='" + tdcode + "'"));




            CTDMNDS c = new CTDMNDS();
            c.TD_CODE = tdcode;
            c.FD_CODE = model.TDV_TDHCFTS.FD_CODE;

            c.DA_CODE = model.AID;

            c.PID = model.PID;

            //c.TID = model.TID;


            c.TD_NAME = model.TDV_TDHCFTS.TD_NAME;//工程名稱
            c.CATY_CODE = model.TDV_TDHCFTS.PJ_CATY_CODE;//分項
            c.ACC_ID = util.getACC_ID(model.PID, c.CATY_CODE);
            c.ADD_DT = now;

            if (!string.IsNullOrEmpty(model.FCAL_FG))
            {
                c.FCAL_FG = true; //經費納入計算
            }
            else
            {
                c.FCAL_FG = false;
            }

            if (!string.IsNullOrEmpty(model.TD_GROUP))
            {
                c.TD_GROUP = true; //下水道永續營運組
            }
            else
            {
                c.TD_GROUP = false;
            }

            if (!string.IsNullOrEmpty(model.F_GROUP))
            {
                c.F_GROUP = true; //F分署
            }
            else
            {
                c.F_GROUP = false;
            }

            if (!string.IsNullOrEmpty(model.B_GROUP))
            {
                c.B_GROUP = true; //B分署
            }
            else
            {
                c.B_GROUP = false;
            }

            //CTDMNDS
            try
            {
                db_pms.CTDMNDS.Add(c);
                db_pms.SaveChanges();
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.StackTrace);
                return 0;
            }

            //經費
            if (model.CTDITFD_SYR != null)
            {
                for (int i = 0; i < model.CTDITFD_SYR.Length; i++)
                {
                    if (!string.IsNullOrEmpty(model.CTDITFD_SYR[i]))
                    {
                        CTDITFD fd = new CTDITFD();
                        fd = new CTDITFD();
                        if (model.CTDITFD_FD_CODE != null)
                        {
                            fd.FIT_CODE = model.CTDITFD_FD_CODE[i];
                        }

                        fd.SYR = model.CTDITFD_SYR[i];
                        fd.FUND = model.CTDITFD_FUND[i];
                        if (model.CTDITFD_FUND_DESC != null)
                        {
                            fd.FUND_DESC = model.CTDITFD_FUND_DESC[i];
                        }

                        fd.TID = c.TID;
                        fd.FID = i + 1;
                        fd.ADD_DT = DateTime.Now;
                        fd.UPD_UID = WebSiteHelper.getUserId();
                        db_pms.CTDITFD.Add(fd);
                    }
                }
                db_pms.SaveChanges();
            }



            //TDMFORM
            if (SqlDao.CMETMSQuery4List("SELECT 1 FROM TDMFORM WHERE TD_CODE='" + c.TD_CODE + "'").Rows.Count == 0)
            {

                Models.CMETMSModels.TDMFORM tdmform = new Models.CMETMSModels.TDMFORM();
                tdmform.TD_CODE = tdcode;
                tdmform.TD_NAME = model.TDV_TDHCFTS.TD_NAME;//工程名稱            

                tdmform.CATY_CODE = model.TDV_TDHCFTS.CATY_CODE;//案件類型
                tdmform.PPID = model.PPID;

                tdmform.TD_CATY = model.TDV_TDHCFTS.TD_CATY;//工程類別

                tdmform.TD_CITY = model.TDV_TDHCFTS.CITY;//CITY
                tdmform.TD_TOWN = model.TDV_TDHCFTS.TOWN;//TOWN

                tdmform.TD_CITY = model.CITY;
                tdmform.TD_TOWN = model.TOWN;

                tdmform.IPJ_CODE = model.TDV_TDHCFTS.IPJ_CODE;//辦理層級
                                                              //tdmform.CATY_CODE = c.CATY_CODE;
                tdmform.EMG_CATY = model.TDV_TDHCFTS.EMG_CATY;//工程型態(雨水用)

                tdmform.ADV_DP = model.TDV_TDHCFTS.ADV_DP;//單位
                                                          //tdmform.ADV_UID = model.TDV_TDHCFTS.ADV_UID;//人員

                tdmform.ADV_DP = model.TDV_TDHCFTS.ADV_DP;        //WebSiteHelper.getUserDept();
                tdmform.ADV_UID = model.TDV_TDHCFTS.ADV_UID;

                tdmform.ADD_DT = DateTime.Now;
                //tdmform.UPD_UID = WebSiteHelper.getUserId();



                if (!string.IsNullOrEmpty(model.UCNT_FG))
                {
                    tdmform.UCNT_FG = true; //不納執行件數
                }
                else
                {
                    tdmform.UCNT_FG = false;
                }
                tdmform.PD_FG = model.TDV_TDHCFTS.YR_WORK;
                if (!string.IsNullOrEmpty(model.TD_GROUP))
                {
                    tdmform.TD_GROUP = true; //下水道永續營運組
                }
                else
                {
                    tdmform.TD_GROUP = false;
                }

                if (!string.IsNullOrEmpty(model.EXE_TYPE))
                {
                    tdmform.EXE_TYPE = "02";//bot案
                }
                else
                {
                    tdmform.EXE_TYPE = "01";
                }

                if (!string.IsNullOrEmpty(model.TD_GROUP))
                {
                    tdmform.TD_GROUP = true; //下水道永續營運組
                }
                else
                {
                    tdmform.TD_GROUP = false;
                }

                try
                {
                    db_tms.TDMFORM.Add(tdmform);
                    db_tms.SaveChanges();
                }
                catch (Exception e)
                {
                    System.Diagnostics.Debug.WriteLine(e);
                    return c.TID;
                }

                //管制點
                util.addTDDSDWB(c, tdmform);


            }


            return c.TID;
        }

        /// <summary>
        /// 管制 - 新增CTDMNDS
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "CTDMNDS_ADD")]
        public ActionResult CTDMNDS_ADD(PJE0030Model model)
        {

            // ==========================================================
            // 1. 【防呆檢核】：檢查「工程編號」與「經費編號」是否重複
            // ==========================================================
            // 將前端傳入的工程編號去除 "-" 符號 (對齊您底層 addCTDMNDS 的處理邏輯)
            string newTdCode = model.TDV_TDHCFTS.TD_CODE.Replace("-", "");
            string newFdCode = model.TDV_TDHCFTS.FD_CODE;

            // 檢查工程編號是否重複 (因為是新增，所以不需像修改時排除自己的 TID)
            bool isTdCodeExist = db_pms.CTDMNDS.Any(x => x.PID == model.PID && x.TD_CODE == newTdCode);
            if (isTdCodeExist)
            {
                ViewBag.msg = "工程編號已存在，請重新輸入！";
                // 如果重複，直接退回原畫面，並帶入原本填寫的資料
                return AddMnd(model.PID, model.TID, model.SEQ_NO, model);
            }

            // 檢查經費編號是否重複
            bool isFdCodeExist = db_pms.CTDMNDS.Any(x => x.FD_CODE == newFdCode);
            if (isFdCodeExist)
            {
                ViewBag.msg = "經費編號已存在，請重新輸入！";
                return AddMnd(model.PID, model.TID, model.SEQ_NO, model);
            }

            // ==========================================================
            // 2. 檢核通過，開始執行新增邏輯
            // ==========================================================

            model.TID = addCTDMNDS(model);
            if (model.TID != 0)
            {
                ViewBag.msg = "已儲存";
            }
            else
            {
                ViewBag.msg = "儲存失敗";
                return AddMnd(model.PID, model.TID, model.SEQ_NO, model);
            }

            return EditMnd(model.PID, model.TID, "");
        }

        /// <summary>
        /// 經費修改記錄
        /// </summary>
        /// <param name="c"></param>
        private void SaveCTDUPRC(CTDITFD c)
        {

            StringBuilder cont = new StringBuilder();
            string from = "";
            switch (c.FIT_CODE)
            {
                case "00":
                    from = c.SYR + "年單位年度預算(中央款):";
                    break;
                case "03":
                    from = c.SYR + "年單位特別預算:";
                    break;
                case "10":
                    from = c.SYR + "年中央部會補助配合款:";
                    break;
                case "11":
                    from = c.SYR + "年縣市自籌配合款(地方款):";
                    break;
                default:
                    from = c.SYR + "中央款:";
                    break;
            }
            from += c.FUND;
            CTDUPRC u = new CTDUPRC();

            u.UPD_CONT = from;
            u.FUND_DESC = c.FUND_DESC;

            u.ADD_DT = DateTime.Now;
            u.UPD_USER = WebSiteHelper.getUserName();
            u.UPD_UID = WebSiteHelper.getUserId();
            u.TID = c.TID;
            db_pms.CTDUPRC.Add(u);
            db_pms.SaveChanges();
        }

        /// <summary>
        /// 新增經費資料
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "CTDITFD_ADD")]
        public ActionResult CTDITFD_ADD(PJE0030Model model)
        {
            decimal fid = 0;
            CTDITFD lastone;
            try
            {
                lastone = db_pms.CTDITFD.Where(w => w.TID == model.TID).OrderByDescending(p => p.FID).FirstOrDefault();
                fid = lastone.FID + 1;
            }
            catch
            {
                fid = 1;
            }

            if (!checkSum(model))
            {
                ViewBag.msg = "金額不符，儲存失敗";
                ViewBag.save2 = true;
                return EditMnd(model.PID, model.TID, "");
            }

            if (model.CTDITFD_FUND[0] != 0)
            {

                CTDITFD fd = new CTDITFD();
                fd.FIT_CODE = model.CTDITFD_FD_CODE[0];
                fd.SYR = model.CTDITFD_SYR[0];
                fd.FUND = model.CTDITFD_FUND[0];
                fd.FUND_DESC = model.CTDITFD_FUND_DESC[0];
                fd.TID = model.TID;
                fd.FID = fid;
                fd.ADD_DT = DateTime.Now;
                fd.UPD_UID = WebSiteHelper.getUserId();
                db_pms.CTDITFD.Add(fd);
                db_pms.SaveChanges();
                SaveCTDUPRC(fd);//經費修改記錄
                ViewBag.msg = "已儲存";
            }
            else
            {
                ViewBag.msg = "儲存失敗";
            }


            ViewBag.save2 = true;

            return EditMnd(model.PID, model.TID, "");
        }

        /// <summary>
        /// 編輯經費資料
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "CTDITFD_SAVE")]
        public ActionResult CTDITFD_SAVE(PJE0030Model model)
        {

            if (!checkSum(model))
            {
                ViewBag.msg = "金額不符，儲存失敗";
                ViewBag.save2 = true;
                return EditMnd(model.PID, model.TID, "");
            }

            decimal tid, fid;
            for (int i = 0; i < model.CTDITFD_TID.Length; i++)
            {
                if (model.CTDITFD_TID[i] != 0 && model.CTDITFD_FID[i] != 0)
                {
                    tid = model.CTDITFD_TID[i];
                    fid = model.CTDITFD_FID[i];
                    CTDITFD fd = db_pms.CTDITFD.Where(f => f.TID == tid && f.FID == fid).FirstOrDefault();
                    fd.FIT_CODE = model.CTDITFD_FD_CODE[i];
                    fd.SYR = model.CTDITFD_SYR[i];
                    fd.FUND = model.CTDITFD_FUND[i];
                    fd.FUND_DESC = model.CTDITFD_FUND_DESC[i];
                    //fd.TID = c.TID;
                    fd.ADD_DT = DateTime.Now;
                    fd.UPD_UID = WebSiteHelper.getUserId();
                    db_pms.SaveChanges();
                    SaveCTDUPRC(fd);//經費修改記錄
                    break;
                }
            }
            ViewBag.save2 = true;
            ViewBag.msg = "已儲存";
            return EditMnd(model.PID, model.TID, "");
        }

        /// <summary>
        /// 刪除經費資料
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "CTDITFD_DEL")]
        public ActionResult CTDITFD_DEL(PJE0030Model model)
        {

            if (!checkSum(model))
            {
                ViewBag.msg = "金額不符，儲存失敗";
                ViewBag.save2 = true;
                return EditMnd(model.PID, model.TID, "");
            }



            CTDITFD fd = db_pms.CTDITFD.Where(f => f.TID == model.DelTid && f.FID == model.DelFid).FirstOrDefault();
            fd.FUND_DESC = "刪除";
            fd.ADD_DT = DateTime.Now;
            fd.UPD_UID = WebSiteHelper.getUserId();
            db_pms.SaveChanges();
            SaveCTDUPRC(fd);//經費修改記錄


            db_pms.CTDITFD.Remove(fd);
            db_pms.SaveChanges();


            ViewBag.save2 = true;
            ViewBag.msg = "已刪除";
            return EditMnd(model.PID, model.TID, "");
        }

        private Boolean checkSum(PJE0030Model model)
        {


            decimal total = 0;
            decimal sum = 0;
            for (int i = 0; i < model.CTDITFD_FUND.Length; i++)
            {
                total += model.CTDITFD_FUND[i];
            }

            if (total > 0)
            {
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_01);
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_02);
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_03);
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_04);
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_05);
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_06);
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_07);
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_08);
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_09);
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_10);
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_11);
                sum += tools.safe2Decimal(model.TDV_TDHCFTS.M_12);
                if (sum > 0)
                {
                    if (sum != total)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// 改為 月分配，不動經費
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "CTDMNDS_FUND_SAVE")]
        public ActionResult CTDMNDS_FUND_SAVE(PJE0030Model model)
        {

            if (!checkSum(model))
            {
                ViewBag.msg = "金額不符，儲存失敗";
                ViewBag.save2 = true;
                return EditMnd(model.PID, model.TID, "");
            }
            DateTime now = DateTime.Now;
            CTDMNDS c;

            if (model.TID == 0)
            {
                c = new CTDMNDS();
                c.PID = model.PID;
            }
            else
            {
                c = db_pms.CTDMNDS.Where(f => f.TID == model.TID).FirstOrDefault();
            }
            c.TD_NAME = model.TDV_TDHCFTS.TD_NAME;//工程名稱
            c.CATY_CODE = model.TDV_TDHCFTS.PJ_CATY_CODE;//分項

            c.M_01 = model.TDV_TDHCFTS.M_01;
            c.M_02 = model.TDV_TDHCFTS.M_02;
            c.M_03 = model.TDV_TDHCFTS.M_03;
            c.M_04 = model.TDV_TDHCFTS.M_04;
            c.M_05 = model.TDV_TDHCFTS.M_05;
            c.M_06 = model.TDV_TDHCFTS.M_06;
            c.M_07 = model.TDV_TDHCFTS.M_07;
            c.M_08 = model.TDV_TDHCFTS.M_08;
            c.M_09 = model.TDV_TDHCFTS.M_09;
            c.M_10 = model.TDV_TDHCFTS.M_10;
            c.M_11 = model.TDV_TDHCFTS.M_11;
            c.M_12 = model.TDV_TDHCFTS.M_12;
            db_pms.SaveChanges();
            /*
            decimal tid, fid = 0;
            //經費            
            if(model.CTDITFD_TID != null)
                for(int i = 0; i < model.CTDITFD_TID.Length; i++) {
                    if(model.CTDITFD_TID[i] > 0) {
                        tid = model.CTDITFD_TID[i];
                        fid = model.CTDITFD_FID[i];
                        //sql =new StringBuilder("UPDATE CTDITFD SET SYR='"+ model.CTDITFD_SYR[i] + "'");                    
                        //sql.Append("WHERE TID='"+tid+"'AND fid='"+fid+"'");

                        CTDITFD fd = db_pms.CTDITFD.Where(f => f.TID == tid && f.FID == fid).FirstOrDefault();
                        fd.FIT_CODE = model.CTDITFD_FD_CODE[i];
                        fd.SYR = model.CTDITFD_SYR[i];
                        fd.FUND = model.CTDITFD_FUND[i];
                        fd.FUND_DESC = model.CTDITFD_FUND_DESC[i];
                        //fd.TID = c.TID;
                        fd.ADD_DT = DateTime.Now;
                        fd.UPD_UID = WebSiteHelper.getUserId();
                        db_pms.SaveChanges();
                    } else {
                        if(!string.IsNullOrEmpty(model.CTDITFD_SYR[i])) {
                            CTDITFD fd = new CTDITFD();
                            fd.TID = model.TID;
                            if(i == 0) {
                                fd.FID = 0;
                            } else {
                                fd.FID = (fid + 1);
                            }

                            fd.FIT_CODE = model.CTDITFD_FD_CODE[i];
                            fd.SYR = model.CTDITFD_SYR[i];
                            fd.FUND = model.CTDITFD_FUND[i];
                            fd.FUND_DESC = model.CTDITFD_FUND_DESC[i];
                            //fd.TID = c.TID;
                            fd.ADD_DT = DateTime.Now;
                            fd.UPD_UID = WebSiteHelper.getUserId();
                            db_pms.CTDITFD.Add(fd);
                            db_pms.SaveChanges();
                        }
                    }
                }
                */
            ViewBag.save2 = true;
            ViewBag.msg = "已儲存";
            return EditMnd(model.PID, model.TID, "");
        }

        /// <summary>
        /// 管制 [儲存CTDMNDS]
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "CTDMNDS_SAVE")]
        public ActionResult CTDMNDS_SAVE(PJE0030Model model)
        {

            // 1. 【防呆檢核】：檢查「工程編號」是否與其他案件重複
            string newTdCode = model.TDV_TDHCFTS.TD_CODE;
            // 條件：TD_CODE 相同，且 TID 不等於當前案件 (代表是別人的案子)
            bool isTdCodeExist = db_pms.CTDMNDS.Any(x => x.TD_CODE == newTdCode && x.TID != model.TID && x.PID == model.PID);
            if (isTdCodeExist)
            {
                ViewBag.msg = "儲存失敗：工程編號 [" + newTdCode + "] 已被其他案件使用，請更換。";
                return EditMnd(model.PID, model.TID, "");
            }

            // 2. 【防呆檢核】：檢查「經費編號」是否與其他案件重複
            string newFdCode = model.TDV_TDHCFTS.FD_CODE;
            bool isFdCodeExist = db_pms.CTDMNDS.Any(x => x.FD_CODE == newFdCode && x.TID != model.TID && x.PID == model.PID);
            if (isFdCodeExist)
            {
                ViewBag.msg = "儲存失敗：經費編號 [" + newFdCode + "] 已被其他案件使用，請更換。";
                return EditMnd(model.PID, model.TID, "");
            }

            // 3. 【安全取得舊資料】：為了判斷 CATY_CODE 是否有變更
            string old_caty_code = "";

            // 先用 TID 找出原本在 CTDMNDS 的舊工程編號
            var oldCtdmnds = db_pms.CTDMNDS.FirstOrDefault(c => c.TID == model.TID);
            if (oldCtdmnds != null)
            {
                // 再用舊的工程編號去 TDMFORM 找出原本的案件類型
                var oldTdmform = db_tms.TDMFORM.FirstOrDefault(t => t.TD_CODE == oldCtdmnds.TD_CODE);
                if (oldTdmform != null)
                {
                    old_caty_code = oldTdmform.CATY_CODE;
                }
            }

            // 4. 呼叫底層儲存邏輯 
            model.TID = saveCTDMNDS(model);

            // 5. 儲存後的狀態判斷
            if (model.TID != 0)
            {
                // 加上 old_caty_code != null 的防呆，避免原資料為空時出錯
                if (old_caty_code != null && old_caty_code.Equals(model.TDV_TDHCFTS.CATY_CODE))
                {
                    ViewBag.msg = "已儲存";
                }
                else
                {
                    // 有變才修
                    util.addTDDSDWB(model.TDV_TDHCFTS.TD_CODE, model.TDV_TDHCFTS.CATY_CODE);
                    ViewBag.msg = "已儲存「請注意：若變更案件類型，工程登記冊之要徑填報資料需重新填寫」";
                }
            }
            else
            {
                ViewBag.msg = "儲存失敗";
                //return AddMnd(model.PID, model.TID, model.SEQ_NO);
            }

            return EditMnd(model.PID, model.TID, "");
        }

        /// <summary>
        /// 共用存 CTDMNDS 基本資料
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private decimal saveCTDMNDS(PJE0030Model model)
        {
            DateTime now = DateTime.Now;
            CTDMNDS c;

            string old_td_code = "";
            string old_fd_code = "";
            bool isNewCtdmnds = (model.TID == 0);
            bool isNewTdmform = false;

            // ==========================================================
            // 1. 處理 CTDMNDS
            // ==========================================================
            if (isNewCtdmnds)
            {
                c = new CTDMNDS();
                c.PID = model.PID;
                // 【修改重點 1】：新增時，可以直接給予 Key 值
                c.TD_CODE = model.TDV_TDHCFTS.TD_CODE;
                c.FD_CODE = model.TDV_TDHCFTS.FD_CODE;
                c.ADD_DT = now;
                db_pms.CTDMNDS.Add(c); // 若您的 EF 版本不用手動 Add 可忽略
            }
            else
            {
                c = db_pms.CTDMNDS.Where(f => f.TID == model.TID).FirstOrDefault();
                if (c != null)
                {
                    // 記錄舊編號，供後續 SQL Update 使用
                    old_td_code = c.TD_CODE;
                    old_fd_code = c.FD_CODE;
                }
                // 【修改重點 2】：嚴禁在這裡寫 c.TD_CODE = ... (會觸發您看到的錯誤)
            }

            c.TD_NAME = model.TDV_TDHCFTS.TD_NAME;
            c.M_01 = model.TDV_TDHCFTS.M_01;
            c.M_02 = model.TDV_TDHCFTS.M_02;
            c.M_03 = model.TDV_TDHCFTS.M_03;
            c.M_04 = model.TDV_TDHCFTS.M_04;
            c.M_05 = model.TDV_TDHCFTS.M_05;
            c.M_06 = model.TDV_TDHCFTS.M_06;
            c.M_07 = model.TDV_TDHCFTS.M_07;
            c.M_08 = model.TDV_TDHCFTS.M_08;
            c.M_09 = model.TDV_TDHCFTS.M_09;
            c.M_10 = model.TDV_TDHCFTS.M_10;
            c.M_11 = model.TDV_TDHCFTS.M_11;
            c.M_12 = model.TDV_TDHCFTS.M_12;
            c.UPD_UID = WebSiteHelper.getUserId();
            c.UPD_DT = now;

            c.FCAL_FG = !string.IsNullOrEmpty(model.FCAL_FG);
            c.TD_GROUP = !string.IsNullOrEmpty(model.TD_GROUP);
            c.F_GROUP = !string.IsNullOrEmpty(model.F_GROUP);
            c.B_GROUP = !string.IsNullOrEmpty(model.B_GROUP);

            // ==========================================================
            // 2. 處理 TDMFORM
            // ==========================================================
            Models.CMETMSModels.TDMFORM tdmform = null;
            if (!string.IsNullOrEmpty(old_td_code))
            {
                tdmform = db_tms.TDMFORM.FirstOrDefault(t => t.TD_CODE == old_td_code);
            }

            if (tdmform == null)
            {
                tdmform = db_tms.TDMFORM.FirstOrDefault(t => t.TD_CODE == model.TDV_TDHCFTS.TD_CODE);
                if (tdmform == null)
                {
                    isNewTdmform = true;
                    tdmform = new Models.CMETMSModels.TDMFORM();
                    // 【修改重點 3】：新增時，可以直接給予 Key 值
                    tdmform.TD_CODE = model.TDV_TDHCFTS.TD_CODE;
                    tdmform.UPD_DT = DateTime.Now;
                    tdmform.UPD_UID = WebSiteHelper.getUserId();
                    db_tms.TDMFORM.Add(tdmform);
                }
            }
            // 【修改重點 4】：嚴禁在這裡寫 tdmform.TD_CODE = ... (會觸發錯誤)

            tdmform.PD_FG = model.TDV_TDHCFTS.PD_FG;
            tdmform.TD_NAME = model.TDV_TDHCFTS.TD_NAME;

            if (tdmform.CATY_CODE != null && !tdmform.CATY_CODE.Equals(model.TDV_TDHCFTS.CATY_CODE))
            {
                util.updateTDDSDWB(c, model.TDV_TDHCFTS.CATY_CODE);
            }
            tdmform.CATY_CODE = model.TDV_TDHCFTS.CATY_CODE;
            tdmform.TD_CATY = model.TDV_TDHCFTS.TD_CATY;
            tdmform.IPJ_CODE = model.TDV_TDHCFTS.IPJ_CODE;
            tdmform.EMG_CATY = model.TDV_TDHCFTS.EMG_CATY;
            tdmform.ADV_DP = model.TDV_TDHCFTS.ADV_DP;
            tdmform.UPD_DT = DateTime.Now;
            tdmform.UPD_UID = WebSiteHelper.getUserId();

            tdmform.UCNT_FG = !string.IsNullOrEmpty(model.UCNT_FG);
            tdmform.EXE_TYPE = !string.IsNullOrEmpty(model.EXE_TYPE) ? "02" : "01";

            // ==========================================================
            // 3. 儲存與 SQL 強制更新 Key 值
            // ==========================================================
            try
            {
                // 先讓 EF 儲存一般屬性的變更
                db_pms.SaveChanges();
                db_tms.SaveChanges();

                // 【修改重點 5】：若為修改模式，且使用者真的改了編號，用 Raw SQL 強制變更 Key 欄位
                if (!isNewCtdmnds)
                {
                    string newTdCode = model.TDV_TDHCFTS.TD_CODE;
                    string newFdCode = model.TDV_TDHCFTS.FD_CODE;

                    if (old_td_code != newTdCode || old_fd_code != newFdCode)
                    {
                        // 變更 CTDMNDS 主鍵
                        db_pms.Database.ExecuteSqlCommand(
                            "UPDATE CTDMNDS SET TD_CODE = @p0, FD_CODE = @p1 WHERE TID = @p2",
                            newTdCode, newFdCode, model.TID);
                    }

                    if (!isNewTdmform && old_td_code != newTdCode)
                    {
                        // 變更 TDMFORM 主鍵
                        db_tms.Database.ExecuteSqlCommand(
                            "UPDATE TDMFORM SET TD_CODE = @p0 WHERE TD_CODE = @p1",
                            newTdCode, old_td_code);
                    }
                }

            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
                return 0;
            }

            return c.TID;
        }

        /// <summary>
        /// 預算統計
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "BudgetCount")]
        public ActionResult BudgetCount(PJE0030Model model)
        {
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)

            if (model.SEQ_NO == 0)
            {
                model.SEQ_NO = util.GetSEQ_NO(model.PID);
            }


            GetSelectList(model.PID);

            DataTable xline;
            DataTable yline;
            GetBudgetCountTables(model.PID, out xline, out yline);

            ViewBag.xline = xline;
            ViewBag.yline = yline;

            return View("PJE0030BudgetCount", model);
        }

        private void GetBudgetCountTables(decimal PID, out DataTable xline, out DataTable yline)
        {
            SqlCommand cmd = new SqlCommand("select CATY_CODE, CATY_NAME from PJDCATY where PID=@PID");
            cmd.Parameters.Add(new SqlParameter("PID", PID));
            xline = SqlDao.CMEPMSQuery4List(cmd);

            cmd = new SqlCommand("select CITY, count(*)as cnt from TDV_TDDCFTS where CITY IS NOT NULL AND PID=@PID group by CITY order by CITY");
            cmd.Parameters.Add(new SqlParameter("PID", PID));
            yline = SqlDao.CMEPMSQuery4List(cmd);
            yline.Columns.Remove("cnt");

            foreach (DataRow x in xline.Rows)
            {
                //yline.Columns.Add(x["CATY_NAME"]+"_NAME");//名
                yline.Columns.Add(x["CATY_CODE"] + "_QUE");//數量
                yline.Columns.Add(x["CATY_CODE"] + "_AMO");//金額
            }


            cmd = new SqlCommand("select CATY_NAME, CITY, COUNT(*)as QUE, FORMAT(SUM(YR_FUND),'###,###.##')as AMO from TDV_TDDCFTS where PID=@PID GROUP BY CITY, CATY_NAME ORDER BY CATY_NAME, city");
            cmd.Parameters.Add(new SqlParameter("PID", PID));
            DataTable myTDDCFTS = SqlDao.CMEPMSQuery4List(cmd);

            foreach (DataRow y in yline.Rows)
            {
                foreach (DataRow x in xline.Rows)
                {
                    try
                    {
                        //y[x["CATY_NAME"] + "_NAME"] = x["CATY_NAME"];

                        foreach (DataRow m in myTDDCFTS.Rows)
                        {
                            if (Convert.ToString(m["CITY"]).Equals(Convert.ToString(y["CITY"])) &&
                                Convert.ToString(m["CATY_NAME"]).Equals(Convert.ToString(x["CATY_NAME"]))

                                )
                            {

                                y[x["CATY_CODE"] + "_QUE"] = m["QUE"];
                                y[x["CATY_CODE"] + "_AMO"] = m["AMO"];
                            }
                        }

                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Debug.WriteLine(e.Message);
                    }
                }
            }



            decimal sum;
            DataRow d = yline.NewRow();
            for (int i = 1; i <= yline.Columns.Count - 1; i++)
            {
                sum = 0;
                foreach (DataRow y in yline.Rows)
                {
                    try
                    {
                        sum += tools.safe2Decimal(y[i]);
                    }
                    catch
                    {
                        sum += 0;
                    }
                }

                d[i] = sum;
            }
            d[0] = "總計";
            yline.Rows.Add(d);
        }

        /// <summary>
        /// 預算統計匯出
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ExportBudgetCount")]
        public ActionResult ExportBudgetCount(PJE0030Model model)
        {
            DataTable xline;
            DataTable yline;
            GetBudgetCountTables(model.PID, out xline, out yline);

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "預算統計";

            int columnCount = 1 + (xline.Rows.Count * 2);
            int column = 1;
            sheet.Cells[0, 0].PutValue("");
            sheet.Cells[1, 0].PutValue("縣市別");

            foreach (DataRow x in xline.Rows)
            {
                sheet.Cells[0, column].PutValue(Convert.ToString(x["CATY_NAME"]));
                sheet.Cells.Merge(0, column, 1, 2);
                sheet.Cells[1, column].PutValue("件數");
                sheet.Cells[1, column + 1].PutValue("金額");
                column += 2;
            }

            int row = 2;
            foreach (DataRow y in yline.Rows)
            {
                for (int i = 0; i < yline.Columns.Count; i++)
                {
                    PutBudgetCountCell(sheet, row, i, y[i]);
                }

                row++;
            }

            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.VerticalAlignment = TextAlignmentType.Center;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            SetThinBorders(headerStyle);

            StyleFlag headerFlag = new StyleFlag();
            headerFlag.FontBold = true;
            headerFlag.HorizontalAlignment = true;
            headerFlag.VerticalAlignment = true;
            headerFlag.CellShading = true;
            headerFlag.Borders = true;
            sheet.Cells.CreateRange(0, 0, 2, columnCount).ApplyStyle(headerStyle, headerFlag);

            Style borderStyle = workbook.CreateStyle();
            SetThinBorders(borderStyle);
            StyleFlag borderFlag = new StyleFlag();
            borderFlag.Borders = true;
            sheet.Cells.CreateRange(0, 0, row, columnCount).ApplyStyle(borderStyle, borderFlag);

            sheet.Cells.SetColumnWidth(0, 16);
            for (int i = 1; i < columnCount; i++)
            {
                sheet.Cells.SetColumnWidth(i, (i % 2 == 0) ? 18 : 10);
            }

            sheet.AutoFitRows();
            sheet.PageSetup.Zoom = 80;

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"PJE0030_預算統計_{timestamp}.xlsx";

            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                ms.Position = 0;
                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        private void PutBudgetCountCell(Worksheet sheet, int row, int column, object value)
        {
            if (value == null || value == DBNull.Value)
            {
                sheet.Cells[row, column].PutValue("");
                return;
            }

            if (column == 0)
            {
                sheet.Cells[row, column].PutValue(Convert.ToString(value));
                return;
            }

            decimal number;
            string text = Convert.ToString(value);
            if (decimal.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out number) ||
                decimal.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out number))
            {
                sheet.Cells[row, column].PutValue(number);
            }
            else
            {
                sheet.Cells[row, column].PutValue(text);
            }
        }

        private void SetThinBorders(Style style)
        {
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
        }

        /// <summary>
        /// [預算調整]
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "BudgetEdit")]
        public ActionResult BudgetEdit(PJE0030Model model)
        {
            //／　取得目前所在的主計畫名稱(+分支計畫名稱)
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.PID);

            DataTable list;

            try
            {
                DataTable catys = SqlDao.CMEPMSQuery4List(" SELECT CATY_CODE FROM PJDCATY WHERE PID = " + model.PID);

                StringBuilder sb = new StringBuilder(" exec SP_PJE0030_PJDFAD " + model.PID + " ,'','");

                foreach (DataRow d in catys.Rows)
                {
                    sb.Append(d["CATY_CODE"] + ",");
                }

                sb.Length--;
                sb.Append("'");

                list = SqlDao.CMEPMSQuery4List(sb.ToString());
            }
            catch (Exception e)
            {
                ViewBag.msg = "預存程序SP_PJE0030_PJDFAD: " + e.Message;

                return Index(model);
            }

            DataTable core = list.Select("AN_CODE IS null").CopyToDataTable();

            core.Columns.Add("prj", typeof(DataTable));

            DataTable citi;
            DataTable prj;

            string keycode;

            foreach (DataRow dr in core.Rows)
            {
                /// 上層
                try
                {
                    prj = list.Select("AN_CODE='" + dr["KEYCODE"] + "'").CopyToDataTable();
                    prj.Columns.Add("citi", typeof(DataTable));
                    dr["prj"] = prj;
                }
                catch
                {
                    prj = list.Copy();
                    prj.Columns.Add("citi", typeof(DataTable));
                    dr["prj"] = prj;
                }

                foreach (DataRow p in prj.Rows)
                {
                    try
                    {
                        keycode = p["KEYCODE"].ToString();

                        citi = list.Select("AN_CODE='" + p["KEYCODE"] + "'AND REC_TP='D'").CopyToDataTable();

                        p["citi"] = citi;
                    }
                    catch (Exception e)
                    {
                        p["citi"] = list.Clone();

                        System.Diagnostics.Debug.WriteLine(e);
                    }
                }
            }

            ViewBag.core = core;

            return View("PJE0030BudgetEdit", model);
        }

        [HttpPost]
        public JsonResult GetTDDCFTS(string PID, string KEYCODE)
        {
            DataTable catys = SqlDao.CMEPMSQuery4List("SELECT CATY_CODE FROM PJDCATY WHERE PID=" + PID);

            StringBuilder sb = new StringBuilder("exec SP_PJE0030_PJDFAD " + PID + " ,'','");

            foreach (DataRow d in catys.Rows)
            {
                sb.Append(d["CATY_CODE"] + ",");
            }

            sb.Length--;
            sb.Append("'");

            DataTable list = SqlDao.CMEPMSQuery4List(sb.ToString());

            DataRow dr = list.Select("KEYCODE='" + KEYCODE + "'")[0];

            string where = dr["WHERE_COND"].ToString();

            string sql = "exec SP_PJE0030_PJDTDD " + PID + " ,\"" + where + "\"";

            DataTable dt = SqlDao.CMEPMSQuery4List(sql);
            dt = tools.ResetNull(dt);

            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 總經費 / 中央款 / 地方款
        /// </summary>
        /// <param name="pid">分支計畫 PID</param>
        /// <param name="syr">年度</param>
        /// <param name="seqNo">修正期別</param>
        /// <returns>
        /// DataTable 欄位：
        /// - YR_FUND：總經費 = 中央款 + 地方款
        /// - CEN_FUND：中央款
        /// - LOC_FUND：地方款
        /// </returns>
        private DataTable GetProjectFundSummary(object pid, string syr, int seqNo, string city = null, string CATY_CODE = null)
        {
            Debug.WriteLine("GetProjectFundSummary");
            Debug.WriteLine($"PID = {pid}");
            Debug.WriteLine($"SYR = {syr}");
            Debug.WriteLine($"SEQ = {seqNo}");
            Debug.WriteLine($"CITY = {city}");

            if (string.IsNullOrWhiteSpace(syr))
            {
                decimal projectPid = Convert.ToDecimal(pid);

                PJMFORM_CMEPMS project = db_pms.PJMFORM.Where(w => w.PID == projectPid).FirstOrDefault();

                if (project != null)
                {
                    syr = project.SYR;
                }
            }

            SqlCommand cmd = new SqlCommand();

            string sql = $@"
                    SELECT
                        ISNULL(SUM(ISNULL(C.FUND, 0) + ISNULL(B.FUND, 0)), 0) AS YR_FUND,
                        ISNULL(SUM(C.FUND), 0) AS CEN_FUND,
                        ISNULL(SUM(B.FUND), 0) AS LOC_FUND
                    FROM CTDMNDS A
                    LEFT JOIN CTDITFD B ON B.TID = A.TID AND B.FIT_CODE = '11' AND B.SYR = @SYR
                    LEFT JOIN CTDITFD C ON C.TID = A.TID AND C.FIT_CODE = '00' AND C.SYR = @SYR
                    WHERE A.PID = @PID ";

            if (!string.IsNullOrWhiteSpace(city))
            {
                sql += @" AND EXISTS ( SELECT 1 FROM TDV_TDDCFTS T WHERE T.PID = A.PID AND T.TID = A.TID AND T.CITY = @CITY )";
            }


            if (!string.IsNullOrWhiteSpace(CATY_CODE))
            {
                sql += @" AND CATY_CODE=@CATY_CODE";
            }

            cmd.CommandText = sql;
            cmd.Parameters.Add(new SqlParameter("PID", pid));
            cmd.Parameters.Add(new SqlParameter("SYR", (object)syr ?? DBNull.Value));
            cmd.Parameters.Add(new SqlParameter("SEQ_NO", seqNo));
            cmd.Parameters.Add(new SqlParameter("CITY", string.IsNullOrWhiteSpace(city) ? (object)DBNull.Value : city.Trim()));
            cmd.Parameters.Add(new SqlParameter("CATY_CODE", string.IsNullOrWhiteSpace(CATY_CODE) ? (object)DBNull.Value : CATY_CODE.Trim()));

            DataTable result = SqlDao.CMEPMSQuery4List(cmd);

            if (result == null || result.Rows.Count == 0)
            {
                return null;
            }

            return result;
        }
    }
}
