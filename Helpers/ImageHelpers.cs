using System;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using WCals.Models;

namespace WCals.Helpers
{
    public static class ImageHelpers
    {

        //public static WCalsEntities db = new WCalsEntities();
        private static WCals.Models.CMBASEModels.CMBASEEntities db = new WCals.Models.CMBASEModels.CMBASEEntities();
        //新增
        public static MvcHtmlString addItem(this HtmlHelper helper, string src)
        {
            return MvcHtmlString.Create(string.Format("<a class='btn btn-primary' href='{0}'  />提案</a>", src));
        }
        //修改
        public static MvcHtmlString updateItem(this HtmlHelper helper, string src)
        {
            return MvcHtmlString.Create(string.Format("<a class='btn btn-default' href='{0}'  />修改</a>", src));
        }
        //結案
        public static MvcHtmlString proposalOutItem(this HtmlHelper helper, string src)
        {
            return MvcHtmlString.Create(string.Format("<a class='btn btn-danger' href='{0}'  />結案</a>", src));
        }
        //刪除
        public static MvcHtmlString deleteItem(this HtmlHelper helper, string src)
        {
            return MvcHtmlString.Create(string.Format("<a class='btn btn-primary' href='{0}'  />提案</a>", src));
        }
        //承辦單位列表
        public static MvcHtmlString SmcscdpItem(this HtmlHelper helper)
        {
            HttpContext context = HttpContext.Current;
            string memeber = Helpers.WebSiteHelper.getUserId(); //(string)context.Session["memberName"];

            var ogpurQry = from p in db.OGCDPUR
                           where p.USR_ID.Contains(memeber)
                           select p.DEPT_NO;
            var smcscdpQry = (from s in db.OGCDEPT
                              where ogpurQry.Contains(s.DEPT_NO)
                              select s).ToList();
            foreach (var smcItem in smcscdpQry)
            {
                if (smcItem.Equals(smcItem))
                {
                    return MvcHtmlString.Create(string.Format("<span>{0}</span><input type='hidden' name='advDp'  type='text' value='{1}' class='form-control' />", smcItem.DEPT_NAME, smcItem.DEPT_NAME));
                }




                // return MvcHtmlString.Create(string.Format("<a class='btn btn-primary' href='{0}'  />提案</a>", src));
            }

            return MvcHtmlString.Create(string.Format("<input disabled type='text' value='{0}' class='form-control' />", "無"));

        }
        //返回
        public static MvcHtmlString backItem(this HtmlHelper helper, string src)
        {
            return MvcHtmlString.Create(string.Format("<a class='btn btn-primary' href='{0}'  />提案</a>", src));
        }
        //提案中
        public static MvcHtmlString proposalItem(this HtmlHelper helper, string src, int count)
        {
            return MvcHtmlString.Create(string.Format("<a class='btn btn-primary' href='{0}'  />提案中<span class='badge'>{1}</span></a>", src, count));
        }
        //審核中
        public static MvcHtmlString underReviewItem(this HtmlHelper helper, string src, int count)
        {
            return MvcHtmlString.Create(string.Format("<a class='btn btn-primary' href='{0}'  />審核中<span class='badge'>{1}</span></a>", src, count));
        }
        //左側功能模組
        public static MvcHtmlString showLeftItem(this HtmlHelper helper, string src)
        {
            


            HttpContext context = HttpContext.Current;
            string memeber = Helpers.WebSiteHelper.getUserId(); //(string)context.Session["memberName"];


            var smcmenu = (from f in db.SMCMENU2
                           where f.AN_SYS_ID == null
                           orderby f.SYS_ORD
                           select f).ToList();

            //var smcmenuAuth = (from f in db.SMCMENU from aa in db.OGCUSER
            //                        let auth = from auth in db.SMCAUTH select auth.SYS_ID where auth.Contains(f.SYS_ID)
            //                        let ce = from ce in db.OGCDPUR select ce.USR_ID where ce.Contains(aa.USR_ID)
            //                        let oguser = from oguser in db.OGCUSER select oguser.USR_ID where oguser.Contains(memeber)
            //                   select f).ToList();

           
            var oguserQry = from c in db.OGCUSER
                            where c.USR_ID == memeber
                            select c.USR_ID;

            var ogpurQry = from p in db.OGCDPUR
                           where oguserQry.Contains(p.USR_ID)
                           select p.DEPT_NO;

            var smcauthrQry = from sm in db.SMCAUTH
                              where ogpurQry.Contains(sm.DEPT_NO)
                              select sm.SYS_ID;
            var smcmenuQry = (from smc in db.SMCMENU2
                              where smcauthrQry.Contains(smc.SYS_ID)

                              select smc).ToList();

            var smcmenuQryDisct = (from smc in db.SMCMENU2
                                   where smcauthrQry.Contains(smc.SYS_ID)

                                   select smc.AN_SYS_ID).Distinct().ToList();





            StringBuilder sb = new StringBuilder();



            foreach (SMCMENU em in smcmenu)
            {
                var smcmenuSub = (from f in db.SMCMENU2
                                  where f.AN_SYS_ID == em.SYS_ID
                                  orderby f.SYS_ORD
                                  select f).ToList();


                var ff = em.SYS_NAME;
                foreach (string emDisc in smcmenuQryDisct)
                {
                    if (em.SYS_ID.Contains(emDisc))
                    {
                        sb.Append("<li>");
                        sb.Append("<a href='javascript:;' data-toggle='collapse' data-target='#" + ff + "'>");
                        sb.Append(ff);
                        sb.Append("<i class='fa fa-fw fa-caret-down'></i>");
                        sb.Append("</a>");
                        sb.Append("<ul id=" + ff + " class='collapse'>");


                        foreach (SMCMENU emAuth in smcmenuQry)
                        {


                            if (em.SYS_ID.Contains(emAuth.AN_SYS_ID))
                            {

                                sb.Append("<li><a href='" + emAuth.APP_PATH + "' > " + emAuth.SYS_NAME + "</a></li>");
                            }






                        }
                        sb.Append("</ul>");
                        sb.Append("</li>");
                    }

                }






            }










            return MvcHtmlString.Create(sb.ToString());
            // return MvcHtmlString.Create(string.Format("<li><a href = 'javascript:;' data-toggle = 'collapse' data - target = '#01計畫管理' >01計畫管理 < i class='fa fa-fw fa-caret-down'></i>s</a>"));
            //return MvcHtmlString.Create(string.Fosrmat("<a class='btn btn-primary' href='{0}'  />審核中<span class='badge'>98</span></a>", src));
        }
    }
}