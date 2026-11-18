using WCals.Models.CMEDOCModels;
using WCals.Models.CMETMSModels;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using PagedList;
using System.Collections.Generic;

namespace WCals.Controllers
{
    public class PictureWordController : Controller
    {
        CMETMSEntities db3 = new CMETMSEntities();
        CMEDOCEntities db5 = new CMEDOCEntities();
        // GET: PictureWord
        public ActionResult Index(string TD_CODE_search, string type, string TD_NAME_search, string KEYWORD_search, int page = 1)
        {
            int currentPage = page < 1 ? 1 : page;
            int pageSize = 10;

            ViewBag.Type = db3.SMCCODE.Where(m => m.SYS_ID == "DC" && m.CATY_ID == "01").ToList();

            var query = db5.TDV_TDDDCOB.AsQueryable();

            if (!string.IsNullOrEmpty(TD_CODE_search))
            {
                query = query.Where(m => m.TD_CODE == TD_CODE_search);
            }

            if (!string.IsNullOrEmpty(type))
            {
                query = query.Where(m => m.IT_CODE_DESC == type);
            }

            if (!string.IsNullOrEmpty(TD_NAME_search))
            {
                query = query.Where(m => m.TD_NAME == TD_NAME_search);
            }

            if (!string.IsNullOrEmpty(KEYWORD_search))
            {
                query = query.Where(m => m.OBJ_DESC == KEYWORD_search);
            }

            if (string.IsNullOrEmpty(TD_CODE_search) && string.IsNullOrEmpty(type) && string.IsNullOrEmpty(TD_NAME_search) && string.IsNullOrEmpty(KEYWORD_search))
            {
                return View(new PagedList<TDV_TDDDCOB>(new List<TDV_TDDDCOB>(), currentPage, pageSize));
            }

            // Add sorting before applying Skip for pagination
            query = query.OrderBy(m => m.TD_CODE); // Replace SomeProperty with the actual property you want to sort by

            var result = query.ToPagedList(currentPage, pageSize);

            ViewBag.TD_CODE_search = TD_CODE_search;
            ViewBag.type1 = type;
            ViewBag.TD_NAME_search = TD_NAME_search;
            ViewBag.KEYWORD_search = KEYWORD_search;

            return View(result);
        }

        public ActionResult GetImage(string name, string name_type )
        {
            string fileName = name + name_type;
            var result = db5.TDV_TDDDCOB.Where(m=>m.OBJ_NAME == fileName).FirstOrDefault();
            if(result != null && result.OBJ_BODY != null) 
            {
                using(System.IO.MemoryStream ms = new System.IO.MemoryStream()) 
                {
                    byte[] image = result.OBJ_BODY;
                    //ms.Write(image, 78, image.Length - 78);
                    //return File(ms.ToArray(), "image/.tiff");
                    Response.Clear();
                    Response.AddHeader("Content-Disposition", "attachment; filename=" +  fileName);
                    Response.ContentType = "application/octet-stream";
                    Response.BinaryWrite(image);
                    Response.End();
                }
            }
            return RedirectToAction("Index");
        }

        private byte[] ResizeImageFile(byte[] imageFile, int targetSize)
        {
            System.Drawing.Image oldImage = System.Drawing.Image.FromStream(new MemoryStream(imageFile));
            Size newSize = CalculateDimensions(oldImage.Size, targetSize);
            Bitmap newImage = new Bitmap(newSize.Width, newSize.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            Graphics canvas = Graphics.FromImage(newImage);
            canvas.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            canvas.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            canvas.DrawImage(oldImage, new Rectangle(new Point(0, 0), newSize));
            MemoryStream m = new MemoryStream();
            newImage.Save(m, System.Drawing.Imaging.ImageFormat.Jpeg);
            return m.GetBuffer();
        }

        private Size CalculateDimensions(Size oldSize, int targetSize)
    {
        Size newSize = new Size();
        if ((oldSize.Height > oldSize.Width))
        {
            newSize.Width = (int)(oldSize.Width * (float)(targetSize / (float)oldSize.Height));
            newSize.Height = targetSize;
        }
        else
        {
            newSize.Width = targetSize;
            newSize.Height = (int)(oldSize.Height * (float)(targetSize / (float)oldSize.Width));
        }
        return newSize;
    }
    }
}