using Syncfusion.Compression;
using Syncfusion.XlsIO.Implementation.Collections;
using Syncfusion.XlsIO.Parser.Biff_Records.MsoDrawing;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.DynamicData;
using System.Xml;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace WCals.classes.john {
    /// <summary>
    /// XML 工具類 - 用於解析工程會電子標單 (ETenderSheet) XML 檔案
    /// </summary>
    public class xmlutil {

        #region ETenderSheet XML 解析說明
/*
        =====================================================================================
        ETenderSheet XML 資料表結構說明
        =====================================================================================
        
     【資料表名稱與用途】
        
   1. TenderInfo (標案基本資訊)
       - 用途：儲存標案的基本資料，如契約編號、機關名稱、工程名稱等
     - 欄位：ContractNo, ProcuringEntity, ProcuringEntityId, ContractTitle, 
      ContractLocation, CorrectRatio, WeightCorrectRatio, FitRatio, WeightFitRatio
        
        2. GrandTotal (總價資訊)
        - 用途：儲存標案的總金額資訊
    - 欄位：TotalAmount, TotalAmountChinese
        
        3. DetailList (詳細價目表/標單明細)
 - 用途：儲存標單的階層式價目明細（主項目、次項目等）
           - 欄位：ItemKey, ParentItemKey, ItemNo, RefItemCode, ItemKind, Description, 
       Unit, Quantity, Price, Amount, Remark, Percent, Level
    - 關聯：透過 ParentItemKey 建立父子階層關係
 
        4. CostBreakdownList (單價分析表)
           - 用途：儲存各工項的單價分析明細（含人工、材料、機具比例）
           - 欄位：ItemCode, ParentItemCode, RefItemNo, ItemKind, AnalysisOutputQuantity,
       Description, Unit, Quantity, Price, Remark, Percent,
        LabourRatio, EquipmentRatio, MaterialRatio, MiscellaneaRatio, Level
  - 關聯：透過 ParentItemCode 建立父子階層關係
        
        5. ResourceList (資源項目表)
           - 用途：儲存工程的資源項目（工項資源、單價資源等）
         - 欄位：ItemCode, ItemKind, Description, Unit, Quantity, Price, Amount, Remark,
       Percent, LabourRatio, EquipmentRatio, MaterialRatio, MiscellaneaRatio
 
        【資料表關聯】
        
        TenderInfo (1) ←──────────────→ (1) GrandTotal
          ↓        
          ↓ ContractNo   
          ↓         
 DetailList (N) ─── ParentItemKey ───→ DetailList (自我關聯/樹狀結構)
              ↓       
    ↓ RefItemCode        
              ↓  
   CostBreakdownList (N) ─ ParentItemCode → CostBreakdownList (自我關聯/樹狀結構)
    ↓     
         ↓ ItemCode        
   ↓     
        ResourceList (N)
  
    =====================================================================================
        */
        #endregion

   /// <summary>
        /// 解析 ETenderSheet XML 檔案並回傳 DataSet
        /// </summary>
        /// <param name="xmlFilePath">XML 檔案路徑</param>
      /// <returns>包含所有解析資料的 DataSet</returns>
        public static DataSet ParseETenderSheetXml(string xmlFilePath) {
            if(!File.Exists(xmlFilePath)) {
      throw new FileNotFoundException("找不到指定的 XML 檔案", xmlFilePath);
      }

string xmlContent = File.ReadAllText(xmlFilePath, System.Text.Encoding.UTF8);
       return ParseETenderSheetXmlContent(xmlContent);
        }

        /// <summary>
 /// 解析 ETenderSheet XML 內容並回傳 DataSet
      /// </summary>
        /// <param name="xmlContent">XML 內容字串</param>
    /// <returns>包含所有解析資料的 DataSet</returns>
   public static DataSet ParseETenderSheetXmlContent(string xmlContent) {
            DataSet ds = new DataSet("ETenderSheet");

  XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmlContent);

        // 設定命名空間管理員 - ETenderSheet 使用默認命名空間
     XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
       
          // 檢查根節點是否有默認命名空間
            string defaultNs = xmlDoc.DocumentElement?.GetAttribute("xmlns");
 bool hasNamespace = !string.IsNullOrEmpty(defaultNs);
            
            if(hasNamespace) {
      nsmgr.AddNamespace("et", defaultNs);
   }

  // 1. 解析標案基本資訊 (TenderInfo)
 DataTable dtTenderInfo = CreateTenderInfoTable();
   ParseTenderInfo(xmlDoc, dtTenderInfo, nsmgr, hasNamespace);
   ds.Tables.Add(dtTenderInfo);

            // 2. 解析總價資訊 (GrandTotal)
  DataTable dtGrandTotal = CreateGrandTotalTable();
   ParseGrandTotal(xmlDoc, dtGrandTotal, nsmgr, hasNamespace);
   ds.Tables.Add(dtGrandTotal);

      // 3. 解析詳細價目表 (DetailList)
 DataTable dtDetailList = CreateDetailListTable();
          string detailListXPath = hasNamespace ? "//et:DetailList" : "//DetailList";
  XmlNode detailListNode = xmlDoc.SelectSingleNode(detailListXPath, nsmgr);
 if(detailListNode != null) {
   ParsePayItems(detailListNode, dtDetailList, null, 0, nsmgr, hasNamespace);
     }
     ds.Tables.Add(dtDetailList);

   // 4. 解析單價分析表 (CostBreakdownList)
 DataTable dtCostBreakdown = CreateCostBreakdownTable();
       string costBreakdownXPath = hasNamespace ? "//et:CostBreakdownList" : "//CostBreakdownList";
   XmlNode costBreakdownNode = xmlDoc.SelectSingleNode(costBreakdownXPath, nsmgr);
  if(costBreakdownNode != null) {
                ParseWorkItems(costBreakdownNode, dtCostBreakdown, null, 0, nsmgr, hasNamespace);
 }
            ds.Tables.Add(dtCostBreakdown);

 // 5. 解析資源項目表 (ResourceList)
    DataTable dtResourceList = CreateResourceListTable();
      string resourceListXPath = hasNamespace ? "//et:ResourceList" : "//ResourceList";
          XmlNode resourceListNode = xmlDoc.SelectSingleNode(resourceListXPath, nsmgr);
    if(resourceListNode != null) {
      ParseResourceItems(resourceListNode, dtResourceList, nsmgr, hasNamespace);
    }
      ds.Tables.Add(dtResourceList);

 return ds;
      }

        #region 建立資料表結構

        /// <summary>
        /// 建立標案基本資訊資料表
        /// </summary>
        private static DataTable CreateTenderInfoTable() {
    DataTable dt = new DataTable("TenderInfo");
       dt.Columns.Add("ContractNo", typeof(string));           // 契約編號
  dt.Columns.Add("ProcuringEntity", typeof(string));      // 招標機關名稱
            dt.Columns.Add("ProcuringEntityId", typeof(string));  // 招標機關代碼
 dt.Columns.Add("ContractTitle", typeof(string));     // 工程名稱
      dt.Columns.Add("ContractLocation", typeof(string));     // 工程地點
     dt.Columns.Add("CorrectRatio", typeof(string));         // 正確率
        dt.Columns.Add("WeightCorrectRatio", typeof(string));   // 加權正確率
    dt.Columns.Add("FitRatio", typeof(string));             // 符合率
            dt.Columns.Add("WeightFitRatio", typeof(string));  // 加權符合率
     dt.Columns.Add("DocumentType", typeof(string));         // 文件類型
            dt.Columns.Add("ProcurementType", typeof(string));      // 採購類型
         dt.Columns.Add("CreatedDate", typeof(string));  // 建立日期
   return dt;
        }

   /// <summary>
        /// 建立總價資訊資料表
    /// </summary>
  private static DataTable CreateGrandTotalTable() {
       DataTable dt = new DataTable("GrandTotal");
            dt.Columns.Add("TotalAmount", typeof(decimal));       // 總金額
            dt.Columns.Add("TotalAmountChinese", typeof(string));   // 總金額中文
    return dt;
        }

      /// <summary>
    /// 建立詳細價目表資料表
        /// </summary>
        private static DataTable CreateDetailListTable() {
      DataTable dt = new DataTable("DetailList");
        dt.Columns.Add("ItemKey", typeof(string));      // 項目 Key (唯一識別)
            dt.Columns.Add("ParentItemKey", typeof(string));        // 父項目 Key
       dt.Columns.Add("ItemNo", typeof(string));               // 項次編號
  dt.Columns.Add("RefItemCode", typeof(string)); // 參考項目代碼
        dt.Columns.Add("ItemKind", typeof(string));    // 項目類型
        dt.Columns.Add("Description", typeof(string));     // 項目說明
 dt.Columns.Add("DescriptionEn", typeof(string));  // 項目說明(英文)
            dt.Columns.Add("Unit", typeof(string));              // 單位
    dt.Columns.Add("UnitEn", typeof(string));               // 單位(英文)
            dt.Columns.Add("Quantity", typeof(decimal)); // 數量
            dt.Columns.Add("Price", typeof(decimal));       // 單價
            dt.Columns.Add("PriceFixed", typeof(bool));        // 單價是否固定
     dt.Columns.Add("Amount", typeof(decimal));           // 複價
         dt.Columns.Add("AmountFixed", typeof(bool)); // 複價是否固定
   dt.Columns.Add("Remark", typeof(string)); // 備註
       dt.Columns.Add("Percent", typeof(decimal));        // 百分比
      dt.Columns.Add("LabourRatio", typeof(decimal));         // 人工比例
            dt.Columns.Add("EquipmentRatio", typeof(decimal));      // 機具比例
            dt.Columns.Add("MaterialRatio", typeof(decimal));    // 材料比例
      dt.Columns.Add("MiscellaneaRatio", typeof(decimal));    // 雜項比例
            dt.Columns.Add("Level", typeof(int));          // 階層深度
            return dt;
      }

      /// <summary>
        /// 建立單價分析表資料表
        /// </summary>
      private static DataTable CreateCostBreakdownTable() {
       DataTable dt = new DataTable("CostBreakdownList");
            dt.Columns.Add("ItemCode", typeof(string));     // 項目代碼
            dt.Columns.Add("ParentItemCode", typeof(string));       // 父項目代碼
    dt.Columns.Add("RefItemNo", typeof(string));       // 參考項次
            dt.Columns.Add("ItemKind", typeof(string));  // 項目類型
     dt.Columns.Add("AnalysisOutputQuantity", typeof(decimal)); // 分析產出數量
dt.Columns.Add("Description", typeof(string));          // 項目說明
            dt.Columns.Add("DescriptionEn", typeof(string));        // 項目說明(英文)
    dt.Columns.Add("Unit", typeof(string));    // 單位
        dt.Columns.Add("UnitEn", typeof(string));          // 單位(英文)
            dt.Columns.Add("Quantity", typeof(decimal));         // 數量
          dt.Columns.Add("Price", typeof(decimal));     // 單價
      dt.Columns.Add("PriceFixed", typeof(bool));     // 單價是否固定
            dt.Columns.Add("Remark", typeof(string));   // 備註
          dt.Columns.Add("Percent", typeof(decimal));             // 百分比
     dt.Columns.Add("LabourRatio", typeof(decimal));         // 人工比例
     dt.Columns.Add("EquipmentRatio", typeof(decimal));      // 機具比例
            dt.Columns.Add("MaterialRatio", typeof(decimal));// 材料比例
            dt.Columns.Add("MiscellaneaRatio", typeof(decimal));    // 雜項比例
    dt.Columns.Add("Level", typeof(int));      // 階層深度
 return dt;
     }

        /// <summary>
  /// 建立資源項目表資料表
        /// </summary>
        private static DataTable CreateResourceListTable() {
            DataTable dt = new DataTable("ResourceList");
  dt.Columns.Add("ItemCode", typeof(string)); // 項目代碼
 dt.Columns.Add("ItemKind", typeof(string));             // 項目類型
            dt.Columns.Add("Description", typeof(string));          // 項目說明
   dt.Columns.Add("DescriptionEn", typeof(string));        // 項目說明(英文)
            dt.Columns.Add("Unit", typeof(string));       // 單位
            dt.Columns.Add("UnitEn", typeof(string));          // 單位(英文)
            dt.Columns.Add("Quantity", typeof(decimal));// 數量
     dt.Columns.Add("Price", typeof(decimal));          // 單價
   dt.Columns.Add("PriceFixed", typeof(bool));             // 單價是否固定
    dt.Columns.Add("Amount", typeof(decimal));              // 複價
   dt.Columns.Add("AmountFixed", typeof(bool));            // 複價是否固定
            dt.Columns.Add("Remark", typeof(string));           // 備註
       dt.Columns.Add("Percent", typeof(decimal)); // 百分比
    dt.Columns.Add("LabourRatio", typeof(decimal));         // 人工比例
 dt.Columns.Add("EquipmentRatio", typeof(decimal));   // 機具比例
        dt.Columns.Add("MaterialRatio", typeof(decimal));       // 材料比例
   dt.Columns.Add("MiscellaneaRatio", typeof(decimal));    // 雜項比例
            return dt;
        }

        #endregion

        #region 解析方法

      /// <summary>
        /// 解析標案基本資訊
        /// </summary>
    private static void ParseTenderInfo(XmlDocument xmlDoc, DataTable dt, XmlNamespaceManager nsmgr, bool hasNamespace) {
      string xpath = hasNamespace ? "//et:TenderInformation" : "//TenderInformation";
       XmlNode tenderInfoNode = xmlDoc.SelectSingleNode(xpath, nsmgr);
   XmlNode rootNode = xmlDoc.DocumentElement;

      if(tenderInfoNode != null) {
          DataRow row = dt.NewRow();
// contractNo 是 TenderInformation 節點的屬性
  row["ContractNo"] = GetAttributeValue(tenderInfoNode, "contractNo");
       row["ProcuringEntity"] = GetDescriptionValue(tenderInfoNode, "ProcuringEntity", nsmgr, hasNamespace);
       row["ProcuringEntityId"] = GetNodeValue(tenderInfoNode, "ProcuringEntityId", nsmgr, hasNamespace);
           row["ContractTitle"] = GetDescriptionValue(tenderInfoNode, "ContractTitle", nsmgr, hasNamespace);
       row["ContractLocation"] = GetNodeValue(tenderInfoNode, "ContractLocation", nsmgr, hasNamespace);
        row["CorrectRatio"] = GetNodeValue(tenderInfoNode, "CorrectRatio", nsmgr, hasNamespace);
            row["WeightCorrectRatio"] = GetNodeValue(tenderInfoNode, "WeightCorrectRatio", nsmgr, hasNamespace);
         row["FitRatio"] = GetNodeValue(tenderInfoNode, "FitRatio", nsmgr, hasNamespace);
            row["WeightFitRatio"] = GetNodeValue(tenderInfoNode, "WeightFitRatio", nsmgr, hasNamespace);

      if(rootNode != null) {
     row["DocumentType"] = GetAttributeValue(rootNode, "documentType");
   row["ProcurementType"] = GetAttributeValue(rootNode, "procurementType");
       row["CreatedDate"] = GetAttributeValue(rootNode, "createdDate");
       }

       dt.Rows.Add(row);
 }
      }

        /// <summary>
        /// 解析總價資訊
  /// </summary>
   private static void ParseGrandTotal(XmlDocument xmlDoc, DataTable dt, XmlNamespaceManager nsmgr, bool hasNamespace) {
   string xpath = hasNamespace ? "//et:GrandTotalInformation" : "//GrandTotalInformation";
      XmlNode grandTotalNode = xmlDoc.SelectSingleNode(xpath, nsmgr);

      if(grandTotalNode != null) {
DataRow row = dt.NewRow();

      string priceXPath = hasNamespace ? "et:TotalAmount" : "TotalAmount";
   XmlNode totalAmountNode = grandTotalNode.SelectSingleNode(priceXPath, nsmgr);
if(totalAmountNode != null) {
        decimal totalAmount = 0;
    decimal.TryParse(totalAmountNode.InnerText?.Trim(), out totalAmount);
      row["TotalAmount"] = totalAmount;
   }

       row["TotalAmountChinese"] = GetNodeValue(grandTotalNode, "TotalAmountChinese", nsmgr, hasNamespace);
         dt.Rows.Add(row);
            }
      }

        /// <summary>
        /// 遞迴解析 PayItem 節點 (用於 DetailList)
     /// </summary>
   private static void ParsePayItems(XmlNode parentNode, DataTable dt, string parentItemKey, int level, XmlNamespaceManager nsmgr, bool hasNamespace) {
           string xpath = hasNamespace ? "et:PayItem" : "PayItem";
 XmlNodeList payItemNodes = parentNode.SelectNodes(xpath, nsmgr);

if(payItemNodes != null) {
   foreach(XmlNode payItemNode in payItemNodes) {
   DataRow row = dt.NewRow();

      string itemKey = GetAttributeValue(payItemNode, "itemKey");
     row["ItemKey"] = itemKey;
      row["ParentItemKey"] = string.IsNullOrEmpty(parentItemKey) ? (object)DBNull.Value : parentItemKey;
       row["ItemNo"] = GetAttributeValue(payItemNode, "itemNo");
     row["RefItemCode"] = GetAttributeValue(payItemNode, "refItemCode")?.Trim();
      row["ItemKind"] = GetAttributeValue(payItemNode, "itemKind");
         row["Description"] = GetDescriptionByLanguage(payItemNode, "Description", "zh-TW", nsmgr, hasNamespace);
      row["DescriptionEn"] = GetDescriptionByLanguage(payItemNode, "Description", "en", nsmgr, hasNamespace);
    row["Unit"] = GetDescriptionByLanguage(payItemNode, "Unit", "zh-TW", nsmgr, hasNamespace);
 row["UnitEn"] = GetDescriptionByLanguage(payItemNode, "Unit", "en", nsmgr, hasNamespace);
         row["Quantity"] = ParseDecimal(GetNodeValue(payItemNode, "Quantity", nsmgr, hasNamespace));
     row["Price"] = ParseDecimal(GetPriceOrAmountValue(payItemNode, "Price", nsmgr, hasNamespace));
          row["PriceFixed"] = GetPriceOrAmountFixed(payItemNode, "Price", nsmgr, hasNamespace);
      row["Amount"] = ParseDecimal(GetPriceOrAmountValue(payItemNode, "Amount", nsmgr, hasNamespace));
         row["AmountFixed"] = GetPriceOrAmountFixed(payItemNode, "Amount", nsmgr, hasNamespace);
              row["Remark"] = GetNodeValue(payItemNode, "Remark", nsmgr, hasNamespace)?.Trim();
       row["Percent"] = ParseDecimal(GetNodeValue(payItemNode, "Percent", nsmgr, hasNamespace));
           row["LabourRatio"] = ParseDecimal(GetNodeValue(payItemNode, "LabourRatio", nsmgr, hasNamespace));
          row["EquipmentRatio"] = ParseDecimal(GetNodeValue(payItemNode, "EquipmentRatio", nsmgr, hasNamespace));
              row["MaterialRatio"] = ParseDecimal(GetNodeValue(payItemNode, "MaterialRatio", nsmgr, hasNamespace));
  row["MiscellaneaRatio"] = ParseDecimal(GetNodeValue(payItemNode, "MiscellaneaRatio", nsmgr, hasNamespace));
        row["Level"] = level;

    dt.Rows.Add(row);

               // 遞迴解析子項目
     ParsePayItems(payItemNode, dt, itemKey, level + 1, nsmgr, hasNamespace);
   }
  }
        }

        /// <summary>
        /// 遞迴解析 WorkItem 節點 (用於 CostBreakdownList)
     /// </summary>
      private static void ParseWorkItems(XmlNode parentNode, DataTable dt, string parentItemCode, int level, XmlNamespaceManager nsmgr, bool hasNamespace) {
           string xpath = hasNamespace ? "et:WorkItem" : "WorkItem";
       XmlNodeList workItemNodes = parentNode.SelectNodes(xpath, nsmgr);

   if(workItemNodes != null) {
     foreach(XmlNode workItemNode in workItemNodes) {
DataRow row = dt.NewRow();

     string itemCode = GetAttributeValue(workItemNode, "itemCode");
 row["ItemCode"] = itemCode;
  row["ParentItemCode"] = string.IsNullOrEmpty(parentItemCode) ? (object)DBNull.Value : parentItemCode;
   row["RefItemNo"] = GetAttributeValue(workItemNode, "refItemNo");
       row["ItemKind"] = GetAttributeValue(workItemNode, "itemKind");
row["AnalysisOutputQuantity"] = ParseDecimal(GetAttributeValue(workItemNode, "analysisOutputQuantity"));
   row["Description"] = GetDescriptionByLanguage(workItemNode, "Description", "zh-TW", nsmgr, hasNamespace);
    row["DescriptionEn"] = GetDescriptionByLanguage(workItemNode, "Description", "en", nsmgr, hasNamespace);
    row["Unit"] = GetDescriptionByLanguage(workItemNode, "Unit", "zh-TW", nsmgr, hasNamespace);
   row["UnitEn"] = GetDescriptionByLanguage(workItemNode, "Unit", "en", nsmgr, hasNamespace);
      row["Quantity"] = ParseDecimal(GetNodeValue(workItemNode, "Quantity", nsmgr, hasNamespace));
 row["Price"] = ParseDecimal(GetPriceOrAmountValue(workItemNode, "Price", nsmgr, hasNamespace));
        row["PriceFixed"] = GetPriceOrAmountFixed(workItemNode, "Price", nsmgr, hasNamespace);
      row["Remark"] = GetNodeValue(workItemNode, "Remark", nsmgr, hasNamespace)?.Trim();
    row["Percent"] = ParseDecimal(GetNodeValue(workItemNode, "Percent", nsmgr, hasNamespace));
        row["LabourRatio"] = ParseDecimal(GetNodeValue(workItemNode, "LabourRatio", nsmgr, hasNamespace));
    row["EquipmentRatio"] = ParseDecimal(GetNodeValue(workItemNode, "EquipmentRatio", nsmgr, hasNamespace));
  row["MaterialRatio"] = ParseDecimal(GetNodeValue(workItemNode, "MaterialRatio", nsmgr, hasNamespace));
            row["MiscellaneaRatio"] = ParseDecimal(GetNodeValue(workItemNode, "MiscellaneaRatio", nsmgr, hasNamespace));
  row["Level"] = level;

         dt.Rows.Add(row);

           // 遞迴解析子項目
    ParseWorkItems(workItemNode, dt, itemCode, level + 1, nsmgr, hasNamespace);
 }
            }
     }

      /// <summary>
        /// 解析 ResourceList 中的項目
 /// </summary>
        private static void ParseResourceItems(XmlNode resourceListNode, DataTable dt, XmlNamespaceManager nsmgr, bool hasNamespace) {
   // 解析 WorkItem
   string workItemXPath = hasNamespace ? "et:WorkItem" : "WorkItem";
       XmlNodeList workItemNodes = resourceListNode.SelectNodes(workItemXPath, nsmgr);
        if(workItemNodes != null) {
foreach(XmlNode workItemNode in workItemNodes) {
     AddResourceRow(dt, workItemNode, nsmgr, hasNamespace);
    }
  }

  // 解析 PayItem
     string payItemXPath = hasNamespace ? "et:PayItem" : "PayItem";
            XmlNodeList payItemNodes = resourceListNode.SelectNodes(payItemXPath, nsmgr);
      if(payItemNodes != null) {
    foreach(XmlNode payItemNode in payItemNodes) {
   AddResourceRow(dt, payItemNode, nsmgr, hasNamespace);
 }
       }
 }

        /// <summary>
     /// 新增資源項目列
   /// </summary>
 private static void AddResourceRow(DataTable dt, XmlNode node, XmlNamespaceManager nsmgr, bool hasNamespace) {
   DataRow row = dt.NewRow();
     row["ItemCode"] = GetAttributeValue(node, "itemCode");
  row["ItemKind"] = GetAttributeValue(node, "itemKind");
       row["Description"] = GetDescriptionByLanguage(node, "Description", "zh-TW", nsmgr, hasNamespace);
            row["DescriptionEn"] = GetDescriptionByLanguage(node, "Description", "en", nsmgr, hasNamespace);
     row["Unit"] = GetDescriptionByLanguage(node, "Unit", "zh-TW", nsmgr, hasNamespace);
     row["UnitEn"] = GetDescriptionByLanguage(node, "Unit", "en", nsmgr, hasNamespace);
          row["Quantity"] = ParseDecimal(GetNodeValue(node, "Quantity", nsmgr, hasNamespace));
            row["Price"] = ParseDecimal(GetPriceOrAmountValue(node, "Price", nsmgr, hasNamespace));
       row["PriceFixed"] = GetPriceOrAmountFixed(node, "Price", nsmgr, hasNamespace);
            row["Amount"] = ParseDecimal(GetPriceOrAmountValue(node, "Amount", nsmgr, hasNamespace));
    row["AmountFixed"] = GetPriceOrAmountFixed(node, "Amount", nsmgr, hasNamespace);
         row["Remark"] = GetNodeValue(node, "Remark", nsmgr, hasNamespace)?.Trim();
   row["Percent"] = ParseDecimal(GetNodeValue(node, "Percent", nsmgr, hasNamespace));
          row["LabourRatio"] = ParseDecimal(GetNodeValue(node, "LabourRatio", nsmgr, hasNamespace));
   row["EquipmentRatio"] = ParseDecimal(GetNodeValue(node, "EquipmentRatio", nsmgr, hasNamespace));
 row["MaterialRatio"] = ParseDecimal(GetNodeValue(node, "MaterialRatio", nsmgr, hasNamespace));
      row["MiscellaneaRatio"] = ParseDecimal(GetNodeValue(node, "MiscellaneaRatio", nsmgr, hasNamespace));
         dt.Rows.Add(row);
      }

        #endregion

        #region 輔助方法

        /// <summary>
        /// 取得節點值 (支援命名空間)
        /// </summary>
      private static string GetNodeValue(XmlNode parentNode, string nodeName, XmlNamespaceManager nsmgr, bool hasNamespace) {
        string xpath = hasNamespace ? "et:" + nodeName : nodeName;
            XmlNode node = parentNode.SelectSingleNode(xpath, nsmgr);
          return node?.InnerText;
  }

        /// <summary>
        /// 取得屬性值
   /// </summary>
     private static string GetAttributeValue(XmlNode node, string attributeName) {
     return node?.Attributes?[attributeName]?.Value;
        }

  /// <summary>
  /// 取得具有 language 屬性的 Description 值 (第一個中文，支援命名空間)
        /// </summary>
       private static string GetDescriptionValue(XmlNode parentNode, string nodeName, XmlNamespaceManager nsmgr, bool hasNamespace) {
       string xpath = hasNamespace ? "et:" + nodeName : nodeName;
  XmlNodeList nodes = parentNode.SelectNodes(xpath, nsmgr);
if(nodes != null && nodes.Count > 0) {
       foreach(XmlNode node in nodes) {
       string lang = GetAttributeValue(node, "language");
  if(lang == "zh-TW" || string.IsNullOrEmpty(lang)) {
       return node.InnerText?.Trim();
    }
     }
           return nodes[0].InnerText?.Trim();
      }
   return null;
        }

        /// <summary>
        /// 依語言取得 Description 或 Unit 的值 (支援命名空間)
        /// </summary>
    private static string GetDescriptionByLanguage(XmlNode parentNode, string nodeName, string language, XmlNamespaceManager nsmgr, bool hasNamespace) {
    string xpath = hasNamespace ? "et:" + nodeName : nodeName;
    XmlNodeList nodes = parentNode.SelectNodes(xpath, nsmgr);
 if(nodes != null) {
    foreach(XmlNode node in nodes) {
        string lang = GetAttributeValue(node, "language");
  if(lang == language) {
        return node.InnerText?.Trim();
    }
    }
   }
   return null;
        }

      /// <summary>
   /// 取得 Price 或 Amount 節點的值 (支援命名空間)
        /// </summary>
        private static string GetPriceOrAmountValue(XmlNode parentNode, string nodeName, XmlNamespaceManager nsmgr, bool hasNamespace) {
     string xpath = hasNamespace ? "et:" + nodeName : nodeName;
         XmlNode node = parentNode.SelectSingleNode(xpath, nsmgr);
          return node?.InnerText;
        }

        /// <summary>
   /// 取得 Price 或 Amount 節點的 fixed 屬性 (支援命名空間)
  /// </summary>
      private static bool GetPriceOrAmountFixed(XmlNode parentNode, string nodeName, XmlNamespaceManager nsmgr, bool hasNamespace) {
      string xpath = hasNamespace ? "et:" + nodeName : nodeName;
        XmlNode node = parentNode.SelectSingleNode(xpath, nsmgr);
    string fixedValue = GetAttributeValue(node, "fixed");
      return fixedValue?.ToLower() == "true";
        }

        /// <summary>
 /// 解析十進位數值
        /// </summary>
        private static object ParseDecimal(string value) {
       if(string.IsNullOrWhiteSpace(value)) {
     return DBNull.Value;
    }

   decimal result;
            if(decimal.TryParse(value.Trim(), out result)) {
    return result;
            }
         return DBNull.Value;
        }

        #endregion

        #region DataSet 輸出方法

        /// <summary>
        /// 將 DataSet 輸出為 JSON 格式的摘要資訊
  /// </summary>
        public static string GetDataSetSummary(DataSet ds) {
            var sb = new System.Text.StringBuilder();
       sb.AppendLine("=== ETenderSheet 解析結果摘要 ===");
            sb.AppendLine();

    foreach(DataTable dt in ds.Tables) {
     sb.AppendLine($"【{dt.TableName}】 共 {dt.Rows.Count} 筆資料");
    sb.AppendLine($"  欄位: {string.Join(", ", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName))}");
           sb.AppendLine();
       }

            return sb.ToString();
        }

        #endregion
    }
}