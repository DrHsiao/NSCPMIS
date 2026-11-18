using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Helpers
{
    public class PriceAnalysisCalculate
    {
        // 依照進位方式取得價差
        public static decimal updatePriceHandler(string Carry, decimal CalculateDecimal, decimal analysisOutputQuantity, ref decimal UpdatedPrice)
        {
            UpdatedPrice = UpdatedPrice / analysisOutputQuantity;
            decimal tempUpdatedPrice = UpdatedPrice;
            decimal priceGap = 0;
            switch (Carry)
            {
                case "0":
                    updatePriceWithCeiling(ref UpdatedPrice, CalculateDecimal);
                    break;
                case "1":
                    updatePriceWithRounding(ref UpdatedPrice, CalculateDecimal);
                    break;
                case "2":
                    updatePriceWithFloor(ref UpdatedPrice, CalculateDecimal);
                    break;
                default:
                    updatePriceWithRounding(ref UpdatedPrice, CalculateDecimal);
                    break;
            }

            priceGap = (UpdatedPrice - tempUpdatedPrice) * analysisOutputQuantity;

            return priceGap;
        }

        // 無條件進位
        private static void updatePriceWithCeiling(ref decimal UpdatedPrice, decimal CalculateDecimal)
        {
            switch (CalculateDecimal)
            {
                case 0:
                    UpdatedPrice = Decimal.Ceiling(UpdatedPrice * 100) / 100;
                    break;
                case 1:
                    UpdatedPrice = Decimal.Ceiling(UpdatedPrice * 10) / 10;
                    break;
                case 2:
                    UpdatedPrice = Decimal.Ceiling(UpdatedPrice);
                    break;
                case 3:
                    UpdatedPrice = Decimal.Ceiling(UpdatedPrice / 10) * 10;
                    break;
                case 4:
                    UpdatedPrice = Decimal.Ceiling(UpdatedPrice / 100) * 100;
                    break;
                case 5:
                    UpdatedPrice = Decimal.Ceiling(UpdatedPrice / 1000) * 1000;
                    break;
            }
        }

        // 無條件捨去
        private static void updatePriceWithFloor(ref decimal UpdatedPrice, decimal CalculateDecimal)
        {
            switch (CalculateDecimal)
            {
                case 0:
                    UpdatedPrice = Decimal.Floor(UpdatedPrice * 100) / 100;
                    break;
                case 1:
                    UpdatedPrice = Decimal.Floor(UpdatedPrice * 10) / 10;
                    break;
                case 2:
                    UpdatedPrice = Decimal.Floor(UpdatedPrice);
                    break;
                case 3:
                    UpdatedPrice = Decimal.Floor(UpdatedPrice / 10) * 10;
                    break;
                case 4:
                    UpdatedPrice = Decimal.Floor(UpdatedPrice / 100) * 100;
                    break;
                case 5:
                    UpdatedPrice = Decimal.Floor(UpdatedPrice / 1000) * 1000;
                    break;
            }
        }

        // 四捨五入
        private static void updatePriceWithRounding(ref decimal UpdatedPrice, decimal CalculateDecimal)
        {
            switch (CalculateDecimal)
            {
                case 0:
                    UpdatedPrice = Decimal.Round(UpdatedPrice, 2, MidpointRounding.AwayFromZero);
                    break;
                case 1:
                    UpdatedPrice = Decimal.Round(UpdatedPrice, 1, MidpointRounding.AwayFromZero);
                    break;
                case 2:
                    UpdatedPrice = Decimal.Round(UpdatedPrice, 0, MidpointRounding.AwayFromZero);
                    break;
                case 3:
                    UpdatedPrice = Decimal.Round(UpdatedPrice / 10, 0, MidpointRounding.AwayFromZero) * 10;
                    break;
                case 4:
                    UpdatedPrice = Decimal.Round(UpdatedPrice / 100, 0, MidpointRounding.AwayFromZero) * 100;
                    break;
                case 5:
                    UpdatedPrice = Decimal.Round(UpdatedPrice / 1000, 0, MidpointRounding.AwayFromZero) * 1000;
                    break;
            }
        }

    }
}