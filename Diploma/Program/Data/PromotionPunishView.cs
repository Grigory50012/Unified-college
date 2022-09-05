using System;

namespace VedomostPropuskovPGEK.Data
{
    public class PromotionPunishView
    {
        public int cn_S { get; set; }
        public DateTime PPDate { get; set; }
        public string PPDescription { get; set; }
        public int id_Promotion { get; set; }
        public string PPName { get; set; }
        public string Category_Name { get; set; }
        public int id_Type { get; set; }
        public int id_Category { get; set; }
    }
}
