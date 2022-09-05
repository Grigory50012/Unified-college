using System;

namespace VedomostPropuskovPGEK.Data
{
    public class AssocialBehavior
    {
      public int ID_Assoc_beh { get; set; }
      public DateTime Date { get; set; }
      public string Content { get; set; }
      public string Nature_Assoc_Beh { get; set; }
      public string Working_with_parents_students { get; set; }
      public string TakenMeasures { get; set; }
      public string Result { get; set; }
      public string PsychologistsRecommendations { get; set; }
      public int cn_S { get; set; }
    }
}
