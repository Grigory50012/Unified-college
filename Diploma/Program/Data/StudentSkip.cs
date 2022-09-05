using System;

namespace VedomostPropuskovPGEK.Data
{
    public class StudentSkip
    {
      public int IdStudentSkip{get; set;}
      public string cn_S{get; set;}
      public int IdCause {get; set;}
      public DateTime Date{get; set;}
      public int Count_hour {get; set;}
      public int IdEmpForn {get; set;}
      public int IdSubject_Teacher {get; set;}
    }
}
