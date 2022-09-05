using System;

namespace VedomostPropuskovPGEK.Data
{
    public class StudentPersonalInfo
    {
        public string Telephone_Mob { get; set; }
        public string Telephone_Home { get; set; }

        public DateTime DateBirth { get; set; }
        public bool RB { get; set; }
        public string PassportSeries { get; set; }
        public string PassportNumber { get; set; }
        public string PasportID { get; set; }

        public string Adress { get; set; }
        public bool FromAnotherTown { get; set; }
        public bool OnFlat { get; set; }
        public bool OnHostel { get; set; }
        public string FlatDescription { get; set; }
        public string RoomNumber { get; set; }

        public int MedicalGroupName { get; set; }
        public bool Budget { get; set; }
        public bool FamilyState { get; set; }
        public bool OnIPA { get; set; }
        public string IPARemarks { get; set; }

        public bool OnSDP { get; set; }
        public string SDPRemarks { get; set; }
        public bool OnNFSP { get; set; }
        public string NFSPRemarks { get; set; }
        public bool IsDisabled { get; set; }
        public string DisabledStudentRemarks { get; set; }

        public bool AnOrphan { get; set; }
        public bool OnGuardianship { get; set; }
        public bool OnTrusteeship { get; set; }
        public bool OnStateSupport { get; set; }
        public bool AnAdopted { get; set; }
        public bool HaveChildren { get; set; }
        public string HaveChildrenRemarks { get; set; }

        public DateTime StateDateOfStudy { get; set; }
        public string PreviousPlaceOfStudy { get; set; }

        public bool OnDisabledParents { get; set; }
        public string DisabledParentsRemarks { get; set; }
    }
}