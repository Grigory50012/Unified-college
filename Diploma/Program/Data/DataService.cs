using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using Dapper;
using System.Configuration;
using System;

namespace VedomostPropuskovPGEK.Data
{
    public partial class DataService
    {
        #region Авторизация
        public static List<CheckCurator> GetCurator(string Login)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<CheckCurator>($"select * from GetGruppaName where cn_C = '{Login}'").ToList();  

            }
        }
        #endregion

        #region Учёт пропусков
        #region Получить список студентов
        public static List<SkipStudentView> StudentSkipV(string cn_G_Student, string month, string YEAR)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SkipStudentView>($"select * from StudentSkipV where cn_G_Student = '{cn_G_Student}' and month([Date]) = '{month}' and YEAR([Date]) = '{YEAR}' order by [Date]" ).ToList();
            }
        }
        #endregion

        #region Фильтры | Режим просмотра
        #region Получить список пропусков конкретного сдутента
        public static List<SkipStudentView> GetPropuskStudent(string cn_S_Student, string month, string YEAR)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SkipStudentView>($"select * from StudentSkipV where cn_S_Student = '{cn_S_Student}' and month([Date]) = '{month}' and YEAR([Date]) = '{YEAR}' order by [Date]").ToList();
            }
        }
        #endregion

        #region Получить список пропусков по причине
        public static List<SkipStudentView> GetPropuskCause(string Gr, string IdCause, string month, string YEAR)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SkipStudentView>($"select * from StudentSkipV where cn_G_Student = '{Gr}' and IdCause_Cause = {IdCause} and month([Date]) = '{month}' and YEAR([Date]) = '{YEAR}' order by [Date]").ToList();
            }
        }
        #endregion

        #region Получить список пропусков по причине и студенту
        public static List<SkipStudentView> GetPropuskCauseAndStudent(string cn_S_Student, string IdCause, string month, string YEAR)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SkipStudentView>($"select * from StudentSkipV where cn_S_Student = '{cn_S_Student}' and IdCause_Cause = {IdCause} and month([Date]) = '{month}' and YEAR([Date]) = '{YEAR}' order by [Date]").ToList();
            }
        }
        #endregion
        #endregion

        #region Получить список причины
        public static List<Cause> GetAllCause()
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<Cause>("select * from Cause").ToList();
            }
        }
        #endregion

        #region Получить список дисциплин
        public static List<GetSubjectView> GetAllSubject(string cn_G)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<GetSubjectView>($"select distinct [Name] from GetSubjectView where cn_G = '{cn_G}'").ToList();
            }
        }
        #endregion
        
        #region Получить список вид занятия
        public static List<EmpForm> GetAllEmpForm()
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<EmpForm>($"select * from EmpForm where [CuratorUsed] = 'True'").ToList();
            }
        }
        #endregion

        #region Получить ФИО преподавателя
        public static List<GetNameTeacher> GetTeacher(string Name_Subject, string cn_G)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<GetNameTeacher>($"select * from GetNameTeacher where Name_Subject = '{Name_Subject}' and cn_G = '{cn_G}' ").ToList();
            }
        }
        #endregion

        #region Получить список студентов определенной группы
        public static List<Student> GetStudentGruppa(string p1)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<Student>("select * From GetStudetGruppa where cn_G = "+"'"+p1+"'").ToList();
            }
        }
        #endregion

        #region Получить ФИО и групу курятора для выхода
        public static List<GetFIOandGroupTeacher> FIOandGroupTeacher(string p1, string p2)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<GetFIOandGroupTeacher>("select * From GetFIOandGroupTeacher where cn_G = " + "'" + p1 + "'" +"and cn_C = "+"'"+p2+"'").ToList();
            }
        }
        #endregion

        #region Добавление пропусков студентов | причина уважительная
        public static List<StudentSkip> AddStudentSkip(string cn_S, string idCause, string date, string count_hour)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<StudentSkip>($"exec AddStudentSkip {cn_S}, {idCause}, '{date}', {count_hour}").ToList();
            }
        }
        #endregion

        #region Добавление пропусков студентов | причина НЕуважительная
        public static void AddStudentSkipNe(string cn_S, string idCause, string date, string count_hour, string idEmpForn, string IdSubject_Teacher)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                 db.Query<StudentSkip>($"exec AddStudentSkip {cn_S}, {idCause}, '{date}', {count_hour}, {idEmpForn}, {IdSubject_Teacher}");
            }
        }
        #endregion

        #region Удаление пропусков студентов
        public static void DeleteStudentSkip(string IdStudentSkip)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                db.Query<StudentSkip>("exec DeleteStudentSkip " + IdStudentSkip);
            }
        }
        #endregion

        #region Изменение пропусков студентов
        public static void UpdateStudentSkipNE(
            string IdStudentSkip, 
            string cn_S, 
            string IdCause, 
            string Date, 
            string Count_hour, 
            string IdEmpForn, 
            string IdSubject_Teacher)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };

                db.Query<StudentSkip>($"exec UpdateStudentSkip {IdStudentSkip}, {cn_S}, {IdCause}, '{Date}', {Count_hour}, {IdEmpForn}, {IdSubject_Teacher}");
            }
        }
        #endregion

        #region Изменение пропусков студентов
        public static void UpdateStudentSkipUvaz(
            string IdStudentSkip,
            string cn_S,
            string IdCause,
            string Date,
            string Count_hour)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };

                db.Query<StudentSkip>($"exec UpdateStudentSkipUvaz {IdStudentSkip}, {cn_S}, {IdCause},'{Date}', {Count_hour}, null, null");
            }
        }
        #endregion

        #region Контекстное меню | Grid
        #region Изменить причину
        public static void SpeedEditCause(
            string IdStudentSkip,
            string IdCause)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                db.Query<StudentSkip>($"update StudentSkip set [IdCause] = {IdCause} where [IdStudentSkip] = {IdStudentSkip}");
            }
        }
        #endregion

        #region Изменить количество часов
        public static void TimeEdit(
            string IdStudentSkip,
            string Count_hour)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                db.Query<StudentSkip>($"update StudentSkip set [Count_hour] = {Count_hour} where [IdStudentSkip] = {IdStudentSkip}");
            }
        }
        #endregion
        #endregion

        #region Отчет за месяц
       
        public static List<SkipStudentView> GetPropuskGruppaZaMes(string cn_G, string Date)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SkipStudentView>($"select * From StudentSkipV where cn_G_Student = '{cn_G}' and Month([Date]) = '{Date}'").ToList();
            }
        }
        #endregion
        #endregion

        #region Сдача СПХ группы
        #region Получить список СПХ 
        public static List<SphSend> GetSPH(string Gr)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SphSend>($"select * from SphSend Where cn_G =N'{Gr}'").ToList();
            }
        }
        #endregion

        #region Загрузка СПХ группы
        public static List<SphSend> GetSPH1(string cn_G)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SphSend>($"EXEC GenerateSphSend '{cn_G}'").ToList();
            }
        }
        #endregion

        #region Добавление СПХ Группы
        public static List<SphSend> AddSphSend(SphSend sph)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SphSend>($"EXEC AddSphSend '{sph.cn_G}', {sph.Students}, {sph.Budget}, {sph.NonBudget}, {sph.Boys}, {sph.Girls}, {sph.Underage}, {sph.Adult}, {sph.Nonresident}, {sph.Hostel}, {sph.Flat},{sph.Foreign_student}, " +
                                                               $"{sph.Incomplete}, {sph.Many_children_family}, {sph.Trusteeship}, {sph.Foster_family}, {sph.Refurgee}, {sph.Have_disabled_parents}, {sph.Low_income_family}, {sph.Family_students}, {sph.Have_children}," +
                                                               $"{sph.State_support_in_college}, {sph.Socially_dangerous_position}, {sph.Need_for_state_protection}, {sph.Individual_prophylactic_accounting}, {sph.Disabled_people}," +
                                                               $"{sph.BRSM_members}, {sph.TradeUnion_members}, {sph.id_period}").ToList();
            }
        }
        #endregion

        #region Получение дат сдачи СПХ
        public static List<ReportingPeriodDate> GetDateSPH()
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<ReportingPeriodDate>($"SELECT DateSPH FROM ReportingPeriodDate").ToList();
            }
        }
        #endregion
        #endregion

        #region Формирование документов
        #region Спрака о посещении семьи учащегося
        #region Получение сведений о семье учащегося
        public static List<FamilyVisitInfo> GetFamilyVisitInfo(string studentId)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<FamilyVisitInfo>($"EXEC GetStudentFamilyVisitInfo '{studentId}'").ToList();
            }
        }
        #endregion

        #region Изменение сведений о семье учащегося
        public static List<FamilyVisitInfo> UpdateFamilyVisitInfo(FamilyVisitInfo familyVisitInfo)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                }
                return db.Query<FamilyVisitInfo>($"EXEC UpdateFamilyVisitInfo '{familyVisitInfo.cn_S}', '{familyVisitInfo.Date_of_visit.Date.ToString("yyyy-MM-dd HH:mm:ss")}', '{familyVisitInfo.House_characteristics}', '{familyVisitInfo.Living_conditions}'").ToList();
            }
        }
        #endregion

        #region Добавление сведений о семье учащегося
        public static List<FamilyVisitInfo> AddFamilyVisitInfo(FamilyVisitInfo familyVisitInfo)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                }
                return db.Query<FamilyVisitInfo>($"EXEC AddFamilyVisitInfo '{familyVisitInfo.cn_S}', '{familyVisitInfo.Date_of_visit.Date.ToString("yyyy-MM-dd HH:mm:ss")}', '{familyVisitInfo.House_characteristics}', '{familyVisitInfo.Living_conditions}'").ToList();
            }
        }
        #endregion

        #region Получение сведений о родителях учащегося
        public static List<FamilyComposition> GetParents(int id)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                }
                return db.Query<FamilyComposition>($"EXEC GetParents '{id}'").ToList();
            }
        }
        #endregion
        #endregion

        #region СПХ Призывника/Характеристика учащегося
        #region Получение сведений об учащемся
        public static List<StudentCharacterization> GetStudentCharacterization(int studentId)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                }
                return db.Query<StudentCharacterization>($"SELECT * FROM StudentCharacterization WHERE cn_S = '{studentId}'").ToList();
            }
        }
        #endregion

        #region Добавление сведений об учащемся
        public static List<StudentCharacterization> AddStudentCharacterization(StudentCharacterization studentCharacterization)
        {
            int inclination_to_withdrawal = 0;
            if (studentCharacterization.Inclination_to_withdrawal == true)
                inclination_to_withdrawal = 1;
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                }
                return db.Query<StudentCharacterization>($"INSERT INTO StudentCharacterization(cn_S, General_development_and_outlook, Hobbies, Academic_performance, Self_assessment, Most_favorite_subjects, Most_disliked_subjects, Long_absences, Attitude_to_physical_culture, Temperament, Psychological_features, Communication_with_peers, Communication_with_teachers, Signs_of_social_neglect, Level_of_neuropsychological_stability, Inclination_to_withdrawal) " +
                    $"VALUES({studentCharacterization.cn_S}, '{studentCharacterization.General_development_and_outlook}', '{studentCharacterization.Hobbies}', '{studentCharacterization.Academic_performance}', '{studentCharacterization.Self_assessment}', '{studentCharacterization.Most_favorite_subjects}', '{studentCharacterization.Most_disliked_subjects}', '{studentCharacterization.Long_absences}', '{studentCharacterization.Attitude_to_physical_culture}', '{studentCharacterization.Temperament}', '{studentCharacterization.Psychological_features}', '{studentCharacterization.Communication_with_peers}', '{studentCharacterization.Communication_with_teachers}', '{studentCharacterization.Signs_of_social_neglect}', '{studentCharacterization.Level_of_neuropsychological_stability}', {inclination_to_withdrawal})").ToList();
            }
        }
        #endregion

        #region Изменение сведений об учащемся
        public static List<StudentCharacterization> UpdateStudentCharacterization(StudentCharacterization studentCharacterization)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                int inclination_to_withdrawal = 0;
                if (studentCharacterization.Inclination_to_withdrawal == true)
                    inclination_to_withdrawal = 1;
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                }
                return db.Query<StudentCharacterization>($"" +
                    $"UPDATE StudentCharacterization " +
                    $"SET " +
                    $"General_development_and_outlook = '{studentCharacterization.General_development_and_outlook}', " +
                    $"Hobbies = '{studentCharacterization.Hobbies}', " +
                    $"Academic_performance = '{studentCharacterization.Academic_performance}', " +
                    $"Self_assessment = '{studentCharacterization.Self_assessment}', " +
                    $"Most_favorite_subjects = '{studentCharacterization.Most_favorite_subjects}', " +
                    $"Most_disliked_subjects = '{studentCharacterization.Most_disliked_subjects}', " +
                    $"Long_absences = '{studentCharacterization.Long_absences}', " +
                    $"Attitude_to_physical_culture = '{studentCharacterization.Attitude_to_physical_culture}', " +
                    $"Temperament = '{studentCharacterization.Temperament}', " +
                    $"Psychological_features = '{studentCharacterization.Psychological_features}', " +
                    $"Communication_with_peers = '{studentCharacterization.Communication_with_peers}', " +
                    $"Communication_with_teachers = '{studentCharacterization.Communication_with_teachers   }', " +
                    $"Signs_of_social_neglect = '{studentCharacterization.Signs_of_social_neglect}', " +
                    $"Level_of_neuropsychological_stability = '{studentCharacterization.Level_of_neuropsychological_stability}', " +
                    $"Inclination_to_withdrawal = {inclination_to_withdrawal} " +
                    $"WHERE cn_S = {studentCharacterization.cn_S}").ToList();
            }
        }
        #endregion

        #region Получение мед групы учащегося
        public static List<MedicalGroup> GetStudentMedicalGroupName(int ID_MG)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                }
                return db.Query<MedicalGroup>($"SELECT * FROM MedicalGroup WHERE ID_MG = '{ID_MG}'").ToList();
            }
        }
        #endregion

        #region Получение названия специальности учащегося
        public static List<string> GetStudentSpecialtyName(string ID)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                }
                return db.Query<string>($"SELECT FullName FROM Specialty WHERE cn_Spec = (Select cn_Spec from[Group] WHERE cn_G = (SELECT cn_G FROM Student WHERE cn_S = '{ID}'))").ToList();
            }
        }
        #endregion
        #endregion
        #endregion

        #region Сведения об учащихся
        #region Личные сведения
        public static List<StudentPersonalInfo> GetStudentPersonalInfo(string studentId)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<StudentPersonalInfo>($"EXEC GetStudentPersonalInfo '{studentId}'").ToList();
            }
        }

        public static List<StudentPersonalInfo> UpdateStudentPersonalInfo(string studentId, StudentPersonalInfo studentPersonalInfo)
        {
            string dateBirth = "NULL";
            if (studentPersonalInfo.DateBirth.Year != 1)
                dateBirth = $"'{studentPersonalInfo.DateBirth.Date.ToString("yyyy-MM-dd HH:mm:ss")}'";

            string stateDateOfStudy = "NULL";
            if (studentPersonalInfo.StateDateOfStudy.Year != 1)
                stateDateOfStudy = $"'{studentPersonalInfo.StateDateOfStudy.Date.ToString("yyyy-MM-dd HH:mm:ss")}'";

            string medicalGroupName = "NULL";
            if (studentPersonalInfo.MedicalGroupName != 0)
                medicalGroupName = studentPersonalInfo.MedicalGroupName.ToString();

            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<StudentPersonalInfo>($"EXEC UpdateStudentPersonalInfo " +
                    $"'{studentId}', " +
                    $"'{studentPersonalInfo.Telephone_Mob}', " +
                    $"'{studentPersonalInfo.Telephone_Home}', " +
                    $"{dateBirth}, " +
                    $"'{studentPersonalInfo.RB}', " +
                    $"'{studentPersonalInfo.PassportSeries}', " +
                    $"'{studentPersonalInfo.PassportNumber}', " +
                    $"'{studentPersonalInfo.PasportID}', " +
                    $"'{studentPersonalInfo.Adress}', " +
                    $"'{studentPersonalInfo.FromAnotherTown}', " +
                    $"'{studentPersonalInfo.OnFlat}', " +
                    $"'{studentPersonalInfo.OnHostel}', " +
                    $"'{studentPersonalInfo.FlatDescription}', " +
                    $"'{studentPersonalInfo.RoomNumber}', " +
                    $"{medicalGroupName}, " +
                    $"'{studentPersonalInfo.Budget}', " +
                    $"'{studentPersonalInfo.OnIPA}', " +
                    $"'{studentPersonalInfo.IPARemarks}', " +
                    $"'{studentPersonalInfo.IsDisabled}', " +
                    $"'{studentPersonalInfo.DisabledStudentRemarks}', " +
                    $"'{studentPersonalInfo.OnSDP}', " +
                    $"'{studentPersonalInfo.SDPRemarks}', " +
                    $"'{studentPersonalInfo.OnNFSP}', " +
                    $"'{studentPersonalInfo.NFSPRemarks}', " +
                    $"'{studentPersonalInfo.HaveChildren}', " +
                    $"'{studentPersonalInfo.HaveChildrenRemarks}', " +
                    $"'{studentPersonalInfo.AnOrphan}', " +
                    $"'{studentPersonalInfo.AnAdopted}', " +
                    $"'{studentPersonalInfo.OnGuardianship}', " +
                    $"'{studentPersonalInfo.OnTrusteeship}', " +
                    $"'{studentPersonalInfo.OnStateSupport}', " +
                    $"'{studentPersonalInfo.FamilyState}', " +
                    $"{stateDateOfStudy}, " +
                    $"'{studentPersonalInfo.PreviousPlaceOfStudy}'").ToList();
            }
        }

        public static List<MedicalGroup> GetAllMedicalGroups()
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<MedicalGroup>($"SELECT * FROM MedicalGroup").ToList();
            }
        }
        #endregion

        #region Сведения о семье
        #region Состав семьи
        #region Получение сведений о составе семьи учащегося
        public static List<FamilyComposition> GetFamily(string studentId)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<FamilyComposition>($"EXEC GetFamily '{studentId}'").ToList();
            }
        }
        #endregion

        #region Получение сведений о видах родства
        public static List<RelativeForm> GetRelativeForms()
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<RelativeForm>($"EXEC GetRelativeForms").ToList();
            }
        }
        #endregion

        #region Добавление родственника
        public static List<FamilyComposition> AddRelative(FamilyComposition relative)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<FamilyComposition>($"EXEC AddRelative '{relative.FIO}', '{relative.YearBirth.Date.ToString("yyyy-MM-dd HH:mm:ss")}', '{relative.Work_Study_Place}', '{relative.Place_of_residence}', '{relative.ID_RF}', '{relative.cn_S}' ").ToList();
            }
        }
        #endregion

        #region Изменение родственника
        public static List<FamilyComposition> EditRelative(FamilyComposition relative)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<FamilyComposition>($"EXEC EditRelative '{relative.FIO}', '{relative.YearBirth.Date.ToString("yyyy-MM-dd HH:mm:ss")}', '{relative.Work_Study_Place}', '{relative.Place_of_residence}', '{relative.ID_RF}', '{relative.Id_Relative}' ").ToList();
            }
        }
        #endregion

        #region Удаление родственника
        public static List<FamilyComposition> DeleteRelative(int Id_Relative)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<FamilyComposition>($"DELETE FROM FamilyComposition WHERE Id_Relative = '{Id_Relative}'").ToList();
            }
        }
        #endregion
        #endregion

        #region Вид семьи
        #region Получение сведений о видах семьи учащегося
        public static List<FamilyType> GetTypesFamilyByStudent(string studentId)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<FamilyType>($"EXEC GetTypeOfFamily '{studentId}'").ToList();
            }
        }
        #endregion

        #region Получение сведений о все возможных видах семьи
        public static List<FamilyType> GetAllTypesFamily()
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<FamilyType>($"SELECT * FROM FamilyType").ToList();
            }
        }
        #endregion

        #region Добавление сведений о виде семьи учащегося
        public static List<FamilyCharacteristics> AddTypeFamily(FamilyCharacteristics characteristics)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<FamilyCharacteristics>($"EXEC AddFamilyCharacteristics '{characteristics.cn_S}', '{characteristics.id_type}'").ToList();
            }
        }
        #endregion

        #region Проверка на совпадение вида семьи
        public static List<FamilyCharacteristics> GetTypeFamilyByRelationAndStudent(FamilyCharacteristics characteristics)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<FamilyCharacteristics>($"EXEC CheckDublicateTypeFamily '{characteristics.id_type}', '{characteristics.cn_S}'").ToList();
            }
        }
        #endregion

        #region Удаление вида семьи
        public static List<FamilyCharacteristics> DeleteTypeFamily(int id)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<FamilyCharacteristics>($"EXEC DeleteFamilyCharacteristics '{id}'").ToList();
            }
        }
        #endregion
        #endregion

        #region Обнавление сведений о виде семьи учащегося
        public static List<StudentPersonalInfo> UpdateInvalidParentsInfo(int id, bool onDisabledParents, string remarks)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<StudentPersonalInfo>($"IF {Convert.ToInt32(onDisabledParents)} = 1 AND NOT EXISTS(SELECT DisabledParents.cn_S FROM DisabledParents WHERE DisabledParents.cn_S = {id}) " +
                    $"INSERT INTO DisabledParents " +
                    $"VALUES({id}, '{remarks}') " +
                    $"ELSE IF {Convert.ToInt32(onDisabledParents)} = 1 AND EXISTS(SELECT DisabledParents.cn_S FROM DisabledParents WHERE DisabledParents.cn_S = {id}) " +
                    $"UPDATE DisabledParents " +
                    $"SET Remarks = '{remarks}' " +
                    $"WHERE cn_S = {id} " +
                    $"ELSE " +
                    $"DELETE FROM DisabledParents " +
                    $"WHERE DisabledParents.cn_S = {id}").ToList();
            }
        }
        #endregion
        #region Родители инвалиды

        #endregion
        #endregion

        #region Занятость ОПД
        #region Занятость ОПД
        #region Получение сведений о занятости ОПД учащегося
        public static List<SUA_Employment> GetSUAEmployment(string studentId)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SUA_Employment>($"EXEC GetSUAEmployment '{studentId}'").ToList();
            }
        }
        #endregion

        #region Добавление сведений о занятости ОПД учащемуся
        public static List<SUA_Employment> AddSUAEmployment(SUA_Employment employment)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SUA_Employment>($"EXEC AddSUAEmployment '{employment.ActivitiesForm}', '{employment.Achievements}','{employment.Note}','{employment.cn_S}'").ToList();
            }
        }
        #endregion

        #region Изменение сведений о занятости ОПД учащегося
        public static List<SUA_Employment> EditSUAEmployment(SUA_Employment employment)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SUA_Employment>($"EXEC EditSUAEmployment '{employment.ActivitiesForm}', '{employment.Achievements}','{employment.Note}','{employment.cn_S}', '{employment.ID_SUA_Emp}'").ToList();
            }
        }
        #endregion

        #region Удаление сведений об занятости ОПД учащегося
        public static List<SUA_Employment> DeleteSUAEmployment(string idEmployment)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SUA_Employment>($"EXEC DeleteSUAEmployment '{idEmployment}'").ToList();
            }
        }
        #endregion
        #endregion

        #region Сектор актива
        #region Получение сведений о секторе актива учащегося
        public static List<ActiveSector> GetActiveSectorByStudent(string studentId)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<ActiveSector>($"EXEC GetActiveSector '{studentId}'").ToList();
            }
        }
        #endregion

        #region Получение сведений о все возможных секторах актива
        public static List<ActiveSector> GetAllActiveSectors()
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<ActiveSector>($"SELECT * FROM ActiveSector").ToList();
            }
        }
        #endregion

        #region Добавление сведений о секторе актива учащегося
        public static List<Assigments> AddActiveSector(Assigments characteristics)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<Assigments>($"EXEC AddActiveSector '{characteristics.ID_ActiveSector}', '{characteristics.cn_S}'").ToList();
            }
        }
        #endregion

        #region Удаление сектора актива
        public static List<Assigments> DeleteActiveSector(int idActiveSector, string idStudent)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<Assigments>($"EXEC DeleteActiveSector '{idActiveSector}', '{idStudent}'").ToList();
            }
        }
        #endregion
        #endregion
        #endregion

        #region Поощрения и асоциальное поведение
        #region Асоциальное поведение
        #region Получение асоциального поведения учащегося
        public static List<AssocialBehavior> GetAssocialBehavior(string studentId)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<AssocialBehavior>($"SELECT * FROM AssocialBehavior WHERE cn_S = '{studentId}'").ToList();
            }
        }
        #endregion

        #region Добавление асоциального поведения учащемуся
        public static List<SUA_Employment> AddAssocialBehavior(AssocialBehavior associalBehavior)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<SUA_Employment>($"INSERT INTO AssocialBehavior([Date], [Content], Nature_Assoc_Beh, Working_with_parents_students, TakenMeasures, Result, PsychologistsRecommendations, cn_S) VALUES('{associalBehavior.Date.Date.ToString("yyyy-MM-dd HH:mm:ss")}', '{associalBehavior.Content}', '{associalBehavior.Nature_Assoc_Beh}', '{associalBehavior.Working_with_parents_students}', '{associalBehavior.TakenMeasures}', '{associalBehavior.Result}', '{associalBehavior.PsychologistsRecommendations}', '{associalBehavior.cn_S}')").ToList();
            }
        }
        #endregion

        #region Изменение асоциального поведения учащегося
        public static List<AssocialBehavior> EditAssocialBehavior(AssocialBehavior associalBehavior)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<AssocialBehavior>($"UPDATE AssocialBehavior " +
                    $"SET [Date] = '{associalBehavior.Date.Date.ToString("yyyy-MM-dd HH:mm:ss")}', Content = '{associalBehavior.Content}', " +
                    $"Nature_Assoc_Beh = '{associalBehavior.Nature_Assoc_Beh}', Working_with_parents_students = '{associalBehavior.Working_with_parents_students}', TakenMeasures = '{associalBehavior.TakenMeasures}', " +
                    $"Result = '{associalBehavior.Result}', PsychologistsRecommendations = '{associalBehavior.PsychologistsRecommendations}'" +
                    $"WHERE ID_Assoc_beh = '{associalBehavior.ID_Assoc_beh}'").ToList();
            }
        }
        #endregion

        #region Удаление асоциального поведения учащегося
        public static List<AssocialBehavior> DeleteAssocialBehavior(int idAssocialBehavior)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<AssocialBehavior>($"DELETE AssocialBehavior Where ID_Assoc_beh = '{idAssocialBehavior}'").ToList();
            }
        }
        #endregion
        #endregion

        #region Поощрения/Взыскания
        #region Получение асоциального поведения учащегося
        public static List<PromotionPunishView> GetPromotionPunish(string studentId)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<PromotionPunishView>($"SELECT * FROM PromotionPunishView WHERE cn_S = '{studentId}'").ToList();
            }
        }
        #endregion

        #region Получение всех категорий
        public static List<Promotion_punish_category> GetAllPromotionPunishCategory()
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<Promotion_punish_category>($"SELECT * FROM Promotion_punish_category").ToList();
            }
        }
        #endregion

        #region Получение всех видов поощрения/взыскания
        public static List<Promotion_Punish_Type> GetAllPromotionPunishType(int idCategory)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<Promotion_Punish_Type>($"SELECT * FROM Promotion_Punish_Type WHERE id_Category = {idCategory}").ToList();
            }
        }
        #endregion

        #region Добавление поощрения/взыскания
        public static List<Promotion_punish> AddPromotionPunish(PromotionPunishView promotionPunish)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                if (promotionPunish.id_Type != 0)
                    return db.Query<Promotion_punish>($"INSERT INTO Promotion_punish(PPDate, PPDescription, cn_S, id_Type) VALUES('{promotionPunish.PPDate.Date.ToString("yyyy-MM-dd HH:mm:ss")}', '{promotionPunish.PPDescription}', '{promotionPunish.cn_S}', {promotionPunish.id_Type})").ToList();
                else
                    return db.Query<Promotion_punish>($"INSERT INTO Promotion_punish(PPDate, PPDescription, cn_S, id_Type) VALUES('{promotionPunish.PPDate.Date.ToString("yyyy-MM-dd HH:mm:ss")}', '{promotionPunish.PPDescription}', '{promotionPunish.cn_S}', NULL)").ToList();
            }
        }
        #endregion

        #region Изменение поощрения/взыскания
        public static List<Promotion_punish> EditPromotionPunish(PromotionPunishView promotionPunish)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<Promotion_punish>($"UPDATE Promotion_punish " +
                    $"SET PPDate = '{promotionPunish.PPDate.Date.ToString("yyyy-MM-dd HH:mm:ss")}', PPDescription = '{promotionPunish.PPDescription}', " +
                    $"id_Type = '{promotionPunish.id_Type}' " +
                    $"WHERE id_Promotion = '{promotionPunish.id_Promotion}'").ToList();
            }
        }
        #endregion

        #region Удаление асоциального поведения учащегося
        public static List<Promotion_punish> DeletePromotionPunish(int idPromotion)
        {
            using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
            {
                if (db.State == ConnectionState.Closed)
                {
                    db.Open();
                };
                return db.Query<Promotion_punish>($"DELETE Promotion_punish Where id_Promotion = '{idPromotion}'").ToList();
            }
        }
        #endregion
        #endregion
        #endregion
        #endregion
    }
}
