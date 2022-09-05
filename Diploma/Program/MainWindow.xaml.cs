using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using VedomostPropuskovPGEK.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace VedomostPropuskovPGEK
{
    public partial class MainWindow : Window
    {
        #region Глобальные переменные
        private List<ListStudents> listStudents;
        private List<SkipStudents> skipStudents;
        private string Gr;
        private string Cnc;
        private string SegodnyaMonth = DateTime.Now.Month.ToString();
        private string SegodnyaYear = DateTime.Now.Year.ToString();
        private string Id_StudentSkip = null;
        private bool StatusEdit = false;
        private bool StatusForm = false;
        private string curator = string.Empty;
        #endregion

        public MainWindow(string grupa, string cn_C)
        {
            Gr = grupa;
            Cnc = cn_C;
            curator = DataService.FIOandGroupTeacher(grupa, cn_C)[0].FIOTeacher;

            InitializeComponent();
            try
            {
                #region Загрузка данных о кураторе
                txtBlock1.Text = curator;
                txtBlock2.Text = "Куратор " + curator;
                #endregion

                #region Загрузка данных в "Учёт пропусков"
                StatusForm = true;

                #region Загрузка данных в listbox
                listStudents = new List<ListStudents>();
                ListStudents list;
                foreach (Student st in DataService.GetStudentGruppa(grupa))
                {
                    list = new ListStudents();
                    list.Obj = st;
                    list.Select = false;
                    listStudents.Add(list);
                }
                ListBoxStudent.ItemsSource = listStudents;
                #endregion

                #region Загрузка данных в DataGrid - режим просмотра - редактирования
                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(grupa, SegodnyaMonth, SegodnyaYear))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                dgStudentSkipView.ItemsSource = skipStudents;
                #endregion

                #region Загрузка данных в ComboBox
                Combo1.ItemsSource = DataService.GetAllCause();
                Combo3.ItemsSource = DataService.GetAllSubject(Gr);
                Combo4.ItemsSource = DataService.GetAllEmpForm();

                cbStudent.ItemsSource = DataService.GetStudentGruppa(Gr);
                cbCause.ItemsSource = DataService.GetAllCause();
                #endregion

                #region Загрузка данных в TextBox
                DateTime nowDays = DateTime.Now;
                int Month = nowDays.Month;
                if (Month < 9)
                    tbUchebniyGod.Text = $"{nowDays.Year - 1}/{nowDays.Year}";
                else
                    tbUchebniyGod.Text = $"{nowDays.Year}/{nowDays.Year + 1}";
                #endregion
                #endregion

                #region Загрузка данных в "Сдача СПХ"
                UpdatngSPHGroup();

                List<ReportingPeriodDate> datesOfPeriods = DataService.GetDateSPH();
                cbPeriod.ItemsSource = datesOfPeriods;

                lbl1.Content = datesOfPeriods[0].dateSPH;
                lbl2.Content = datesOfPeriods[1].dateSPH;
                lbl3.Content = datesOfPeriods[2].dateSPH;
                #endregion

                #region Загрузка данных в "Формирование документов" и "Сведения об учащихся"
                List<Student> allStudentsGroup = DataService.GetStudentGruppa(Gr);

                lbGroupListToCreateDocuments.ItemsSource = allStudentsGroup;
                lbGroupListForStudentDetails.ItemsSource = allStudentsGroup;

                cbRelationshipKind.ItemsSource = DataService.GetRelativeForms();
                cbFamilyTipe.ItemsSource = DataService.GetAllTypesFamily();
                cbActiveSector.ItemsSource = DataService.GetAllActiveSectors();

                cbPromotionPunishCategory.ItemsSource = DataService.GetAllPromotionPunishCategory();
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка загрузки данных!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool Warning(string warningString)
        {
            MessageBox.Show(warningString, "Предупреждене", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        private bool Information(string inforationString)
        {
            MessageBox.Show(inforationString, "Внимание", MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }

        #region Учёт пропусков
        #region Класс для ListBox
        public class ListStudents
        {
            private bool select;
            private Object obj;

            public bool Select
            {
                get { return select; }
                set { select = value; }
            }
            public Object Obj
            {
                get { return obj; }
                set { obj = value; }
            }
        }
        #endregion

        #region Для того что бы удалять несколько записей
        public class SkipStudents
        {
            private bool select;
            private Object obj;

            public bool Select
            {
                get {  return select; }
                set { select = value; }
            }
            public Object Obj
            {
                get { return obj; }
                set { obj = value; }
            }
        }
        #endregion

        #region Режим просмотра
        #region Календарь 
        private void calendar1_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            Mouse.Capture(null);
        }

        //Фильтр по месяцу в Datagrid через календарь
        private void calendar1_DisplayDateChanged(object sender, CalendarDateChangedEventArgs e)
        {
            if (IsLoaded && calendar1.DisplayDate != null)
            {
                #region Загрузка данных в DataGrid просмотр
                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, calendar1.DisplayDate.ToString("MM"), calendar1.DisplayDate.ToString("yyyy")))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipView.ItemsSource = skipStudents;
                #endregion
            }
        }
        #endregion

        #region Buttom | Отчет за месяц
        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            #region Переменные дат
            var today = calendar1.DisplayDate;
            int daysCount = DateTime.DaysInMonth(today.Year, today.Month);

            var first = new DateTime(today.Year, today.Month, 1);
            var last = new DateTime(today.Year, today.Month, daysCount);
            #endregion
           
            try
            {
                OtchetZaMesyaz.Foreground = new SolidColorBrush(Colors.White);
                OtchetZaMesyaz.Content = "Отчет формируется...";

                int countprBar = 0;
                pgProgress.Visibility = Visibility.Visible;

                await Task.Delay(1);
                List<Student> listStudents = new List<Student>(DataService.GetStudentGruppa(Gr));

                pgProgress.Maximum = 164;

                List<SkipStudentView> skipStudentViews = new List<SkipStudentView>(DataService.GetPropuskGruppaZaMes(Gr, calendar1.DisplayDate.ToString("MM")));

                //Объявляем приложение
                Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
                //Отобразить Excel
                ex.Visible = false;
                // ex.Windows.Application.Visible = false;
                //Количество листов в рабочей книге
                ex.SheetsInNewWorkbook = 1;
                //Добавить рабочую книгу
                Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
                //Отключить отображение окон с сообщениями

                ex.DisplayAlerts = false;
                //Получаем первый лист документа (счет начинается с 1)
                Excel.Worksheet page1 = (Excel.Worksheet)ex.Worksheets.get_Item(1);
                page1.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                #region Ведомость за месяц
                //Название листа (вкладки снизу)
                page1.Name = "Отчёт за месяц";
                Excel.Range range1 = (Excel.Range)page1.Range["A1", "AL1"];
                range1.Cells.Merge();
                range1.Rows[1].RowHeight = 19.50;
                page1.Cells[1, 1] = String.Format($"Ведомость учета пропусков занятий за {calendar1.DisplayDate.ToString("MMMM")} {calendar1.DisplayDate.ToString("yyyy")} учебного года Группа {DataService.FIOandGroupTeacher(Gr, Cnc)[0].Name_Group}");
                range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range1.Font.Size = 14;
                range1.Cells.Font.Name = "Calibri";

                Excel.Range range2 = (Excel.Range)page1.Range["A2", "AL2"];
                range2.HorizontalAlignment = Excel.Constants.xlLeft;

                Excel.Range fonts = (Excel.Range)page1.Range["AG2", "AL2"];
                fonts.Orientation = 90;
                page1.Cells[2, 33] = String.Format("По болезни");
                page1.Cells[2, 34] = String.Format("По заяв.");
                page1.Cells[2, 35] = String.Format("Сл. зап.");
                page1.Cells[2, 36] = String.Format("Прочие");
                page1.Cells[2, 37] = String.Format("Неуваж");
                page1.Cells[2, 38] = String.Format("Всего");

                range1.Rows[2].RowHeight = 48;
                range1.Columns[1].ColumnWidth = 13.50;
                page1.Cells[2, 1] = String.Format("№ Учащийся");
                int d = 1;
                for (int i = 2; i <= 38; i++)
                {
                    countprBar++;

                    pgProgress.Value = pgProgress.Value + 1;
                    await Task.Delay(1);

                    range1.Columns[i].ColumnWidth = 2.43;
                    if (i <= 32)
                    {
                        page1.Cells[2, i] = String.Format(d.ToString());
                        d++;
                    }
                }
                for (int i = 3; i <= 39; i++)
                {
                    countprBar++;

                    pgProgress.Value = pgProgress.Value + 1;
                    await Task.Delay(1);

                    range1.Rows[i].RowHeight = 11.25;
                }

                //Заполенние учащихсся
                int dd = 3;
                int s = 1;
                for (int d1 = 0; d1 < DataService.GetStudentGruppa(Gr).Count; d1++)
                {
                    page1.Cells[dd, 1] = String.Format($"{s}    {DataService.GetStudentGruppa(Gr)[d1].Uchashchiysya}"); ;
                    dd++;
                    s++;

                    countprBar++;

                    pgProgress.Value = pgProgress.Value + 1;
                    await Task.Delay(1);
                }
                page1.Cells[dd, 1] = String.Format($"Итоги за {calendar1.DisplayDate.ToString("MMMM")}");

                page1.Cells[dd + 1, 2] = String.Format("б-");
                page1.Cells[dd + 1, 3] = String.Format("по болезни (справка, талон,больничный)");
                page1.Cells[dd + 1, 13] = String.Format("з-");
                page1.Cells[dd + 1, 14] = String.Format("по заявлению;");
                page1.Cells[dd + 1, 18] = String.Format("с-");
                page1.Cells[dd + 1, 19] = String.Format("служебная записка;");
                page1.Cells[dd + 1, 24] = String.Format("п-");
                page1.Cells[dd + 1, 25] = String.Format("прочие причины;");
                page1.Cells[dd + 1, 29] = String.Format("н-");
                page1.Cells[dd + 1, 30] = String.Format("неуважительная причина");

                Excel.Range fonts1 = (Excel.Range)page1.Range[page1.Cells[2, 1], $"AL{dd + 1}"];
                fonts1.Font.Size = 8;
                fonts1.Cells.Font.Name = "Calibri";

                Excel.Range settings1 = (Excel.Range)page1.Range[page1.Cells[dd + 2, 1], $"AL{dd + 2}"];
                page1.Cells[dd + 2, 1] = String.Format($"Дата сдачи: __________________   Куратор: __________________{DataService.FIOandGroupTeacher(Gr, Cnc)[0].SurName} {DataService.FIOandGroupTeacher(Gr, Cnc)[0].Name} {DataService.FIOandGroupTeacher(Gr, Cnc)[0].FatherName} ");
                settings1.Cells.Merge();
                settings1.Font.Size = 11;
                settings1.Cells.Font.Name = "Calibri";
                range1.Rows[dd + 2].RowHeight = 15;

                Excel.Range border = (Excel.Range)page1.Range["A2", $"AL{dd}"];
                border.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                border.HorizontalAlignment = Excel.Constants.xlLeft;

                Excel.Range border1 = (Excel.Range)page1.Range["AG2", "AL2"];
                border1.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                border1.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                border1.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                border1.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                border1.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                border1.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                border1.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                Excel.Range border2 = (Excel.Range)page1.Range[$"B{dd + 1}", $"B{dd + 1}"];
                border2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                border2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                border2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                border2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                border2.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                Excel.Range border3 = (Excel.Range)page1.Range[$"M{dd + 1}", $"M{dd + 1}"];
                border3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                border3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                border3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                border3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                border3.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                Excel.Range border4 = (Excel.Range)page1.Range[$"R{dd + 1}", $"R{dd + 1}"];
                border4.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                border4.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                border4.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                border4.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                border4.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                Excel.Range border5 = (Excel.Range)page1.Range[$"X{dd + 1}", $"X{dd + 1}"];
                border5.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                border5.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                border5.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                border5.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                border5.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                System.Drawing.ColorConverter cc = new System.Drawing.ColorConverter();
                Excel.Range border6 = (Excel.Range)page1.Range[$"AC{dd + 1}", $"AC{dd + 1}"];
                border6.Interior.Color = ColorTranslator.ToOle((System.Drawing.Color)cc.ConvertFromString("#fabf8f"));
                border6.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                border6.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                border6.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                border6.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                border6.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                Excel.Range border7 = (Excel.Range)page1.Range["AK2", $"AK{dd}"];

                border7.Interior.Color = ColorTranslator.ToOle((System.Drawing.Color)cc.ConvertFromString("#fabf8f")); //FromArgb(250, 191, 143); 

                Excel.Range merg = (Excel.Range)page1.Range[$"C{dd + 1}", $"L{dd + 1}"];
                merg.Cells.Merge();
                Excel.Range merg1 = (Excel.Range)page1.Range[$"N{dd + 1}", $"Q{dd + 1}"];
                merg1.Cells.Merge();
                Excel.Range merg2 = (Excel.Range)page1.Range[$"S{dd + 1}", $"W{dd + 1}"];
                merg2.Cells.Merge();
                Excel.Range merg3 = (Excel.Range)page1.Range[$"AD{dd + 1}", $"AI{dd + 1}"];
                merg3.Cells.Merge();

                int kk = 0;
                int studentID = 0;

                for (int i = 3; i < listStudents.Count + 3; i++)
                {
                    countprBar++;

                    pgProgress.Value = pgProgress.Value + 1;
                    await Task.Delay(1);
                    for (int j = 2; j < daysCount + 2; j++)
                    {
                        List<SkipStudentView> skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && item.Date == first.AddDays(kk)));

                        #region Вывод пропусков за дни
                        if (skip == null)
                        {
                            kk++;
                            continue;
                        }
                        if (skip.Count == 1)
                        {
                            page1.Cells[i, j] = String.Format($"{skip[0].Count_hour}{skip[0].SmallName}");
                        }
                        if (skip.Count >= 2)
                        {
                            page1.Cells[i, j] = String.Format($"{skip[0].Count_hour}{skip[0].SmallName}\n{skip[1].Count_hour}{skip[1].SmallName}");
                            range1.Rows[i].RowHeight = 22.50;
                        }
                        #endregion
                        kk++;
                    }
                    kk = 0;
                    studentID++;
                }

                int sumVsego1 = 0;
                int sumVsego2 = 0;
                int sumVsego3 = 0;
                int sumVsego4 = 0;
                int sumVsego5 = 0;
                int sumVsego6 = 0;
                int sumVsegoZaMes1 = 0;
                int sumVsegoZaMes2 = 0;
                int sumVsegoZaMes3 = 0;
                int sumVsegoZaMes4 = 0;
                int sumVsegoZaMes5 = 0;
                int sumVsegoZaMes6 = 0;
                studentID = 0;

                for (int i = 3; i < listStudents.Count + 3; i++)
                {
                    countprBar++;

                    pgProgress.Value = pgProgress.Value + 1;
                    await Task.Delay(1);

                    //Всего
                    List<SkipStudentView> skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last)));
                    sumVsego1 = 0;
                    for (int a = 0; a < skip.Count; a++)
                    {
                        sumVsego1 = sumVsego1 + Convert.ToInt32(skip[a].Count_hour);
                        sumVsegoZaMes1 = sumVsegoZaMes1 + Convert.ToInt32(skip[a].Count_hour);
                    }
                    page1.Cells[i, 38] = sumVsego1.ToString();
                    skip.Clear();

                    //Неуважительная
                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last) && item.SmallName == "н"));
                    sumVsego2 = 0;
                    for (int a = 0; a < skip.Count; a++)
                    {
                        sumVsego2 = sumVsego2 + Convert.ToInt32(skip[a].Count_hour);
                        sumVsegoZaMes2 = sumVsegoZaMes2 + Convert.ToInt32(skip[a].Count_hour);
                    }
                    page1.Cells[i, 37] = sumVsego2.ToString();
                    skip.Clear();

                    //Прочие
                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last) && item.SmallName == "п"));
                    sumVsego3 = 0;
                    for (int a = 0; a < skip.Count; a++)
                    {
                        sumVsego3 = sumVsego3 + Convert.ToInt32(skip[a].Count_hour);
                        sumVsegoZaMes3 = sumVsegoZaMes3 + Convert.ToInt32(skip[a].Count_hour);
                    }
                    page1.Cells[i, 36] = sumVsego3.ToString();
                    skip.Clear();

                    //Сл.зап
                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last) && item.SmallName == "с"));
                    sumVsego4 = 0;
                    for (int a = 0; a < skip.Count; a++)
                    {
                        sumVsego4 = sumVsego4 + Convert.ToInt32(skip[a].Count_hour);
                        sumVsegoZaMes4 = sumVsegoZaMes4 + Convert.ToInt32(skip[a].Count_hour);
                    }
                    page1.Cells[i, 35] = sumVsego4.ToString();
                    skip.Clear();

                    //По. зая
                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last) && item.SmallName == "з"));
                    sumVsego5 = 0;
                    for (int a = 0; a < skip.Count; a++)
                    {
                        sumVsego5 = sumVsego5 + Convert.ToInt32(skip[a].Count_hour);
                        sumVsegoZaMes5 = sumVsegoZaMes5 + Convert.ToInt32(skip[a].Count_hour);
                    }
                    page1.Cells[i, 34] = sumVsego5.ToString();
                    skip.Clear();

                    //По болезни
                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last) && item.SmallName == "б"));
                    sumVsego6 = 0;
                    for (int a = 0; a < skip.Count; a++)
                    {
                        sumVsego6 = sumVsego6 + Convert.ToInt32(skip[a].Count_hour);
                        sumVsegoZaMes6 = sumVsegoZaMes6 + Convert.ToInt32(skip[a].Count_hour);
                    }
                    page1.Cells[i, 33] = sumVsego6.ToString();
                    skip.Clear();

                    studentID++;
                }

                page1.Cells[dd, 38] = sumVsegoZaMes1.ToString();
                page1.Cells[dd, 37] = sumVsegoZaMes2.ToString();
                page1.Cells[dd, 36] = sumVsegoZaMes3.ToString();
                page1.Cells[dd, 35] = sumVsegoZaMes4.ToString();
                page1.Cells[dd, 34] = sumVsegoZaMes5.ToString();
                page1.Cells[dd, 33] = sumVsegoZaMes6.ToString();
                #endregion

                ex.Visible = true;
                ex.WindowState = Excel.XlWindowState.xlMaximized;

                await Task.Delay(1);

                MessageBoxResult result = MessageBox.Show("Отчет сформирован!", "Информация!", MessageBoxButton.OK, MessageBoxImage.Information);
                if (result == MessageBoxResult.OK)
                {
                    pgProgress.Visibility = Visibility.Hidden;
                    pgProgress.Value = 0;
                }

                var bc = new BrushConverter();
                OtchetZaMesyaz.Foreground = (System.Windows.Media.Brush)bc.ConvertFrom("#FF329A93");
                OtchetZaMesyaz.Content = "Сформировать отчет за месяц";

            }
            catch (Exception ex)
            {
                var bc = new BrushConverter();
                OtchetZaMesyaz.Foreground = (System.Windows.Media.Brush)bc.ConvertFrom("#FF329A93");
                OtchetZaMesyaz.Content = "Сформировать отчет за месяц";
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region Button | Сводный отчет и анализ
        static bool Predicat(char chr)
        {
            if (chr == '_') return true;
            else return false;
        }

        private async void Consolidated(object sender, RoutedEventArgs e)
        {
            try
            {
                #region textbox проверка на год
                char[] chars = tbUchebniyGod.Text.ToCharArray();
                char[] findChars = Array.FindAll(chars, Predicat);
                #endregion
                List<Student> listStudents = new List<Student>(DataService.GetStudentGruppa(Gr));

                if (false != cbSeptember.IsChecked ||
                    false != cbOctober.IsChecked ||
                    false != cbNovember.IsChecked ||
                    false != cbDecember.IsChecked ||
                    false != cbJanuary.IsChecked ||
                    false != cbFebruary.IsChecked ||
                    false != cbMarch.IsChecked ||
                    false != cbApril.IsChecked ||
                    false != cbMay.IsChecked ||
                    false != cbJune.IsChecked ||
                    false != cbJuly.IsChecked ||
                    false != cbAugust.IsChecked)
                {

                    #region Создания блока месяца
                    int m = 0;
                    int September = 0;
                    int October = 0;
                    int November = 0;
                    int December = 0;
                    int January = 0;
                    int February = 0;
                    int March = 0;
                    int April = 0;
                    int May = 0;
                    int June = 0;
                    int July = 0;
                    int August = 0;

                    if (cbSeptember.IsChecked == true)
                    {
                        September = 6;
                    }
                    if (cbOctober.IsChecked == true)
                    {
                        October = 6;
                    }
                    if (cbNovember.IsChecked == true)
                    {
                        November = 6;
                    }
                    if (cbDecember.IsChecked == true)
                    {
                        December = 6;
                    }
                    if (cbJanuary.IsChecked == true)
                    {
                        January = 6;
                    }
                    if (cbFebruary.IsChecked == true)
                    {
                        February = 6;
                    }
                    if (cbMarch.IsChecked == true)
                    {
                        March = 6;
                    }
                    if (cbApril.IsChecked == true)
                    {
                        April = 6;
                    }
                    if (cbMay.IsChecked == true)
                    {
                        May = 6;
                    }
                    if (cbJune.IsChecked == true)
                    {
                        June = 6;
                    }
                    if (cbJuly.IsChecked == true)
                    {
                        July = 6;
                    }
                    if (cbAugust.IsChecked == true)
                    {
                        August = 6;
                    }

                    List<Month> ms = new List<Month>();
                    ms.Add(new Month() { kolvo = September, NameMonth = "Сентябрь", nomerMes = 9 });
                    ms.Add(new Month() { kolvo = October, NameMonth = "Октябрь", nomerMes = 10 });
                    ms.Add(new Month() { kolvo = November, NameMonth = "Ноябрь", nomerMes = 11 });
                    ms.Add(new Month() { kolvo = December, NameMonth = "Декабрь", nomerMes = 12 });
                    ms.Add(new Month() { kolvo = January, NameMonth = "Январь", nomerMes = 1 });
                    ms.Add(new Month() { kolvo = February, NameMonth = "Февраль", nomerMes = 2 });
                    ms.Add(new Month() { kolvo = March, NameMonth = "Март", nomerMes = 3 });
                    ms.Add(new Month() { kolvo = April, NameMonth = "Апрель", nomerMes = 4 });
                    ms.Add(new Month() { kolvo = May, NameMonth = "Май", nomerMes = 5 });
                    ms.Add(new Month() { kolvo = June, NameMonth = "Июнь", nomerMes = 6 });
                    ms.Add(new Month() { kolvo = July, NameMonth = "Июль", nomerMes = 7 });
                    ms.Add(new Month() { kolvo = August, NameMonth = "Август", nomerMes = 8 });

                    m = September + October + November + December + January + February + March + April + May + June + July + August;
                    #endregion

                    if (findChars.Length == 0)
                    {
                        int countprBar = 0;
                        pgProgress.Visibility = Visibility.Visible;
                        pgProgress.Maximum = 84;

                        int block = m + 2;

                        //Объявляем приложение
                        Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
                        //Отобразить Excel
                        ex.Visible = false;
                        //Количество листов в рабочей книге
                        ex.SheetsInNewWorkbook = 2;
                        //Добавить рабочую книгу
                        Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
                        //Отключить отображение окон с сообщениями
                        ex.DisplayAlerts = false;

                        //Получаем первый лист документа (счет начинается с 1)
                        Excel.Worksheet page1 = (Excel.Worksheet)ex.Worksheets.get_Item(1);
                        page1.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                        System.Drawing.ColorConverter cc = new System.Drawing.ColorConverter();

                        #region Лист 1
                        page1.PageSetup.TopMargin = 20;
                        page1.PageSetup.BottomMargin = 5;

                        //Название листа (вкладки снизу)
                        page1.Name = "Группа";
                        Excel.Range NazvanieStr1 = (Excel.Range)page1.Range["A1", page1.Cells[1, block + 6]];
                        Excel.Range NazvanieStr2 = (Excel.Range)page1.Range["A2", page1.Cells[2, block + 6]];
                        Excel.Range NazvanieStr3 = (Excel.Range)page1.Range["A3", page1.Cells[3, block + 6]];
                        NazvanieStr1.Cells.Merge();
                        NazvanieStr2.Cells.Merge();
                        NazvanieStr3.Cells.Merge();
                        NazvanieStr1.ColumnWidth = 2.57;
                        page1.Rows[1].RowHeight = 15;
                        page1.Rows[2].RowHeight = 15.75;
                        page1.Rows[3].RowHeight = 15.75;
                        page1.Rows[4].RowHeight = 15;
                        page1.Rows[5].RowHeight = 16.50;
                        page1.Rows[6].RowHeight = 16.50;
                        page1.Rows[7].RowHeight = 15;
                        page1.Rows[8].RowHeight = 58.50;
                        page1.Columns[1].ColumnWidth = 2.86;
                        page1.Columns[2].ColumnWidth = 14.14;

                        NazvanieStr1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        NazvanieStr2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        NazvanieStr3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        NazvanieStr1.Cells.Font.Name = "Arial";
                        NazvanieStr1.Font.Size = 11;
                        NazvanieStr1.Font.Bold = true;
                        NazvanieStr2.Cells.Font.Name = "Arial";
                        NazvanieStr2.Font.Size = 12;
                        NazvanieStr2.Font.Bold = true;
                        NazvanieStr3.Cells.Font.Name = "Arial";
                        NazvanieStr3.Font.Size = 12;
                        NazvanieStr3.Font.Bold = true;

                        var str = DataService.GetStudentGruppa(Gr)[0].Name_Group.ToCharArray();
                        page1.Cells[1, 1] = String.Format($"ВЕДОМОСТЬ");
                        page1.Cells[2, 1] = String.Format($"учета пропусков занятий за {tbUchebniyGod.Text} учебный год");
                        page1.Cells[3, 1] = String.Format($"курс {str[1].ToString()} группа {DataService.GetStudentGruppa(Gr)[0].Name_Group}");

                        page1.Cells[7, 1] = String.Format($"№\nп/п");
                        page1.Cells[7, 1].Font.Size = 8;
                        page1.Cells[7, 1].HorizontalAlignment = Excel.Constants.xlCenter;
                        page1.Cells[7, 1].VerticalAlignment = Excel.Constants.xlCenter;
                        page1.Cells[7, 1].Font.Name = "Arial";

                        page1.Cells[7, 2] = String.Format($"ФИО");
                        page1.Cells[7, 2].Font.Size = 8;
                        page1.Cells[7, 2].HorizontalAlignment = Excel.Constants.xlCenter;
                        page1.Cells[7, 2].VerticalAlignment = Excel.Constants.xlCenter;
                        page1.Cells[7, 2].Font.Name = "Arial";

                        Excel.Range ss1 = (Excel.Range)page1.Range["A7", "A8"];
                        ss1.Cells.Merge();
                        Excel.Range ss2 = (Excel.Range)page1.Range["B7", "B8"];
                        ss2.Cells.Merge();
                        Excel.Range ss3 = (Excel.Range)page1.Range["A9", "B8"];
                        ss2.Cells.Merge();

                        Excel.Range borderTitul = (Excel.Range)page1.Range["A6", "B6"];
                        borderTitul.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
                        borderTitul.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                        borderTitul.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                        borderTitul.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;

                        Excel.Range borderTitul2 = (Excel.Range)page1.Range["A7", "B8"];
                        borderTitul2.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
                        borderTitul2.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                        borderTitul2.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                        borderTitul2.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                        borderTitul2.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                        int dd = 9;
                        int s = 1;
                        for (int d1 = 0; d1 < listStudents.Count; d1++)
                        {
                            countprBar++;

                            pgProgress.Value = pgProgress.Value + 1;
                            await Task.Delay(1);

                            page1.Cells[dd, 1] = String.Format($"{s}");
                            page1.Cells[dd, 1].HorizontalAlignment = Excel.Constants.xlCenter;
                            page1.Cells[dd, 1].VerticalAlignment = Excel.Constants.xlCenter;
                            page1.Cells[dd, 2] = String.Format($"{listStudents[d1].Uchashchiysya}");
                            dd++;
                            s++;
                        }

                        Excel.Range borderTitul3 = (Excel.Range)page1.Range["A7", page1.Cells[dd, 2]];
                        borderTitul3.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
                        borderTitul3.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                        borderTitul3.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                        borderTitul3.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                        borderTitul3.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                        borderTitul3.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                        #region Блок месяца
                        int blockMes = 2;
                        for (int i = 0; i < ms.Count; i++)
                        {
                            countprBar++;
                            pgProgress.Value = pgProgress.Value + 1;
                            await Task.Delay(1);

                            if (ms[i].kolvo > 0)
                            {
                                Excel.Range Mesyz = (Excel.Range)page1.Range[page1.Cells[6, blockMes + 1], page1.Cells[6, blockMes + 6]];
                                Mesyz.Cells.Merge();
                                page1.Cells[6, blockMes + 1] = String.Format($"{ms[i].NameMonth}");
                                Mesyz.HorizontalAlignment = Excel.Constants.xlCenter;
                                Mesyz.Font.Size = 12;
                                Mesyz.Font.Bold = true;
                                Mesyz.Cells.Font.Name = "Arial";
                                Mesyz.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                                Excel.Range uvzprich = (Excel.Range)page1.Range[page1.Cells[7, blockMes + 1], page1.Cells[7, blockMes + 4]];
                                uvzprich.Cells.Merge();
                                page1.Cells[7, blockMes + 1] = String.Format($"ув. причины");
                                uvzprich.Cells.Font.Name = "Arial";
                                uvzprich.Font.Size = 8;
                                uvzprich.HorizontalAlignment = Excel.Constants.xlCenter;
                                uvzprich.VerticalAlignment = Excel.Constants.xlCenter;

                                page1.Cells[8, blockMes + 1] = String.Format($"по болезни");
                                page1.Cells[8, blockMes + 1].Orientation = 90;
                                page1.Cells[8, blockMes + 1].Font.Size = 8;
                                page1.Cells[8, blockMes + 1].Font.Name = "Arial";
                                page1.Cells[8, blockMes + 1].HorizontalAlignment = Excel.Constants.xlCenter;
                                page1.Cells[8, blockMes + 1].VerticalAlignment = Excel.Constants.xlCenter;

                                page1.Cells[8, blockMes + 2] = String.Format($"по заявлению");
                                page1.Cells[8, blockMes + 2].Orientation = 90;
                                page1.Cells[8, blockMes + 2].Font.Size = 8;
                                page1.Cells[8, blockMes + 2].Font.Name = "Arial";
                                page1.Cells[8, blockMes + 2].HorizontalAlignment = Excel.Constants.xlCenter;
                                page1.Cells[8, blockMes + 2].VerticalAlignment = Excel.Constants.xlCenter;

                                page1.Cells[8, blockMes + 3] = String.Format($"сл. записка");
                                page1.Cells[8, blockMes + 3].Orientation = 90;
                                page1.Cells[8, blockMes + 3].Font.Size = 8;
                                page1.Cells[8, blockMes + 3].Font.Name = "Arial";
                                page1.Cells[8, blockMes + 3].HorizontalAlignment = Excel.Constants.xlCenter;
                                page1.Cells[8, blockMes + 3].VerticalAlignment = Excel.Constants.xlCenter;

                                page1.Cells[8, blockMes + 4] = String.Format($"прочие");
                                page1.Cells[8, blockMes + 4].Orientation = 90;
                                page1.Cells[8, blockMes + 4].Font.Size = 8;
                                page1.Cells[8, blockMes + 4].Font.Name = "Arial";
                                page1.Cells[8, blockMes + 4].HorizontalAlignment = Excel.Constants.xlCenter;
                                page1.Cells[8, blockMes + 4].VerticalAlignment = Excel.Constants.xlCenter;

                                Excel.Range Mesyz1 = (Excel.Range)page1.Range[page1.Cells[8, blockMes + 5], page1.Cells[7, blockMes + 5]];
                                Mesyz1.Cells.Merge();
                                page1.Cells[7, blockMes + 5] = String.Format($"неуваж. причина");
                                Mesyz1.Cells.Font.Name = "Arial";
                                Mesyz1.Font.Size = 8;
                                Mesyz1.Orientation = 90;
                                Mesyz1.HorizontalAlignment = Excel.Constants.xlCenter;
                                Mesyz1.VerticalAlignment = Excel.Constants.xlCenter;
                                Excel.Range Mesyz11 = (Excel.Range)page1.Range[page1.Cells[dd, blockMes + 5], page1.Cells[9, blockMes + 5]];
                                Mesyz11.Cells.Font.Name = "Arial";
                                Mesyz11.Font.Size = 8;
                                Mesyz11.Interior.Color = ColorTranslator.ToOle((System.Drawing.Color)cc.ConvertFromString("#d1ffd1"));
                                Mesyz11.HorizontalAlignment = Excel.Constants.xlCenter;
                                Mesyz11.VerticalAlignment = Excel.Constants.xlCenter;

                                Excel.Range Mesyz2 = (Excel.Range)page1.Range[page1.Cells[8, blockMes + 6], page1.Cells[7, blockMes + 6]];
                                Mesyz2.Cells.Merge();
                                page1.Cells[7, blockMes + 6] = String.Format($"всего");
                                Mesyz2.Cells.Font.Name = "Arial";
                                Mesyz2.Font.Size = 8;
                                Mesyz2.Orientation = 90;
                                Mesyz2.HorizontalAlignment = Excel.Constants.xlCenter;
                                Mesyz2.VerticalAlignment = Excel.Constants.xlCenter;

                                Excel.Range cvet1 = (Excel.Range)page1.Range[page1.Cells[dd, blockMes + 6], page1.Cells[7, blockMes + 6]];
                                cvet1.Cells.Font.Name = "Arial";
                                cvet1.Font.Size = 8;
                                cvet1.HorizontalAlignment = Excel.Constants.xlCenter;
                                cvet1.VerticalAlignment = Excel.Constants.xlCenter;
                                cvet1.Interior.Color = ColorTranslator.ToOle((System.Drawing.Color)cc.ConvertFromString("#fabf8f"));
                                cvet1.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                                cvet1.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
                                cvet1.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                                cvet1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;

                                Excel.Range cvet2 = (Excel.Range)page1.Range[page1.Cells[8, blockMes + 6], page1.Cells[7, blockMes + 6]];
                                cvet2.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                                cvet2.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
                                cvet2.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;



                                Excel.Range border3 = (Excel.Range)page1.Range[page1.Cells[dd, blockMes + 5], page1.Cells[9, blockMes + 1]];
                                border3.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                                border3.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
                                border3.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                                border3.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                                border3.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                                Excel.Range border4 = (Excel.Range)page1.Range[page1.Cells[7, blockMes + 1], page1.Cells[8, blockMes + 5]];
                                border4.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                                border4.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
                                border4.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                                border4.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                                border4.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                                int sumVsego1 = 0;
                                int sumVsego2 = 0;
                                int sumVsego3 = 0;
                                int sumVsego4 = 0;
                                int sumVsego5 = 0;
                                int sumVsego6 = 0;
                                int sumVsegoZaMes1 = 0;
                                int sumVsegoZaMes2 = 0;
                                int sumVsegoZaMes3 = 0;
                                int sumVsegoZaMes4 = 0;
                                int sumVsegoZaMes5 = 0;
                                int sumVsegoZaMes6 = 0;
                                int studentID = 0;

                                string god;
                                DateTime nowDays = DateTime.Now;
                                if (ms[i].nomerMes < 9)
                                {
                                    god = $"{nowDays.Year}";
                                }
                                else
                                {
                                    god = $"{nowDays.Year - 1}";
                                }
                                
                                int daysCount = DateTime.DaysInMonth(Convert.ToInt32(god), ms[i].nomerMes);
                                var first = new DateTime(Convert.ToInt32(god), ms[i].nomerMes, 1);
                                var last = new DateTime(Convert.ToInt32(god), ms[i].nomerMes, daysCount);

                                List<SkipStudentView> skipStudentViews = new List<SkipStudentView>(DataService.GetPropuskGruppaZaMes(Gr, ms[i].nomerMes.ToString()));

                                for (int aas = 9; aas < listStudents.Count + 9; aas++)
                                {
                                    //Всего
                                    List<SkipStudentView> skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last)));
                                    sumVsego1 = 0;
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego1 = sumVsego1 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoZaMes1 = sumVsegoZaMes1 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    page1.Cells[aas, blockMes + 6] = sumVsego1.ToString();
                                    skip.Clear();

                                    //Неуважительная
                                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last) && item.SmallName == "н"));
                                    sumVsego2 = 0;
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego2 = sumVsego2 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoZaMes2 = sumVsegoZaMes2 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    page1.Cells[aas, blockMes + 5] = sumVsego2.ToString();
                                    skip.Clear();

                                    //Прочие
                                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last) && item.SmallName == "п"));
                                    sumVsego3 = 0;
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego3 = sumVsego3 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoZaMes3 = sumVsegoZaMes3 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    page1.Cells[aas, blockMes + 4] = sumVsego3.ToString();
                                    skip.Clear();

                                    //Сл.зап
                                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last) && item.SmallName == "с"));
                                    sumVsego4 = 0;
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego4 = sumVsego4 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoZaMes4 = sumVsegoZaMes4 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    page1.Cells[aas, blockMes + 3] = sumVsego4.ToString();
                                    skip.Clear();

                                    //По. зая
                                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last) && item.SmallName == "з"));
                                    sumVsego5 = 0;
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego5 = sumVsego5 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoZaMes5 = sumVsegoZaMes5 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    page1.Cells[aas, blockMes + 2] = sumVsego5.ToString();
                                    skip.Clear();

                                    //По болезни
                                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentID].cn_S && (item.Date >= first && item.Date <= last) && item.SmallName == "б"));
                                    sumVsego6 = 0;
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego6 = sumVsego6 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoZaMes6 = sumVsegoZaMes6 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    page1.Cells[aas, blockMes + 1] = sumVsego6.ToString();
                                    skip.Clear();

                                    studentID++;
                                }

                                page1.Cells[dd, blockMes + 6] = sumVsegoZaMes1.ToString();
                                page1.Cells[dd, blockMes + 5] = sumVsegoZaMes2.ToString();
                                page1.Cells[dd, blockMes + 4] = sumVsegoZaMes3.ToString();
                                page1.Cells[dd, blockMes + 3] = sumVsegoZaMes4.ToString();
                                page1.Cells[dd, blockMes + 2] = sumVsegoZaMes5.ToString();
                                page1.Cells[dd, blockMes + 1] = sumVsegoZaMes6.ToString();
                                blockMes = blockMes + 6;
                            }
                        }

                        Excel.Range range = (Excel.Range)page1.Range["A9", page1.Cells[dd, block + 6]];
                        range.Rows.RowHeight = 12;
                        range.Font.Size = 8;
                        range.Font.Name = "Arial";
                        #endregion

                        #region Блок __-й семестр
                        Excel.Range Semest1 = (Excel.Range)page1.Range[page1.Cells[6, block + 1], page1.Cells[6, block + 6]];
                        Semest1.Cells.Merge();
                        ComboBoxItem typeItem = (ComboBoxItem)cbSemestr.SelectedItem;
                        page1.Cells[6, block + 1] = String.Format($"{typeItem.Content.ToString()}-й семестр");
                        Semest1.HorizontalAlignment = Excel.Constants.xlCenter;
                        Semest1.Font.Size = 12;
                        Semest1.Font.Bold = true;
                        Semest1.Cells.Font.Name = "Arial";
                        Semest1.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                        Excel.Range Semest2 = (Excel.Range)page1.Range[page1.Cells[7, block + 1], page1.Cells[7, block + 4]];
                        Semest2.Cells.Merge();
                        page1.Cells[7, block + 1] = String.Format($"ув. причины");
                        Semest2.Cells.Font.Name = "Arial";
                        Semest2.Font.Size = 8;
                        Semest2.HorizontalAlignment = Excel.Constants.xlCenter;
                        Semest2.VerticalAlignment = Excel.Constants.xlCenter;

                        page1.Cells[8, block + 1] = String.Format($"по болезни");
                        page1.Cells[8, block + 1].Orientation = 90;
                        page1.Cells[8, block + 1].Font.Size = 8;
                        page1.Cells[8, block + 1].Font.Name = "Arial";
                        page1.Cells[8, block + 1].HorizontalAlignment = Excel.Constants.xlCenter;
                        page1.Cells[8, block + 1].VerticalAlignment = Excel.Constants.xlCenter;

                        page1.Cells[8, block + 2] = String.Format($"по заявлению");
                        page1.Cells[8, block + 2].Orientation = 90;
                        page1.Cells[8, block + 2].Font.Size = 8;
                        page1.Cells[8, block + 2].Font.Name = "Arial";
                        page1.Cells[8, block + 2].HorizontalAlignment = Excel.Constants.xlCenter;
                        page1.Cells[8, block + 2].VerticalAlignment = Excel.Constants.xlCenter;

                        page1.Cells[8, block + 3] = String.Format($"сл. записка");
                        page1.Cells[8, block + 3].Orientation = 90;
                        page1.Cells[8, block + 3].Font.Size = 8;
                        page1.Cells[8, block + 3].Font.Name = "Arial";
                        page1.Cells[8, block + 3].HorizontalAlignment = Excel.Constants.xlCenter;
                        page1.Cells[8, block + 3].VerticalAlignment = Excel.Constants.xlCenter;

                        page1.Cells[8, block + 4] = String.Format($"прочие");
                        page1.Cells[8, block + 4].Orientation = 90;
                        page1.Cells[8, block + 4].Font.Size = 8;
                        page1.Cells[8, block + 4].Font.Name = "Arial";
                        page1.Cells[8, block + 4].HorizontalAlignment = Excel.Constants.xlCenter;
                        page1.Cells[8, block + 4].VerticalAlignment = Excel.Constants.xlCenter;

                        Excel.Range Semest3 = (Excel.Range)page1.Range[page1.Cells[8, block + 5], page1.Cells[7, block + 5]];
                        Semest3.Cells.Merge();
                        page1.Cells[7, block + 5] = String.Format($"неуваж. причина");
                        Semest3.Cells.Font.Name = "Arial";
                        Semest3.Font.Size = 8;
                        Semest3.Orientation = 90;
                        Semest3.HorizontalAlignment = Excel.Constants.xlCenter;
                        Semest3.VerticalAlignment = Excel.Constants.xlCenter;

                        Excel.Range Semest31 = (Excel.Range)page1.Range[page1.Cells[dd, block + 5], page1.Cells[9, block + 5]];
                        Semest31.Cells.Font.Name = "Arial";
                        Semest31.Font.Size = 8;
                        Semest31.Interior.Color = ColorTranslator.ToOle((System.Drawing.Color)cc.ConvertFromString("#d1ffd1"));

                        Excel.Range Semest4 = (Excel.Range)page1.Range[page1.Cells[8, block + 6], page1.Cells[7, block + 6]];
                        Semest4.Cells.Merge();
                        page1.Cells[7, block + 6] = String.Format($"всего");
                        Semest4.Cells.Font.Name = "Arial";
                        Semest4.Font.Size = 8;
                        Semest4.Orientation = 90;
                        Semest4.HorizontalAlignment = Excel.Constants.xlCenter;
                        Semest4.VerticalAlignment = Excel.Constants.xlCenter;

                        Excel.Range border = (Excel.Range)page1.Range[page1.Cells[dd, block + 6], page1.Cells[9, block + 1]];
                        border.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                        border.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                        border.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                        border.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;

                        Excel.Range border1 = (Excel.Range)page1.Range[page1.Cells[7, block + 1], page1.Cells[8, block + 5]];
                        border1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        border1.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium;

                        Excel.Range cvet3 = (Excel.Range)page1.Range[$"A{dd}", page1.Cells[dd, block + 6]];
                        cvet3.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                        cvet3.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
                        cvet3.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                        cvet3.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                        cvet3.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                        cvet3.Interior.Color = ColorTranslator.ToOle((System.Drawing.Color)cc.ConvertFromString("#FABF8F"));

                        Excel.Range cvet = (Excel.Range)page1.Range[page1.Cells[dd, block + 6], page1.Cells[7, block + 6]];
                        cvet.Cells.Font.Name = "Arial";
                        cvet.Font.Size = 8;
                        cvet.HorizontalAlignment = Excel.Constants.xlCenter;
                        cvet.VerticalAlignment = Excel.Constants.xlCenter;
                        cvet.Interior.Color = ColorTranslator.ToOle((System.Drawing.Color)cc.ConvertFromString("#FABF8F"));
                        cvet.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                        int studentIDs = 0;
                        int sumVsegoSemestr1 = 0;
                        int sumVsegoSemestr2 = 0;
                        int sumVsegoSemestr3 = 0;
                        int sumVsegoSemestr4 = 0;
                        int sumVsegoSemestr5 = 0;
                        int sumVsegoSemestr6 = 0;

                        for (int aas = 9; aas < listStudents.Count + 9; aas++)
                        {
                            countprBar++;
                            pgProgress.Value = pgProgress.Value + 1;
                            await Task.Delay(1);

                            int sumVsego1 = 0;
                            int sumVsego2 = 0;
                            int sumVsego3 = 0;
                            int sumVsego4 = 0;
                            int sumVsego5 = 0;
                            int sumVsego6 = 0;

                            for (int i = 0; i < ms.Count; i++)
                            {
                                if (ms[i].kolvo > 0)
                                {
                                    string gods;
                                    DateTime nowDays = DateTime.Now;
                                    if (ms[i].nomerMes < 9)
                                    {
                                        gods = $"{nowDays.Year}";
                                    }
                                    else
                                    {
                                        gods = $"{nowDays.Year - 1}";
                                    }

                                    int daysCounts = DateTime.DaysInMonth(Convert.ToInt32(gods), ms[i].nomerMes);
                                    var firsts = new DateTime(Convert.ToInt32(gods), ms[i].nomerMes, 1);
                                    var lasts = new DateTime(Convert.ToInt32(gods), ms[i].nomerMes, daysCounts);

                                    List<SkipStudentView> skipStudentViews = new List<SkipStudentView>(DataService.GetPropuskGruppaZaMes(Gr, ms[i].nomerMes.ToString()));

                                    List<SkipStudentView> skip = new List<SkipStudentView>();

                                    //По болезни
                                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentIDs].cn_S && (item.Date >= firsts && item.Date <= lasts) && item.SmallName == "б"));
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego6 = sumVsego6 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoSemestr6 = sumVsegoSemestr6 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    skip.Clear();

                                    //По. зая
                                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentIDs].cn_S && (item.Date >= firsts && item.Date <= lasts) && item.SmallName == "з"));
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego5 = sumVsego5 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoSemestr5 = sumVsegoSemestr5 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    skip.Clear();

                                    //Сл.зап
                                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentIDs].cn_S && (item.Date >= firsts && item.Date <= lasts) && item.SmallName == "с"));
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego4 = sumVsego4 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoSemestr4 = sumVsegoSemestr4 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    skip.Clear();

                                    //Прочие
                                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentIDs].cn_S && (item.Date >= firsts && item.Date <= lasts) && item.SmallName == "п"));
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego3 = sumVsego3 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoSemestr3 = sumVsegoSemestr3 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    skip.Clear();

                                    //Неуважительная
                                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentIDs].cn_S && (item.Date >= firsts && item.Date <= lasts) && item.SmallName == "н"));
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego2 = sumVsego2 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoSemestr2 = sumVsegoSemestr2 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    skip.Clear();

                                    //Всего
                                    skip = skipStudentViews.FindAll(item => (item.cn_S_Student == listStudents[studentIDs].cn_S && (item.Date >= firsts && item.Date <= lasts)));
                                    for (int a = 0; a < skip.Count; a++)
                                    {
                                        sumVsego1 = sumVsego1 + Convert.ToInt32(skip[a].Count_hour);
                                        sumVsegoSemestr1 = sumVsegoSemestr1 + Convert.ToInt32(skip[a].Count_hour);
                                    }
                                    skip.Clear();
                                }

                                page1.Cells[aas, block + 6] = sumVsego1.ToString();
                                page1.Cells[aas, block + 5] = sumVsego2.ToString();
                                page1.Cells[aas, block + 4] = sumVsego3.ToString();
                                page1.Cells[aas, block + 3] = sumVsego4.ToString();
                                page1.Cells[aas, block + 2] = sumVsego5.ToString();
                                page1.Cells[aas, block + 1] = sumVsego6.ToString();
                            }
                            studentIDs++;
                        }
                        page1.Cells[dd, block + 6] = sumVsegoSemestr1.ToString();
                        page1.Cells[dd, block + 5] = sumVsegoSemestr2.ToString();
                        page1.Cells[dd, block + 4] = sumVsegoSemestr3.ToString();
                        page1.Cells[dd, block + 3] = sumVsegoSemestr4.ToString();
                        page1.Cells[dd, block + 2] = sumVsegoSemestr5.ToString();
                        page1.Cells[dd, block + 1] = sumVsegoSemestr6.ToString();
                        #endregion
                        Excel.Range aligncenter = (Excel.Range)page1.Range[page1.Cells[dd, block + 6], "C9"];
                        aligncenter.HorizontalAlignment = Excel.Constants.xlCenter;
                        aligncenter.VerticalAlignment = Excel.Constants.xlCenter;
                        #endregion

                        #region Лист 2
                        //Получаем первый лист документа (счет начинается с 1)
                        Excel.Worksheet page2 = (Excel.Worksheet)ex.Worksheets.get_Item(2);
                        page2.Select();
                        page2.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                        page2.Name = "Анализ";
                        //  page2.PageSetup.Zoom = 85;

                        page2.Cells[1, 1] = String.Format($"Анализ");
                        page2.Cells[1, 1].Font.Size = 11;
                        page2.Cells[1, 1].Font.Name = "Arial";
                        page2.Cells[1, 1].Font.Bold = true;
                        page2.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
                        page2.Cells[1, 1].VerticalAlignment = Excel.Constants.xlCenter;
                        page2.Rows[1].RowHeight = 15;

                        if (typeItem.Content.ToString() == "Год")
                        {
                            page2.Cells[2, 1] = String.Format($"пропусков занятий за {tbUchebniyGod.Text} учебный год");
                        }
                        else
                        {
                            page2.Cells[2, 1] = String.Format($"пропусков занятий за {tbUchebniyGod.Text} учебный год {typeItem.Content.ToString()}-й семестр");
                        }

                        page2.Cells[2, 1].Font.Size = 12;
                        page2.Cells[2, 1].Font.Name = "Arial";
                        page2.Cells[2, 1].Font.Bold = true;
                        page2.Cells[2, 1].HorizontalAlignment = Excel.Constants.xlCenter;
                        page2.Cells[2, 1].VerticalAlignment = Excel.Constants.xlCenter;
                        page2.Rows[2].RowHeight = 15.75;

                        page2.Cells[3, 1] = String.Format($"курс {str[1].ToString()} группа {DataService.GetStudentGruppa(Gr)[0].Name_Group}");
                        page2.Cells[3, 1].Font.Size = 12;
                        page2.Cells[3, 1].Font.Name = "Arial";
                        page2.Cells[3, 1].Font.Bold = true;
                        page2.Cells[3, 1].HorizontalAlignment = Excel.Constants.xlCenter;
                        page2.Cells[3, 1].VerticalAlignment = Excel.Constants.xlCenter;
                        page2.Rows[3].RowHeight = 15.75;

                        page2.Cells[4, 5] = String.Format($"количество человек: ");
                        page2.Cells[4, 8] = String.Format($"{listStudents.Count}");
                        page2.Cells[4, 5].Font.Size = 12;
                        page2.Cells[4, 8].Font.Size = 11;
                        page2.Cells[4, 5].Font.Name = "Arial";
                        page2.Cells[4, 8].Font.Name = "Arial";
                        page2.Cells[4, 5].HorizontalAlignment = Excel.Constants.xlCenter;
                        page2.Cells[4, 5].VerticalAlignment = Excel.Constants.xlCenter;
                        page2.Cells[4, 8].HorizontalAlignment = Excel.Constants.xlLeft;
                        page2.Rows[4].RowHeight = 15;

                        Excel.Range List2Nazvanie1 = (Excel.Range)page2.Range["A1", "M1"];
                        List2Nazvanie1.Cells.Merge();
                        Excel.Range List2Nazvanie2 = (Excel.Range)page2.Range["A2", "M2"];
                        List2Nazvanie2.Cells.Merge();
                        Excel.Range List2Nazvanie3 = (Excel.Range)page2.Range["A3", "M3"];
                        List2Nazvanie3.Cells.Merge();
                        Excel.Range List2Nazvanie4 = (Excel.Range)page2.Range["E4", "G4"];
                        List2Nazvanie4.Cells.Merge();
                        List2Nazvanie4.Columns.ColumnWidth = 8.43;
                        page2.Columns[2].ColumnWidth = 11.71;
                        page2.Columns[5].ColumnWidth = 9.43;

                        page2.Rows[5].RowHeight = 15;
                        page2.Rows[6].RowHeight = 15.75;
                        page2.Rows[7].RowHeight = 63;

                        Excel.Range List2Nazvanie5 = (Excel.Range)page2.Range["C6", "D6"];
                        List2Nazvanie5.Cells.Merge();
                        page2.Cells[6, 3] = String.Format($"Пропуски");
                        page2.Cells[7, 3] = String.Format($"всего,\nкол-во");
                        page2.Cells[7, 4] = String.Format($"на 1\nуч-ся,\nкол-во");

                        Excel.Range List2Nazvanie6 = (Excel.Range)page2.Range["F6", "I6"];
                        List2Nazvanie6.Cells.Merge();
                        page2.Cells[6, 6] = String.Format($"Из них, кол-во");
                        page2.Cells[7, 6] = String.Format($"по\nболез-\nни");
                        page2.Cells[7, 7] = String.Format($"по\nзаявле-\nнию");
                        page2.Cells[7, 8] = String.Format($"по\nслуже-\nной\nзаписке");
                        page2.Cells[7, 9] = String.Format($"прочие\nпри-\nчины");
                        Excel.Range List2Nazvanie7 = (Excel.Range)page2.Range["J6", "L6"];
                        List2Nazvanie7.Cells.Merge();
                        page2.Cells[6, 10] = String.Format($"Без уважительной");
                        page2.Cells[7, 10] = String.Format($"коли\nчество");
                        page2.Cells[7, 11] = String.Format($"на 1\nуч-ся,\nкол-во");
                        page2.Cells[7, 12] = String.Format($"%");

                        Excel.Range List2Nazvanie8 = (Excel.Range)page2.Range["B6", "B7"];
                        List2Nazvanie8.Cells.Merge();
                        page2.Cells[6, 2] = String.Format($"Проме-\nжуток");

                        Excel.Range List2Nazvanie10 = (Excel.Range)page2.Range["E6", "E7"];
                        List2Nazvanie10.Cells.Merge();
                        page2.Cells[6, 5] = String.Format($"По уважи-\nтельной\nпричине,\nкол-во");

                        int x = 7;

                        double SummaKolVoSemestrSchet = 0;
                        double SummaPoBolesniSemestrSchet = 0;
                        double SummaPoZayavleniySemestrSchet = 0;
                        double SummaPoSluzebZapisSemestrSchet = 0;
                        double SummaProchiePrichiniSemestrSchet = 0;
                        double SummaBezUvazizSemestrSchet = 0;

                        for (int i = 0; i < ms.Count; i++)
                        {
                            countprBar++;
                            pgProgress.Value = pgProgress.Value + 1;
                            await Task.Delay(1);

                            if (ms[i].kolvo > 0)
                            {

                                string gods;
                                DateTime nowDays = DateTime.Now;
                                if (ms[i].nomerMes < 9)
                                {
                                    gods = $"{nowDays.Year}";
                                }
                                else
                                {
                                    gods = $"{nowDays.Year - 1}";
                                }

                                int daysCount = DateTime.DaysInMonth(Convert.ToInt32(gods), ms[i].nomerMes);
                                var first = new DateTime(Convert.ToInt32(gods), ms[i].nomerMes, 1);
                                var last = new DateTime(Convert.ToInt32(gods), ms[i].nomerMes, daysCount);

                                List<SkipStudentView> skipStudentViews = new List<SkipStudentView>(DataService.GetPropuskGruppaZaMes(Gr, ms[i].nomerMes.ToString()));

                                page2.Cells[x + 1, 2] = String.Format($"{ms[i].NameMonth}");

                                //Всего
                                List<SkipStudentView> vsegoKolVo = skipStudentViews.FindAll(item => (item.cn_G_Student == Gr && (item.Date >= first && item.Date <= last)));
                                double vsegoKolVoSchet = 0;
                                for (int kolVo = 0; kolVo < vsegoKolVo.Count; kolVo++)
                                {
                                    vsegoKolVoSchet = vsegoKolVoSchet + Convert.ToInt32(vsegoKolVo[kolVo].Count_hour);
                                    SummaKolVoSemestrSchet = SummaKolVoSemestrSchet + Convert.ToInt32(vsegoKolVo[kolVo].Count_hour);
                                }
                                if (vsegoKolVo.Count != 0)
                                {
                                    double clna1 = vsegoKolVoSchet / Convert.ToDouble(listStudents.Count);
                                    page2.Cells[x + 1, 3] = String.Format($"{vsegoKolVoSchet}");
                                    page2.Cells[x + 1, 4] = String.Format("{0:0.##}", clna1);
                                    clna1 = 0;
                                }
                                else
                                {
                                    page2.Cells[x + 1, 3] = String.Format($"0");
                                    page2.Cells[x + 1, 4] = String.Format("0");
                                }

                                //По болезни
                                List<SkipStudentView> poBolesni = skipStudentViews.FindAll(item => (item.cn_G_Student == Gr && (item.Date >= first && item.Date <= last) && item.SmallName == "б"));
                                double poBolesniSchet = 0;
                                for (int kolVo = 0; kolVo < poBolesni.Count; kolVo++)
                                {
                                    poBolesniSchet = poBolesniSchet + Convert.ToInt32(poBolesni[kolVo].Count_hour);
                                    SummaPoBolesniSemestrSchet = SummaPoBolesniSemestrSchet + Convert.ToInt32(poBolesni[kolVo].Count_hour);
                                }
                                if (poBolesni.Count != 0)
                                {
                                    page2.Cells[x + 1, 6] = String.Format($"{poBolesniSchet}");
                                }
                                else
                                {
                                    page2.Cells[x + 1, 6] = String.Format($"0");
                                }

                                //По заявлению
                                List<SkipStudentView> poZayavleniy = skipStudentViews.FindAll(item => (item.cn_G_Student == Gr && (item.Date >= first && item.Date <= last) && item.SmallName == "з"));
                                double poZayavleniySchet = 0;
                                for (int kolVo = 0; kolVo < poZayavleniy.Count; kolVo++)
                                {
                                    poZayavleniySchet = poZayavleniySchet + Convert.ToInt32(poZayavleniy[kolVo].Count_hour);
                                    SummaPoZayavleniySemestrSchet = SummaPoZayavleniySemestrSchet + Convert.ToInt32(poZayavleniy[kolVo].Count_hour);
                                }
                                if (poZayavleniy.Count != 0)
                                {
                                    page2.Cells[x + 1, 7] = String.Format($"{poZayavleniySchet}");
                                }
                                else
                                {
                                    page2.Cells[x + 1, 7] = String.Format($"0");
                                }

                                //По служебной записке
                                List<SkipStudentView> poSluzebZapis = skipStudentViews.FindAll(item => (item.cn_G_Student == Gr && (item.Date >= first && item.Date <= last) && item.SmallName == "с"));
                                double poSluzebZapisSchet = 0;
                                for (int kolVo = 0; kolVo < poSluzebZapis.Count; kolVo++)
                                {
                                    poSluzebZapisSchet = poSluzebZapisSchet + Convert.ToInt32(poSluzebZapis[kolVo].Count_hour);
                                    SummaPoSluzebZapisSemestrSchet = SummaPoSluzebZapisSemestrSchet + Convert.ToInt32(poSluzebZapis[kolVo].Count_hour);
                                }
                                if (poSluzebZapis.Count != 0)
                                {
                                    page2.Cells[x + 1, 8] = String.Format($"{poSluzebZapisSchet}");
                                }
                                else
                                {
                                    page2.Cells[x + 1, 8] = String.Format($"0");
                                }

                                //Порочие причины
                                List<SkipStudentView> prochiePrichini = skipStudentViews.FindAll(item => (item.cn_G_Student == Gr && (item.Date >= first && item.Date <= last) && item.SmallName == "п"));
                                double prochiePrichiniSchet = 0;
                                for (int kolVo = 0; kolVo < prochiePrichini.Count; kolVo++)
                                {
                                    prochiePrichiniSchet = prochiePrichiniSchet + Convert.ToInt32(prochiePrichini[kolVo].Count_hour);
                                    SummaProchiePrichiniSemestrSchet = SummaProchiePrichiniSemestrSchet + Convert.ToInt32(prochiePrichini[kolVo].Count_hour);
                                }
                                if (prochiePrichini.Count != 0)
                                {
                                    page2.Cells[x + 1, 9] = String.Format($"{prochiePrichiniSchet}");
                                }
                                else
                                {
                                    page2.Cells[x + 1, 9] = String.Format($"0");
                                }

                                //Подсчет по уважительное причине
                                int poUvasPrich = 0;
                                poUvasPrich = Convert.ToInt32(poBolesniSchet + poZayavleniySchet + poSluzebZapisSchet + prochiePrichiniSchet);
                                page2.Cells[x + 1, 5] = String.Format($"{poUvasPrich}");

                                //количество без уважительной 
                                List<SkipStudentView> neuvazPrichini = skipStudentViews.FindAll(item => (item.cn_G_Student == Gr && (item.Date >= first && item.Date <= last) && item.SmallName == "н"));
                                double neuvazPrichiniSchet = 0;
                                for (int kolVo = 0; kolVo < neuvazPrichini.Count; kolVo++)
                                {
                                    neuvazPrichiniSchet = neuvazPrichiniSchet + Convert.ToInt32(neuvazPrichini[kolVo].Count_hour);
                                    SummaBezUvazizSemestrSchet = SummaBezUvazizSemestrSchet + Convert.ToInt32(neuvazPrichini[kolVo].Count_hour);
                                }
                                if (neuvazPrichini.Count != 0)
                                {
                                    page2.Cells[x + 1, 10] = String.Format($"{neuvazPrichiniSchet}");
                                    double clna2 = neuvazPrichiniSchet / Convert.ToDouble(listStudents.Count);
                                    page2.Cells[x + 1, 11] = String.Format("{0:0.##}", clna2);
                                    page2.Cells[x + 1, 12] = String.Format("{0:0.##}", (neuvazPrichiniSchet / vsegoKolVoSchet) * 100);
                                }
                                else
                                {
                                    page2.Cells[x + 1, 10] = String.Format($"0");
                                    page2.Cells[x + 1, 11] = String.Format($"0");
                                    page2.Cells[x + 1, 12] = String.Format($"0");
                                }

                                vsegoKolVo.Clear();
                                poBolesni.Clear();
                                poZayavleniy.Clear();
                                poSluzebZapis.Clear();
                                prochiePrichini.Clear();
                                neuvazPrichini.Clear();

                                x++;
                            }
                        }
                        page2.Cells[x + 1, 3] = String.Format($"{SummaKolVoSemestrSchet}");
                        page2.Cells[x + 1, 4] = String.Format("{0:0.##}", SummaKolVoSemestrSchet / listStudents.Count);
                        page2.Cells[x + 1, 6] = String.Format($"{SummaPoBolesniSemestrSchet}");
                        page2.Cells[x + 1, 7] = String.Format($"{SummaPoZayavleniySemestrSchet}");
                        page2.Cells[x + 1, 8] = String.Format($"{SummaPoSluzebZapisSemestrSchet}");
                        page2.Cells[x + 1, 9] = String.Format($"{SummaProchiePrichiniSemestrSchet}");
                        page2.Cells[x + 1, 5] = String.Format($"{SummaProchiePrichiniSemestrSchet + SummaPoSluzebZapisSemestrSchet + SummaPoZayavleniySemestrSchet + SummaPoBolesniSemestrSchet}");
                        page2.Cells[x + 1, 10] = String.Format($"{SummaBezUvazizSemestrSchet}");
                        page2.Cells[x + 1, 11] = String.Format("{0:0.##}", SummaBezUvazizSemestrSchet / listStudents.Count);
                        page2.Cells[x + 1, 12] = String.Format("{0:0.##}", (SummaBezUvazizSemestrSchet / SummaKolVoSemestrSchet) * 100);

                        Excel.Range List2Nazvanie11 = (Excel.Range)page2.Range[page2.Cells[x + 3, 2], page2.Cells[x + 3, 12]];
                        List2Nazvanie11.Cells.Merge();
                        List2Nazvanie11.HorizontalAlignment = Excel.Constants.xlLeft;
                        List2Nazvanie11.VerticalAlignment = Excel.Constants.xlTop;
                        page2.Cells[x + 3, 2] = String.Format($"Анализ:");
                        page2.Rows[x + 3].RowHeight = 120;

                        page2.Cells[x + 1, 2] = String.Format($"{typeItem.Content.ToString()}-й\nсеместр");
                        page2.Cells[x + 1, 2].Font.Bold = true;

                        Excel.Range List2Nazvanie9 = (Excel.Range)page2.Range["B6", page2.Cells[x + 1, 12]];
                        List2Nazvanie9.HorizontalAlignment = Excel.Constants.xlCenter;
                        List2Nazvanie9.VerticalAlignment = Excel.Constants.xlCenter;
                        List2Nazvanie9.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        #endregion

                        ex.Visible = true;
                        ex.WindowState = Excel.XlWindowState.xlMaximized;

                        await Task.Delay(1);

                        MessageBoxResult result = MessageBox.Show("Отчет сформирован!", "Информация!", MessageBoxButton.OK, MessageBoxImage.Information);
                        if (result == MessageBoxResult.OK)
                        {
                            pgProgress.Visibility = Visibility.Hidden;
                            pgProgress.Value = 0;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Укажите учебный год!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Не выбран месяц!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region Меню грида
        #region Перейти в режим редактирования
        private void EditRezim_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TabItemEdit.IsSelected = true;

                SkipStudents skip = (SkipStudents)dgStudentSkipView.SelectedItem;
                SkipStudentView skipobj = (SkipStudentView)skip.Obj;

                Id_StudentSkip = Convert.ToString(skipobj.IdStudentSkip);
                dgStudentSkipView.Items.IndexOf(Id_StudentSkip);

                Combo1.SelectedValue = skipobj.IdCause;
                Combo2.Text = skipobj.Count_hour.ToString();
                Combo3.SelectedValue = skipobj.IdSubject_Teacher;
                Combo4.SelectedValue = skipobj.IdEmpForn;

                Combo3.SelectedIndex = 0;
                Combo4.SelectedIndex = 0;
                if (skipobj.IdSubject_Teacher != 0)
                {
                    if (DataService.GetTeacher(Combo3.SelectedValue.ToString(), Gr) != null)
                    {
                        Combo5.ItemsSource = DataService.GetTeacher(Combo3.SelectedValue.ToString(), Gr);
                        Combo5.SelectedIndex = 0;
                    }
                }
                calendar2.DisplayDate = skipobj.Date;
                calendar2.SelectedDate = skipobj.Date;

                #region listbox
                listStudents = new List<ListStudents>();
                ListStudents lists;
                foreach (Student str in DataService.GetStudentGruppa(Gr))
                {
                    lists = new ListStudents();
                    lists.Obj = str;
                    if (skipobj.cn_S == str.cn_S)
                    {
                        lists.Select = true;
                    }
                    listStudents.Add(lists);
                }
                ListBoxStudent.ItemsSource = listStudents;

                #endregion

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView sts in DataService.StudentSkipV(Gr, calendar1.DisplayDate.ToString("MM"), calendar1.DisplayDate.ToString("yyyy")))
                {
                    students = new SkipStudents();
                    students.Obj = sts;

                    if (skipobj.IdStudentSkip == sts.IdStudentSkip)
                    {
                        students.Select = true;
                    }

                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;

                dgStudentSkipViewVnesenie.UpdateLayout();
                dgStudentSkipViewVnesenie.ScrollIntoView(dgStudentSkipViewVnesenie.Items[dgStudentSkipViewVnesenie.Items.Count - 1]);
                dgStudentSkipViewVnesenie.ScrollIntoView(dgStudentSkipViewVnesenie.Items[0]);
                dgStudentSkipViewVnesenie.UpdateLayout();
         
                SelectRowByIndex(dgStudentSkipViewVnesenie, Convert.ToInt32(dgStudentSkipView.SelectedIndex));

                BtnUpdate.Visibility = Visibility.Visible;
                BtnUndoEdit.Visibility = Visibility.Visible;
                GroupInfo.Visibility = Visibility.Visible;
                BtnAdd.Visibility = Visibility.Hidden;

                StatusEdit = true;
            }
            catch
            {

            }
        }
       
        public static void SelectRowByIndex(DataGrid dataGrid, int rowIndex)
        {
            if (!dataGrid.SelectionUnit.Equals(DataGridSelectionUnit.FullRow))
                throw new ArgumentException("Для SelectionUnit DataGrid должно быть установлено значение FullRow.");

            if (rowIndex < 0 || rowIndex > (dataGrid.Items.Count - 1))
                throw new ArgumentException(string.Format("{0} неверный индекс строки.", rowIndex));

            dataGrid.SelectedItems.Clear();
          
            object item = dataGrid.Items[rowIndex];
            dataGrid.SelectedItem = item;

            DataGridRow row = dataGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex) as DataGridRow;
            if (row == null)
            {
                dataGrid.ScrollIntoView(item);
                row = dataGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex) as DataGridRow;
            }
        }
        #endregion
        #endregion

        #region Фильтры
        #region Кнопка применить | Учащийся 
        private void btnFilterStudent_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (checkBox.IsChecked == false)
                {
                    #region Загрузка данных в DataGrid - режим просмотра
                    skipStudents = new List<SkipStudents>();
                    SkipStudents students;
                    foreach (SkipStudentView st in DataService.GetPropuskStudent(cbStudent.SelectedValue.ToString(), calendar1.DisplayDate.ToString("MM"), calendar1.DisplayDate.ToString("yyyy")))
                    {
                        students = new SkipStudents();
                        students.Obj = st;
                        students.Select = false;
                        skipStudents.Add(students);
                    }
                    dgStudentSkipView.ItemsSource = skipStudents;
                    #endregion
                }
                else
                {
                    #region Загрузка данных в DataGrid - режим просмотра


                    skipStudents = new List<SkipStudents>();
                    SkipStudents students;
                    foreach (SkipStudentView st in DataService.GetPropuskCauseAndStudent(cbStudent.SelectedValue.ToString(), cbCause.SelectedValue.ToString(), calendar1.DisplayDate.ToString("MM"), calendar1.DisplayDate.ToString("yyyy")))
                    {
                        students = new SkipStudents();
                        students.Obj = st;
                        students.Select = false;
                        skipStudents.Add(students);
                    }
                    dgStudentSkipView.ItemsSource = skipStudents;

                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
        #endregion

        #region Кнопка применить | Причина 
        private void btnFilterCause_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (checkBox.IsChecked == false)
                {
                    #region Загрузка данных в DataGrid - режим просмотра
                    skipStudents = new List<SkipStudents>();
                    SkipStudents students;
                    foreach (SkipStudentView st in DataService.GetPropuskCause(Gr, cbCause.SelectedValue.ToString(), calendar1.DisplayDate.ToString("MM"), calendar1.DisplayDate.ToString("yyyy")))
                    {
                        students = new SkipStudents();
                        students.Obj = st;
                        students.Select = false;
                        skipStudents.Add(students);
                    }
                    dgStudentSkipView.ItemsSource = skipStudents;
                    #endregion
                }
                else
                {
                    #region Загрузка данных в DataGrid - режим просмотра


                    skipStudents = new List<SkipStudents>();
                    SkipStudents students;
                    foreach (SkipStudentView st in DataService.GetPropuskCauseAndStudent(cbStudent.SelectedValue.ToString(), cbCause.SelectedValue.ToString(), calendar1.DisplayDate.ToString("MM"), calendar1.DisplayDate.ToString("yyyy")))
                    {
                        students = new SkipStudents();
                        students.Obj = st;
                        students.Select = false;
                        skipStudents.Add(students);
                    }
                    dgStudentSkipView.ItemsSource = skipStudents;

                    #endregion
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
           
        }

        #endregion

        #region Очистить фильтры
        private void btnFilterClear_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                #region Загрузка данных в DataGrid - режим просмотра - редактирования

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, calendar1.DisplayDate.ToString("MM"), calendar1.DisplayDate.ToString("yyyy")))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipView.ItemsSource = skipStudents;

                #endregion
                cbStudent.SelectedIndex = -1;
                cbCause.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
        #endregion
        #endregion
        #endregion

        #region  Режим редактирование
        #region Combobox Дисциплина
        private void Combo3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Convert.ToInt32(Combo1.SelectedValue) == 5)
            {
                try
                {
                    if (Combo3 != null & Combo5 != null & Combo3.SelectedIndex != -1)
                    {
                        Combo5.ItemsSource = DataService.GetTeacher(Combo3.SelectedValue.ToString(), Gr);
                        Combo5.SelectedIndex = 0;
                    }
                }
                catch
                {

                }
            }
        }
        #endregion

        #region Combobox Причина
        private void Combo1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Convert.ToInt32(Combo1.SelectedValue) == 5)
            {
                Label3.Visibility = Visibility.Visible;
                Label4.Visibility = Visibility.Visible;
                Label5.Visibility = Visibility.Visible;

                Combo3.Visibility = Visibility.Visible;
                Combo3.SelectedIndex = 0;
                Combo4.SelectedIndex = 0;
                try
                {
                    if (Convert.ToInt32(Combo1.SelectedValue) == 5)
                    {
                        Combo5.ItemsSource = DataService.GetTeacher(Combo3.SelectedValue.ToString(), Gr);
                        Combo5.SelectedIndex = 0;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                Combo4.Visibility = Visibility.Visible;
                Combo5.Visibility = Visibility.Visible;
            }
            else
            {
                Label3.Visibility = Visibility.Hidden;
                Label4.Visibility = Visibility.Hidden;
                Label5.Visibility = Visibility.Hidden;

                Combo3.Visibility = Visibility.Hidden;
                Combo4.Visibility = Visibility.Hidden;
                Combo5.Visibility = Visibility.Hidden;
            }
        }
        #endregion

        #region Календарь | Внесение сведений
        private void calendar2_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            Mouse.Capture(null);
        }

        //Фильтр по месяцу в Datagrid через календарь
        private void calendar2_DisplayDateChanged(object sender, CalendarDateChangedEventArgs e)
        {
            if (IsLoaded && calendar2.DisplayDate != null)
            {
                #region Загрузка данных в DataGrid Внесение сведений
                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, calendar2.DisplayDate.ToString("MM"), calendar2.DisplayDate.ToString("yyyy")))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion
            }

            if (StatusEdit == false)
            {
                calendar2.SelectedDates.Clear();

                #region Загрузка данных в listbox
                listStudents = new List<ListStudents>();
                ListStudents list;
                foreach (Student st in DataService.GetStudentGruppa(Gr))
                {
                    list = new ListStudents();
                    list.Obj = st;
                    list.Select = false;
                    listStudents.Add(list);
                }
                ListBoxStudent.ItemsSource = listStudents;
                #endregion
            }
        }
        #endregion

        #region Кнопка | Внести в базу
        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            int n = 0;
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");

            try
            {
                if (calendar2.SelectedDates.Count != 0)
                {
                    for (int lst = 0; lst < listStudents.Count; lst++)
                    {
                        if (listStudents[lst].Select)
                        {
                            n++;
                        }
                    }

                    if (n != 0)
                    {
                        if (Convert.ToInt32(Combo1.SelectedValue) != 5)
                        {
                            for (int i = 0; i < calendar2.SelectedDates.Count; i++)
                            {
                                for (int u = 0; u < listStudents.Count; u++)
                                {
                                    if (listStudents[u].Select)
                                    {
                                        Student st = (Student)listStudents[u].Obj;
                                        DataService.AddStudentSkip(st.cn_S, Combo1.SelectedValue.ToString(), calendar2.SelectedDates[i].ToString("yyyy-MM-dd"), Combo2.Text);
                                    }
                                }
                            }
                            skipStudents = new List<SkipStudents>();
                            SkipStudents students;
                            foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                            {
                                students = new SkipStudents();
                                students.Obj = st;
                                students.Select = false;
                                skipStudents.Add(students);
                            }
                            dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                            dgStudentSkipViewVnesenie.Items.Refresh();

                            MessageBox.Show($"Информация добавлена!", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        else
                        {
                            for (int i = 0; i < calendar2.SelectedDates.Count; i++)
                            {
                                for (int u = 0; u < listStudents.Count; u++)
                                {
                                    if (listStudents[u].Select)
                                    {
                                        Student st = (Student)listStudents[u].Obj;
                                        DataService.AddStudentSkipNe(st.cn_S, Combo1.SelectedValue.ToString(), calendar2.SelectedDates[i].ToString("yyyy-MM-dd"), Combo2.Text, Combo4.SelectedValue.ToString(), Combo5.SelectedValue.ToString());

                                    }
                                }
                            }

                            skipStudents = new List<SkipStudents>();
                            SkipStudents students;
                            foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                            {
                                students = new SkipStudents();
                                students.Obj = st;
                                students.Select = false;
                                skipStudents.Add(students);
                            }
                            dgStudentSkipViewVnesenie.ItemsSource = skipStudents;

                            dgStudentSkipViewVnesenie.Items.Refresh();
                            MessageBox.Show("Информация добавлена!", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Выберите учащегося!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Дата не выбрана!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region Кнопка | Применить изменения
        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgStudentSkipViewVnesenie.SelectedItems.Count != 0 && Id_StudentSkip != null)
                {
                    string Month = calendar2.DisplayDate.ToString("MM");
                    string Year = calendar2.DisplayDate.ToString("yyyy");
                    if (Convert.ToInt32(Combo1.SelectedValue) != 5)
                    {
                        for (int i = 0; i < calendar2.SelectedDates.Count; i++)
                        {
                            for (int u = 0; u < listStudents.Count; u++)
                            {
                                if (listStudents[u].Select)
                                {
                                    Student st = (Student)listStudents[u].Obj;
                                    DataService.UpdateStudentSkipUvaz(Id_StudentSkip, st.cn_S, Combo1.SelectedValue.ToString(), calendar2.SelectedDates[i].ToString("yyyy-MM-dd"), Combo2.Text);
                                }
                            }
                        }

                        skipStudents = new List<SkipStudents>();
                        SkipStudents students;
                        foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                        {
                            students = new SkipStudents();
                            students.Obj = st;
                            students.Select = false;
                            skipStudents.Add(students);
                        }
                        dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                        dgStudentSkipViewVnesenie.Items.Refresh();

                        MessageBox.Show("Информация изменена!", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                        Id_StudentSkip = null;
                    }
                    else
                    {
                        for (int i = 0; i < calendar2.SelectedDates.Count; i++)
                        {
                            for (int u = 0; u < listStudents.Count; u++)
                            {
                                if (listStudents[u].Select)
                                {
                                    Student st = (Student)listStudents[u].Obj;
                                    DataService.UpdateStudentSkipNE(Id_StudentSkip, st.cn_S, Combo1.SelectedValue.ToString(), calendar2.SelectedDates[i].ToString("yyyy-MM-dd"), Combo2.Text, Combo4.SelectedValue.ToString(), Combo5.SelectedValue.ToString());
                                }
                            }
                        }

                        skipStudents = new List<SkipStudents>();
                        SkipStudents students;
                        foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                        {
                            students = new SkipStudents();
                            students.Obj = st;
                            students.Select = false;
                            skipStudents.Add(students);
                        }
                        dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                        dgStudentSkipViewVnesenie.Items.Refresh();

                        MessageBox.Show("Информация изменена!", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);

                        Id_StudentSkip = null;
                    }
                    calendar2.SelectedDate = null;

                    #region Загрузка данных в listbox

                    listStudents = new List<ListStudents>();
                    ListStudents list;
                    foreach (Student st in DataService.GetStudentGruppa(Gr))
                    {
                        list = new ListStudents();
                        list.Obj = st;
                        list.Select = false;
                        listStudents.Add(list);
                    }
                    ListBoxStudent.ItemsSource = listStudents;

                    Id_StudentSkip = null;

                    BtnUpdate.Visibility = Visibility.Hidden;
                    BtnUndoEdit.Visibility = Visibility.Hidden;

                    BtnAdd.Visibility = Visibility.Visible;

                    Combo1.SelectedIndex = 1;
                    Combo2.SelectedIndex = 5;

                    StatusEdit = false;
                    #endregion

                    BtnUpdate.Visibility = Visibility.Hidden;
                    BtnUndoEdit.Visibility = Visibility.Hidden;
                }
                else
                {
                    MessageBox.Show("Не выбрана строка для изменения!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                StatusEdit = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region Кнопка | Отменить изменение
        private void BtnUndoEdit_Click(object sender, RoutedEventArgs e)
        {
            calendar2.SelectedDates.Clear();

            #region Загрузка данных в listbox
            listStudents = new List<ListStudents>();
            ListStudents list;
            foreach (Student st in DataService.GetStudentGruppa(Gr))
            {
                list = new ListStudents();
                list.Obj = st;
                list.Select = false;
                listStudents.Add(list);
            }
            ListBoxStudent.ItemsSource = listStudents;

            Id_StudentSkip = null;

            BtnUpdate.Visibility = Visibility.Hidden;
            BtnUndoEdit.Visibility = Visibility.Hidden;
            BtnAdd.Visibility = Visibility.Visible;

            Combo1.SelectedIndex = 1;
            Combo2.SelectedIndex = 5;
            #endregion

            #region Загрузка данных в DataGrid - режим просмотра - редактирования
            skipStudents = new List<SkipStudents>();
            SkipStudents students;
            foreach (SkipStudentView st in DataService.StudentSkipV(Gr, calendar2.DisplayDate.ToString("MM"), calendar2.DisplayDate.ToString("yyyy")))
            {
                students = new SkipStudents();
                students.Obj = st;
                students.Select = false;
                skipStudents.Add(students);
            }
            dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
            dgStudentSkipView.ItemsSource = skipStudents;
            #endregion

            StatusEdit = false;
        }
        #endregion

        #region Меню грида
        #region Меню грида | Кнопка Удалить
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string Month = calendar2.DisplayDate.ToString("MM");
                string Year = calendar2.DisplayDate.ToString("yyyy");
                MessageBoxResult messageBoxResult = MessageBox.Show("Вы действительное хотите удалить запись?", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (messageBoxResult == MessageBoxResult.Yes)
                {
                    int n = 0;

                    for (int i = 0; i < skipStudents.Count; i++)
                    {
                        n++;
                        if (skipStudents[i].Select)
                        {
                            SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                            DataService.DeleteStudentSkip(st.IdStudentSkip.ToString());
                        }
                    }

                    if (n != 0)
                    {
                        MessageBox.Show("Удаление прошло успешно!", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Строка не выбрана!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                    }

                    skipStudents = new List<SkipStudents>();
                    SkipStudents students;
                    foreach (SkipStudentView str in DataService.StudentSkipV(Gr, Month, Year))
                    {
                        students = new SkipStudents();
                        students.Obj = str;
                        students.Select = false;
                        skipStudents.Add(students);
                    }
                    dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region Меню грида | Кнопка Изменить
        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
                {
                    for (int i = 0; i < skipStudents.Count; i++)
                    {
                        if (skipStudents[i].Select)
                        {
                            SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                            Combo1.SelectedValue = st.IdCause;
                            Combo2.Text = st.Count_hour;
                            Combo3.SelectedValue = st.IdSubject;
                            Combo4.SelectedValue = st.IdEmpForn;

                            Combo3.SelectedIndex = 0;
                            Combo4.SelectedIndex = 0;
                            if (st.IdSubject != 0)
                            {
                                if (DataService.GetTeacher(Combo3.SelectedValue.ToString(), Gr) != null)
                                {
                                    Combo5.ItemsSource = DataService.GetTeacher(Combo3.SelectedValue.ToString(), Gr);
                                    Combo5.SelectedIndex = 0;
                                }
                            }

                            calendar2.SelectedDate = st.Date;

                            listStudents = new List<ListStudents>();
                            ListStudents lists;
                            foreach (Student str in DataService.GetStudentGruppa(Gr))
                            {
                                lists = new ListStudents();
                                lists.Obj = str;
                                if (st.cn_S == str.cn_S)
                                {

                                    lists.Select = true;

                                }
                                listStudents.Add(lists);
                            }

                            ListBoxStudent.ItemsSource = listStudents;

                            Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                        }
                    }

                    BtnUpdate.Visibility = Visibility.Visible;
                    BtnUndoEdit.Visibility = Visibility.Visible;

                    GroupInfo.Visibility = Visibility.Visible;

                    BtnAdd.Visibility = Visibility.Hidden;

                    StatusEdit = true;
                }
                else
                {
                    calendar2.SelectedDates.Clear();

                    #region Загрузка данных в listbox

                    listStudents = new List<ListStudents>();
                    ListStudents list;
                    foreach (Student st in DataService.GetStudentGruppa(Gr))
                    {
                        list = new ListStudents();
                        list.Obj = st;
                        list.Select = false;
                        listStudents.Add(list);
                    }
                    ListBoxStudent.ItemsSource = listStudents;
                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region Меню грида | Кнопка "По болезни"
        private void MDisease(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.SpeedEditCause(Id_StudentSkip, 2.ToString());

                #region Загрузка данных в DataGrid Внесение сведений

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion
            }
        }
        #endregion

        #region Меню грида | Кнопка "По заявлению"
        private void Mstatement(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.SpeedEditCause(Id_StudentSkip, 4.ToString());

                #region Загрузка данных в DataGrid Внесение сведений
                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion
            }
        }
        #endregion

        #region Меню грида | Кнопка "Прочие"
        private void Others(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.SpeedEditCause(Id_StudentSkip, 3.ToString());

                #region Загрузка данных в DataGrid Внесение сведений
                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion
            }
        }
        #endregion

        #region Меню грида | Кнопка "Служебная записка"
        private void SluzhebnayaZapiska(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.SpeedEditCause(Id_StudentSkip, 1.ToString());

                #region Загрузка данных в DataGrid Внесение сведений
                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion
            }
        }
        #endregion

        #region Меню грида | Кнопка "Неуважительная причина"
        private void neuVaz(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
                {

                    for (int i = 0; i < skipStudents.Count; i++)
                    {
                        if (skipStudents[i].Select)
                        {
                            SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;

                            Combo2.Text = st.Count_hour;
                            Combo1.SelectedValue = 5;
                            Combo3.SelectedValue = st.IdSubject;
                            Combo4.SelectedValue = st.IdEmpForn;

                            Combo3.SelectedIndex = 0;
                            Combo4.SelectedIndex = 0;
                            if (st.IdSubject != 0)
                            {
                                if (DataService.GetTeacher(Combo3.SelectedValue.ToString(), Gr) != null)
                                {
                                    Combo5.ItemsSource = DataService.GetTeacher(Combo3.SelectedValue.ToString(), Gr);
                                    Combo5.SelectedIndex = 0;
                                }
                            }

                            calendar2.SelectedDate = st.Date;

                            listStudents = new List<ListStudents>();
                            ListStudents lists;
                            foreach (Student str in DataService.GetStudentGruppa(Gr))
                            {
                                lists = new ListStudents();
                                lists.Obj = str;
                                if (st.cn_S == str.cn_S)
                                {

                                    lists.Select = true;

                                }
                                listStudents.Add(lists);
                            }

                            ListBoxStudent.ItemsSource = listStudents;

                            Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                        }
                    }

                    BtnUpdate.Visibility = Visibility.Visible;
                    BtnUndoEdit.Visibility = Visibility.Visible;
                    GroupInfo.Visibility = Visibility.Visible;
                    BtnAdd.Visibility = Visibility.Hidden;

                    StatusEdit = true;
                }
                else
                {
                    calendar2.SelectedDates.Clear();

                    #region Загрузка данных в listbox

                    listStudents = new List<ListStudents>();
                    ListStudents list;
                    foreach (Student st in DataService.GetStudentGruppa(Gr))
                    {
                        list = new ListStudents();
                        list.Obj = st;
                        list.Select = false;
                        listStudents.Add(list);
                    }
                    ListBoxStudent.ItemsSource = listStudents;
                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region Количество часов 1
        private void ClickTime_1(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.TimeEdit(Id_StudentSkip, 1.ToString());

                #region Загрузка данных в DataGrid Внесение сведений

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion
            }
        }
        #endregion

        #region Количество часов 2
        private void ClickTime_2(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.TimeEdit(Id_StudentSkip, 2.ToString());

                #region Загрузка данных в DataGrid Внесение сведений

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion

            }
        }
        #endregion

        #region Количество часов 3
        private void ClickTime_3(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.TimeEdit(Id_StudentSkip, 3.ToString());

                #region Загрузка данных в DataGrid Внесение сведений

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion

            }
        }
        #endregion

        #region Количество часов 4
        private void ClickTime_4(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.TimeEdit(Id_StudentSkip, 4.ToString());

                #region Загрузка данных в DataGrid Внесение сведений

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion

            }
        }
        #endregion

        #region Количество часов 5
        private void ClickTime_5(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.TimeEdit(Id_StudentSkip, 5.ToString());

                #region Загрузка данных в DataGrid Внесение сведений

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion

            }
        }
        #endregion

        #region Количество часов 6
        private void ClickTime_6(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.TimeEdit(Id_StudentSkip, 6.ToString());

                #region Загрузка данных в DataGrid Внесение сведений

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion

            }
        }
        #endregion

        #region Количество часов 7
        private void ClickTime_7(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.TimeEdit(Id_StudentSkip, 7.ToString());

                #region Загрузка данных в DataGrid Внесение сведений

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion

            }
        }
        #endregion

        #region Количество часов 8
        private void ClickTime_8(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.TimeEdit(Id_StudentSkip, 8.ToString());

                #region Загрузка данных в DataGrid Внесение сведений

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion

            }
        }
        #endregion

        #region Количество часов 9
        private void ClickTime_9(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.TimeEdit(Id_StudentSkip, 9.ToString());

                #region Загрузка данных в DataGrid Внесение сведений

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion

            }
        }
        #endregion

        #region Количество часов 10
        private void ClickTime_10(object sender, RoutedEventArgs e)
        {
            string Month = calendar2.DisplayDate.ToString("MM");
            string Year = calendar2.DisplayDate.ToString("yyyy");
            if (dgStudentSkipViewVnesenie.SelectedItems.Count == 1)
            {
                for (int i = 0; i < skipStudents.Count; i++)
                {
                    if (skipStudents[i].Select)
                    {
                        SkipStudentView st = (SkipStudentView)skipStudents[i].Obj;
                        Id_StudentSkip = Convert.ToString(st.IdStudentSkip);
                    }
                }
                DataService.TimeEdit(Id_StudentSkip, 10.ToString());

                #region Загрузка данных в DataGrid Внесение сведений

                skipStudents = new List<SkipStudents>();
                SkipStudents students;
                foreach (SkipStudentView st in DataService.StudentSkipV(Gr, Month, Year))
                {
                    students = new SkipStudents();
                    students.Obj = st;
                    students.Select = false;
                    skipStudents.Add(students);
                }
                dgStudentSkipViewVnesenie.ItemsSource = skipStudents;
                #endregion

            }
        }
        #endregion
        #endregion
        #endregion

        #region События TabControl
        private void TabControlProect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StatusForm == true)
            {
                if (BtnUndoEdit.Visibility == Visibility.Visible)
                {
                    if (TabControlProect.SelectedIndex != 3)
                    {
                        BtnUndoEdit.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                    }
                }
            }
        }
        #endregion
        #endregion

        #region Сдача СПХ группы
        #region Кнопка Добавить СПХ
        private void BtnAddSPH_Click(object sender, RoutedEventArgs e)
        {
            CleaningTableSPHGroup();
            TransitionToCreateSPHForm();
            IsReadOnliDataGridSPH(false);
        }

        private void CleaningTableSPHGroup()
        {
            dgGroupSPH1.ItemsSource = null;
            dgGroupSPH2.ItemsSource = null;
            dgGroupSPH3.ItemsSource = null;
            dgGroupSPH4.ItemsSource = null;
        }

        private void TransitionToCreateSPHForm()
        {
            panelTableSPHHeder.Visibility = Visibility.Hidden;
            panelTableSPHHederPeriod.Visibility = Visibility.Visible;
            groupBoxCreateSPH.Visibility = Visibility.Visible;
            BtnAddSPH.Visibility = Visibility.Hidden;
        }
        #endregion

        #region Радио кнопки (Просмотр СПХ)
        private void radiobtnSPH1_Checked(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityTableSPHGroup();
            dgGroupSPH1.Visibility = Visibility.Visible;
        }

        private void radiobtnSPH2_Checked(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityTableSPHGroup();
            dgGroupSPH2.Visibility = Visibility.Visible;
        }

        private void radiobtnSPH3_Checked(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityTableSPHGroup();
            dgGroupSPH3.Visibility = Visibility.Visible;
        }

        private void radiobtnSPH4_Checked(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityTableSPHGroup();
            dgGroupSPH4.Visibility = Visibility.Visible;
        }

        private void ChengeVisibilityTableSPHGroup()
        {
            dgGroupSPH1.Visibility = Visibility.Hidden;
            dgGroupSPH2.Visibility = Visibility.Hidden;
            dgGroupSPH3.Visibility = Visibility.Hidden;
            dgGroupSPH4.Visibility = Visibility.Hidden;
        }
        #endregion

        private void IsReadOnliDataGridSPH(bool yeasOrNot)
        {
            dgGroupSPH1.IsReadOnly = yeasOrNot;
            dgGroupSPH2.IsReadOnly = yeasOrNot;
            dgGroupSPH3.IsReadOnly = yeasOrNot;
            dgGroupSPH4.IsReadOnly = yeasOrNot;
        }

        #region Кнопка Рассчитать СПХ
        private void BtnCalculateSPH_Click(object sender, RoutedEventArgs e)
        {
            CalculateSPH();
            BtnHandSPH.IsEnabled = true;
        }

        private void CalculateSPH()
        {
            List<SphSend> groupSPH = DataService.GetSPH1(Gr);
            dgGroupSPH1.ItemsSource = groupSPH;
            dgGroupSPH2.ItemsSource = groupSPH;
            dgGroupSPH3.ItemsSource = groupSPH;
            dgGroupSPH4.ItemsSource = groupSPH;
        }
        #endregion

        #region Кнопка Сдать СПХ
        private void BtnHandSPH_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (checkInputDataFromSPH(SetSph()))
                {
                    dgGroupSPH4.ItemsSource = DataService.AddSphSend(SetSph());

                    BtnHandSPH.IsEnabled = false;
                    MessageBox.Show("СПХ успешно сохранено!", "Внимание", MessageBoxButton.OK, MessageBoxImage.Information);
                    UpdatngSPHGroup();
                    IsReadOnliDataGridSPH(true);
                    GoToViewSPHForm();
                }
            }
            catch(System.Data.SqlClient.SqlException)
            {
                MessageBox.Show("СПХ с таким периодом уже присутствует!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool checkInputDataFromSPH(SphSend sph)
        {
            if (!CheckCountCharacteristics(sph))
                return false;
            if (!CheckFamilyCharacteristics(sph))
                return false;
            if (!CheckStudentCharacteristics(sph))
                return false;
            if (!CheckPublicEmploymentCharacteristics(sph))
                return false;
            return true;
        }

        private bool CheckCountCharacteristics(SphSend sph)
        {
            if (sph.Students < sph.Budget || 0 > sph.Budget)
                return Warning("Кол-во учащихся на бюджете не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.NonBudget || 0 > sph.NonBudget)
                return Warning("Кол-во учащихся вне бюджета не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Boys || 0 > sph.Boys)
                return Warning("Кол-во юношей не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Girls || 0 > sph.Girls)
                return Warning("Кол-во девушек не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Adult || 0 > sph.Adult)
                return Warning("Кол-во совершеннолетних учащихся не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Underage || 0 > sph.Underage)
                return Warning("Кол-во несовершеннолетних учащихся не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Nonresident || 0 > sph.Nonresident)
                return Warning("Кол-во иногородних учащихся не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Hostel || 0 > sph.Hostel)
                return Warning("Кол-во учащихся проживающих в общежитии не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Flat || 0 > sph.Flat)
                return Warning("Кол-во учащихся проживающих на квартире не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Foreign_student || 0 > sph.Foreign_student)
                return Warning("Кол-во иностранных учащихся не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            return true;
        }

        private bool CheckFamilyCharacteristics(SphSend sph)
        {
            if (sph.Students < sph.Incomplete || 0 > sph.Incomplete)
                return Warning("Кол-во учащихся с неполной семьёй не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Many_children_family || 0 > sph.Many_children_family)
                return Warning("Кол-во учащихся с многодетной семьёй не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Trusteeship || 0 > sph.Trusteeship)
                return Warning("Кол-во учащихся на попечительстве не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Foster_family || 0 > sph.Foster_family)
                return Warning("Кол-во учащихся с приёмной семьёй не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Refurgee || 0 > sph.Refurgee)
                return Warning("Кол-во учащихся беженцев/переселенцев не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Have_disabled_parents || 0 > sph.Have_disabled_parents)
                return Warning("Кол-во учащихся имеющих родителей инвалидов не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Low_income_family || 0 > sph.Low_income_family)
                return Warning("Кол-во малообеспеченных учащихся не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Family_students || 0 > sph.Family_students)
                return Warning("Кол-во семейных учащихся не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Have_children || 0 > sph.Have_children)
                return Warning("Кол-во учащихся имеющих детей не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            return true;
        }

        private bool CheckStudentCharacteristics(SphSend sph)
        {
            if (sph.Students < sph.State_support_in_college || 0 > sph.State_support_in_college)
                return Warning("Кол-во учащихся на гос. обеспечении в колледже не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Socially_dangerous_position || 0 > sph.Socially_dangerous_position)
                return Warning("Кол-во учащихся СОП не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Need_for_state_protection || 0 > sph.Need_for_state_protection)
                return Warning("Кол-во учащихся НГЗ не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Individual_prophylactic_accounting || 0 > sph.Individual_prophylactic_accounting)
                return Warning("Кол-во учащихся состоящих на индивидуально-профилактическом учёте не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.Disabled_people || 0 > sph.Disabled_people)
                return Warning("Кол-во учащихся инвалидов не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            return true;
        }

        private bool CheckPublicEmploymentCharacteristics(SphSend sph)
        {
            if (sph.Students < sph.BRSM_members || 0 > sph.BRSM_members)
                return Warning("Кол-во учащихся БРСМ не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            if (sph.Students < sph.TradeUnion_members || 0 > sph.TradeUnion_members)
                return Warning("Кол-во учащихся являющимися членами профсоюза не должно иметь отрицательное значение или превышать общее кол-во учащихся!");
            return true;
        }

        private SphSend SetSph()
        {
            dgGroupSPH1.SelectedIndex = 0;
            SphSend sph = (SphSend)dgGroupSPH1.SelectedItem;
            sph.cn_G = Gr;
            sph.id_period = cbPeriod.SelectedIndex + 1;
            dgGroupSPH1.SelectedItem = null;
            return sph;
        }

        private void UpdatngSPHGroup()
        {
            List<SphSend> groupSPH = GetSPHGroup();
            dgGroupSPH1.ItemsSource = groupSPH;
            dgGroupSPH2.ItemsSource = groupSPH;
            dgGroupSPH3.ItemsSource = groupSPH;
            dgGroupSPH4.ItemsSource = groupSPH;
        }

        private List<SphSend> GetSPHGroup()
        {
            return DataService.GetSPH(Gr);
        }

        private void GoToViewSPHForm()
        {
            panelTableSPHHeder.Visibility = Visibility.Visible;
            panelTableSPHHederPeriod.Visibility = Visibility.Hidden;
            groupBoxCreateSPH.Visibility = Visibility.Hidden;
            BtnAddSPH.Visibility = Visibility.Visible;
        }
        #endregion

        #region Кнопка Отмена
        private void BtnCansel_Click(object sender, RoutedEventArgs e)
        {
            UpdatngSPHGroup();
            IsReadOnliDataGridSPH(true);
            GoToViewSPHForm();

            BtnHandSPH.IsEnabled = false;
        }
        #endregion
        #endregion

        #region Формирование документов
        private void lbGroupListToCreateDocuments_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadingFamilyVisitInfo();
            LoadingStudentCharacterization();
        }

        private Student GetSelectedStudentForCreateDocuments()
        {
            return (Student)lbGroupListToCreateDocuments.SelectedItem;
        }

        private bool CheckIsSelectedStudentForCreateDocuments()
        {
            if (lbGroupListToCreateDocuments.SelectedItem != null)
                return true;
            return Information("Выберите студента!");
        }

        #region Спрака о посещении семьи учащегося
        private void ChangingUnVisibilityButtonsForGeneratingDocumentAboutStudentFamily()
        {
            CreateDocumentsCalendar.IsEnabled = false;
            tbAccommodations.IsReadOnly = true;
            tbHousingCharacteristic.IsReadOnly = true;
            BtnForCreateStudntFamilyDocument.Visibility = Visibility.Visible;
            BtnForSaveDataStudntFamily.Visibility = Visibility.Hidden;
            BtnForCanselEditStudntFamily.Visibility = Visibility.Hidden;
            BtnForEditStudntFamily.Visibility = Visibility.Visible;
        }

        private void ChangingVisibilityButtonsForGeneratingDocumentAboutStudentFamily()
        {
            CreateDocumentsCalendar.IsEnabled = true;
            tbAccommodations.IsReadOnly = false;
            tbHousingCharacteristic.IsReadOnly = false;
            BtnForCreateStudntFamilyDocument.Visibility = Visibility.Hidden;
            BtnForSaveDataStudntFamily.Visibility = Visibility.Visible;
            BtnForCanselEditStudntFamily.Visibility = Visibility.Visible;
            BtnForEditStudntFamily.Visibility = Visibility.Hidden;
        }

        #region Загрузка данных о посещении семьи учащегося
        private void LoadingFamilyVisitInfo()
        {
            if (ChekIsStudentHaveFamilyVisitInfo())
                GetFamilyVisitInfo();
            else
                ClearFamilyVisitInfoInForm();
        }

        private bool ChekIsStudentHaveFamilyVisitInfo()
        {
            if (DataService.GetFamilyVisitInfo(GetSelectedStudentForCreateDocuments().cn_S).Count != 0)
                return true;
            return false;
        }

        private void GetFamilyVisitInfo()
        {
            FamilyVisitInfo studentfamilyVisitInfo = DataService.GetFamilyVisitInfo(GetSelectedStudentForCreateDocuments().cn_S)[0];

            CreateDocumentsCalendar.SelectedDate = studentfamilyVisitInfo.Date_of_visit;
            tbAccommodations.Text = studentfamilyVisitInfo.Living_conditions;
            tbHousingCharacteristic.Text = studentfamilyVisitInfo.House_characteristics;
        }

        private void ClearFamilyVisitInfoInForm()
        {
            CreateDocumentsCalendar.SelectedDate = null;
            tbAccommodations.Text = string.Empty;
            tbHousingCharacteristic.Text = string.Empty;
        }
        #endregion

        #region Кнопа изменить сведения о посещении семьи учащегося
        private void BtnForEditStudntFamily_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedStudentForCreateDocuments())
                ChangingVisibilityButtonsForGeneratingDocumentAboutStudentFamily();
        }
        #endregion

        #region Кнопа отменить измениения сведений о посещении семьи учащегося
        private void BtnForCanselEditStudntFamily_Click(object sender, RoutedEventArgs e)
        {
            ChangingUnVisibilityButtonsForGeneratingDocumentAboutStudentFamily();
            LoadingFamilyVisitInfo();
        }
        #endregion

        #region Кнопка сохранить измениения сведений о посещении семьи учащегося
        private void BtnForSaveDataStudntFamily_Click(object sender, RoutedEventArgs e)
        {
            if (ChekIsStudentHaveFamilyVisitInfo())
                EditFamilyVisitInfo();
            else
                AddFamilyVisitInfo();
        }

        private void EditFamilyVisitInfo()
        {
            if (CheckIsFilledFamilyVisit())
            {
                DataService.UpdateFamilyVisitInfo(SetFamilyVisit());
                ChangingUnVisibilityButtonsForGeneratingDocumentAboutStudentFamily();
                LoadingFamilyVisitInfo();
            } 
        }

        private void AddFamilyVisitInfo()
        {
            if (CheckIsFilledFamilyVisit())
            {
                DataService.AddFamilyVisitInfo(SetFamilyVisit());
                ChangingUnVisibilityButtonsForGeneratingDocumentAboutStudentFamily();
                LoadingFamilyVisitInfo();
            }
        }

        private bool CheckIsFilledFamilyVisit()
        {
            if (CreateDocumentsCalendar.SelectedDate == null)
                return Warning("Выберите дату посещения семьи учащегося!");
            if (string.IsNullOrEmpty(tbHousingCharacteristic.Text))
                return Warning("Заполните характеристику жилья!");
            if (string.IsNullOrEmpty(tbAccommodations.Text))
                return Warning("Заполните условия проживания!");
            return true;
        }

        private FamilyVisitInfo SetFamilyVisit()
        {
            FamilyVisitInfo familyVisitInfo = new FamilyVisitInfo();
            familyVisitInfo.cn_S = GetSelectedStudentForCreateDocuments().cn_S;
            familyVisitInfo.Date_of_visit = GetDateOfVisitStudentFamily();
            familyVisitInfo.House_characteristics = tbHousingCharacteristic.Text;
            familyVisitInfo.Living_conditions = tbAccommodations.Text;

            return familyVisitInfo;
        }

        private DateTime GetDateOfVisitStudentFamily()
        {
            return (DateTime)CreateDocumentsCalendar.SelectedDate;
        }
        #endregion

        #region Кнопа сформировать справку о посещении семьи учащегося
        private void BtnForCreateStudntFamilyDocument_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedStudentForCreateDocuments())
                if (CheckIsFilledFamilyVisit())
                    CreateCertificateOfAttendanceAtTheStudentFamily();
        }

        private void CreateCertificateOfAttendanceAtTheStudentFamily()
        {
            Student currentStudent = GetSelectedStudentForCreateDocuments();
            DateTime dateVisitFamilyStudent = GetDateOfVisitStudentFamily();
            try
            {
                // Настройка документа
                Excel.Application excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.Visible = false;
                excelApplication.SheetsInNewWorkbook = 1;
                Excel.Workbook workBook = excelApplication.Workbooks.Add(Type.Missing);
                excelApplication.DisplayAlerts = false;
                Excel.Worksheet firstPage = (Excel.Worksheet)excelApplication.Worksheets.get_Item(1);
                firstPage.Name = "Справка о посещ. семьи учащ.";

                // Заголовок
                Excel.Range rangeForTitle = (Excel.Range)firstPage.Range["A2", "N2"];
                rangeForTitle.Cells.Merge();
                rangeForTitle.Rows[1].RowHeight = 19.50;
                rangeForTitle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangeForTitle.Font.Size = 12;
                rangeForTitle.Cells.Font.Name = "Times New Roman";

                DateTime dateBirth = currentStudent.DateBirth;
                firstPage.Cells[2, 1] = String.Format($"Справка о посещении семьи учащегося {currentStudent.Uchashchiysya}, {dateBirth.Day}.{dateBirth.Month}.{dateBirth.Year} года рожд.");

                // Таблица
                Excel.Range rangeForOptionsTable = (Excel.Range)firstPage.Range["B4", "M10"];
                rangeForOptionsTable.RowHeight = 19.50;
                rangeForOptionsTable.Font.Size = 12;
                rangeForOptionsTable.Cells.Font.Name = "Times New Roman";
                rangeForOptionsTable.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                rangeForOptionsTable.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                rangeForOptionsTable.WrapText = true;
                rangeForOptionsTable.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeForOptionsTable.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeForOptionsTable.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeForOptionsTable.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeForOptionsTable.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeForOptionsTable.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;

                Excel.Range rangeRow1Column1 = (Excel.Range)firstPage.Range["B4", "E4"];
                rangeRow1Column1.Cells.Merge();
                firstPage.Cells[4, 2] = String.Format("Дата посещения");

                Excel.Range rangeRow1Column2 = (Excel.Range)firstPage.Range["F4", "M4"];
                rangeRow1Column2.Cells.Merge();
                firstPage.Cells[4, 6] = String.Format($"{dateVisitFamilyStudent.Day}.{dateVisitFamilyStudent.Month}.{dateVisitFamilyStudent.Year}г.");

                Excel.Range rangeRow2Column1 = (Excel.Range)firstPage.Range["B5", "E5"];
                rangeRow2Column1.Cells.Merge();
                firstPage.Cells[5, 2] = String.Format("Адрес проживания");

                Excel.Range rangeRow2Column2 = (Excel.Range)firstPage.Range["F5", "M5"];
                rangeRow2Column2.Cells.Merge();
                firstPage.Cells[5, 6] = String.Format($"{currentStudent.Adress} ");

                Excel.Range rangeRow3Column1 = (Excel.Range)firstPage.Range["B6", "E6"];
                rangeRow3Column1.Cells.Merge();
                rangeRow3Column1.Rows[1].RowHeight = 33;
                firstPage.Cells[6, 2] = String.Format("С кем проживает несовершеннолетний");

                Excel.Range rangeRow3Column3 = (Excel.Range)firstPage.Range["F6", "M6"];
                rangeRow3Column3.Cells.Merge();
                rangeRow3Column3.Rows[1].RowHeight = 33;
                firstPage.Cells[6, 6] = String.Format($"{GetParentsFIO(Convert.ToInt32(currentStudent.cn_S))} ");

                Excel.Range rangeRow4Column1 = (Excel.Range)firstPage.Range["B7", "E7"];
                rangeRow4Column1.Cells.Merge();
                rangeRow4Column1.Rows[1].RowHeight = 66;
                firstPage.Cells[7, 2] = String.Format("Сведения о родителях: мать, отец (ФИО, место работы, должность, место проживания, состоят ли в браке)");

                Excel.Range rangeRow4Column4 = (Excel.Range)firstPage.Range["F7", "M7"];
                rangeRow4Column4.Cells.Merge();
                rangeRow4Column4.Rows[1].RowHeight = 66;
                firstPage.Cells[7, 6] = String.Format($"{GetParentsAllInfo(Convert.ToInt32(currentStudent.cn_S))}");

                Excel.Range rangeRow5Column1 = (Excel.Range)firstPage.Range["B8", "E8"];
                rangeRow5Column1.Cells.Merge();
                rangeRow5Column1.Rows[1].RowHeight = 81;
                firstPage.Cells[8, 2] = String.Format("Характеристика жилья (квартира, дом, количество жилых комнат) Принадлежность жилья (в собственности, съемное, арендное, принадлежит родственникам и т.д..)");

                Excel.Range rangeRow5Column5 = (Excel.Range)firstPage.Range["F8", "M8"];
                rangeRow5Column5.Cells.Merge();
                rangeRow5Column5.Rows[1].RowHeight = 81;
                firstPage.Cells[8, 6] = String.Format(tbHousingCharacteristic.Text);

                Excel.Range rangeRow6Column1 = (Excel.Range)firstPage.Range["B9", "E9"];
                rangeRow6Column1.Cells.Merge();
                rangeRow6Column1.Rows[1].RowHeight = 130;
                firstPage.Cells[9, 2] = String.Format("Условия проживания несовершеннолетнего (наличие отдельного спального места, места для подготовки к занятиям, наличие одежды, обуви по сезону,  соблюдение санитарно-гигиенических норм, соблюдение режима дня несовершеннолетним)");

                Excel.Range rangeRow6Column6 = (Excel.Range)firstPage.Range["F9", "M9"];
                rangeRow6Column6.Cells.Merge();
                rangeRow6Column6.Rows[1].RowHeight = 130;
                firstPage.Cells[9, 6] = String.Format(tbAccommodations.Text);

                Excel.Range rangeRow7Column1 = (Excel.Range)firstPage.Range["B10", "E10"];
                rangeRow7Column1.Cells.Merge();
                firstPage.Cells[10, 2] = String.Format("Выводы");

                Excel.Range rangeRow7Column7 = (Excel.Range)firstPage.Range["F10", "M10"];
                rangeRow7Column7.Cells.Merge();
                firstPage.Cells[10, 6] = String.Format("");

                // Подпись куратора
                Excel.Range rangeForCuratorsPainting = (Excel.Range)firstPage.Range["B12", "M12"];
                rangeForCuratorsPainting.Cells.Merge();
                rangeForCuratorsPainting.Rows[1].RowHeight = 19.50;
                rangeForCuratorsPainting.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                rangeForCuratorsPainting.Font.Size = 12;
                rangeForCuratorsPainting.Cells.Font.Name = "Times New Roman";
                rangeForCuratorsPainting.Cells[1, 1] = String.Format($"Куратор группы    ____________________ {DataService.FIOandGroupTeacher(Gr, Cnc)[0].FIOTeacher}");

                excelApplication.Visible = true;
                excelApplication.WindowState = Excel.XlWindowState.xlMaximized;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string GetParentsAllInfo(int idStudent)
        {
            string parentsInfo = string.Empty;
            List<FamilyComposition> parentsList = DataService.GetParents(idStudent);
            for (int i = 0; i <= parentsList.Count-1; i++)
            {
                parentsInfo += $"{parentsList[i].Name}: {parentsList[i].FIO}, Место работы: {parentsList[i].Work_Study_Place}," +
                               $" Место проживание: {parentsList[i].Place_of_residence}\n";
            }
            return parentsInfo;
        }

        private string GetParentsFIO(int idStudent)
        {
            string parentsName = string.Empty;
            List<FamilyComposition> parentsList = DataService.GetParents(idStudent);
            for (int i = 0; i <= parentsList.Count - 1; i++)
            {
                parentsName += $"{parentsList[i].Name} - {parentsList[i].FIO}\n";
            }
            return parentsName;
        }
        #endregion
        #endregion

        #region СПХ Призывника/Характеристика учащегося
        private void ChangingVisibilityButtonsForGeneratingCharacterizationsAboutStudents()
        {
            StudentPersonalInformationForDocuments.IsEnabled = true;
            BtnForEditStudntInfoForDocuments.Visibility = Visibility.Hidden;
            BtnForCanselEditStudntInfoForDocuments.Visibility = Visibility.Visible;
            BtnForSaveDataStudntInfoForDocuments.Visibility = Visibility.Visible;
            BtnForCreateStudentProfile.Visibility = Visibility.Hidden;
            BtnForCreateSPCofTheConscript.Visibility = Visibility.Hidden;
        }

        private void ChangingUnVisibilityButtonsForGeneratingCharacterizationsAboutStudents()
        {
            StudentPersonalInformationForDocuments.IsEnabled = false;
            BtnForEditStudntInfoForDocuments.Visibility = Visibility.Visible;
            BtnForCanselEditStudntInfoForDocuments.Visibility = Visibility.Hidden;
            BtnForSaveDataStudntInfoForDocuments.Visibility = Visibility.Hidden;
            BtnForCreateStudentProfile.Visibility = Visibility.Visible;
            BtnForCreateSPCofTheConscript.Visibility = Visibility.Visible;
        }

        #region Загрузка характеристики учащегося
        private void LoadingStudentCharacterization()
        {
            if (ChekIsStudentHaveCharacterization())
            {
                GetStudentCharacterization();
            }
            else
            {
                ClearStudentCharacterization();
            }
        }

        private bool ChekIsStudentHaveCharacterization()
        {
            if (DataService.GetStudentCharacterization(Convert.ToInt32(GetSelectedStudentForCreateDocuments().cn_S)).Count != 0)
                return true;
            return false;
        }

        private void GetStudentCharacterization()
        {
            StudentCharacterization studentCharacterization = DataService.GetStudentCharacterization(Convert.ToInt32(GetSelectedStudentForCreateDocuments().cn_S))[0];

            tbStudentGeneralDevelopmentAndOutlook.Text = studentCharacterization.General_development_and_outlook;
            tbStudentHobbies.Text = studentCharacterization.Hobbies;
            tbStudentAttitudeToPhysicalCulture.Text = studentCharacterization.Attitude_to_physical_culture;
            tbStudentAcademicPerformance.Text = studentCharacterization.Academic_performance;
            tbStudentMostFavoriteSubjects.Text = studentCharacterization.Most_favorite_subjects;
            tbStudentMostDislikedSubjects.Text = studentCharacterization.Most_disliked_subjects;
            tbStudentLongAbsences.Text = studentCharacterization.Long_absences;
            tbStudentCommunicationWithPeers.Text = studentCharacterization.Communication_with_peers;
            tbStudentCommunicationWithTeachers.Text = studentCharacterization.Communication_with_teachers;
            tbStudentPsychologicalFeatures.Text = studentCharacterization.Psychological_features;
            tbStudentSignsOfSocialNeglect.Text = studentCharacterization.Signs_of_social_neglect;
            tbStudentLevelOfNeuropsychologicalStability.Text = studentCharacterization.Level_of_neuropsychological_stability;
            cbStudentInclinationToWithdrawal.IsChecked = studentCharacterization.Inclination_to_withdrawal;
            tbStudentTemperament.Text = studentCharacterization.Temperament;
            tbStudentSelfAssessment.Text = studentCharacterization.Self_assessment;
        }

        private void ClearStudentCharacterization()
        {
            tbStudentGeneralDevelopmentAndOutlook.Text = string.Empty;
            tbStudentHobbies.Text = string.Empty;
            tbStudentAttitudeToPhysicalCulture.Text = string.Empty;
            tbStudentAcademicPerformance.Text = string.Empty;
            tbStudentMostFavoriteSubjects.Text = string.Empty;
            tbStudentMostDislikedSubjects.Text = string.Empty;
            tbStudentLongAbsences.Text = string.Empty;
            tbStudentCommunicationWithPeers.Text = string.Empty;
            tbStudentCommunicationWithTeachers.Text = string.Empty;
            tbStudentPsychologicalFeatures.Text = string.Empty;
            tbStudentSignsOfSocialNeglect.Text = string.Empty;
            tbStudentLevelOfNeuropsychologicalStability.Text = string.Empty;
            cbStudentInclinationToWithdrawal.IsChecked = false;
            tbStudentTemperament.Text = string.Empty;
            tbStudentSelfAssessment.Text = string.Empty;
        }
        #endregion

        #region Кнопа изменить характеристику учащегося
        private void BtnForEditStudntInfoForDocuments_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedStudentForCreateDocuments())
                ChangingVisibilityButtonsForGeneratingCharacterizationsAboutStudents();
        }
        #endregion

        #region Кнопа отменить изменения характеристики учащегося
        private void BtnForCanselEditStudntInfoForDocuments_Click(object sender, RoutedEventArgs e)
        {
            ChangingUnVisibilityButtonsForGeneratingCharacterizationsAboutStudents();
            LoadingStudentCharacterization();
        }
        #endregion

        #region Кнопа сохранить изменения о характеристике учащегося
        private void BtnForSaveDataStudntInfoForDocuments_Click(object sender, RoutedEventArgs e)
        {
            if (ChekIsStudentHaveCharacterization())
                EditStudentCharacterization();
            else
                AddStudentCharacterization();
        }

        private void EditStudentCharacterization()
        {
            if (CheckIsFilledStudentCharacterization())
            {
                DataService.UpdateStudentCharacterization(SetStudentCharacterization());
                ChangingUnVisibilityButtonsForGeneratingCharacterizationsAboutStudents();
                LoadingStudentCharacterization();
            }
        }

        private void AddStudentCharacterization()
        {
            if (CheckIsFilledStudentCharacterization())
            {
                DataService.AddStudentCharacterization(SetStudentCharacterization());
                ChangingUnVisibilityButtonsForGeneratingCharacterizationsAboutStudents();
                LoadingStudentCharacterization();
            }
        }

        private bool CheckIsFilledStudentCharacterization()
        {
            if (string.IsNullOrEmpty(tbStudentGeneralDevelopmentAndOutlook.Text))
                return Warning("Заполните поле Общее развитие и кругозор!");
            if (string.IsNullOrEmpty(tbStudentHobbies.Text))
                return Warning("Заполните поле Увлечения!");
            if (string.IsNullOrEmpty(tbStudentAttitudeToPhysicalCulture.Text))
                return Warning("Заполните поле Отношение к физической культуре!");
            if (string.IsNullOrEmpty(tbStudentAcademicPerformance.Text))
                return Warning("Заполните поле Успеваемость!");
            if (string.IsNullOrEmpty(tbStudentCommunicationWithPeers.Text))
                return Warning("Заполните поле Общение со сверстниками!");
            if (string.IsNullOrEmpty(tbStudentCommunicationWithTeachers.Text))
                return Warning("Заполните поле Общение с преподователями!");
            if (string.IsNullOrEmpty(tbStudentPsychologicalFeatures.Text))
                return Warning("Заполните поле Психологические особенности!");
            if (string.IsNullOrEmpty(tbStudentSignsOfSocialNeglect.Text))
                return Warning("Заполните поле Признаки соц. запущенности!");
            if (string.IsNullOrEmpty(tbStudentLevelOfNeuropsychologicalStability.Text))
                return Warning("Заполните поле Уровень нервно-псих. устойчивости!");
            if (string.IsNullOrEmpty(tbStudentTemperament.Text))
                return Warning("Заполните поле Темперамент!");
            if (string.IsNullOrEmpty(tbStudentSelfAssessment.Text))
                return Warning("Заполните поле Самооценка!");
            return true;
        }

        private StudentCharacterization SetStudentCharacterization()
        {
            StudentCharacterization studentCharacterization = new StudentCharacterization();

            studentCharacterization.cn_S = Convert.ToInt32(GetSelectedStudentForCreateDocuments().cn_S);
            studentCharacterization.General_development_and_outlook = tbStudentGeneralDevelopmentAndOutlook.Text;
            studentCharacterization.Hobbies = tbStudentHobbies.Text;
            studentCharacterization.Attitude_to_physical_culture = tbStudentAttitudeToPhysicalCulture.Text;
            studentCharacterization.Academic_performance = tbStudentAcademicPerformance.Text;
            studentCharacterization.Most_favorite_subjects = tbStudentMostFavoriteSubjects.Text;
            studentCharacterization.Most_disliked_subjects = tbStudentMostDislikedSubjects.Text;
            studentCharacterization.Long_absences = tbStudentLongAbsences.Text;
            studentCharacterization.Communication_with_peers = tbStudentCommunicationWithPeers.Text;
            studentCharacterization.Communication_with_teachers = tbStudentCommunicationWithTeachers.Text;
            studentCharacterization.Psychological_features = tbStudentPsychologicalFeatures.Text;
            studentCharacterization.Signs_of_social_neglect = tbStudentSignsOfSocialNeglect.Text;
            studentCharacterization.Level_of_neuropsychological_stability = tbStudentLevelOfNeuropsychologicalStability.Text;
            studentCharacterization.Inclination_to_withdrawal = cbStudentInclinationToWithdrawal.IsChecked.Value;
            studentCharacterization.Temperament = tbStudentTemperament.Text;
            studentCharacterization.Self_assessment = tbStudentSelfAssessment.Text;

            return studentCharacterization;
        }
        #endregion

        #region Кнопа для формирования СПХ призывника
        private void BtnForCreateSPCofTheConscript_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedStudentForCreateDocuments())
                PrintCharacteristic();
        }

        private void ReplaceStringInDocuments(Word.Document document, string strForReplace, string str)
        {
            var range = document.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: strForReplace, ReplaceWith: str);
        }

        private void PrintCharacteristic()
        {
            var newWordApp = new Word.Application();
            newWordApp.Visible = false;

            var templateCharacteristic = newWordApp.Documents.Open(Environment.CurrentDirectory + "\\Print Templates\\Socio-psychological characteristics.docx");
            var characteristic = newWordApp.Documents.Add();
            characteristic.Content.Delete();

            templateCharacteristic.Range(templateCharacteristic.Content.Start, templateCharacteristic.Content.End).Copy();
            Word.Range rng = characteristic.Range(characteristic.Content.Start, characteristic.Content.End);
            rng.Paste();

            templateCharacteristic.Close();
            newWordApp.Visible = false;

            string studentID = GetSelectedStudentForCreateDocuments().cn_S;

            StudentPersonalInfo studentPersonalInfo = DataService.GetStudentPersonalInfo(studentID)[0];
            StudentCharacterization studentCharacterization = SetStudentCharacterization();
            Student student = GetSelectedStudentForCreateDocuments();
            string medGroupName = string.Empty;
            if (DataService.GetStudentMedicalGroupName(studentPersonalInfo.MedicalGroupName).Count != 0)
            {
                medGroupName = DataService.GetStudentMedicalGroupName(studentPersonalInfo.MedicalGroupName)[0].Name;
            }
            List<SUA_Employment> sUA_Employments = DataService.GetSUAEmployment(studentID);
            string employments = string.Empty;
            for(int i = 0; i < sUA_Employments.Count; i++)
            {
                employments += sUA_Employments[i].ActivitiesForm;
                if(i == sUA_Employments.Count-1)
                    employments += ".";
                else
                    employments += ", ";
            }
            string inclination_to_withdrawal = "не имеет";
            if(studentCharacterization.Inclination_to_withdrawal == true)
            {
                inclination_to_withdrawal = "имеет склонность к замкнутости";
            }
            string strFamilyTypes = string.Empty;
            List<FamilyType> familyTypes = DataService.GetTypesFamilyByStudent(studentID);
            for (int i = 0; i < familyTypes.Count; i++)
            {
                strFamilyTypes += familyTypes[i].FamilyType_Name;
                if (i == familyTypes.Count - 1)
                    strFamilyTypes += ".";
                else
                    strFamilyTypes += ", ";
            }
            string[] curatorFIO = curator.Split(' ');
            string curatorName = $"{curatorFIO[1]}{curatorFIO[2]} {curatorFIO[0]}";

            ReplaceStringInDocuments(characteristic, "{student}", $"{student.SurName} {student.Name} {student.FatherName}");
            ReplaceStringInDocuments(characteristic, "{dateBirth}", student.DateBirth.ToString("dd.MM.yyyy"));
            ReplaceStringInDocuments(characteristic, "{outlook}", studentCharacterization.General_development_and_outlook);
            ReplaceStringInDocuments(characteristic, "{hobbies}", studentCharacterization.Hobbies);
            ReplaceStringInDocuments(characteristic, "{medGroup}", medGroupName);
            ReplaceStringInDocuments(characteristic, "{employments}", employments);
            ReplaceStringInDocuments(characteristic, "{physical}", studentCharacterization.Attitude_to_physical_culture);
            ReplaceStringInDocuments(characteristic, "{performance}", studentCharacterization.Academic_performance);
            ReplaceStringInDocuments(characteristic, "{favoriteSubjects}", studentCharacterization.Most_favorite_subjects);
            ReplaceStringInDocuments(characteristic, "{dislikedSubjects}", studentCharacterization.Most_disliked_subjects);
            ReplaceStringInDocuments(characteristic, "{absences}", studentCharacterization.Long_absences);
            ReplaceStringInDocuments(characteristic, "{peers}", studentCharacterization.Communication_with_peers);
            ReplaceStringInDocuments(characteristic, "{teachers}", studentCharacterization.Communication_with_teachers);
            ReplaceStringInDocuments(characteristic, "{social}", studentCharacterization.Signs_of_social_neglect);
            ReplaceStringInDocuments(characteristic, "{psychological}", studentCharacterization.Level_of_neuropsychological_stability);
            ReplaceStringInDocuments(characteristic, "{temperament}", studentCharacterization.Temperament); 
            ReplaceStringInDocuments(characteristic, "{assessment}", studentCharacterization.Self_assessment);
            ReplaceStringInDocuments(characteristic, "{inclination}", inclination_to_withdrawal);
            ReplaceStringInDocuments(characteristic, "{familyTypes}", strFamilyTypes);
            ReplaceStringInDocuments(characteristic, "{kurator}", curatorName);

            MessageBox.Show("Печать выполнена");
            newWordApp.Visible = true;
        }
        #endregion

        #region Кнопа для формирования Характеристики учащегося
        private void BtnForCreateStudentProfile_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedStudentForCreateDocuments())
                PrintSPHCharacteristic();
        }

        private void PrintSPHCharacteristic()
        {
            var newWordApp = new Word.Application();
            newWordApp.Visible = false;

            var templateCharacteristic = newWordApp.Documents.Open(Environment.CurrentDirectory + "\\Print Templates\\Characteristic.docx");
            var characteristic = newWordApp.Documents.Add();
            characteristic.Content.Delete();

            templateCharacteristic.Range(templateCharacteristic.Content.Start, templateCharacteristic.Content.End).Copy();
            Word.Range rng = characteristic.Range(characteristic.Content.Start, characteristic.Content.End);
            rng.Paste();

            templateCharacteristic.Close();
            newWordApp.Visible = false;

            string studentID = GetSelectedStudentForCreateDocuments().cn_S;

            StudentPersonalInfo studentPersonalInfo = DataService.GetStudentPersonalInfo(studentID)[0];
            StudentCharacterization studentCharacterization = SetStudentCharacterization();
            Student student = GetSelectedStudentForCreateDocuments();
            string medGroupName = string.Empty;
            if (DataService.GetStudentMedicalGroupName(studentPersonalInfo.MedicalGroupName).Count != 0)
            {
                medGroupName = DataService.GetStudentMedicalGroupName(studentPersonalInfo.MedicalGroupName)[0].Name;
            }
            List<SUA_Employment> sUA_Employments = DataService.GetSUAEmployment(studentID);
            string employments = string.Empty;
            for (int i = 0; i < sUA_Employments.Count; i++)
            {
                employments += sUA_Employments[i].ActivitiesForm;
                if (i == sUA_Employments.Count - 1)
                    employments += ".";
                else
                    employments += ", ";
            }
            string inclination_to_withdrawal = "не имеет";
            if (studentCharacterization.Inclination_to_withdrawal == true)
            {
                inclination_to_withdrawal = "имеет";
            }
            string strFamilyTypes = string.Empty;
            List<FamilyType> familyTypes = DataService.GetTypesFamilyByStudent(studentID);
            for (int i = 0; i < familyTypes.Count; i++)
            {
                strFamilyTypes += familyTypes[i].FamilyType_Name;
                if (i == familyTypes.Count - 1)
                    strFamilyTypes += ".";
                else
                    strFamilyTypes += ", ";
            }
            string[] curatorFIO = curator.Split(' ');
            string curatorName = $"{curatorFIO[1]}{curatorFIO[2]} {curatorFIO[0]}";
            string strActiveSectors = string.Empty;
            List<ActiveSector> activeSectors = DataService.GetActiveSectorByStudent(studentID);
            for (int i = 0; i < activeSectors.Count; i++)
            {
                strActiveSectors += activeSectors[i].Name;
                if (i != activeSectors.Count - 1)
                    strActiveSectors += ", ";
            }

            ReplaceStringInDocuments(characteristic, "{student}", $"{student.SurName} {student.Name} {student.FatherName}");
            ReplaceStringInDocuments(characteristic, "{studentFIO}", $"{student.SurName} {student.Name[0]}.{student.FatherName[0]}.");
            ReplaceStringInDocuments(characteristic, "{studentFIO}", $"{student.SurName} {student.Name[0]}.{student.FatherName[0]}.");
            ReplaceStringInDocuments(characteristic, "{studentFIO}", $"{student.SurName} {student.Name[0]}.{student.FatherName[0]}.");
            ReplaceStringInDocuments(characteristic, "{studentFIO}", $"{student.SurName} {student.Name[0]}.{student.FatherName[0]}.");
            ReplaceStringInDocuments(characteristic, "{studentFIO}", $"{student.SurName} {student.Name[0]}.{student.FatherName[0]}.");
            ReplaceStringInDocuments(characteristic, "{yearStudy}", student.StateDateOfStudy.ToString("dd.MM.yyyy"));
            ReplaceStringInDocuments(characteristic, "{dateBirth}", student.DateBirth.ToString("dd.MM.yyyy"));
            ReplaceStringInDocuments(characteristic, "{outlook}", studentCharacterization.General_development_and_outlook);
            ReplaceStringInDocuments(characteristic, "{hobbies}", studentCharacterization.Hobbies);
            ReplaceStringInDocuments(characteristic, "{medGroup}", medGroupName);
            ReplaceStringInDocuments(characteristic, "{physical}", studentCharacterization.Attitude_to_physical_culture);
            ReplaceStringInDocuments(characteristic, "{favoriteSubjects}", studentCharacterization.Most_favorite_subjects);
            ReplaceStringInDocuments(characteristic, "{dislikedSubjects}", studentCharacterization.Most_disliked_subjects);
            ReplaceStringInDocuments(characteristic, "{absences}", studentCharacterization.Long_absences);
            ReplaceStringInDocuments(characteristic, "{peers}", studentCharacterization.Communication_with_peers);
            ReplaceStringInDocuments(characteristic, "{teachers}", studentCharacterization.Communication_with_teachers);
            ReplaceStringInDocuments(characteristic, "{assessment}", studentCharacterization.Self_assessment);
            ReplaceStringInDocuments(characteristic, "{inclination}", inclination_to_withdrawal);
            ReplaceStringInDocuments(characteristic, "{familyTypes}", strFamilyTypes); 
            ReplaceStringInDocuments(characteristic, "{sector}", strActiveSectors);
            ReplaceStringInDocuments(characteristic, "{special}", DataService.GetStudentSpecialtyName(studentID)[0]);
            ReplaceStringInDocuments(characteristic, "{kurator}", curatorName);

            MessageBox.Show("Печать выполнена");
            newWordApp.Visible = true;
        }
        #endregion
        #endregion
        #endregion

        #region Сведения об учащихся
        private void lbGroupListForStudentDetails_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CheckIsSelectedStudent())
            {
                string selectedStudentId = GetSelectedStudent().cn_S;

                dgStudentTipeOfFamily.ItemsSource = DataService.GetTypesFamilyByStudent(selectedStudentId);
                dgStudentFamilyInfo.ItemsSource = DataService.GetFamily(selectedStudentId);

                dgStudentSUAEmployment.ItemsSource = DataService.GetSUAEmployment(selectedStudentId);
                dgStudentActiveSector.ItemsSource = DataService.GetActiveSectorByStudent(selectedStudentId);

                dgStudentAssocialBehavior.ItemsSource = DataService.GetAssocialBehavior(selectedStudentId);
                dgStudentPromotionPunish.ItemsSource = DataService.GetPromotionPunish(selectedStudentId);

                cbStudentMedicalGroup.ItemsSource = DataService.GetAllMedicalGroups();

                GetSudentPersonalInfo(selectedStudentId);
            }
        }

        private void GetSudentPersonalInfo(string selectedStudentId)
        {
            StudentPersonalInfo studentInfo = DataService.GetStudentPersonalInfo(selectedStudentId)[0];

            tbStudentTelephone.Text = studentInfo.Telephone_Mob;
            tbStudentTelephoneHome.Text = studentInfo.Telephone_Home;
            if (studentInfo.DateBirth.Year != 1)
                dpStudentDateBirth.SelectedDate = studentInfo.DateBirth;
            else
                dpStudentDateBirth.SelectedDate = null;
            chbCitizenRB.IsChecked = studentInfo.RB;
            tbStudentPosportSeries.Text = studentInfo.PassportSeries;
            if (studentInfo.PassportNumber == null || studentInfo.PassportNumber == "0")
                tbStudentPosportNumber.Text = string.Empty;
            else
                tbStudentPosportNumber.Text = studentInfo.PassportNumber.ToString();
            tbStudentPosportIdentificNumber.Text = studentInfo.PasportID;
            tbStudentAdress.Text = studentInfo.Adress;
            if (studentInfo.FromAnotherTown == true)
            {
                chbStudentInnigites.IsChecked = true;

                if(studentInfo.OnHostel == true)
                {
                    chbStudentLivesHostal.IsChecked = studentInfo.OnHostel;
                    tbStudentRoomNumber.Text = studentInfo.RoomNumber.ToString();
                }
                else
                {
                    chbStudentLivesApartment.IsChecked = studentInfo.OnFlat;
                    tbStudentDescriptionHousing.Text = studentInfo.FlatDescription;
                }
            }
            else
            {
                chbStudentInnigites.IsChecked = false;
            }

            cbStudentMedicalGroup.SelectedValue = studentInfo.MedicalGroupName;
            chbStudentFormEducation.IsChecked = studentInfo.Budget;
            chbStudentMaritalStatus.IsChecked = studentInfo.FamilyState;
            if (studentInfo.OnIPA == true)
            {
                chbStudentIndividualProtAccounting.IsChecked = true;
                tbStudentDescripnIndividualProtAccounting.Text = studentInfo.IPARemarks;
            }
            else
            {
                chbStudentIndividualProtAccounting.IsChecked = false;
            }

            if (studentInfo.OnSDP == true)
            {
                chbStudentSOP.IsChecked = true;
                tbStudentDescripnSOP.Text = studentInfo.SDPRemarks;
            }
            else
            {
                chbStudentSOP.IsChecked = false;
            }

            if (studentInfo.OnNFSP == true)
            {
                chbStudentNGZ.IsChecked = true;
                tbStudentDescripnNGZ.Text = studentInfo.NFSPRemarks;
            }
            else
            {
                chbStudentNGZ.IsChecked = false;
            }
            if (studentInfo.IsDisabled == true)
            {
                chbStudentInvalid.IsChecked = true;
                tbStudentDescripnInvalid.Text = studentInfo.DisabledStudentRemarks;
            }
            else
            {
                chbStudentInvalid.IsChecked = false;
            }

            if (studentInfo.AnOrphan == true)
            {
                chbStudentOrphan.IsChecked = true;
                chbStudentGuardianship.IsChecked = studentInfo.OnGuardianship;
                chbStudentCustody.IsChecked = studentInfo.OnTrusteeship;
                chbStudentStateSecurity.IsChecked = studentInfo.OnStateSupport;
                chbStudentFoster.IsChecked = studentInfo.AnAdopted;
            }
            else
            {
                chbStudentOrphan.IsChecked = false;
            }
            if (studentInfo.HaveChildren == true)
            {
                chbStudentHaveChild.IsChecked = studentInfo.HaveChildren;
                tbStudentHaveChild.Text = studentInfo.HaveChildrenRemarks;
            }
            else
            {
                chbStudentHaveChild.IsChecked = false;
            }

            if (studentInfo.StateDateOfStudy.Year != 1)
                dtStudentDateStartEducation.SelectedDate = studentInfo.StateDateOfStudy;
            else
                dtStudentDateStartEducation.SelectedDate = null;
            tbStudentPreviousPlaseStudy.Text = studentInfo.PreviousPlaceOfStudy;

            if (studentInfo.OnDisabledParents == true)
            {
                chbStudentInvalidParents.IsChecked = true;
                tbStudentDescripnInvalidParents.Text = studentInfo.DisabledParentsRemarks;
            }
            else
            {
                chbStudentInvalidParents.IsChecked = false;
            }
        }

        private bool CheckIsSelectedStudent()
        {
            if (lbGroupListForStudentDetails.SelectedItem != null)
                return true;
            return Information("Выберите студента для редактирования!");
        }

        private Student GetSelectedStudent()
        {
            return (Student)lbGroupListForStudentDetails.SelectedItem;
        }

        private bool ConfirmationRemove()
        {
            MessageBoxResult result = MessageBox.Show("Вы действеительно хотите удалить выбранное поле?", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
                return true;
            return false;
        }

        #region Личные сведения
        private void BtnEditStudentInfo_Click(object sender, RoutedEventArgs e)
        {
            if(CheckIsSelectedStudent())
                ChengeVisibilityButtonsForStudentInfo();
        }

        private void BtnSaveStudentInfo_Click(object sender, RoutedEventArgs e)
        {
            StudentPersonalInfo personalInfo = SetStudentPersonalInfo();
            string idStudent = GetSelectedStudent().cn_S;
            DataService.UpdateStudentPersonalInfo(idStudent, personalInfo);
            ChengeVisibilityButtonsForStudentInfo();
            GetSudentPersonalInfo(idStudent);
        }

        private StudentPersonalInfo SetStudentPersonalInfo()
        {
            StudentPersonalInfo studentInfo = new StudentPersonalInfo();

            studentInfo.Telephone_Mob = tbStudentTelephone.Text;
            studentInfo.Telephone_Home = tbStudentTelephoneHome.Text;

            if(dpStudentDateBirth.SelectedDate != null)
                studentInfo.DateBirth = (DateTime)dpStudentDateBirth.SelectedDate;
            studentInfo.RB = chbCitizenRB.IsChecked.Value;
            studentInfo.PassportSeries = tbStudentPosportSeries.Text;
            studentInfo.PassportNumber = tbStudentPosportNumber.Text;
            studentInfo.PasportID = tbStudentPosportIdentificNumber.Text;

            studentInfo.Adress = tbStudentAdress.Text;
            studentInfo.FromAnotherTown = chbStudentInnigites.IsChecked.Value;
            studentInfo.OnHostel = chbStudentLivesHostal.IsChecked.Value;
            studentInfo.RoomNumber = tbStudentRoomNumber.Text;
            studentInfo.OnFlat = chbStudentLivesApartment.IsChecked.Value;
            studentInfo.FlatDescription = tbStudentDescriptionHousing.Text;

            studentInfo.MedicalGroupName = Convert.ToInt32(cbStudentMedicalGroup.SelectedValue);
            studentInfo.Budget = chbStudentFormEducation.IsChecked.Value;
            studentInfo.FamilyState = chbStudentMaritalStatus.IsChecked.Value;
            studentInfo.OnIPA = chbStudentIndividualProtAccounting.IsChecked.Value;
            studentInfo.IPARemarks = tbStudentDescripnIndividualProtAccounting.Text;

            studentInfo.OnSDP = chbStudentSOP.IsChecked.Value;
            studentInfo.SDPRemarks = tbStudentDescripnSOP.Text;
            studentInfo.OnNFSP = chbStudentNGZ.IsChecked.Value;
            studentInfo.NFSPRemarks = tbStudentDescripnNGZ.Text;
            studentInfo.IsDisabled = chbStudentInvalid.IsChecked.Value;
            studentInfo.DisabledStudentRemarks = tbStudentDescripnInvalid.Text;

            studentInfo.AnOrphan = chbStudentOrphan.IsChecked.Value;
            studentInfo.OnGuardianship = chbStudentGuardianship.IsChecked.Value;
            studentInfo.OnTrusteeship = chbStudentCustody.IsChecked.Value;
            studentInfo.OnStateSupport = chbStudentStateSecurity.IsChecked.Value;
            studentInfo.AnAdopted = chbStudentFoster.IsChecked.Value;
            studentInfo.HaveChildren = chbStudentHaveChild.IsChecked.Value;
            studentInfo.HaveChildrenRemarks = tbStudentHaveChild.Text;

            if (dtStudentDateStartEducation.SelectedDate != null)
                studentInfo.StateDateOfStudy = (DateTime)dtStudentDateStartEducation.SelectedDate;
            studentInfo.PreviousPlaceOfStudy = tbStudentPreviousPlaseStudy.Text;

            return studentInfo;
        }

        private void BtnCanselEditStudentInfo_Click(object sender, RoutedEventArgs e)
        {
            string selectedStudentId = GetSelectedStudent().cn_S;
            GetSudentPersonalInfo(selectedStudentId);
            ChengeVisibilityButtonsForStudentInfo();
        }

        private void ChengeVisibilityButtonsForStudentInfo()
        {
            if (BtnEditStudentInfo.Visibility == Visibility.Visible)
            {
                StudentPersonalInformation.IsEnabled = true;
                BtnSaveStudentInfo.Visibility = Visibility.Visible;
                BtnCanselEditStudentInfo.Visibility = Visibility.Visible;
                BtnEditStudentInfo.Visibility = Visibility.Hidden;
            }
            else
            {
                StudentPersonalInformation.IsEnabled = false;
                BtnSaveStudentInfo.Visibility = Visibility.Hidden;
                BtnCanselEditStudentInfo.Visibility = Visibility.Hidden;
                BtnEditStudentInfo.Visibility = Visibility.Visible;
            }
        }

        #region Чекбоксы
        private void chbStudentOrphan_Checked(object sender, RoutedEventArgs e)
        {
            chbStudentGuardianship.Visibility = Visibility.Visible;
            chbStudentCustody.Visibility = Visibility.Visible;
            chbStudentStateSecurity.Visibility = Visibility.Visible;
            chbStudentFoster.Visibility = Visibility.Visible;
        }

        private void chbStudentOrphan_Unchecked(object sender, RoutedEventArgs e)
        {
            chbStudentGuardianship.Visibility = Visibility.Hidden;
            chbStudentCustody.Visibility = Visibility.Hidden;
            chbStudentStateSecurity.Visibility = Visibility.Hidden;
            chbStudentFoster.Visibility = Visibility.Hidden;

            chbStudentGuardianship.IsChecked = false;
            chbStudentCustody.IsChecked = false;
            chbStudentStateSecurity.IsChecked = false;
            chbStudentFoster.IsChecked = false;
        }

        private void chbStudentSOP_Checked(object sender, RoutedEventArgs e)
        {
            lblStudentDescripnSOP.Visibility = Visibility.Visible;
            tbStudentDescripnSOP.Visibility = Visibility.Visible;
        }

        private void chbStudentSOP_Unchecked(object sender, RoutedEventArgs e)
        {
            lblStudentDescripnSOP.Visibility = Visibility.Hidden;
            tbStudentDescripnSOP.Visibility = Visibility.Hidden;
            tbStudentDescripnSOP.Text = string.Empty;
        }

        private void chbStudentNGZ_Checked(object sender, RoutedEventArgs e)
        {
            lblStudentDescripnNGZ.Visibility = Visibility.Visible;
            tbStudentDescripnNGZ.Visibility = Visibility.Visible;
        }

        private void chbStudentNGZ_Unchecked(object sender, RoutedEventArgs e)
        {
            lblStudentDescripnNGZ.Visibility = Visibility.Hidden;
            tbStudentDescripnNGZ.Visibility = Visibility.Hidden;
            tbStudentDescripnNGZ.Text = string.Empty;
        }

        private void chbStudentInvalid_Checked(object sender, RoutedEventArgs e)
        {
            lblStudentDescripnInvalid.Visibility = Visibility.Visible;
            tbStudentDescripnInvalid.Visibility = Visibility.Visible;
        }

        private void chbStudentInvalid_Unchecked(object sender, RoutedEventArgs e)
        {
            lblStudentDescripnInvalid.Visibility = Visibility.Hidden;
            tbStudentDescripnInvalid.Visibility = Visibility.Hidden;
            tbStudentDescripnInvalid.Text = string.Empty;
        }

        private void chbStudentIndividualProtAccounting_Checked(object sender, RoutedEventArgs e)
        {
            lblStudentDescripnIndividualProtAccounting.Visibility = Visibility.Visible;
            tbStudentDescripnIndividualProtAccounting.Visibility = Visibility.Visible;
        }

        private void chbStudentIndividualProtAccounting_Unchecked(object sender, RoutedEventArgs e)
        {
            lblStudentDescripnIndividualProtAccounting.Visibility = Visibility.Hidden;
            tbStudentDescripnIndividualProtAccounting.Visibility = Visibility.Hidden;
            tbStudentDescripnIndividualProtAccounting.Text = string.Empty;
        }

        private void chbStudentInnigites_Checked(object sender, RoutedEventArgs e)
        {
            chbStudentLivesApartment.Visibility = Visibility.Visible;
            chbStudentLivesHostal.Visibility = Visibility.Visible;
        }

        private void chbStudentInnigites_Unchecked(object sender, RoutedEventArgs e)
        {
            chbStudentLivesApartment.Visibility = Visibility.Hidden;
            chbStudentLivesHostal.Visibility = Visibility.Hidden;
            lblStudentRoomNumber.Visibility = Visibility.Hidden;
            tbStudentRoomNumber.Visibility = Visibility.Hidden;
            lblStudentApartmentDescription.Visibility = Visibility.Hidden;
            tbStudentDescriptionHousing.Visibility = Visibility.Hidden;

            chbStudentLivesApartment.IsChecked = false;
            chbStudentLivesHostal.IsChecked = false;
            tbStudentRoomNumber.Text = string.Empty;
            tbStudentDescriptionHousing.Text = string.Empty;
        }

        private void chbStudentLivesHostal_Checked(object sender, RoutedEventArgs e)
        {
            lblStudentRoomNumber.Visibility = Visibility.Visible;
            tbStudentRoomNumber.Visibility = Visibility.Visible;
            chbStudentLivesApartment.IsChecked = false;
        }

        private void chbStudentLivesHostal_Unchecked(object sender, RoutedEventArgs e)
        {
            lblStudentRoomNumber.Visibility = Visibility.Hidden;
            tbStudentRoomNumber.Visibility = Visibility.Hidden;
            tbStudentRoomNumber.Text = string.Empty;
        }

        private void chbStudentLivesApartment_Checked(object sender, RoutedEventArgs e)
        {
            lblStudentApartmentDescription.Visibility = Visibility.Visible;
            tbStudentDescriptionHousing.Visibility = Visibility.Visible;
            chbStudentLivesHostal.IsChecked = false;
        }

        private void chbStudentLivesApartment_Unchecked(object sender, RoutedEventArgs e)
        {
            lblStudentApartmentDescription.Visibility = Visibility.Hidden;
            tbStudentDescriptionHousing.Visibility = Visibility.Hidden;
            tbStudentDescriptionHousing.Text = string.Empty;
        }

        private void chbStudentHaveChild_Checked(object sender, RoutedEventArgs e)
        {
            lblStudentHaveChild.Visibility = Visibility.Visible;
            tbStudentHaveChild.Visibility = Visibility.Visible;
        }

        private void chbStudentHaveChild_Unchecked(object sender, RoutedEventArgs e)
        {
            lblStudentHaveChild.Visibility = Visibility.Hidden;
            tbStudentHaveChild.Visibility = Visibility.Hidden;
            tbStudentHaveChild.Text = string.Empty;
        }
        #endregion
        #endregion

        #region Сведения о семье
        private void BtnEditStudentFamilyInfo_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedStudent())
                ChengeVisibilityButtonsForFamily();
        }

        private void BtnSaveStudentFamilyInfo_Click(object sender, RoutedEventArgs e)
        {
            VisibilityButtonsForKindOfFamily();
            UnVisibilityButtonsForRelative();
            ChengeVisibilityButtonsForFamily();
            SaveStudentInvalidParents();
        }

        private void ChengeVisibilityButtonsForFamily()
        {
            if (BtnEditStudentFamilyInfo.Visibility == Visibility.Hidden)
                UnVisibilityButtonsForFamily();
            else
                VisibilityButtonsForFamily();
        }

        private void VisibilityButtonsForFamily()
        {
            BtnForAddRelative.Visibility = Visibility.Visible;
            BtnForEditRelative.Visibility = Visibility.Visible;
            BtnForDeleteRelative.Visibility = Visibility.Visible;

            BtnForAddKindOfFamily.Visibility = Visibility.Visible;
            BtnForDeleteKindOfFamily.Visibility = Visibility.Visible;

            BtnEditStudentFamilyInfo.Visibility = Visibility.Hidden;
            BtnSaveStudentFamilyInfo.Visibility = Visibility.Visible;

            chbStudentInvalidParents.IsEnabled = true;
            tbStudentDescripnInvalidParents.IsEnabled = true;
        }

        private void UnVisibilityButtonsForFamily()
        {
            BtnForAddRelative.Visibility = Visibility.Hidden;
            BtnForEditRelative.Visibility = Visibility.Hidden;
            BtnForDeleteRelative.Visibility = Visibility.Hidden;

            BtnForAddKindOfFamily.Visibility = Visibility.Hidden;
            BtnForDeleteKindOfFamily.Visibility = Visibility.Hidden;

            BtnEditStudentFamilyInfo.Visibility = Visibility.Visible;
            BtnSaveStudentFamilyInfo.Visibility = Visibility.Hidden;

            chbStudentInvalidParents.IsEnabled = false;
            tbStudentDescripnInvalidParents.IsEnabled = false;
        }

        #region Состав семьи
        private void BtnForDeleteRelative_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedRelative())
            {
                if (ConfirmationRemove())
                {
                    DataService.DeleteRelative(GetSelectRelative().Id_Relative);
                    dgStudentFamilyInfo.ItemsSource = DataService.GetFamily(GetSelectedStudent().cn_S);
                }
            }
        }

        private bool AddOrEditRelative = false;

        private void BtnForAddRelative_Click(object sender, RoutedEventArgs e)
        {
            AddOrEditRelative = true;
            ChengeVisibilityButtonsForRelative();
        }

        private void BtnForEditRelative_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedRelative())
            {
                AddOrEditRelative = false;
                FillingRelativePanel();
                ChengeVisibilityButtonsForRelative();
            }
        }

        private bool CheckIsSelectedRelative()
        {
            if (dgStudentFamilyInfo.SelectedItem != null)
                return true;
            return Information("Выберите родственника!");
        }

        private void FillingRelativePanel()
        {
            FamilyComposition relative = GetSelectRelative();

            tbRelativeName.Text = relative.FIO;
            tbRelativePlaceOfWork.Text = relative.Work_Study_Place;
            dpRelativeDateOfBirth.SelectedDate = relative.YearBirth;
            tbRelativePlaceOfResidence.Text = relative.Place_of_residence;
            cbRelationshipKind.SelectedValue = relative.ID_RF;
        }

        private void BtnForCanselAddOrEditRelative_Click(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityButtonsForRelative();
            ClearPanelForAddOrEditRelative();
        }

        private void BtnForAppliEditOrAddRelative_Click(object sender, RoutedEventArgs e)
        {
            if (AddOrEditRelative)
                AddRelative();
            else
                EditRelative();
        }

        private void AddRelative()
        {
            if (CheckingTheCorrecnessOfInformationAboutARelative())
            {
                DataService.AddRelative(GetInfoRelative());
                dgStudentFamilyInfo.ItemsSource = DataService.GetFamily(GetSelectedStudent().cn_S);
                ChengeVisibilityButtonsForRelative();
                ClearPanelForAddOrEditRelative();
            }
        }

        private void EditRelative()
        {
            if (CheckingTheCorrecnessOfInformationAboutARelative())
            {
                FamilyComposition relative = GetInfoRelative();
                relative.Id_Relative = GetSelectRelative().Id_Relative;

                DataService.EditRelative(relative);
                dgStudentFamilyInfo.ItemsSource = DataService.GetFamily(GetSelectedStudent().cn_S);
                ChengeVisibilityButtonsForRelative();
                ClearPanelForAddOrEditRelative();
            }
        }

        private bool CheckingTheCorrecnessOfInformationAboutARelative()
        {
            if (string.IsNullOrEmpty(tbRelativeName.Text))
                return Warning("Заполните поле Ф.И.О. родственника!");
            if (string.IsNullOrEmpty(tbRelativePlaceOfWork.Text))
                return Warning("Заполните поле Место работы/учёбы родственника!");
            if (dpRelativeDateOfBirth.SelectedDate == null)
                return Warning("Выберите дату рождения родственника!");
            if (string.IsNullOrEmpty(tbRelativePlaceOfResidence.Text))
                return Warning("Заполните поле Место проживание родственника!");
            if (Convert.ToInt32(cbRelationshipKind.SelectedValue) == 0)
                return Warning("Выберите форму родства!");
            return true;
        }

        private FamilyComposition GetInfoRelative()
        {
            FamilyComposition relative = new FamilyComposition();

            relative.cn_S = Convert.ToInt32(GetSelectedStudent().cn_S);
            relative.FIO = tbRelativeName.Text;
            relative.Work_Study_Place = tbRelativePlaceOfWork.Text;
            relative.YearBirth = (DateTime)dpRelativeDateOfBirth.SelectedDate;
            relative.Place_of_residence = tbRelativePlaceOfResidence.Text;
            relative.ID_RF = Convert.ToInt32(cbRelationshipKind.SelectedValue);

            return relative;
        }

        private FamilyComposition GetSelectRelative()
        {
            return (FamilyComposition)dgStudentFamilyInfo.SelectedItem;
        }

        private void ChengeVisibilityButtonsForRelative()
        {
            if(BtnForAddRelative.Visibility == Visibility.Visible)
                VisibilityButtonsForRelative();
            else
                UnVisibilityButtonsForRelative();
        }

        private void VisibilityButtonsForRelative()
        {
            borderForAddOrEditCompositionFamily.Visibility = Visibility.Visible;

            BtnForDeleteRelative.Visibility = Visibility.Hidden;
            BtnForEditRelative.Visibility = Visibility.Hidden;
            BtnForAddRelative.Visibility = Visibility.Hidden;

            BtnForAppliEditOrAddRelative.Visibility = Visibility.Visible;
            BtnForCanselAddOrEditRelative.Visibility = Visibility.Visible;
        }

        private void UnVisibilityButtonsForRelative()
        {
            borderForAddOrEditCompositionFamily.Visibility = Visibility.Hidden;

            BtnForDeleteRelative.Visibility = Visibility.Visible;
            BtnForEditRelative.Visibility = Visibility.Visible;
            BtnForAddRelative.Visibility = Visibility.Visible;

            BtnForAppliEditOrAddRelative.Visibility = Visibility.Hidden;
            BtnForCanselAddOrEditRelative.Visibility = Visibility.Hidden;
        }

        private void ClearPanelForAddOrEditRelative()
        {
            tbRelativeName.Clear();
            tbRelativePlaceOfResidence.Clear();
            tbRelativePlaceOfWork.Clear();
            cbRelationshipKind.SelectedItem = null;
            dpRelativeDateOfBirth.SelectedDate = null;
        }
        #endregion

        #region Вид семьи
        #region Удаление вида семьи
        private void BtnForDeleteKindOfFamily_Click(object sender, RoutedEventArgs e)
        {
            if (CheckForSelectedTypeFamilyForDeleteFamilyCharacteristics())
            {
                if (ConfirmationRemove())
                {
                    DataService.DeleteTypeFamily(GetSelectIdFamilyCharacteristics());
                    dgStudentTipeOfFamily.ItemsSource = DataService.GetTypesFamilyByStudent(GetSelectedStudent().cn_S);
                }
            }
        }

        private bool CheckForSelectedTypeFamilyForDeleteFamilyCharacteristics()
        {
            if (dgStudentTipeOfFamily.SelectedItem != null)
                return true;
            return Information("Выберите тип семьи!");
        }

        private int GetSelectIdFamilyCharacteristics()
        {
            return DataService.GetTypeFamilyByRelationAndStudent(GetFamilyCharacteristics())[0].id_Family;
        }

        private FamilyCharacteristics GetFamilyCharacteristics()
        {
            FamilyType familyType = (FamilyType)dgStudentTipeOfFamily.SelectedItem;

            FamilyCharacteristics familyCharacteristics = new FamilyCharacteristics();
            familyCharacteristics.cn_S = Convert.ToInt32(GetSelectedStudent().cn_S);
            familyCharacteristics.id_type = familyType.id_type;

            return familyCharacteristics;
        }
        #endregion

        #region Добавление вида семьи
        private void BtnForAddKindOfFamily_Click(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityButtonsForKindOfFamily();
        }

        private void BtnForCanselAddKindOfFamily_Click(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityButtonsForKindOfFamily();
        }

        private void BtnForAppliAddKindOfFamily_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedTypeFamilyForAddFamilyCharacteristics())
            {
                if (CheckForDublicateTypeFamily())
                {
                    DataService.AddTypeFamily(GetSelectTipeFamilyForAddFamilyCharacteristics());
                    dgStudentTipeOfFamily.ItemsSource = DataService.GetTypesFamilyByStudent(GetSelectedStudent().cn_S);
                    ChengeVisibilityButtonsForKindOfFamily();
                }
            }
        }

        private bool CheckIsSelectedTypeFamilyForAddFamilyCharacteristics()
        {
            if (cbFamilyTipe.SelectedItem == null)
                return Information("Выберите тип семьи!");
            return true;
        }

        private bool CheckForDublicateTypeFamily()
        {
            if (DataService.GetTypeFamilyByRelationAndStudent(GetSelectTipeFamilyForAddFamilyCharacteristics()).Count != 0)
                return Warning("Такой вид семьи уже присутствует у данного студента!\nВыберите дрогой вид семьи.");
            return true;
        }

        private FamilyCharacteristics GetSelectTipeFamilyForAddFamilyCharacteristics()
        {
            FamilyCharacteristics familyCharacteristics = new FamilyCharacteristics();

            familyCharacteristics.cn_S = Convert.ToInt32(GetSelectedStudent().cn_S);
            familyCharacteristics.id_type = Convert.ToInt32(cbFamilyTipe.SelectedValue);

            return familyCharacteristics;
        }

        private void ChengeVisibilityButtonsForKindOfFamily()
        {
            if (BtnForAddKindOfFamily.Visibility == Visibility.Visible)
                UnVisibilityButtonsForKindOfFamily();
            else
                VisibilityButtonsForKindOfFamily();
        }

        private void VisibilityButtonsForKindOfFamily()
        {
            BtnForAddKindOfFamily.Visibility = Visibility.Visible;
            BtnForDeleteKindOfFamily.Visibility = Visibility.Visible;

            BtnForCanselAddKindOfFamily.Visibility = Visibility.Hidden;
            BtnForAppliAddKindOfFamily.Visibility = Visibility.Hidden;

            borderForAddOrEditFamilyTipe.Visibility = Visibility.Hidden;
        }

        private void UnVisibilityButtonsForKindOfFamily()
        {
            BtnForAddKindOfFamily.Visibility = Visibility.Hidden;
            BtnForDeleteKindOfFamily.Visibility = Visibility.Hidden;

            BtnForCanselAddKindOfFamily.Visibility = Visibility.Visible;
            BtnForAppliAddKindOfFamily.Visibility = Visibility.Visible;

            borderForAddOrEditFamilyTipe.Visibility = Visibility.Visible;
        }
        #endregion
        #endregion

        private void chbStudentInvalidParents_Unchecked(object sender, RoutedEventArgs e)
        {
            lblStudentDescripnInvalidParents.Visibility = Visibility.Hidden;
            tbStudentDescripnInvalidParents.Visibility = Visibility.Hidden;
            tbStudentDescripnInvalidParents.Text = string.Empty;
        }

        private void chbStudentInvalidParents_Checked(object sender, RoutedEventArgs e)
        {
            lblStudentDescripnInvalidParents.Visibility = Visibility.Visible;
            tbStudentDescripnInvalidParents.Visibility = Visibility.Visible;
        }

        private void SaveStudentInvalidParents()
        {
            bool onDisabledParents = chbStudentInvalidParents.IsChecked.Value;
            string disabledParentsRemarks = tbStudentDescripnInvalidParents.Text;
            int idStudent = Convert.ToInt32(GetSelectedStudent().cn_S);
            DataService.UpdateInvalidParentsInfo(idStudent, onDisabledParents, disabledParentsRemarks);
        }
        #endregion

        #region Занятость ОПД
        private void BtnEditStudentSUAEmployment_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedStudent())
                ChengeVisibilityButtonsForSUAEmployment();
        }

        private void BtnSaveStudentSUAEmployment_Click(object sender, RoutedEventArgs e)
        {
            UnVisibilityButtonsForEmployment();
            VisibilityButtonsForActiveSector();
            ChengeVisibilityButtonsForSUAEmployment();
        }

        private void ChengeVisibilityButtonsForSUAEmployment()
        {
            if (BtnEditStudentSUAEmployment.Visibility == Visibility.Visible)
                VisibilityButtonsForSUAEmployment();
            else
                UnVisibilityButtonsForSUAEmployment();
        }

        private void VisibilityButtonsForSUAEmployment()
        {
            BtnForAddEmployment.Visibility = Visibility.Visible;
            BtnForEditEmployment.Visibility = Visibility.Visible;
            BtnForDeleteEmployment.Visibility = Visibility.Visible;

            BtnForAddActiveSector.Visibility = Visibility.Visible;
            BtnForDeleteActiveSector.Visibility = Visibility.Visible;

            BtnSaveStudentSUAEmployment.Visibility = Visibility.Visible;
            BtnEditStudentSUAEmployment.Visibility = Visibility.Hidden;
        }

        private void UnVisibilityButtonsForSUAEmployment()
        {
            BtnForAddEmployment.Visibility = Visibility.Hidden;
            BtnForEditEmployment.Visibility = Visibility.Hidden;
            BtnForDeleteEmployment.Visibility = Visibility.Hidden;

            BtnForAddActiveSector.Visibility = Visibility.Hidden;
            BtnForDeleteActiveSector.Visibility = Visibility.Hidden;

            BtnSaveStudentSUAEmployment.Visibility = Visibility.Hidden;
            BtnEditStudentSUAEmployment.Visibility = Visibility.Visible;
        }

        #region Формирование списка занятостей ОПД
        private void BtnForDeleteEmployment_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedEmployment())
            {
                if (ConfirmationRemove())
                {
                    DataService.DeleteSUAEmployment(GetSelectEmployment().ID_SUA_Emp);
                    dgStudentSUAEmployment.ItemsSource = DataService.GetSUAEmployment(GetSelectedStudent().cn_S);
                }
            }
        }

        private bool AddOrEditEmployment = false;

        private void BtnForAddEmployment_Click(object sender, RoutedEventArgs e)
        {
            AddOrEditEmployment = true;
            ChengeVisibilityButtonsForEmployment();
        }

        private void BtnForEditEmployment_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedEmployment())
            {
                AddOrEditEmployment = false;
                FillingEmploymentPanel();
                ChengeVisibilityButtonsForEmployment();
            }
        }

        private bool CheckIsSelectedEmployment()
        {
            if (dgStudentSUAEmployment.SelectedItem != null)
                return true;
            return Information("Выберите занятость!");
        }

        private void FillingEmploymentPanel()
        {
            SUA_Employment employment = GetSelectEmployment();

            tbEmploymentActivitiesForm.Text = employment.ActivitiesForm;
            tbEmploymenAchievements.Text = employment.Achievements;
            tbEmploymenNote.Text = employment.Note;
        }

        private void BtnForAppliEditOrAddEmployment_Click(object sender, RoutedEventArgs e)
        {
            if (AddOrEditEmployment)
                AddEmployment();
            else
                EditEmployment();
        }

        private void BtnForCanselAddOrEditEmployment_Click(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityButtonsForEmployment();
            ClearPanelForAddOrEditEmployment();
        }

        private void AddEmployment()
        {
            if (CheckingTheCorrecnessOfInforationAboutAEmployment())
            {
                DataService.AddSUAEmployment(GetInfoEmployment());
                dgStudentSUAEmployment.ItemsSource = DataService.GetSUAEmployment(GetSelectedStudent().cn_S);
                ChengeVisibilityButtonsForEmployment();
                ClearPanelForAddOrEditEmployment();
            }
        }

        private void EditEmployment()
        {
            if (CheckingTheCorrecnessOfInforationAboutAEmployment())
            {
                SUA_Employment employment = GetInfoEmployment();
                employment.ID_SUA_Emp = GetSelectEmployment().ID_SUA_Emp;

                DataService.EditSUAEmployment(employment);
                dgStudentSUAEmployment.ItemsSource = DataService.GetSUAEmployment(GetSelectedStudent().cn_S);
                ChengeVisibilityButtonsForEmployment();
                ClearPanelForAddOrEditEmployment();
            }
        }

        private bool CheckingTheCorrecnessOfInforationAboutAEmployment()
        {
            if (string.IsNullOrEmpty(tbEmploymentActivitiesForm.Text))
                return Warning("Заполните форму занятости ОПД!");
            return true;
        }

        private SUA_Employment GetInfoEmployment()
        {
            SUA_Employment employment = new SUA_Employment();

            employment.ActivitiesForm = tbEmploymentActivitiesForm.Text;
            employment.Achievements = tbEmploymenAchievements.Text;
            employment.Note = tbEmploymenNote.Text;
            employment.cn_S = GetSelectedStudent().cn_S;

            return employment;
        }

        private SUA_Employment GetSelectEmployment()
        {
            return (SUA_Employment)dgStudentSUAEmployment.SelectedItem;
        }

        private void ChengeVisibilityButtonsForEmployment()
        {
            if (BtnForAddEmployment.Visibility == Visibility.Visible)
                VisibilityButtonsForEmployment();
            else
                UnVisibilityButtonsForEmployment();
        }

        private void VisibilityButtonsForEmployment()
        {
            borderForAddOrEditSUA_Employment.Visibility = Visibility.Visible;

            BtnForDeleteEmployment.Visibility = Visibility.Hidden;
            BtnForEditEmployment.Visibility = Visibility.Hidden;
            BtnForAddEmployment.Visibility = Visibility.Hidden;

            BtnForAppliEditOrAddEmployment.Visibility = Visibility.Visible;
            BtnForCanselAddOrEditEmployment.Visibility = Visibility.Visible;
        }

        private void UnVisibilityButtonsForEmployment()
        {
            borderForAddOrEditSUA_Employment.Visibility = Visibility.Hidden;

            BtnForDeleteEmployment.Visibility = Visibility.Visible;
            BtnForEditEmployment.Visibility = Visibility.Visible;
            BtnForAddEmployment.Visibility = Visibility.Visible;

            BtnForAppliEditOrAddEmployment.Visibility = Visibility.Hidden;
            BtnForCanselAddOrEditEmployment.Visibility = Visibility.Hidden;
        }

        private void ClearPanelForAddOrEditEmployment()
        {
            tbEmploymentActivitiesForm.Clear();
            tbEmploymenAchievements.Clear();
            tbEmploymenNote.Clear();
        }
        #endregion

        #region Сектор актива
        #region Удаление сектора актива
        private void BtnForDeleteActiveSector_Click(object sender, RoutedEventArgs e)
        {
            if (CheckForSelectedActiveSectorForDelete())
            {
                if (ConfirmationRemove())
                {
                    DataService.DeleteActiveSector(GetActiveSector().ID_ActiveSector, GetSelectedStudent().cn_S);
                    dgStudentActiveSector.ItemsSource = DataService.GetActiveSectorByStudent(GetSelectedStudent().cn_S);
                }
            }
        }

        private bool CheckForSelectedActiveSectorForDelete()
        {
            if (dgStudentActiveSector.SelectedItem != null)
                return true;
            return Information("Выберите сектор актива!");
        }

        private ActiveSector GetActiveSector()
        {
            return (ActiveSector)dgStudentActiveSector.SelectedItem;
        }
        #endregion

        #region Добавление сектора актива
        private void BtnForAddActiveSector_Click(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityButtonsForActiveSector();
        }

        private void BtnForCanselAddActiveSector_Click(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityButtonsForActiveSector();
        }

        private void BtnForAppliAddActiveSector_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedActiveSectorForAdd())
                AddActiveSector();
        }

        private void AddActiveSector()
        {
            try
            {
                DataService.AddActiveSector(GetSelectTipeFamilyForAddActiveSector());
                dgStudentActiveSector.ItemsSource = DataService.GetActiveSectorByStudent(GetSelectedStudent().cn_S);
                ChengeVisibilityButtonsForActiveSector();
            }
            catch
            {
                MessageBox.Show("Учащийся уже состоит в выбранном секторе актива!" +
                    "\nВыберите дрогой сектор актива.", "Предуприждеие", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private bool CheckIsSelectedActiveSectorForAdd()
        {
            if (cbActiveSector.SelectedItem == null)
                return Warning("Выберите сектор актива!");
            return true;
        }

        private Assigments GetSelectTipeFamilyForAddActiveSector()
        {
            Assigments assigments = new Assigments();

            assigments.cn_S = Convert.ToInt32(GetSelectedStudent().cn_S);
            assigments.ID_ActiveSector = Convert.ToInt32(cbActiveSector.SelectedValue);

            return assigments;
        }

        private void ChengeVisibilityButtonsForActiveSector()
        {
            if (BtnForAddActiveSector.Visibility == Visibility.Visible)
                UnVisibilityButtonsForActiveSector();
            else
                VisibilityButtonsForActiveSector();
        }

        private void VisibilityButtonsForActiveSector()
        {
            BtnForAddActiveSector.Visibility = Visibility.Visible;
            BtnForDeleteActiveSector.Visibility = Visibility.Visible;

            BtnForCanselAddActiveSector.Visibility = Visibility.Hidden;
            BtnForAppliAddActiveSector.Visibility = Visibility.Hidden;

            borderForAddOrEditActiveSector.Visibility = Visibility.Hidden;
        }

        private void UnVisibilityButtonsForActiveSector()
        {
            BtnForAddActiveSector.Visibility = Visibility.Hidden;
            BtnForDeleteActiveSector.Visibility = Visibility.Hidden;

            BtnForCanselAddActiveSector.Visibility = Visibility.Visible;
            BtnForAppliAddActiveSector.Visibility = Visibility.Visible;

            borderForAddOrEditActiveSector.Visibility = Visibility.Visible;
        }
        #endregion
        #endregion
        #endregion

        #region Поощрения и асоциальное поведение
        private void BtnEditStudentAssocialBehavior_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedStudent())
                ChengeVisibilityButtonsForAssocialBehavior();
        }

        private void BtnSaveStudentAssocialBehavior_Click(object sender, RoutedEventArgs e)
        {
            UnVisibilityButtonsForBehavior();
            ChengeVisibilityButtonsForAssocialBehavior();
        }

        private void ChengeVisibilityButtonsForAssocialBehavior()
        {
            if (BtnEditStudentAssocialBehavior.Visibility == Visibility.Visible)
                VisibilityButtonsForAssocialBehavior();
            else
                UnVisibilityButtonsForAssocialBehavior();
        }

        private void VisibilityButtonsForAssocialBehavior()
        {
            BtnForAddAssocialBehavior.Visibility = Visibility.Visible;
            BtnForEditAssocialBehavior.Visibility = Visibility.Visible;
            BtnForDeleteAssocialBehavior.Visibility = Visibility.Visible;

            BtnForAddPromotionPunish.Visibility = Visibility.Visible;
            BtnForEditPromotionPunish.Visibility = Visibility.Visible;
            BtnForDeletePromotionPunish.Visibility = Visibility.Visible;

            BtnSaveStudentAssocialBehavior.Visibility = Visibility.Visible;
            BtnEditStudentAssocialBehavior.Visibility = Visibility.Hidden;
        }

        private void UnVisibilityButtonsForAssocialBehavior()
        {
            BtnForAddAssocialBehavior.Visibility = Visibility.Hidden;
            BtnForEditAssocialBehavior.Visibility = Visibility.Hidden;
            BtnForDeleteAssocialBehavior.Visibility = Visibility.Hidden;

            BtnForAddPromotionPunish.Visibility = Visibility.Hidden;
            BtnForEditPromotionPunish.Visibility = Visibility.Hidden;
            BtnForDeletePromotionPunish.Visibility = Visibility.Hidden;

            BtnSaveStudentAssocialBehavior.Visibility = Visibility.Hidden;
            BtnEditStudentAssocialBehavior.Visibility = Visibility.Visible;
        }

        #region Асоциальное поведение
        private bool AddOrEditAssocialBehavior = false;

        private void BtnForAddAssocialBehavior_Click(object sender, RoutedEventArgs e)
        {
            AddOrEditAssocialBehavior = true;
            ChengeVisibilityButtonsForBehavior();
        }

        private void BtnForEditAssocialBehavior_Click(object sender, RoutedEventArgs e)
        {
            if(CheckIsSelectedAssocialBehavior())
            {
                AddOrEditAssocialBehavior = false;
                FillingAssocialBehaviorPanel();
                ChengeVisibilityButtonsForBehavior();
            }
        }

        private bool CheckIsSelectedAssocialBehavior()
        {
            if (dgStudentAssocialBehavior.SelectedItem != null)
                return true;
            return Information("Выберите поведение!");
        }

        private void FillingAssocialBehaviorPanel()
        {
            AssocialBehavior associalBehavior = GetSelectAssocialBehavior();

            tbAssocialBehaviorContent.Text = associalBehavior.Content;
            tbAssocialBehaviorNature_Assoc_Beh.Text = associalBehavior.Nature_Assoc_Beh;
            tbAssocialBehaviorWorking_with_parents_students.Text = associalBehavior.Working_with_parents_students;
            tbAssocialBehaviorTakenMeasures.Text = associalBehavior.TakenMeasures;
            tbAssocialBehaviorResult.Text = associalBehavior.Result;
            tbAssocialBehaviorPsychologistsRecommendations.Text = associalBehavior.PsychologistsRecommendations;
            dpAssocialBehaviorDate.SelectedDate = associalBehavior.Date;
        }

        private AssocialBehavior GetSelectAssocialBehavior()
        {
            return (AssocialBehavior)dgStudentAssocialBehavior.SelectedItem;
        }

        private AssocialBehavior GetInfoAssocialBehavior()
        {
            AssocialBehavior associalBehavior = new AssocialBehavior();

            associalBehavior.Date = (DateTime)dpAssocialBehaviorDate.SelectedDate;
            associalBehavior.Content = tbAssocialBehaviorContent.Text;
            associalBehavior.Nature_Assoc_Beh = tbAssocialBehaviorNature_Assoc_Beh.Text;
            associalBehavior.Working_with_parents_students = tbAssocialBehaviorWorking_with_parents_students.Text;
            associalBehavior.TakenMeasures = tbAssocialBehaviorTakenMeasures.Text;
            associalBehavior.Result = tbAssocialBehaviorResult.Text;
            associalBehavior.PsychologistsRecommendations = tbAssocialBehaviorPsychologistsRecommendations.Text;
            associalBehavior.cn_S = Convert.ToInt32(GetSelectedStudent().cn_S);

            return associalBehavior;
        }

        private void BtnForDeleteAssocialBehavior_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedAssocialBehavior())
            {
                if (ConfirmationRemove())
                {
                    DataService.DeleteAssocialBehavior(GetSelectAssocialBehavior().ID_Assoc_beh);
                    dgStudentAssocialBehavior.ItemsSource = DataService.GetAssocialBehavior(GetSelectedStudent().cn_S);
                }
            }
        }

        private void BtnForAppliEditOrAddAssocialBehavior_Click(object sender, RoutedEventArgs e)
        {
            if (AddOrEditAssocialBehavior)
                AddAssocialBehavior();
            else
                EditAssocialBehavior();
        }

        private void AddAssocialBehavior()
        {
            if (CheckingTheCorrecnessOfInforationAboutAssocialBehavior())
            {
                DataService.AddAssocialBehavior(GetInfoAssocialBehavior());
                dgStudentAssocialBehavior.ItemsSource = DataService.GetAssocialBehavior(GetSelectedStudent().cn_S);
                ChengeVisibilityButtonsForBehavior();
                ClearPanelForAddOrEditAssocialBehavior();
            }
        }

        private void EditAssocialBehavior()
        {
            if (CheckingTheCorrecnessOfInforationAboutAssocialBehavior())
            {
                AssocialBehavior associalBehavior = GetInfoAssocialBehavior();
                associalBehavior.ID_Assoc_beh = GetSelectAssocialBehavior().ID_Assoc_beh;

                DataService.EditAssocialBehavior(associalBehavior);
                dgStudentAssocialBehavior.ItemsSource = DataService.GetAssocialBehavior(GetSelectedStudent().cn_S);
                ChengeVisibilityButtonsForBehavior();
                ClearPanelForAddOrEditAssocialBehavior();
            }
        }

        private bool CheckingTheCorrecnessOfInforationAboutAssocialBehavior()
        {
            if(string.IsNullOrEmpty(tbAssocialBehaviorContent.Text))
                return Warning("Заполните поле Содержание!");
            if(string.IsNullOrEmpty(tbAssocialBehaviorNature_Assoc_Beh.Text))
                return Warning("Заполните поле Хар-р проявления!");
            return true;
        }

        private void BtnForCanselAddOrEditAssocialBehavior_Click(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityButtonsForBehavior();
            ClearPanelForAddOrEditAssocialBehavior();
        }

        private void ChengeVisibilityButtonsForBehavior()
        {
            if (BtnForAddAssocialBehavior.Visibility == Visibility.Visible)
                VisibilityButtonsForBehavior();
            else
                UnVisibilityButtonsForBehavior();
        }

        private void VisibilityButtonsForBehavior()
        {
            borderForAddOrEditAssocialBehavior.Visibility = Visibility.Visible;

            BtnForDeleteAssocialBehavior.Visibility = Visibility.Hidden;
            BtnForEditAssocialBehavior.Visibility = Visibility.Hidden;
            BtnForAddAssocialBehavior.Visibility = Visibility.Hidden;

            BtnForAppliEditOrAddAssocialBehavior.Visibility = Visibility.Visible;
            BtnForCanselAddOrEditAssocialBehavior.Visibility = Visibility.Visible;
        }

        private void UnVisibilityButtonsForBehavior()
        {
            borderForAddOrEditAssocialBehavior.Visibility = Visibility.Hidden;

            BtnForDeleteAssocialBehavior.Visibility = Visibility.Visible;
            BtnForEditAssocialBehavior.Visibility = Visibility.Visible;
            BtnForAddAssocialBehavior.Visibility = Visibility.Visible;

            BtnForAppliEditOrAddAssocialBehavior.Visibility = Visibility.Hidden;
            BtnForCanselAddOrEditAssocialBehavior.Visibility = Visibility.Hidden;
        }

        private void ClearPanelForAddOrEditAssocialBehavior()
        {
            tbAssocialBehaviorContent.Clear();
            tbAssocialBehaviorNature_Assoc_Beh.Clear();
            tbAssocialBehaviorWorking_with_parents_students.Clear();
            tbAssocialBehaviorTakenMeasures.Clear();
            tbAssocialBehaviorResult.Clear();
            tbAssocialBehaviorPsychologistsRecommendations.Clear();
            dpAssocialBehaviorDate.SelectedDate = null;
        }
        #endregion

        #region Поощрения/Взыскания
        private bool AddOrEditPromotionPunish = false;

        private void BtnForAddPromotionPunish_Click(object sender, RoutedEventArgs e)
        {
            AddOrEditPromotionPunish = true;
            ChengeVisibilityButtonsForPromotionPunish();
        }

        private void BtnForAppliEditOrAddPromotionPunish_Click(object sender, RoutedEventArgs e)
        {
            if (AddOrEditPromotionPunish)
                AddPromotionPunish();
            else
                EditPromotionPunish();
        }

        private void AddPromotionPunish()
        {
            if (CheckingTheCorrecnessOfInforationAboutPromotionPunish())
            {
                DataService.AddPromotionPunish(GetInfoPromotionPunish());
                dgStudentPromotionPunish.ItemsSource = DataService.GetPromotionPunish(GetSelectedStudent().cn_S);
                ChengeVisibilityButtonsForPromotionPunish();
                ClearPanelForAddOrEditPromotionPunish();
            }
        }

        private void EditPromotionPunish()
        {
            if (CheckingTheCorrecnessOfInforationAboutPromotionPunish())
            {
                PromotionPunishView promotionPunish = GetInfoPromotionPunish();
                promotionPunish.id_Promotion = GetSelectPromotionPunish().id_Promotion;

                DataService.EditPromotionPunish(promotionPunish);
                dgStudentPromotionPunish.ItemsSource = DataService.GetPromotionPunish(GetSelectedStudent().cn_S);
                ChengeVisibilityButtonsForPromotionPunish();
                ClearPanelForAddOrEditPromotionPunish();
            }
        }

        private bool CheckingTheCorrecnessOfInforationAboutPromotionPunish()
        {
            if (string.IsNullOrEmpty(tbPromotionPunishDescription.Text))
                return Warning("Заполните поле Описание!");
            if (dpPromotionPunishDate.SelectedDate == null)
                return Warning("Заполните поле Дата!");
            return true;
        }

        private PromotionPunishView GetInfoPromotionPunish()
        {
            PromotionPunishView promotionPunish = new PromotionPunishView();

            promotionPunish.PPDescription = tbPromotionPunishDescription.Text;
            promotionPunish.PPDate = (DateTime)dpPromotionPunishDate.SelectedDate;
            if(cbPromotionPunishTypeName.SelectedValue != null)
                promotionPunish.id_Type = (int)cbPromotionPunishTypeName.SelectedValue;
            promotionPunish.cn_S = Convert.ToInt32(GetSelectedStudent().cn_S);

            return promotionPunish;
        }

        private PromotionPunishView GetSelectPromotionPunish()
        {
            return (PromotionPunishView)dgStudentPromotionPunish.SelectedItem;
        }

        private void BtnForEditPromotionPunish_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedPromotionPunish())
            {
                AddOrEditPromotionPunish = false;
                FillingPromotionPunishPanel();
                ChengeVisibilityButtonsForPromotionPunish();
            }
        }

        private bool CheckIsSelectedPromotionPunish()
        {
            if (dgStudentPromotionPunish.SelectedItem != null)
                return true;
            return Information("Выберите поощрение/взыскание");
        }

        private void FillingPromotionPunishPanel()
        {
            PromotionPunishView promotionPunish = GetSelectPromotionPunish();

            tbPromotionPunishDescription.Text = promotionPunish.PPDescription;
            cbPromotionPunishCategory.SelectedValue = promotionPunish.id_Category;
            cbPromotionPunishTypeName.SelectedValue = promotionPunish.id_Type;
            dpPromotionPunishDate.SelectedDate = promotionPunish.PPDate;
        }

        private void BtnForDeletePromotionPunish_Click(object sender, RoutedEventArgs e)
        {
            if (CheckIsSelectedPromotionPunish())
            {
                if (ConfirmationRemove())
                {
                    DataService.DeletePromotionPunish(GetSelectPromotionPunish().id_Promotion);
                    dgStudentPromotionPunish.ItemsSource = DataService.GetPromotionPunish(GetSelectedStudent().cn_S);
                }
            }
        }

        private void BtnForCanselAddOrEditPromotionPunish_Click(object sender, RoutedEventArgs e)
        {
            ChengeVisibilityButtonsForPromotionPunish();
            ClearPanelForAddOrEditPromotionPunish();
        }

        private void cbPromotionPunishCategory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cbPromotionPunishTypeName.ItemsSource = DataService.GetAllPromotionPunishType(Convert.ToInt32(cbPromotionPunishCategory.SelectedValue));
        }

        private void ChengeVisibilityButtonsForPromotionPunish()
        {
            if (BtnForAddPromotionPunish.Visibility == Visibility.Visible)
                VisibilityButtonsForPromotionPunish();
            else
                UnVisibilityButtonsForPromotionPunish();
        }

        private void VisibilityButtonsForPromotionPunish()
        {
            borderForAddOrEditPromotionPunish.Visibility = Visibility.Visible;

            BtnForAddPromotionPunish.Visibility = Visibility.Hidden;
            BtnForEditPromotionPunish.Visibility = Visibility.Hidden;
            BtnForDeletePromotionPunish.Visibility = Visibility.Hidden;

            BtnForAppliEditOrAddPromotionPunish.Visibility = Visibility.Visible;
            BtnForCanselAddOrEditPromotionPunish.Visibility = Visibility.Visible;
        }

        private void UnVisibilityButtonsForPromotionPunish()
        {
            borderForAddOrEditPromotionPunish.Visibility = Visibility.Hidden;

            BtnForAddPromotionPunish.Visibility = Visibility.Visible;
            BtnForEditPromotionPunish.Visibility = Visibility.Visible;
            BtnForDeletePromotionPunish.Visibility = Visibility.Visible;

            BtnForAppliEditOrAddPromotionPunish.Visibility = Visibility.Hidden;
            BtnForCanselAddOrEditPromotionPunish.Visibility = Visibility.Hidden;
        }

        private void ClearPanelForAddOrEditPromotionPunish()
        {
            tbPromotionPunishDescription.Clear();
            cbPromotionPunishCategory.SelectedItem = null;
            cbPromotionPunishTypeName.SelectedItem = null;
            dpPromotionPunishDate.SelectedDate = null;
        }
        #endregion
        #endregion
        #endregion

        #region О программе
        private void Oprogramme_Click(object sender, RoutedEventArgs e)
        {
            AboutTheProgram aboutTheProgram = new AboutTheProgram();
            aboutTheProgram.Show();
        }
        #endregion

        #region Выход
        private void bExit_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult messageBoxResult = MessageBox.Show("Вы действительное хотите выйти?", "Выход", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (messageBoxResult == MessageBoxResult.Yes)
            {
                Authorization authorization = new Authorization();
                authorization.Show();
                Close();
            }
        }
        #endregion
    }
}