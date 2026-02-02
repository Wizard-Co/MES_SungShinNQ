using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using WizMes_SungShinNQ.PopUp;
using WizMes_SungShinNQ.PopUP;
using System.Windows.Input;
using System.Threading;

namespace WizMes_SungShinNQ
{
    /// <summary>
    /// Win_Prd_ProcessResult_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Prd_WorkLog_Q : UserControl
    {
        Lib lib = new Lib();
        public Win_Prd_WorkLog_Q()
        {
            InitializeComponent();
            
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            lib.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            setComboBox();

            cboProcess.SelectedIndex = 0;
            cboMachine.SelectedIndex = 0;
            cboGubun.SelectedIndex = 0;
        }

        #region 콤보박스 세팅 setComboBox

        private void setComboBox()
        {
            ObservableCollection<CodeView> ovcProcess = ComboBoxUtil.Instance.GetWorkProcess(0, "");
            this.cboProcess.ItemsSource = ovcProcess;
            this.cboProcess.DisplayMemberPath = "code_name";
            this.cboProcess.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcMachine = GetMachineByProcessID("");
            this.cboMachine.ItemsSource = ovcMachine;
            this.cboMachine.DisplayMemberPath = "code_name";
            this.cboMachine.SelectedValuePath = "code_id";

            //라벨발행품 여부(입력)
            List<string[]> lstGubun = new List<string[]>();
            lstGubun.Add(new string[] { "0", "전체" });
            lstGubun.Add(new string[] { "1", "실적처리건" });
            lstGubun.Add(new string[] { "2", "실적처리 대기건" });
            lstGubun.Add(new string[] { "3", "오류건" });
            lstGubun.Add(new string[] { "4", "오류비포함" });

            ObservableCollection<CodeView> ovcGugunSearch = ComboBoxUtil.Instance.Direct_SetComboBox(lstGubun);
            this.cboGubun.ItemsSource = ovcGugunSearch;
            this.cboGubun.DisplayMemberPath = "code_name";
            this.cboGubun.SelectedValuePath = "code_id";
        }

        #endregion // 콤보박스 세팅 setComboBox

        #region mt_Machine - 호기 세팅

        /// <summary>
        /// 호기ID 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetMachineByProcessID(string value)
        {
            ObservableCollection<CodeView> ovcMachine = new ObservableCollection<CodeView>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("sProcessID", value);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sMachineForComboBoxAndUsing", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    CodeView CV = new CodeView();
                    CV.code_id = "";
                    CV.code_name = "전체";
                    ovcMachine.Add(CV);

                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {
                        CodeView mCodeView = new CodeView()
                        {
                            code_id = dr["Code"].ToString().Trim(),
                            code_name = dr["Name"].ToString().Trim()
                        };

                        ovcMachine.Add(mCodeView);
                    }
                }
            }

            return ovcMachine;
        }

        #endregion // mt_Machine - 호기 세팅

        #region 날짜버튼 클릭 이벤트

        // 전일 금일 전월 금월 버튼
        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpEDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpSDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }


        #endregion

        #region 검색조건 - 공정 콤보박스 선택 이벤트

        // 공정 콤보박스 선택 이벤트
        private void cboProcess_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboProcess.SelectedValue != null)
            {
                ObservableCollection<CodeView> ovcMachine = GetMachineByProcessID(cboProcess.SelectedValue.ToString());
                this.cboMachine.ItemsSource = ovcMachine;
                this.cboMachine.DisplayMemberPath = "code_name";
                this.cboMachine.SelectedValuePath = "code_id";

                cboMachine.SelectedIndex = 0;
            }
        }

        #endregion

        #region 버튼 클릭 이벤트

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                using (Loading lw = new Loading(beSearch))
                {
                    lw.ShowDialog();
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void beSearch()
        {
            FillGrid();

            if (dgdMain.Items.Count < 1)
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        // 엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "3차원 측정기 수집 이력";
            dgdStr[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMain);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;
                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
                    else
                        return;
                }
                else
                {
                    if (dt != null)
                    {
                        dt.Clear();
                    }
                }
            }
        }

        #endregion

        #region 주요 메서드 - 조회 FillGrid

        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {

                // 공정 호기 세팅
                string ProcessID = "";
                string MachineID = "";

                // 공정을 전체나 선택하지 않았을시 → 호기는 공정 + 호기로 출력 → 공정과 호기를 검색하기 위해서
                if (chkMachine.IsChecked == true
                    && cboMachine.SelectedValue != null
                    && cboMachine.SelectedValue.ToString().Trim().Length == 6)
                {
                    ProcessID = cboMachine.SelectedValue.ToString().Trim().Substring(0, 4);
                    MachineID = cboMachine.SelectedValue.ToString().Trim().Substring(4, 2);
                }
                else
                {
                    ProcessID = chkProcess.IsChecked == true && cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "";
                    MachineID = chkMachine.IsChecked == true && cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "";
                }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sFromDate", dtpSDate.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("sToDate", dtpEDate.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("sProcessID", ProcessID);
                sqlParameter.Add("sMachineID", MachineID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sWorkLog", sqlParameter, true);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WinR = new Win_Prd_WorkLog_Q_CodeView()
                            {
                                Num = i.ToString(),

                                LogID = dr["LogID"].ToString(),
                                WorkDate = DatePickerFormat(dr["WorkDate"].ToString()),
                                WorkQty = Convert.ToDouble(dr["WorkQty"]),
                                DefectQty = Convert.ToDouble(dr["DefectQty"]),
                                WorkTime = ConvertTimeFormat(dr["WorkTime"].ToString()),
                                StationNO = dr["StationNO"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),

                                Process = dr["Process"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                Machine = dr["Machine"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                                Comments = dr["Comments"].ToString(),

                                // 2026.02.02 
                                M1DayWorkQty = Convert.ToDouble(dr["M1DayWorkQty"]),
                                M1NightWorkQty = Convert.ToDouble(dr["M1NightWorkQty"]),
                                M2DayWorkQty = Convert.ToDouble(dr["M2DayWorkQty"]),
                                M2NightWorkQty = Convert.ToDouble(dr["M2NightWorkQty"]),
                                M3DayWorkQty = Convert.ToDouble(dr["M3DayWorkQty"]),
                                M3NightWorkQty = Convert.ToDouble(dr["M3NightWorkQty"]),
                            };

                            dgdMain.Items.Add(WinR);
                        }

                        tblCnt.Text = "▶검색 결과 : " + i.ToString() + "건";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        #endregion // 주요 메서드 - 조회 FillGrid

        private string StartTimeAndEndTime(string SDate, string STime, string EDate, string ETime)
        {
            string STandET = string.Empty;
            
            STandET += STime.Substring(0, 2) + ":" + STime.Substring(2, 2) + " ~ ";
            STandET += ETime.Substring(0, 2) + ":" + ETime.Substring(2, 2);

            return STandET;
        }


        #region 기타 메서드 모음

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        private string stringFormatN1(object obj)
        {
            return string.Format("{0:N1}", obj);
        }

        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
        }

        private string stringFormatNN(object obj, int length)
        {
            return string.Format("{0:N" + length + "}", obj);
        }

        // 데이터피커 포맷으로 변경
        private string DatePickerFormat(string str)
        {
            string result = "";

            if (str.Length == 8)
            {
                if (!str.Trim().Equals(""))
                {
                    result = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
                }
            }

            return result;
        }

        // Int로 변환
        private int ConvertInt(string str)
        {
            int result = 0;
            int chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");

                if (Int32.TryParse(str, out chkInt) == true)
                {
                    result = Int32.Parse(str);
                }
            }

            return result;
        }

        // 소수로 변환 가능한지 체크 이벤트
        private bool CheckConvertDouble(string str)
        {
            bool flag = false;
            double chkDouble = 0;

            if (!str.Trim().Equals(""))
            {
                if (Double.TryParse(str, out chkDouble) == true)
                {
                    flag = true;
                }
            }

            return flag;
        }

        // 숫자로 변환 가능한지 체크 이벤트
        private bool CheckConvertInt(string str)
        {
            bool flag = false;
            int chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace(",", "");

                if (Int32.TryParse(str, out chkInt) == true)
                {
                    flag = true;
                }
            }

            return flag;
        }

        // 소수로 변환
        private double ConvertDouble(string str)
        {
            double result = 0;
            double chkDouble = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");

                if (Double.TryParse(str, out chkDouble) == true)
                {
                    result = Double.Parse(str);
                }
            }

            return result;
        }

        // 시간 : 분 으로 변환
        private string ConvertTimeFormat(string str)
        {
            string result = "";

            str = str.Trim().Replace(":", "");
            if (str.Length > 5)
            {
                string hour = str.Substring(0, 2);
                string min = str.Substring(2, 2);
                string sec = str.Substring(4, 2);

                result = hour + ":" + min;
            }
            else if (str.Length > 3 && str.Length < 5)
            {
                string hour = str.Substring(0, 2);
                string min = str.Substring(2, 2);

                result = hour + ":" + min;
            }

            return result;
        }

        #endregion

    }

    class Win_Prd_WorkLog_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }
        public string Num { get; set; }
        public string LogID { get; set; }
        public string WorkDate { get; set; }
        public string WorkTime { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public double WorkQty { get; set; }
        public double DefectQty { get; set; }
        public string MachineID { get; set; }
        public string MachineNo { get; set; }
        public string Machine { get; set; }
        public string StationNO { get; set; }
        public string Comments { get; set; }

        public double M1DayWorkQty { get; set; }
        public double M1NightWorkQty { get; set; }
        public double M2DayWorkQty { get; set; }
        public double M2NightWorkQty { get; set; }
        public double M3DayWorkQty { get; set; }
        public double M3NightWorkQty { get; set; }


    }

   
}
