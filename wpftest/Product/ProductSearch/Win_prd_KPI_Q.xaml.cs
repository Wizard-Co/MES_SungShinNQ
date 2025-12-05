using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WizMes_SungShinNQ.PopUP;
using WizMes_SungShinNQ.PopUp;
using System.Text.RegularExpressions;

namespace WizMes_SungShinNQ
{
    /// <summary>
    /// Win_prd_KPI_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_KPI_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        int rowNum = 0;

        public Win_prd_KPI_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            lib.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        #region 상단 검색조건
        //전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dtpSDate.SelectedDate != null)
                {
                    DateTime ThatMonth1 = dtpSDate.SelectedDate.Value.AddDays(-(dtpSDate.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                    DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                    dtpSDate.SelectedDate = LastMonth1;
                    dtpEDate.SelectedDate = LastMonth31;
                }
                else
                {
                    DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                    DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                    dtpSDate.SelectedDate = LastMonth1;
                    dtpEDate.SelectedDate = LastMonth31;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnLastMonth_Click : " + ee.ToString());
            }
        }

        //전일
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dtpSDate.SelectedDate != null)
                {
                    dtpSDate.SelectedDate = dtpSDate.SelectedDate.Value.AddDays(-1);
                    dtpEDate.SelectedDate = dtpSDate.SelectedDate;
                }
                else
                {
                    dtpSDate.SelectedDate = DateTime.Today.AddDays(-1);
                    dtpEDate.SelectedDate = DateTime.Today.AddDays(-1);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnYesterday_Click : " + ee.ToString());
            }
        }
        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }
        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        #endregion

        #region Re_Search
        private void re_Search(int selectedIndex)
        {
            try
            {
                if (dgdOut.Items.Count > 0)
                {
                    dgdOut.Items.Clear();
                }

                if (dgdGonsu.Items.Count > 0)
                {
                    dgdGonsu.Items.Clear();
                }

                FillGrid();

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        #endregion

        #region 공수조회
        private void FillGrid()
        {
            try
            {       

                if (dgdOut.Items.Count > 0)
                {
                    dgdOut.Items.Clear();
                }
                if (dgdGonsu.Items.Count > 0)
                {
                    dgdGonsu.Items.Clear();
                }

        


                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("FromDate", dtpSDate.SelectedDate == null ? "" : dtpSDate.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("ToDate", dtpEDate.SelectedDate == null ? "" : dtpEDate.SelectedDate.Value.ToString("yyyyMMdd"));

                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true? 1:0); //품번
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true? txtBuyerArticleNoSrh.Tag != null ?  txtBuyerArticleNoSrh.Tag.ToString() :"":""); //품번

                sqlParameter.Add("ChkArticleID", ChkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", ChkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "":"" ); 
                ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sKPI_KPI", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var WPKQC = new Win_prd_KPI_Q_CodeView()
                            {
                                Num = i,
                                gbn = dr["GBN"].ToString(),
                                WorkMonth = DateTypeHyphen(dr["WorkMonth"].ToString()),
                                WorkQty = stringFormatN0(dr["WorkQty"]),
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                DefectRate = stringFormatN2Truncate(dr["DefectRate"]),
                                WorkTime = stringFormatN1(dr["WorkTime"]),
                                PerDayWorkQty = stringFormatN1(dr["PerDayWorkQty"]),
                                //WorkDays = stringFormatN0(dr["WorkDays"]),
                                WorkGoalRate = stringFormatN1(dr["WorkGoalRate"]),
                                WorkUpRate = stringFormatN1(dr["WorkUpRate"]),
                                Goal = dr["GBN"].ToString().Equals("P") ? stringFormatN0(dr["Goal"]) : stringFormatN2(dr["Goal"]),
                                MonthSort = dr["MonthSort"].ToString(),
                            };

                            if (WPKQC.MonthSort.Equals("999999"))
                            {
                                WPKQC.SetColor = true;
                                WPKQC.WorkMonth = "기간 합계";
                            }
                            

                            if (WPKQC.gbn.Equals("P"))
                            {
                                dgdGonsu.Items.Add(WPKQC);                                                             
                            }
                            else if (WPKQC.gbn.Equals("Q"))
                            {
                                dgdOut.Items.Add(WPKQC);                            
                            }



                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }
        #endregion

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                try
                {
                    rowNum = 0;
                    using (Loading lw = new Loading(FillGrid))
                    {
                        lw.ShowDialog();
                        
                        if (dgdGonsu.Items.Count <= 0 || dgdOut.Items.Count <= 0)
                        {
                            MessageBox.Show("조회된 내용이 없습니다.");
                        }
                        btnSearch.IsEnabled = true;
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show("예외처리 - " + ee.ToString());
                }

            }), System.Windows.Threading.DispatcherPriority.Background);


        }

        private void btiClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                lib.ChildMenuClose(this.ToString());
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void btiExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //if(dgdOut.Items.Count == 0 && dgdGonsu.Items.Count == 0)
                //{
                //    MessageBox.Show("먼저 검색해 주세요.");
                //    return;
                //}

                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[4];
                lst[0] = "생산성 향상";
                lst[1] = "품질 향상";
                lst[2] = dgdGonsu.Name;
                lst[3] = dgdOut.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {

                    if (ExpExc.choice.Equals(dgdGonsu.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdGonsu);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdGonsu);

                        Name = dgdGonsu.Name;
                        Lib.Instance.GenerateExcel(dt, Name);
                        Lib.Instance.excel.Visible = true;
                    }
                    else if (ExpExc.choice.Equals(dgdOut.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdOut);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdOut);

                        Name = dgdOut.Name;
                        Lib.Instance.GenerateExcel(dt, Name);
                        Lib.Instance.excel.Visible = true;
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
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void lblBuyerArticleNoSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuyerArticleNoSrh.IsChecked == true)
            {
                chkBuyerArticleNoSrh.IsChecked = false;
            }
            else
            {
                chkBuyerArticleNoSrh.IsChecked = true;
            }
        }
        // 거래처 체크박스 이벤트
        private void chkBuyerArticleNoSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkBuyerArticleNoSrh.IsChecked = true;
            txtBuyerArticleNoSrh.IsEnabled = true;
            btnBuyerArticleNoSrh.IsEnabled = true;
        }
        private void chkBuyerArticleNoSrh_UnChecked(object sender, RoutedEventArgs e)
        {
            chkBuyerArticleNoSrh.IsChecked = false;
            txtBuyerArticleNoSrh.IsEnabled = false;
            btnBuyerArticleNoSrh.IsEnabled = false;
        }
        // 거래처 텍스트박스 엔터 → 플러스파인더
        private void txtBuyerArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtBuyerArticleNoSrh, 76, "");
            }
        }
        // 거래처 플러스파인더 이벤트
        private void btnBuyerArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            // 거래처 : 0
            MainWindow.pf.ReturnCode(txtBuyerArticleNoSrh, 76, "");
        }

        //품명 라벨 클릭
        private void lblArticleIDSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (ChkArticleIDSrh.IsChecked == true)
            {
                ChkArticleIDSrh.IsChecked = false;
            }
            else
            {
                ChkArticleIDSrh.IsChecked = true;
            }
        }

        private void ChkArticleIDSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtArticleIDSrh.IsEnabled = true;
            btnArticleIDSrh.IsEnabled = true;
        }

        private void ChkArticleIDSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticleIDSrh.IsEnabled = false;
            btnArticleIDSrh.IsEnabled = false;
        }

        private void txtArticleIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void btnArticleIDSrh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천마리 콤마, 소수점 두자리
        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
        }
        // 천마리 콤마, 소수점 한자리
        private string stringFormatN1(object obj)
        {
            return string.Format("{0:N1}", obj);
        }

        private string stringFormatN2Truncate(object obj)
        {
            // 먼저 객체를 double로 변환
            double value = Convert.ToDouble(obj);

            // 소수점 2자리에서 버림 수행
            value = Math.Truncate(value * 100) / 100;

            // N2 형식으로 출력 (천 단위 구분자와 소수점 두 자리 포함)
            return value.ToString("N2");
        }

        private string DateTypeHyphen(string DigitsDate)
        {
            string pattern1 = @"(\d{4})(\d{2})(\d{2})";
            string pattern2 = @"(\d{4})(\d{2})(\d{2})(\d{4})(\d{2})(\d{2})";
            string pattern3 = @"(\d{4})(\d{2})";

            if (DigitsDate.Length == 8)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern1, "$1-$2-$3");
            }
            else if (DigitsDate.Length == 6)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern3, "$1-$2");
            }
            else if (DigitsDate.Length == 16)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern2, "$1-$2-$3 ~ $4-$5-$6");
            }          
            else if (DigitsDate.Length == 0)
            {
                DigitsDate = string.Empty;
            }

            return DigitsDate;
        }

        private string SetTimeColon(string time)
        {
            string conlonTime = string.Empty;

            if (time.Length == 4)
            {
                conlonTime = time.Substring(0, 2) + ":" + time.Substring(1, 2);
            }

            return conlonTime;
        }

    }

    #region CodeView
    class Win_prd_KPI_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }

        public string GbnName { get; set; }
        public string WorkDate { get; set; }
        public string ArticleNo { get; internal set; }
        public string Article { get; internal set; }
        public string WorkQty { get; internal set; }
        public string WorkMonth { get; set; }
        public string WorkTime { get; internal set; }
        public string WorkDays { get; set; }
        public string WorkQtyPerHour { get; internal set; }
        public string PerDayWorkQty { get; set; }
        public string WorkMan { get; set; }
        public string WorkUpRate { get; set; }
        public string WorkGoalRate { get; set; }
        public string DefectQty { get; set; }
        public string DefectWorkQty { get; set; }
        public string DefectRate { get; set; }
        public string DefectUpRate { get; set; }
        public string DefectGoalRate { get; set; }
        public string gbn { get; set; }
        public string Sort { get; set; }
        public string Goal { get; set; }
        public string MonthSort { get; set; }
        public bool SetColor { get; set; } = false;

    }

    #endregion

}