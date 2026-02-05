using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Shapes;
using System.Windows.Threading;
using WizMes_SungShinNQ.PopUp;
using WizMes_SungShinNQ.PopUP;
using WPF.MDI;

namespace WizMes_SungShinNQ
{
    /// <summary>
    /// Win_ord_InOutSum_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_InOutSum_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        // 그리드 셀렉트 도전(2018_08_09)
        int Clicked_row = 0;
        int Clicked_col = 0;
        List<Rectangle> PreRect = new List<Rectangle>();

        //전역변수는 이럴때 쓰는거 아니겠어??!!?
        private DataTable PeriodDataTable = null;
        private DataTable DaysDataTable = null;
        private DataTable MonthDataTable = null;
        private DataTable SpreadMonthDataTable = null;


        public Win_ord_InOutSum_Q()
        {
            InitializeComponent();
        }

        // 화면 첫 시작.
        private void Window_InOutTotalGrid_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            First_Step();
            ComboBoxSetting();
        }

        #region  첫 스텝 // 일자버튼 // 초기설정 // 조회용 체크박스 컨트롤 
        private void First_Step()
        {
            // 월별 가로집계 최근 3개월 지정하기.
            List<MonthChange> MC = new List<MonthChange>();
            MC.Add(new MonthChange()
            {
                H_MON1 = DateTime.Now.ToString("yyyy-MM"),
                H_MON2 = DateTime.Now.AddMonths(-1).ToString("yyyy-MM"),
                H_MON3 = DateTime.Now.AddMonths(-2).ToString("yyyy-MM"),
            });

            this.DataContext = MC;
            //////////////////////////////////////


            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            txtblMessage.Visibility = Visibility.Hidden;


        }

        // 어제.(전일)
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            //string[] receiver = lib.BringYesterdayDatetime();

            //dtpFromDate.Text = receiver[0];
            //dtpToDate.Text = receiver[1];

            if (dtpFromDate.SelectedDate != null)
            {
                dtpFromDate.SelectedDate = dtpFromDate.SelectedDate.Value.AddDays(-1);
                dtpToDate.SelectedDate = dtpFromDate.SelectedDate;
            }
            else
            {
                dtpFromDate.SelectedDate = DateTime.Today.AddDays(-1);
                dtpToDate.SelectedDate = DateTime.Today.AddDays(-1);
            }
        }
        // 오늘(금일)
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }
        // 지난 달(전월)
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //string[] receiver = lib.BringLastMonthDatetime();

            //dtpFromDate.Text = receiver[0];
            //dtpToDate.Text = receiver[1];

            if (dtpFromDate.SelectedDate != null)
            {
                DateTime ThatMonth1 = dtpFromDate.SelectedDate.Value.AddDays(-(dtpFromDate.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                dtpFromDate.SelectedDate = LastMonth1;
                dtpToDate.SelectedDate = LastMonth31;
            }

        }
        // 이번 달(금월)
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            string[] receiver = lib.BringThisMonthDatetime();

            dtpFromDate.SelectedDate = DateTime.Parse(receiver[0]);
            dtpToDate.SelectedDate = DateTime.Parse(receiver[1]);
        }


        #endregion

        #region 콤보박스 세팅
        // 콤보박스 세팅.
        private void ComboBoxSetting()
        {
            cboInOutGubunSrh.Items.Clear();
            cboInInspectGubunSrh.Items.Clear();

            string[] DirectCombo = new string[2];
            DirectCombo[0] = "Y";
            DirectCombo[1] = "합격";
            string[] DirectCombo1 = new string[2];
            DirectCombo1[0] = "N";
            DirectCombo1[1] = "불합격";

            List<string[]> DirectCombOList = new List<string[]>();
            DirectCombOList.Add(DirectCombo.ToArray());
            DirectCombOList.Add(DirectCombo1.ToArray());

            ObservableCollection<CodeView> cbInInspectGubunSrh = ComboBoxUtil.Instance.Direct_SetComboBox(DirectCombOList);

            DirectCombo = new string[2];
            DirectCombo[0] = "1";
            DirectCombo[1] = "입고";
            DirectCombo1 = new string[2];
            DirectCombo1[0] = "2";
            DirectCombo1[1] = "출고";

            DirectCombOList = new List<string[]>();
            DirectCombOList.Add(DirectCombo.ToArray());
            DirectCombOList.Add(DirectCombo1.ToArray());

            ObservableCollection<CodeView> cbInOutGubunSrh = ComboBoxUtil.Instance.Direct_SetComboBox(DirectCombOList);

            this.cboInOutGubunSrh.ItemsSource = cbInOutGubunSrh;
            this.cboInOutGubunSrh.DisplayMemberPath = "code_name";
            this.cboInOutGubunSrh.SelectedValuePath = "code_id";
            this.cboInOutGubunSrh.SelectedIndex = 0;

            this.cboInInspectGubunSrh.ItemsSource = cbInInspectGubunSrh;
            this.cboInInspectGubunSrh.DisplayMemberPath = "code_name";
            this.cboInInspectGubunSrh.SelectedValuePath = "code_id";
            this.cboInInspectGubunSrh.SelectedIndex = 0;

        }
        #endregion

        #region 플러스 파인더
        //플러스 파인더

        //거래처
        private void btnCustomIDSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomIDSrh, 0, "");
        }

        // 품명
        private void btnArticleIDSrh_click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticleIDSrh, 77, "");
        }

        #endregion


        // 검색(조회) 버튼 클릭
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if (tiMonth_H.IsSelected || tiMonth_V.IsSelected || lib.DatePickerCheck(dtpFromDate, dtpToDate, chkDateSrh) )
            {
                using (Loading ld = new Loading(beSearch))
                {
                    ld.ShowDialog();
                }
            }

        }

        private void beSearch()
        {
            //검색버튼 비활성화   
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {


                Console.WriteLine();

                DataStore.Instance.InsertLogByForm(this.GetType().Name, "R");
                TabItem NowTI = tabconGrid.SelectedItem as TabItem;

                if (NowTI.Header.ToString() == "기간집계") { FillGrid_Period(); }
                else if (NowTI.Header.ToString() == "일일집계") { FillGrid_Day(); }
                else if (NowTI.Header.ToString() == "월별집계(세로)") { FillGrid_Month_V(); }
                else if (NowTI.Header.ToString() == "월별집계(가로)") { FillGrid_Month_H(); }

            }), System.Windows.Threading.DispatcherPriority.Background);      

            btnSearch.IsEnabled = true;
        }

        #region 기간집계 조회
        //기간집계 조회
        private void FillGrid_Period()
        {
            //grdPeriod.Items.Clear();
            grdPeriod.ItemsSource = null;
            dgdPeriodTotal.Items.Clear();



            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", !lib.IsDatePickerNull(dtpFromDate) ? lib.ConvertDate(dtpFromDate) : "");
                sqlParameter.Add("EDate", !lib.IsDatePickerNull(dtpToDate) ? lib.ConvertDate(dtpToDate) : "");

                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : string.Empty : string.Empty);

                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkOrder", chkOrderIDSrh.IsChecked == true ? rbnOrderNOSrh.IsChecked == true ? 1 : 2 : 0);
                sqlParameter.Add("Order", chkOrderIDSrh.IsChecked == true ? !string.IsNullOrEmpty(txtOrderIDSrh.Text) ? txtOrderIDSrh.Text : "" : "");



                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_Period", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    PeriodDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {

                        int i = 0;                    
                        DataRowCollection drc = dt.Rows;

                        Win_ord_InOutSum_Total_QView total = new Win_ord_InOutSum_Total_QView();
                        List<Win_ord_InOutSum_QView> lstRows = new List<Win_ord_InOutSum_QView>();

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var PeriodItem = new Win_ord_InOutSum_QView
                            {
                                P_NUM = i,
                                P_Gbn = dr["Gbn"].ToString(),
                                P_CustomName = dr["KCustom"].ToString(),
                                P_OrderID = dr["OrderID"].ToString(),
                                P_OrderNo = dr["OrderNo"].ToString(),
                                P_BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                P_Article = dr["Article"].ToString(),
                                P_Spec = dr["Spec"].ToString(),
                                P_Roll = Lib.Instance.ToDecimal(dr["Roll"]),
                                P_Qty = Lib.Instance.ToDecimal(dr["TotQty"]),
                                P_UnitClssName = dr["UnitClssName"].ToString(),
                                P_CustomRate = Lib.Instance.ToDecimal(dr["CustomRate"]),

                            };

                            if (PeriodItem.P_Gbn.Equals("1"))
                            {
                                PeriodItem.P_Gbn = "입고";
                                total.P_TotalStuffRoll += PeriodItem.P_Roll;
                                total.P_TotalStuffQty += PeriodItem.P_Qty;
                                lstRows.Add(PeriodItem);

                            }
                            else if (PeriodItem.P_Gbn.Equals("2"))
                            {
                                PeriodItem.P_Gbn = "출고";
                                total.P_TotalOutRoll += PeriodItem.P_Roll;
                                total.P_TotalOutQty += PeriodItem.P_Qty;
                                lstRows.Add(PeriodItem);
                            }
                            else if (PeriodItem.P_Gbn.Equals("3"))
                            {
                                PeriodItem.P_Color1 = true;
                                PeriodItem.P_Gbn = string.Empty;
                                PeriodItem.P_Article = "거래처 계";
                                PeriodItem.P_CustomName = string.Empty;
                                PeriodItem.P_Article_TextAlignment = TextAlignment.Center;
                                lstRows.Add(PeriodItem);
                            }
                            else if (PeriodItem.P_Gbn.Equals("4"))
                            {
                                PeriodItem.P_Color2 = true;
                                PeriodItem.P_Gbn = string.Empty;
                                PeriodItem.P_CustomName = string.Empty;
                                PeriodItem.P_Article_TextAlignment = TextAlignment.Center;
                                lstRows.Add(PeriodItem);
                            }
                        }

                        grdPeriod.ItemsSource = lstRows;
                        dgdPeriodTotal.Items.Add(total);

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

        #endregion

        #region 일일집계 조회
        //일일집계 조회
        private void FillGrid_Day()
        {

            grdMergeDays.ItemsSource = null;
            dgdDaysOutTotal.Items.Clear();
            dgdDaysStuffTotal.Items.Clear();

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", !lib.IsDatePickerNull(dtpFromDate) ? lib.ConvertDate(dtpFromDate) : "");
                sqlParameter.Add("EDate", !lib.IsDatePickerNull(dtpToDate) ? lib.ConvertDate(dtpToDate) : "");

                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : string.Empty : string.Empty);

                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkOrder", chkOrderIDSrh.IsChecked == true ? rbnOrderNOSrh.IsChecked == true ? 1 : 2 : 0);
                sqlParameter.Add("Order", chkOrderIDSrh.IsChecked == true ? !string.IsNullOrEmpty(txtOrderIDSrh.Text) ? txtOrderIDSrh.Text : "" : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_Day", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    DaysDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;


                        int i = 0;

                        Win_ord_InOutSum_Total_QView StuffTotal = new Win_ord_InOutSum_Total_QView();
                        Win_ord_InOutSum_Total_QView OutTotal = new Win_ord_InOutSum_Total_QView();
                        List<Win_ord_InOutSum_QView> lstRows = new List<Win_ord_InOutSum_QView>();

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var DayItem = new Win_ord_InOutSum_QView
                            {
                                D_NUM = i,
                                D_IODate = lib.DateTypeHyphen(dr["IODate"].ToString()),
                                D_Gbn = dr["Gbn"].ToString(),
                                D_CustomName = dr["KCustom"].ToString(),
                                D_OrderID = dr["OrderID"].ToString(),
                                D_OrderNo = dr["OrderNo"].ToString(),
                                D_BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                D_Article = dr["Article"].ToString(),
                                D_Spec = dr["Spec"].ToString(),
                                D_Roll = Lib.Instance.ToDecimal(dr["Roll"]),
                                D_Qty = Lib.Instance.ToDecimal(dr["TotQty"]),
                                D_UnitClssName = dr["UnitClssName"].ToString(),
                                D_Amount = Lib.Instance.ToDecimal(dr["Amount"]),
                                D_VatAmount = Lib.Instance.ToDecimal(dr["VatAmount"]),
                                D_TotAmount = Lib.Instance.ToDecimal(dr["TotalAmount"]),
                                D_CustomRate = Lib.Instance.ToDecimal(dr["CustomRate"])

                            };

                            if (DayItem.D_Gbn.Equals("1"))
                            {
                                DayItem.D_Gbn = "입고";
                                StuffTotal.D_TotalStuffRoll += DayItem.D_Roll;
                                StuffTotal.D_TotalStuffQty += DayItem.D_Qty;
                                StuffTotal.D_TotalStuffAmount += DayItem.D_Amount;
                                StuffTotal.D_TotalStuffVatAmount += DayItem.D_VatAmount;
                                StuffTotal.D_TotalStuffPrice += DayItem.D_TotAmount;
                                lstRows.Add(DayItem);
                            }
                            else if (DayItem.D_Gbn.Equals("2"))
                            {
                                DayItem.D_Gbn = "출고";
                                OutTotal.D_TotalOutRoll += DayItem.D_Roll;
                                OutTotal.D_TotalOutQty += DayItem.D_Qty;
                                OutTotal.D_TotalOutAmount += DayItem.D_Amount;
                                OutTotal.D_TotalOutVatAmount += DayItem.D_VatAmount;
                                OutTotal.D_TotalStuffPrice += DayItem.D_TotAmount;
                                lstRows.Add(DayItem);


                            }
                            else if (DayItem.D_Gbn.Equals("3"))
                            {
                                DayItem.D_Color1 = true;
                                DayItem.D_Gbn = string.Empty;
                                DayItem.D_CustomName_TextAlignment = TextAlignment.Center;
                                lstRows.Add(DayItem);

                            }

                        }


                        grdMergeDays.ItemsSource = lstRows;
                        dgdDaysStuffTotal.Items.Add(StuffTotal);
                        dgdDaysOutTotal.Items.Add(OutTotal);

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

        #endregion

        #region 월별집계 (세로) 조회
        //월별집계 (세로) 조회
        private void FillGrid_Month_V()
        {

            grdMergeMonth_V.ItemsSource = null;
            dgdMonthVtotal.Items.Clear();

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", !lib.IsDatePickerNull(dtpFromDate) ? lib.ConvertDate(dtpFromDate) : "");
                sqlParameter.Add("EDate", !lib.IsDatePickerNull(dtpToDate) ? lib.ConvertDate(dtpToDate) : "");

                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : string.Empty : string.Empty);

                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkOrder", chkOrderIDSrh.IsChecked == true ? rbnOrderNOSrh.IsChecked == true ? 1 : 2 : 0);
                sqlParameter.Add("Order", chkOrderIDSrh.IsChecked == true ? !string.IsNullOrEmpty(txtOrderIDSrh.Text) ? txtOrderIDSrh.Text : "" : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_Month", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    MonthDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        //grdMerge_Month_V.RowDefinitions.Clear();

                        DataRowCollection drc = dt.Rows;
                        int i = 0;                  

                        List<Win_ord_InOutSum_QView> lstRows = new List<Win_ord_InOutSum_QView>();
                        Win_ord_InOutSum_Total_QView total = new Win_ord_InOutSum_Total_QView();
                        foreach (DataRow dr in drc)
                        {
                            i++; ;
                            var MonthHItem = new Win_ord_InOutSum_QView
                            {
                                V_NUM = i,
                                V_IODate = lib.DateTypeHyphen(dr["IODate"].ToString()),
                                V_Gbn = dr["Gbn"].ToString(),
                                V_CustomName = dr["KCustom"].ToString(),
                                V_BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                V_Article = dr["Article"].ToString(),
                                V_Spec = dr["Spec"].ToString(),
                                V_Roll = Lib.Instance.ToDecimal(dr["Roll"]),
                                V_Qty = Lib.Instance.ToDecimal(dr["TotQty"]),
                                V_UnitClssName = dr["UnitClssName"].ToString(),
                                V_CustomRate = Lib.Instance.ToDecimal(dr["CustomRate"])
                            };

                            if (MonthHItem.V_Gbn.Equals("1"))
                            {
                                MonthHItem.V_Gbn = "입고";
                                total.V_TotalStuffRoll += MonthHItem.V_Roll;
                                total.V_TotalStuffQty += MonthHItem.V_Qty;
                                lstRows.Add(MonthHItem);
                            }
                            else if (MonthHItem.V_Gbn.Equals("2"))
                            {
                                MonthHItem.V_Gbn = "출고";
                                total.V_TotalOutRoll += MonthHItem.V_Roll;
                                total.V_TotalOutQty += MonthHItem.V_Qty;
                                lstRows.Add(MonthHItem);
                            }
                            else if (MonthHItem.V_Gbn.Equals("3"))
                            {
                                MonthHItem.V_Color1 = true;
                                MonthHItem.V_Gbn = string.Empty;
                                MonthHItem.V_Article = "거래처 계";
                                MonthHItem.V_Article_TextAlignment = TextAlignment.Center;
                                MonthHItem.V_CustomName = string.Empty;
                                lstRows.Add(MonthHItem);
                            }

                        }

                        grdMergeMonth_V.ItemsSource = lstRows;
                        dgdMonthVtotal.Items.Add(total);


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

        #endregion

        #region 월별집계 최근 3개월 가로집계
        // 월별집계 (가로) (최근 3개월)
        private void FillGrid_Month_H()
        {

            try
            {
                grdMergeMonth_H.ItemsSource = null;
                dgdMonthHOutTotal.Items.Clear();
                dgdMonthHStuffTotal.Items.Clear();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", !lib.IsDatePickerNull(dtpFromDate) ? lib.ConvertDate(dtpFromDate) : "");
                sqlParameter.Add("EDate", !lib.IsDatePickerNull(dtpToDate) ? lib.ConvertDate(dtpToDate) : "");

                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : string.Empty : string.Empty);

                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkOrder", chkOrderIDSrh.IsChecked == true ? rbnOrderNOSrh.IsChecked == true ? 1 : 2 : 0);
                sqlParameter.Add("OrderID", chkOrderIDSrh.IsChecked == true ? !string.IsNullOrEmpty(txtOrderIDSrh.Text) ? txtOrderIDSrh.Text : "" : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_MonthSpread3", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    SpreadMonthDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        //grdMerge_Month_H.RowDefinitions.Clear();

                        DataRowCollection drc = dt.Rows;

                        int i = 0; 

                        List<Win_ord_InOutSum_QView> lstRows = new List<Win_ord_InOutSum_QView>();
                        Win_ord_InOutSum_Total_QView StuffTotal = new Win_ord_InOutSum_Total_QView();
                        Win_ord_InOutSum_Total_QView OutTotal = new Win_ord_InOutSum_Total_QView();
                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var MonthHItem = new Win_ord_InOutSum_QView
                            {
                                H_NUM = i,
                                H_Gbn = dr["Gbn"].ToString(),
                                H_CustomName = dr["KCustom"].ToString(),
                                H_Article = dr["Article"].ToString(),
                                H_BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                H_Spec = dr["Spec"].ToString(),
                                H_UnitClssName = dr["UnitClssName"].ToString(),

                                H_TotalMonthRoll = Lib.Instance.ToDecimal(dr["TotalRoll"]),
                                H_TotalMonthQty = Lib.Instance.ToDecimal(dr["TotalQty"]),
                                H_TotalMonthAmount = Lib.Instance.ToDecimal(dr["TotalAmount"]),
                                H_BaseMonthRoll = Lib.Instance.ToDecimal(dr["BaseMonthRoll"]),
                                H_BaseMonthQty = Lib.Instance.ToDecimal(dr["BaseMonthQty"]),
                                H_BaseMonthAmount = Lib.Instance.ToDecimal(dr["BaseMonthAmount"]),
                                H_Add1MonthRoll = Lib.Instance.ToDecimal(dr["Add1MonthRoll"]),
                                H_Add1MonthQty = Lib.Instance.ToDecimal(dr["Add1MonthQty"]),
                                H_Add1MonthAmount = Lib.Instance.ToDecimal(dr["Add1MonthAmount"]),
                                H_Add2MonthRoll = Lib.Instance.ToDecimal(dr["Add2MonthRoll"]),
                                H_Add2MonthQty = Lib.Instance.ToDecimal(dr["Add2MonthQty"]),
                                H_Add2MonthAmount = Lib.Instance.ToDecimal(dr["Add2MonthAmount"]),
                            };

                            if (MonthHItem.H_Gbn.Equals("2"))       //화면디자인이 출고가 먼저 나와야 하기에..
                            {
                                MonthHItem.H_Gbn = "출고";
                                OutTotal.H_TotalOutRoll += MonthHItem.H_TotalMonthRoll;
                                OutTotal.H_TotalOutQty += MonthHItem.H_TotalMonthQty;

                                OutTotal.H_TotalBaseOutRoll += MonthHItem.H_BaseMonthRoll;
                                OutTotal.H_TotalBaseOutQty += MonthHItem.H_BaseMonthQty;

                                OutTotal.H_TotalAdd1OutRoll += MonthHItem.H_Add1MonthRoll;
                                OutTotal.H_TotalAdd1OutQty += MonthHItem.H_Add1MonthQty;
                                OutTotal.H_TotalAdd2OutRoll += MonthHItem.H_Add2MonthRoll;
                                OutTotal.H_TotalAdd2OutQty += MonthHItem.H_Add2MonthQty;
                                lstRows.Add(MonthHItem);

                            }
                            else if (MonthHItem.H_Gbn.Equals("1"))
                            {
                                MonthHItem.H_Gbn = "입고";
                                StuffTotal.H_TotalStuffRoll += MonthHItem.H_TotalMonthRoll;
                                StuffTotal.H_TotalStuffQty += MonthHItem.H_TotalMonthQty;

                                StuffTotal.H_TotalBaseStuffRoll += MonthHItem.H_BaseMonthRoll;
                                StuffTotal.H_TotalBaseStuffQty += MonthHItem.H_BaseMonthQty;

                                StuffTotal.H_TotalAdd1StuffRoll += MonthHItem.H_Add1MonthRoll;
                                StuffTotal.H_TotalAdd1StuffQty += MonthHItem.H_Add1MonthQty;
                                StuffTotal.H_TotalAdd2StuffRoll += MonthHItem.H_Add2MonthRoll;
                                StuffTotal.H_TotalAdd2StuffQty += MonthHItem.H_Add2MonthQty;
                                lstRows.Add(MonthHItem);

                            }
                            else if (MonthHItem.H_Gbn.Equals("3") || MonthHItem.H_Gbn.Equals("4"))
                            {
                                MonthHItem.H_Color1 = true;
                                MonthHItem.H_Gbn = string.Empty;
                                MonthHItem.H_CustomName = string.Empty;
                                MonthHItem.H_Article_TextAlignment = TextAlignment.Center;
                                lstRows.Add(MonthHItem);
                            }                    
                        }

                        grdMergeMonth_H.ItemsSource = lstRows;                       

                        dgdMonthHOutTotal.Items.Add(OutTotal);
                        dgdMonthHStuffTotal.Items.Add(StuffTotal);
                  
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

        #endregion


        #region 월별 가로집계 셀렉션 체인지 이벤트
        // 탬 컨트롤 셀렉션 체인지 이벤트.
        private void tabconGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.Source != sender)
                return;
            if (dtpFromDate == null || dtpToDate == null)
                return;

            // 먼저 닫기
            dtpFromDate.IsDropDownOpen = false;
            dtpToDate.IsDropDownOpen = false;

            string sNowTI = ((sender as TabControl).SelectedItem as TabItem)?.Header as string;
            if (string.IsNullOrEmpty(sNowTI))
                return;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                Style newStyle;

                switch (sNowTI)
                {
                    case "기간집계":  
                    case "일일집계":
                        txtblMessage.Visibility = Visibility.Hidden;
                        dtpFromDate.IsEnabled = true;
                        dtpToDate.IsEnabled = true;
                        dtpToDate.Visibility = Visibility.Visible;
                        newStyle = FindResource("DatePickerSearch") as Style;
                        // 일별 버튼 표시
                        btnToday.Visibility = Visibility.Visible;
                        btnYesterday.Visibility = Visibility.Visible;
                        btnThisMonth.Visibility = Visibility.Visible;
                        btnLastMonth.Visibility = Visibility.Visible;
                        break;

                    case "월별집계(세로)":
                    case "월별집계(가로)":
                        txtblMessage.Visibility = (sNowTI == "월별집계(가로)") ? Visibility.Visible : Visibility.Hidden;
                        dtpFromDate.IsEnabled = true;
                        dtpToDate.IsEnabled = true;
                        dtpToDate.Visibility = Visibility.Hidden;
                        newStyle = FindResource("DatePickerMonthYearSearch") as Style;
                        // 일별 버튼 숨김
                        btnToday.Visibility = Visibility.Hidden;
                        btnYesterday.Visibility = Visibility.Hidden;
                        btnThisMonth.Visibility = Visibility.Visible;
                        btnLastMonth.Visibility = Visibility.Visible;
                        break;

                    default:
                        return;
                }

                dtpFromDate.Style = newStyle;
                dtpToDate.Style = newStyle;

                // Popup 바깥 클릭 감지
                var popup1 = dtpFromDate.Template.FindName("PART_Popup", dtpFromDate) as Popup;
                var popup2 = dtpToDate.Template.FindName("PART_Popup", dtpToDate) as Popup;
                if (popup1 != null)
                    popup1.StaysOpen = false;
                if (popup2 != null)
                    popup2.StaysOpen = false;

            }), DispatcherPriority.Background);
        }

        //private void tabconGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    string sNowTI = ((sender as TabControl).SelectedItem as TabItem).Header as string;

        //    switch (sNowTI)
        //    {
        //        case "기간집계":
        //            txtblMessage.Visibility = Visibility.Hidden;
        //            dtpFromDate.IsEnabled = true;
        //            dtpToDate.IsEnabled = true;
        //            break;
        //        case "일일집계":
        //            txtblMessage.Visibility = Visibility.Hidden;
        //            dtpFromDate.IsEnabled = true;
        //            dtpToDate.IsEnabled = true;
        //            break;
        //        case "월별집계(세로)":
        //            txtblMessage.Visibility = Visibility.Hidden;
        //            dtpFromDate.IsEnabled = true;
        //            dtpToDate.IsEnabled = true;
        //            break;
        //        case "월별집계(가로)":
        //            txtblMessage.Visibility = Visibility.Visible;
        //            dtpFromDate.IsEnabled = true;
        //            dtpToDate.IsEnabled = true;
        //            break;
        //        default: return;
        //    }
        //}


        #endregion


        //닫기 버튼 클릭.
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");

            int i = 0;
            foreach (MenuViewModel mvm in MainWindow.mMenulist)
            {
                if (mvm.subProgramID.ToString().Contains("MDI"))
                {
                    if (this.ToString().Equals((mvm.subProgramID as MdiChild).Content.ToString()))
                    {
                        (MainWindow.mMenulist[i].subProgramID as MdiChild).Close();
                        break;
                    }
                }
                i++;
            }
        }


        #region 엑셀

        // 엑셀 버튼 클릭.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            string sNowTI = (tabconGrid.SelectedItem as TabItem).Header as string;
            string Listname1 = string.Empty;
            string Listname2 = string.Empty;
            DataTable choicedt = null;
            Lib lib2 = new Lib();

            if (PeriodDataTable != null)
            {
                switch (sNowTI)
                {
                    case "기간집계":
                        if (PeriodDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 기간집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "기간집계";
                        Listname2 = "PeriodData";
                        choicedt = PeriodDataTable;
                        break;
                    case "일일집계":
                        if (DaysDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 일일집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "일일집계";
                        Listname2 = "DayData";
                        choicedt = DaysDataTable;
                        break;
                    case "월별집계(세로)":
                        if (MonthDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 월별(세로)집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "월(세로)집계";
                        Listname2 = "MonthData";
                        choicedt = MonthDataTable;
                        break;
                    case "월별집계(가로)":
                        if (SpreadMonthDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 월별(가로)집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "월(가로)집계";
                        Listname2 = "SpreadMonthData";
                        choicedt = SpreadMonthDataTable;
                        break;
                    default: return;
                }

                string Name = string.Empty;

                string[] lst = new string[4];
                lst[0] = Listname1;
                lst[2] = Listname2;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                // 어쨋든 머든 여기서 dt로 만들어서 주면 된다는 거네.
                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(Listname2))
                    {
                        Name = Listname2;
                        if (lib2.GenerateExcel(choicedt, Name))
                        {
                            DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                            lib2.excel.Visible = true;
                            lib2.ReleaseExcelObject(lib2.excel);
                        }
                    }
                    else
                    {
                        if (choicedt != null)
                        {
                            choicedt.Clear();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("엑설로 변환할 자료가 없습니다.");
            }

            lib2 = null;
        }





        #endregion







        private void DataGrid_SizeChange(object sender, SizeChangedEventArgs e)
        {
            DataGrid dgs = sender as DataGrid;
            if (dgs.ColumnHeaderHeight == 0)
            {
                dgs.ColumnHeaderHeight = 1;
            }
            double a = e.NewSize.Height / 100;
            double b = e.PreviousSize.Height / 100;
            double c = a / b;

            if (c != double.PositiveInfinity && c != 0 && double.IsNaN(c) == false)
            {
                dgs.ColumnHeaderHeight = dgs.ColumnHeaderHeight * c;
                dgs.FontSize = dgs.FontSize * c;
            }
        }

        private void CommonPlusfinder_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    TextBox txtbox = sender as TextBox;
                    if (txtbox != null)
                    {
                        if (txtbox.Name.Contains("CustomID"))
                        {
                            pf.ReturnCode(txtbox, 0, "");

                        }
                        else if (txtbox.Name.Contains("BuyerArticleNo"))
                        {
                            pf.ReturnCode(txtbox, 76, "");
                        }
                        else if (txtbox.Name.Contains("ArticleID"))
                        {
                            pf.ReturnCode(txtbox, 77, "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void CommonPlusfinder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TextBox txtbox = Lib.Instance.FindSiblingControl<TextBox>(sender as Button);
                if (txtbox != null)
                {
                    if (txtbox.Name.Contains("CustomID"))
                    {
                        pf.ReturnCode(txtbox, 0, "");

                    }
                    else if (txtbox.Name.Contains("BuyerArticleNo"))
                    {
                        pf.ReturnCode(txtbox, 76, "");
                    }
                    else if (txtbox.Name.Contains("ArticleID"))
                    {
                        pf.ReturnCode(txtbox, 77, "");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }


        private void CommonControl_Click(object sender, MouseButtonEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void CommonControl_Click(object sender, RoutedEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void rbnOrderNOSrh_Click(object sender, RoutedEventArgs e)
        {
            tblOrderIDSrh.Text = "OrderNo";
            dtcOrderID_P.Visibility = Visibility.Hidden;
            dtcOrderID_D.Visibility = Visibility.Hidden;
            dtcOrderNo_P.Visibility = Visibility.Visible;
            dtcOrderNo_D.Visibility = Visibility.Visible;
        }

        private void rbnOrderIDSrh_Click(object sender, RoutedEventArgs e)
        {
            tblOrderIDSrh.Text = "관리번호";
            dtcOrderID_P.Visibility = Visibility.Visible;
            dtcOrderID_D.Visibility = Visibility.Visible;
            dtcOrderNo_P.Visibility = Visibility.Hidden;
            dtcOrderNo_D.Visibility = Visibility.Hidden;
        }



        // 천 단위 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천 단위 콤마, 소수점 두자리
        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
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


    }





    /// <summary>
    /// /////////////////////////////////////////////////////////////////////
    /// </summary>


    public class MonthChange
    {
        //SpreadMonth 월 기간 확인용
        public string H_MON1 { get; set; }
        public string H_MON2 { get; set; }
        public string H_MON3 { get; set; }

    }



    class Win_ord_InOutSum_QView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // 조회 - 기간집계용 ( P_ (Period))
        public int P_NUM { get; set; }
        public string P_cls { get; set; }
        public string P_Gbn { get; set; }
        public string P_IODate { get; set; }
        public string P_CustomID { get; set; }
        public string P_CustomName { get; set; }
        public string P_OrderID { get; set; }
        public string P_OrderNo { get; set; }
        public string P_Sabun { get; set; }

        public string P_BuyerArticleNo { get; set; }
        public string P_ArticleID { get; set; }
        public string P_Article { get; set; }
        public TextAlignment P_Article_TextAlignment { get; set; } = TextAlignment.Left;
        public string P_Spec { get; set; }
        public decimal? P_Roll { get; set; }
        public decimal? P_Qty { get; set; }
        public string P_UnitClss { get; set; }

        public string P_UnitClssName { get; set; }
        public decimal? P_UnitPrice { get; set; }
        public string P_PriceClss { get; set; }
        public string P_PriceClssName { get; set; }
        public decimal? P_Amount { get; set; }

        public decimal? P_VatAmount { get; set; }
        public decimal? P_TotAmount { get; set; }
        public decimal? P_CustomRate { get; set; }
        public decimal? P_CustomRateOrder { get; set; }
        public bool P_Color1 { get; set; } = false;
        public bool P_Color2 { get; set; } = false;



        // 조회 - 일별집계용 ( D_ (Day))
        public int D_NUM { get; set; }
        public string D_cls { get; set; }
        public string D_Gbn { get; set; }
        public string D_IODate { get; set; }
        public string D_CustomID { get; set; }
        public string D_CustomName { get; set; }
        public TextAlignment D_CustomName_TextAlignment { get; set; } = TextAlignment.Left;
        public string D_OrderID { get; set; }
        public string D_OrderNo { get; set; }

        public string D_BuyerArticleNo { get; set; }
        public string D_ArticleID { get; set; }
        public string D_Article { get; set; }
        public string D_Spec { get; set; }
        public decimal? D_Roll { get; set; }
        public decimal? D_Qty { get; set; }
        public string D_UnitClss { get; set; }

        public string D_Sabun { get; set; }

        public string D_UnitClssName { get; set; }
        public decimal? D_UnitPrice { get; set; }
        public string D_PriceClss { get; set; }
        public string D_PriceClssName { get; set; }
        public decimal? D_Amount { get; set; }

        public decimal? D_VatAmount { get; set; }
        public decimal? D_TotAmount { get; set; }
        public decimal? D_CustomRate { get; set; }
        public decimal? D_CustomRateOrder { get; set; }
        public bool D_Color1 { get; set; } = false;
        public bool D_Color2 { get; set; } = false;


        // 조회 - 월별집계용 _V ( V_ (V_Month))
        public int V_NUM { get; set; }
        public string V_cls { get; set; }
        public string V_Gbn { get; set; }
        public string V_IODate { get; set; }
        public string V_CustomID { get; set; }
        public string V_CustomName { get; set; }
        public string V_OrderID { get; set; }
        public string V_OrderNo { get; set; }

        public string V_BuyerArticleNo { get; set; }
        public string V_ArticleID { get; set; }
        public string V_Article { get; set; }
        public TextAlignment V_Article_TextAlignment { get; set; } = TextAlignment.Left;
        public string V_Spec { get; set; }  
        public decimal? V_Roll { get; set; }
        public decimal? V_Qty { get; set; }
        public string V_UnitClss { get; set; }

        public string V_Sabun { get; set; }

        public string V_UnitClssName { get; set; }
        public decimal? V_UnitPrice { get; set; }
        public string V_PriceClss { get; set; }
        public string V_PriceClssName { get; set; }
        public decimal? V_Amount { get; set; }

        public decimal? V_VatAmount { get; set; }
        public decimal? V_TotAmount { get; set; }
        public decimal? V_CustomRate { get; set; }
        public decimal? V_CustomRateOrder { get; set; }
        public string V_RN { get; set; }

        public bool V_Color1 { get; set; } = false;
        public bool V_Color2 { get; set; } = false;



        // 조회 - 월별집계용 _H ( H_ (H_Month))
        public int H_NUM { get; set; }
        public string H_cls { get; set; }
        public string H_Gbn { get; set; }
        public string H_CustomID { get; set; }
        public string H_CustomName { get; set; }
        public string H_BuyerArticleNo { get; set; }
        public string H_ArticleID { get; set; }
        public string H_Article { get; set; }
        public TextAlignment H_Article_TextAlignment { get; set; } = TextAlignment.Left;
        public string H_Spec { get; set; }
        public string H_UnitClss { get; set; }
        public string H_UnitClssName { get; set; }
        public decimal? H_UnitPrice { get; set; }

        public string H_Sabun { get; set; }

        public string H_PriceClss { get; set; }
        public string H_PriceClssName { get; set; }
        public string H_YYYYMM1 { get; set; }
        public string H_YYYYMM2 { get; set; }
        public string H_YYYYMM3 { get; set; }

        public string H_YYYYMM4 { get; set; }
        public string H_YYYYMM5 { get; set; }
        public string H_YYYYMM6 { get; set; }
        public string H_YYYYMM7 { get; set; }
        public string H_YYYYMM8 { get; set; }

        public string H_YYYYMM9 { get; set; }
        public string H_YYYYMM10 { get; set; }
        public string H_roll10 { get; set; }
        public string H_Qty10 { get; set; }
        public string H_Amount10 { get; set; }

        public string H_VatAmount10 { get; set; }
        public string H_YYYYMM11 { get; set; }
        public string H_roll11 { get; set; }
        public string H_Qty11 { get; set; }
        public string H_Amount11 { get; set; }

        public string H_VatAmount11 { get; set; }
        public string H_YYYYMM12 { get; set; }
        public string H_roll12 { get; set; }
        public string H_Qty12 { get; set; }
        public string H_Amount12 { get; set; }

        public string H_VatAmount12 { get; set; }
        public string H_YYYYMM13 { get; set; }
        public string H_roll13 { get; set; }
        public string H_Qty13 { get; set; }
        public string H_Amount13 { get; set; }

        public string H_VatAmount13 { get; set; }
        public string H_RN { get; set; }
        public string H_CustomRate { get; set; }
        public string H_CustomAmount { get; set; }
        public string H_AllTotalAmount { get; set; }

        public decimal? H_TotalMonthQty { get; set; }
        public decimal? H_TotalMonthRoll { get; set; }
        public decimal? H_TotalMonthAmount { get; set; }
        public decimal? H_BaseMonthQty { get; set; }
        public decimal? H_BaseMonthRoll { get; set; }
        public decimal? H_BaseMonthAmount { get; set; }
        public decimal? H_Add1MonthQty { get; set; }
        public decimal? H_Add1MonthRoll { get; set; }
        public decimal? H_Add1MonthAmount { get; set; }
        public decimal? H_Add2MonthQty { get; set; }
        public decimal? H_Add2MonthRoll { get; set; }
        public decimal? H_Add2MonthAmount { get; set; }
        public bool H_Color1 { get; set; } = false;
        public bool H_Color2 { get; set; } = false;


        public List<P_listmodel> P_listmodel { get; set; }
        public List<D_gbnmodel> D_gbnmodel { get; set; }
        public List<V_gbnmodel> V_gbnmodel { get; set; }
        public List<H_custommodel> H_custommodel { get; set; }


    }

    public class D_gbnmodel
    {
        public string D_Gbn { get; set; }
        public string D_YesColor { get; set; }

        public List<D_custommodel> D_custommodel { get; set; }
    }

    public class V_gbnmodel
    {
        public string V_Gbn { get; set; }
        public List<V_custommodel> V_custommodel { get; set; }
    }



    public class D_custommodel
    {
        public string D_CustomName { get; set; }
        public List<D_listmodel> D_listmodel { get; set; }
    }

    public class V_custommodel
    {
        public string V_CustomName { get; set; }
        public string V_YesColor { get; set; }
        public List<V_listmodel> V_listmodel { get; set; }
    }

    public class H_custommodel
    {
        public string H_CustomName { get; set; }
        public List<H_listmodel> H_listmodel { get; set; }
    }



    public class D_listmodel
    {
        public string D_ArticleID { get; set; }
        public string D_Article { get; set; }
        public string D_Roll { get; set; }
        public string D_Qty { get; set; }
        public string D_UnitClssName { get; set; }
        public string D_PriceClssName { get; set; }

        public string D_VatAmount { get; set; }
        public string D_TotAmount { get; set; }
        public string D_CustomRate { get; set; }

    }

    public class P_listmodel
    {
        public string P_ArticleID { get; set; }
        public string P_Article { get; set; }
        public string P_Roll { get; set; }
        public string P_Qty { get; set; }
        public string P_UnitClssName { get; set; }
        public string P_CustomRate { get; set; }

        public string P_YesColor { get; set; }

    }

    public class V_listmodel
    {
        public string V_ArticleID { get; set; }
        public string V_Article { get; set; }
        public string V_Roll { get; set; }
        public string V_Qty { get; set; }
        public string V_UnitClssName { get; set; }
        public string V_CustomRate { get; set; }
    }



    public class H_listmodel
    {
        public string H_ArticleID { get; set; }
        public string H_Article { get; set; }
        public string H_UnitClssName { get; set; }
        public string H_PriceClssName { get; set; }
        public string H_roll10 { get; set; }
        public string H_Qty10 { get; set; }
        public string H_CustomRate { get; set; }

        public string H_roll11 { get; set; }
        public string H_Qty11 { get; set; }
        public string H_roll12 { get; set; }
        public string H_Qty12 { get; set; }
        public string H_roll13 { get; set; }
        public string H_Qty13 { get; set; }

    }

    public class Win_ord_InOutSum_Total_QView : BaseView
    {
        public decimal? P_TotalOutRoll { get; set; } = 0m;
        public decimal? P_TotalOutQty { get; set; } = 0m;
        public decimal? P_TotalStuffRoll { get; set; } = 0m;
        public decimal? P_TotalStuffQty { get; set; } = 0m;

        public decimal? D_TotalOutRoll { get; set; } = 0m; 
        public decimal? D_TotalOutQty { get; set; } = 0m;
        public decimal? D_TotalOutAmount { get; set; } = 0m;
        public decimal? D_TotalOutVatAmount { get; set; } = 0m;
        public decimal? D_TotalOutPrice { get; set; } = 0m;
        public decimal? D_TotalStuffRoll { get; set; } = 0m;
        public decimal? D_TotalStuffQty { get; set; } = 0m;
        public decimal? D_TotalStuffAmount { get; set; } = 0m;
        public decimal? D_TotalStuffVatAmount { get; set; } = 0m;
        public decimal? D_TotalStuffPrice { get; set; } = 0m;
        public bool D_Color1 { get; set; } = false;
        public bool D_Color2 { get; set; } = false;

        public decimal? V_TotalOutRoll { get; set; } = 0m;
        public decimal? V_TotalOutQty { get; set; } = 0m;
        public decimal? V_TotalStuffRoll { get; set; } = 0m;
        public decimal? V_TotalStuffQty { get; set; } = 0m;
        public bool V_Color1 { get; set; } = false;
        public bool V_Color2 { get; set; } = false;

        public decimal? H_TotalOutRoll { get; set; } = 0m;
        public decimal? H_TotalOutQty { get; set; } = 0m;
        public decimal? H_TotalStuffRoll { get; set; } = 0m;
        public decimal? H_TotalStuffQty { get; set; } = 0m;

        public decimal? H_TotalBaseOutRoll { get; set; } = 0m;
        public decimal? H_TotalBaseOutQty { get; set; } = 0m;
        public decimal? H_TotalBaseStuffRoll { get; set; } = 0m;
        public decimal? H_TotalBaseStuffQty { get; set; } = 0m;

        public decimal? H_TotalAdd1OutRoll { get; set; } = 0m;
        public decimal? H_TotalAdd1OutQty { get; set; } = 0m;
        public decimal? H_TotalAdd1StuffRoll { get; set; } = 0m;
        public decimal? H_TotalAdd1StuffQty { get; set; } = 0m;

        public decimal? H_TotalAdd2OutRoll { get; set; } = 0m;
        public decimal? H_TotalAdd2OutQty { get; set; } = 0m;
        public decimal? H_TotalAdd2StuffRoll { get; set; } = 0m;
        public decimal? H_TotalAdd2StuffQty { get; set; } = 0m;
        public bool H_Color1 { get; set; } = false;
        public bool H_Color2 { get; set; } = false;
    }


}

