using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
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
using WPF.MDI;

namespace WizMes_SungShinNQ
{
    /// <summary>
    /// Win_prd_Warning_S.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_Warning_S : UserControl
    {
        public Win_prd_Warning_S()
        {
            InitializeComponent();
        }
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);

            chkDateSrh.IsChecked = true;
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        #region 상단 날짜

        private void lblDateSrh_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkDateSrh.IsChecked == true)
            {
                chkDateSrh.IsChecked = false;
            }
            else
            {
                chkDateSrh.IsChecked = true;
            }
        }

        private void chkDateSrh_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        private void chkDateSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

    
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = Lib.Instance.BringLastMonthContinue(dtpSDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        #endregion

        #region 버튼

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            FillGrid_Tool();
            FillGrid_Stock();
        }

        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[6];
            lst[0] = "툴 교체주기";
            lst[1] = "원부자재 재고 부족";
            lst[2] = "하위결합정보";
            lst[3] = dgdTool.Name;
            lst[4] = dgdStock.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdTool.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdTool);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdTool);

                    Name = dgdTool.Name;

                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdStock.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdStock);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdStock);

                    Name = dgdStock.Name;

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

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }
        #endregion

        private void FillGrid_Tool()
        {
            if (dgdTool.Items.Count > 0)
            {
                dgdTool.Items.Clear();
            }
            if (dgdToolSum.Items.Count > 0)
            {
                dgdToolSum.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkDate", chkDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkDateSrh.IsChecked == true && dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", chkDateSrh.IsChecked == true && dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Warning_sTool", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        var sum = new Warning();

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var warning = new Warning()
                            {
                                Num = i,
                                MCPartName = dr["MCPartName"].ToString(),
                                ChangeDate = DatePickerFormat(dr["ChangeDate"].ToString()),
                                UseQty = dr["UseQty"].ToString(),
                                PerProdQty = dr["PerProdQty"].ToString(),
                                SetProdQty = dr["SetProdQty"].ToString(),
                            };
                            sum.UseQty += warning.UseQty;
                            dgdTool.Items.Add(warning);
                        }

                        sum.Num = i;
                        
                        dgdToolSum.Items.Add(sum);

                    } else
                    {
                        MessageBox.Show("조회된 데이터가 없습니다");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private void FillGrid_Stock()
        {

            if (dgdStock.Items.Count > 0)
            {
                dgdStock.Items.Clear();
            }
            if (dgdStockSum.Items.Count > 0)
            {           
                dgdStockSum.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("nChkDate", chkDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sSDate", chkDateSrh.IsChecked == true && dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sEDate", chkDateSrh.IsChecked == true && dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nChkCustom", 0);
                sqlParameter.Add("sCustomID", "");

                sqlParameter.Add("nChkArticleID", 0);
                sqlParameter.Add("sArticleID", "");
                sqlParameter.Add("nChkOrder", 0);
                sqlParameter.Add("sOrder", "");
                sqlParameter.Add("ArticleGrpID", "");

                sqlParameter.Add("sFromLocID", "");
                sqlParameter.Add("sToLocID", "");
                sqlParameter.Add("nChkOutClss", 0);
                sqlParameter.Add("sOutClss", "");
                sqlParameter.Add("nChkInClss", 0);

                sqlParameter.Add("sInClss", "");
                sqlParameter.Add("nChkReqID", 0);
                sqlParameter.Add("sReqID", "");
                sqlParameter.Add("incNotApprovalYN", "N");
                sqlParameter.Add("incAutoInOutYN", "N");

                sqlParameter.Add("sArticleIDS", "");
                sqlParameter.Add("sMissSafelyStockQty", "");
                sqlParameter.Add("sProductYN", "");
                sqlParameter.Add("nMainItem", 0);
                sqlParameter.Add("nCustomItem", 0);

                sqlParameter.Add("nSupplyType", 0);
                sqlParameter.Add("sSupplyType", "");

                sqlParameter.Add("nBuyerArticleNo", 0);
                sqlParameter.Add("BuyerArticleNo", "");


                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Subul_sStockList_Mtr", sqlParameter, false);


                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        var sum = new Win_sbl_Stock_Q_View();

                        foreach (DataRow dr in drc)
                        {

                            if (dr["cls"].ToString().Equals("1"))
                            {
                                i++;
                                var stock = new Win_sbl_Stock_Q_View()
                                {
                                    NUM = i,
                                    Article = dr["Article"].ToString(),
                                    BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                    StockQty = Convert.ToInt32(dr["StockQty"]),             //수량 
                                    NeedstockQty = Convert.ToInt32(dr["NeedstockQty"]),     //적정재고
                                    OverQty = Convert.ToInt32(dr["OverQty"]),               //부족량
                                    StockRate = Convert.ToInt32(dr["StockRate"]),           //부족비율 
                                    UnitClssName = dr["UnitClssName"].ToString(),           //단위
                                };
                                sum.OverQty += stock.OverQty;
                                sum.StockQty += stock.StockQty;
                                sum.NeedstockQty += stock.NeedstockQty;
                                sum.NUM = i;
                               dgdStock.Items.Add(stock);
                            }
                        } 
                        if(sum.NeedstockQty != 0) sum.StockRate = sum.StockQty / sum.NeedstockQty * 100;

                        dgdStockSum.Items.Add(sum);
                        
                    }
                    else
                    {
                        MessageBox.Show("조회된 데이터가 없습니다");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

           

        }

        #region 기타 메서드 모음

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

        #endregion
    }

    public class Warning
    {
        public int Num { get; set; }
        public String MCPartName { get; set; } //툴명
        public String ChangeDate { get; set; } //최종 교체일
        public String UseQty { get; set; } //사용횟수
        public String SetProdQty { get; set; } //수명한계
        public String PerProdQty { get; set; } //퍼센트로 
    }



}
