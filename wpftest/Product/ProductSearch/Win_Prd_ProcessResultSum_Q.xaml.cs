
using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_SungShinNQ;
using WizMes_SungShinNQ.PopUp;
using WizMes_SungShinNQ.PopUP;
using WPF.MDI;


namespace WizMes_SungShinNQ
{
    /// <summary>
    /// Win_Prd_ProcessResultSum_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Prd_ProcessResultSum_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();

        // 인쇄 활용 용도 (프린트)
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;

        NoticeMessage msg = new NoticeMessage();

        public Win_Prd_ProcessResultSum_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            Lib.Instance.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            SetComboBox();

            cboProcess.SelectedIndex = 0;
            cboMachine.SelectedIndex = 0;

            rbnOrderID.IsChecked = true;
        }

        private void SetComboBox()
        {
            ObservableCollection<CodeView> cbWork = ComboBoxUtil.Instance.GetWorkProcess(0, "");

            this.cboProcess.ItemsSource = cbWork;
            this.cboProcess.DisplayMemberPath = "code_name";
            this.cboProcess.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcMachine = GetMachineByProcessID("");
            this.cboMachine.ItemsSource = ovcMachine;
            this.cboMachine.DisplayMemberPath = "code_name";
            this.cboMachine.SelectedValuePath = "code_id";

        }

        #region mt_Machine - 호기 세팅

        /// <summary>
        /// 호기ID 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetMachineByProcessID(string value)
        {
            //2021-10-25 공정 콤보박스에 전체가 선택되면 호기 공정 콤보박스 안되게 막기
            if (value.Equals(""))
            {
                cboMachine.IsEnabled = false;
            }
            else
            {
                cboMachine.IsEnabled = true;
            }

            ObservableCollection<CodeView> ovcMachine = new ObservableCollection<CodeView>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("sProcessID", value);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sMachineForComboBox", sqlParameter, false);

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

        #region 날짜 관련 이벤트

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

        #region 체크 등 이벤트

        //최종거래처
        private void lbInCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkInCustom.IsChecked == true)
            {
                chkInCustom.IsChecked = false;
            }
            else
            {
                chkInCustom.IsChecked = true;
            }
        }

        //최종거래처
        private void chkInCustom_Checked(object sender, RoutedEventArgs e)
        {
            txtInCustom.IsEnabled = true;
            btnPfInCustom.IsEnabled = true;
            txtInCustom.Focus();
        }

        //최종거래처
        private void chkInCustom_Unchecked(object sender, RoutedEventArgs e)
        {
            txtInCustom.IsEnabled = false;
            btnPfInCustom.IsEnabled = false;
        }

        //최종거래처
        private void txtInCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtInCustom, 72, "");
            }
        }

        //최종거래처
        private void btnPfInCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtInCustom, 72, "");
        }

        //작업자
        private void lblPerson_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkPerson.IsChecked == true) { chkPerson.IsChecked = false; }
            else { chkPerson.IsChecked = true; }
        }

        //작업자
        private void chkPerson_Checked(object sender, RoutedEventArgs e)
        {
            txtPerson.IsEnabled = true;
            txtPerson.Focus();
        }

        //작업자
        private void chkPerson_Unchecked(object sender, RoutedEventArgs e)
        {
            txtPerson.IsEnabled = false;
        }

        //관리번호
        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {
            if (rbnOrderNo.IsChecked == true)
            {
                tbkOrder.Text = " Order No.";
            }
        }

        //관리번호
        private void rbnOrderID_Click(object sender, RoutedEventArgs e)
        {
            if (rbnOrderID.IsChecked == true)
            {
                tbkOrder.Text = " 관리번호";
            }
        }

        //관리번호
        private void lblOrder_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrder.IsChecked == true) { chkOrder.IsChecked = false; }
            else { chkOrder.IsChecked = true; }
        }

        //관리번호
        private void chkOrder_Checked(object sender, RoutedEventArgs e)
        {
            txtOrder.IsEnabled = true;
            btnPfOrder.IsEnabled = true;
            txtOrder.Focus();
        }

        //관리번호
        private void chkOrder_Unchecked(object sender, RoutedEventArgs e)
        {
            txtOrder.IsEnabled = false;
            btnPfOrder.IsEnabled = false;
        }

        //관리번호
        private void txtOrder_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtOrder, (int)Defind_CodeFind.DCF_ORDER, "");
            }
        }

        //관리번호
        private void btnPfOrder_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtOrder, (int)Defind_CodeFind.DCF_ORDER, "");
        }

        //거래처
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustom.IsChecked == true) { chkCustom.IsChecked = false; }
            else { chkCustom.IsChecked = true; }
        }

        //거래처
        private void chkCustom_Checked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = true;
            btnPfCustom.IsEnabled = true;
            txtCustom.Focus();
        }

        private void chkCustom_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = false;
            btnPfCustom.IsEnabled = false;
        }

        //거래처
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //pf.ReturnCode(txtCustom, 0, "");
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //거래처
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            //pf.ReturnCode(txtCustom, 0, "");
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //품명
        private void lblArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true) { chkArticle.IsChecked = false; }
            else { chkArticle.IsChecked = true; }
        }

        //품명
        private void chkArticle_Checked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = true;
            btnPfArticle.IsEnabled = true;
            txtArticle.Focus();
        }

        private void chkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = false;
            btnPfArticle.IsEnabled = false;
        }

        //품명
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 77, "");
            }
        }

        //품명
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 77, "");
        }

        // 품명대분류
        private void lblCategory_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCategory.IsChecked == true) { chkCategory.IsChecked = false; }
            else { chkCategory.IsChecked = true; }
        }

        // 품명대분류
        private void chkCategory_Checked(object sender, RoutedEventArgs e)
        {
            cboCategory.IsEnabled = true;
        }

        // 품명대분류
        private void chkCategory_Unchecked(object sender, RoutedEventArgs e)
        {
            cboCategory.IsEnabled = false;
        }

        #endregion

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {

            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                //로직
                using (Loading lw = new Loading(re_Search))
                {
                    lw.ShowDialog();
                }

            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSearch.IsEnabled = true;
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        #region 엑셀 버튼 이벤트
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            TabItem nowTab = tabconGrid.SelectedItem as TabItem;

            if (nowTab.Header.ToString().Equals("공정별 호기별 집계"))
            {
                string[] lst = new string[2];
                lst[0] = "공정별 호기별 집계";
                lst[1] = dgdByProcess.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdByProcess.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdByProcess);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdByProcess);

                        Name = dgdByProcess.Name;

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
            else if (nowTab.Header.ToString().Equals("품번별 집계"))
            {
                string[] lst = new string[2];
                lst[0] = "품번별 집계";
                lst[1] = dgdByArticle.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdByArticle.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdByArticle);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdByArticle);

                        Name = dgdByArticle.Name;

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
            else if (nowTab.Header.ToString().Equals("작업자별 집계"))
            {
                string[] lst = new string[2];
                lst[0] = "작업자별 집계";
                lst[1] = dgdByWorker.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdByWorker.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdByWorker);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdByWorker);

                        Name = dgdByWorker.Name;

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
            else if (nowTab.Header.ToString().Equals("일별 집계"))
            {
                string[] lst = new string[2];
                lst[0] = "일별 집계";
                lst[1] = DataGridThisMonth.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(DataGridThisMonth.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(DataGridThisMonth);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(DataGridThisMonth);

                        Name = DataGridThisMonth.Name;

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
        }
        #endregion // 엑셀 버튼 이벤트

        private bool CheckData()
        {
            bool flag = true;

            if (cboProcess.SelectedValue == null)
            {
                MessageBox.Show("공정이 선택되지 않았습니다. 선택해주세요");
                flag = false;
                return flag;
            }

            if (cboMachine.SelectedValue == null)
            {
                MessageBox.Show("호기가 선택되지 않았습니다. 선택해주세요");
                flag = false;
                return flag;
            }

            return flag;
        }

        private void re_Search()
        {
            if (CheckData())
            {
                TabItem nowTab = tabconGrid.SelectedItem as TabItem;

                if (nowTab.Name == "tabProcess")
                {
                    FillGridProcessMachine();

                    if (dgdByProcess.Items.Count <= 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        return;
                    }
                }
                else if (nowTab.Name == "tabArticle")
                {
                    FillGridArticle();

                    if (dgdByArticle.Items.Count <= 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        return;
                    }
                }
                else if (nowTab.Name == "tabWorker")
                {
                    FillGridWorker();

                    if (dgdByWorker.Items.Count <= 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        return;
                    }
                }
                else if (nowTab.Name == "tabThisMonth")
                {
                    FillGrid_ThisMonth();

                    if (DataGridThisMonth.Items.Count <= 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        return;
                    }
                }
            }
        }

        #region 주요 메서드 - 공정별 호기별 집계 조회 FillGridProcessMachine
        private void FillGridProcessMachine()
        {
            dgdByProcessTotal.Items.Clear();

            if (dgdByProcess.Items.Count > 0)
            {
                dgdByProcess.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("sFromDate", dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sProcessIDS", cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "");
                sqlParameter.Add("sMachineIDS", cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "");
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");

                sqlParameter.Add("CustomID", chkCustom.IsChecked == true && txtCustom.Tag != null ? txtCustom.Tag.ToString() : "");
                sqlParameter.Add("nOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? 1 : 2) : 0);
                sqlParameter.Add("sOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? txtOrder.Tag.ToString() : txtOrder.Text) : "");
                sqlParameter.Add("sWorker", chkPerson.IsChecked == true ? txtPerson.Text : "");
                sqlParameter.Add("nBuySaleMainYN", chkMainItem.IsChecked == true ? 1 : 0);

                sqlParameter.Add("BuyerArticleNoID", CheckBoxBuyerArticleNoSearch.IsChecked == true && TextBoxBuyerArticleNoSearch.Tag != null ? TextBoxBuyerArticleNoSearch.Tag.ToString() : "");
                sqlParameter.Add("ChkInCustom", chkInCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("InCustomID", chkInCustom.IsChecked == true ? (txtInCustom.Tag != null ? txtInCustom.Tag.ToString() : "") : "");
  

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sWKResultByProcessMachine", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int cnt = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var cls = dr["cls"].ToString();

                            if (cls == "1")
                            {
                                cnt++;   
                            }

                            var WinM = new Win_Prd_ProcessResultSum_Q_ByProcessMachine()
                            {
                                cls = dr["cls"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                BuyerModel = dr["BuyerModel"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                QtyPerBox = stringFormatN0(dr["QtyPerBox"]),
                                WorkQty = stringFormatN0(dr["WorkQty"]),
                                UnitPrice = stringFormatN0(dr["UnitPrice"]),
                                Amount = stringFormatN0(dr["Amount"]),
                                WorkTime = stringFormatN1(dr["WorkTime"]),
                                Num = i,
                                Cnt = ""   // 기본값
                            };

                            
                            if (WinM.cls.Equals("2")) // 호기계
                            {
                                WinM.BuyerModel = "호기계";

                                WinM.QtyPerBox = "";
                                dgdByProcess.Items.Add(WinM);
                            }
                            else if (WinM.cls.Equals("3")) // 공정계
                            {
                                WinM.MachineNo = "공정계";

                                WinM.QtyPerBox = "";
                                dgdByProcess.Items.Add(WinM);
                            }
                            else if (WinM.cls.Equals("9")) // 총계
                            {
                                WinM.Process = "총계";

                                WinM.QtyPerBox = "";

                                dgdByProcessTotal.Items.Add(WinM);
                                WinM.Cnt = cnt.ToString();
                            }
                            else
                            {
                                dgdByProcess.Items.Add(WinM);
                            }

       


                        }
                        setGraphProcessMachine(dgdByProcess);
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

        #endregion // 공정별 호기별 집계

        #region 주요 메서드 - 품명별 집계 조회 FillGridArticle
        private void FillGridArticle()  //2021-06-10 GLS는 품번별로 변경
        {
            dgdByArticleTotal.Items.Clear();

            if (dgdByArticle.Items.Count > 0)
            {
                dgdByArticle.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("sFromDate", dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sProcessIDS", cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "");
                sqlParameter.Add("sMachineIDS", cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "");
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true && txtCustom.Tag != null ? txtCustom.Tag.ToString() : "");
                sqlParameter.Add("nOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? 1 : 2) : 0);
                sqlParameter.Add("sOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? txtOrder.Tag.ToString() : txtOrder.Text) : "");
                sqlParameter.Add("sWorker", chkPerson.IsChecked == true ? txtPerson.Text : "");
                sqlParameter.Add("nBuySaleMainYN", chkMainItem.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNoID", CheckBoxBuyerArticleNoSearch.IsChecked == true && TextBoxBuyerArticleNoSearch.Tag != null ? TextBoxBuyerArticleNoSearch.Tag.ToString() : "");
                sqlParameter.Add("ChkInCustom", chkInCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("InCustomID", chkInCustom.IsChecked == true ? (txtInCustom.Tag != null ? txtInCustom.Tag.ToString() : "") : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sWKResultByArticle", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;
                        int cnt = 0;


                        foreach (DataRow dr in drc)
                        {

                            var cls = dr["cls"].ToString().Trim();
                            if (cls == "1")
                            {
                                cnt++;
                            }

                            i++;
                            var WinA = new Win_Prd_ProcessResultSum_Q_ByArticle()
                            {
                                Num = i,
                                cls = dr["cls"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                WorkQty = stringFormatN0(dr["WorkQty"]),
                                ProdQtyPerBox = stringFormatN0(dr["ProdQtyPerBox"]),
                                Cnt = ""

                            };

                            if (WinA.cls.Trim().Equals("3")) // 품명계
                            {
                                WinA.BuyerArticleNo = "품명계";
                                dgdByArticle.Items.Add(WinA);
                            }
                            else if (WinA.cls.Trim().Equals("9")) // 총계
                            {
                                WinA.Process = "총계";
                                WinA.BuyerArticleNo = "";
                                dgdByArticleTotal.Items.Add(WinA);
                                WinA.Cnt = cnt.ToString();
                            }
                            else
                            {
                                dgdByArticle.Items.Add(WinA);
                            }


                        }

                        setGraphArticle(dgdByArticle);
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
        #endregion // 품명별 집계

        #region 주요 메서드 - 작업자별 집계 조회 FillGridWorker

        private void FillGridWorker()
        {
            dgdByWorkerTotal.Items.Clear();

            if (dgdByWorker.Items.Count > 0)
            {
                dgdByWorker.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("sFromDate", dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sProcessIDS", cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "");
                sqlParameter.Add("sMachineIDS", cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "");
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");

                sqlParameter.Add("CustomID", chkCustom.IsChecked == true && txtCustom.Tag != null ? txtCustom.Tag.ToString() : "");
                sqlParameter.Add("nOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? 1 : 2) : 0);
                sqlParameter.Add("sOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? txtOrder.Tag.ToString() : txtOrder.Text) : "");
                sqlParameter.Add("sWorker", chkPerson.IsChecked == true ? txtPerson.Text : "");
                sqlParameter.Add("nBuySaleMainYN", chkMainItem.IsChecked == true ? 1 : 0);

                sqlParameter.Add("BuyerArticleNoID", CheckBoxBuyerArticleNoSearch.IsChecked == true && TextBoxBuyerArticleNoSearch.Tag != null ? TextBoxBuyerArticleNoSearch.Tag.ToString() : "");
                sqlParameter.Add("ChkInCustom", chkInCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("InCustomID", chkInCustom.IsChecked == true ? (txtInCustom.Tag != null ? txtInCustom.Tag.ToString() : "") : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sWKResultByWorker", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;
                        int cnt = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var cls = dr["cls"].ToString().Trim();
                            if (cls == "1")
                            {
                                cnt++;
                            }
                            var WinW = new Win_Prd_ProcessResultSum_Q_ByWorker()
                            {
                                Num = i,
                                cls = dr["cls"].ToString().Trim(),

                                WorkPersonID = dr["WorkPersonID"].ToString(),
                                Name = dr["Name"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                MachineID = dr["MachineID"].ToString(),

                                Machine = dr["Machine"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                Model = dr["Model"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),

                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                WorkQty = stringFormatN0(dr["WorkQty"]),

                                ProdQtyPerBox = stringFormatN0(dr["ProdQtyPerBox"]),

                            };

                            if (WinW.cls.Trim().Equals("3")) // 작업자계
                            {
                                WinW.Process = "작업자계";
                                dgdByWorker.Items.Add(WinW);
                            }
                            else if (WinW.cls.Trim().Equals("9")) // 총계
                            {
                                WinW.Process = "총계";
                                WinW.Name = "";

                                dgdByWorkerTotal.Items.Add(WinW);
                            }
                            else
                            {
                                dgdByWorker.Items.Add(WinW);
                            }

                        }

                        setGraphWorker(dgdByWorker);

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

        #endregion // 주요 메서드 - 작업자별 집계 조회 FillGridWorker

        #region 조회 일자별 집계
        private void FillGrid_ThisMonth()
        {
            if (DataGridThisMonth.Items.Count > 0)
            {
                DataGridThisMonth.Items.Clear();
            }

            //int chkDate = 0;
            //string sFromDate = string.Empty;
            //string sToDate = string.Empty;
            //int chkProcessID = 0;
            //string sProcessID = string.Empty;
            //int chkMachineID = 0;
            //string sMachineID = string.Empty;
            //int chkWorker = 0;
            //string sWorker = string.Empty;
            //int chkOrderID = 0;
            //string sOrderID = string.Empty;
            //int chkCustomID = 0;
            //string sCustomID = string.Empty;
            //int chkArticleID = 0;
            //string sArticleID = string.Empty;
            //int chkBuySaleMainYN = 0;
            //int chkBuyerArticleNo = 0;
            //string buyerArticleNo = string.Empty;


            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkDate", chkDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromDate", chkDateSrh.IsChecked == true ? (dtpSDate.SelectedDate == null ? "" : dtpSDate.SelectedDate.Value.ToString("yyyyMMdd")) : "");
                sqlParameter.Add("sToDate", chkDateSrh.IsChecked == true ? (dtpEDate.SelectedDate == null ? "" : dtpEDate.SelectedDate.Value.ToString("yyyyMMdd")) : "");
                sqlParameter.Add("ChkProcessID", CheckBoxProcessSearch.IsChecked == true && cboProcess.SelectedValue.ToString() != "" ? 1 : 0);
                sqlParameter.Add("sProcessID", CheckBoxProcessSearch.IsChecked == true ? (cboProcess.SelectedValue == null ? "" : cboProcess.SelectedValue.ToString()) : "");
                sqlParameter.Add("ChkMachineID", CheckBoxMachineSearch.IsChecked == true && cboMachine.SelectedValue.ToString() != "" ? 1 : 0);
                sqlParameter.Add("sMachineID", CheckBoxMachineSearch.IsChecked == true ? (cboMachine.SelectedValue == null ? "" : cboMachine.SelectedValue.ToString()) : "");
                sqlParameter.Add("ChkWorker", chkPerson.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sWorker", chkPerson.IsChecked == true ? (txtPerson.Text == string.Empty ? "" : txtPerson.Text) : "");
                sqlParameter.Add("ChkOrderID", chkOrder.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sOrderID", chkOrder.IsChecked == true ? (txtOrder.Tag == null ? "" : txtOrder.Tag.ToString()) : "");
                sqlParameter.Add("ChkCustomID", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sCustomID", chkCustom.IsChecked == true ? (txtCustom.Tag == null ? "" : txtCustom.Tag.ToString()) : "");
                sqlParameter.Add("ChkArticleID", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");
                sqlParameter.Add("ChkBuySaleMainYN", chkMainItem.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ChkBuyerArticleNo", CheckBoxBuyerArticleNoSearch.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", (CheckBoxBuyerArticleNoSearch.IsChecked == true && TextBoxBuyerArticleNoSearch.Tag != null) ? TextBoxBuyerArticleNoSearch.Tag.ToString() : "");
                sqlParameter.Add("ChkInCustom", chkInCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("InCustomID", chkInCustom.IsChecked == true ? (txtInCustom.Tag != null ? txtInCustom.Tag.ToString() : "") : "");
         

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sWKResult_Article_ThisMonth", sqlParameter, false);

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
                            var WPPQCT = new Win_Prd_ProcessResultSum_Q_CodeView_ThisMonth()
                            {
                                Num = i,

                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),

                                SDay01 = Convert.ToDouble(dr["SDay01"]),
                                SDay02 = Convert.ToDouble(dr["SDay02"]),
                                SDay03 = Convert.ToDouble(dr["SDay03"]),
                                SDay04 = Convert.ToDouble(dr["SDay04"]),
                                SDay05 = Convert.ToDouble(dr["SDay05"]),
                                SDay06 = Convert.ToDouble(dr["SDay06"]),
                                SDay07 = Convert.ToDouble(dr["SDay07"]),
                                SDay08 = Convert.ToDouble(dr["SDay08"]),
                                SDay09 = Convert.ToDouble(dr["SDay09"]),
                                SDay10 = Convert.ToDouble(dr["SDay10"]),
                                SDay11 = Convert.ToDouble(dr["SDay11"]),
                                SDay12 = Convert.ToDouble(dr["SDay12"]),
                                SDay13 = Convert.ToDouble(dr["SDay13"]),
                                SDay14 = Convert.ToDouble(dr["SDay14"]),
                                SDay15 = Convert.ToDouble(dr["SDay15"]),
                                SDay16 = Convert.ToDouble(dr["SDay16"]),
                                SDay17 = Convert.ToDouble(dr["SDay17"]),
                                SDay18 = Convert.ToDouble(dr["SDay18"]),
                                SDay19 = Convert.ToDouble(dr["SDay19"]),
                                SDay20 = Convert.ToDouble(dr["SDay20"]),
                                SDay21 = Convert.ToDouble(dr["SDay21"]),
                                SDay22 = Convert.ToDouble(dr["SDay22"]),
                                SDay23 = Convert.ToDouble(dr["SDay23"]),
                                SDay24 = Convert.ToDouble(dr["SDay24"]),
                                SDay25 = Convert.ToDouble(dr["SDay25"]),
                                SDay26 = Convert.ToDouble(dr["SDay26"]),
                                SDay27 = Convert.ToDouble(dr["SDay27"]),
                                SDay28 = Convert.ToDouble(dr["SDay28"]),
                                SDay29 = Convert.ToDouble(dr["SDay29"]),
                                SDay30 = Convert.ToDouble(dr["SDay30"]),
                                SDay31 = Convert.ToDouble(dr["SDay31"]),
                            };

                            double sum = WPPQCT.SDay01 + WPPQCT.SDay02 + WPPQCT.SDay03 + WPPQCT.SDay04 + WPPQCT.SDay05
                                + WPPQCT.SDay06 + WPPQCT.SDay07 + WPPQCT.SDay08 + WPPQCT.SDay09 + WPPQCT.SDay10
                                + WPPQCT.SDay11 + WPPQCT.SDay12 + WPPQCT.SDay13 + WPPQCT.SDay14 + WPPQCT.SDay15
                                + WPPQCT.SDay16 + WPPQCT.SDay17 + WPPQCT.SDay18 + WPPQCT.SDay19 + WPPQCT.SDay20
                                + WPPQCT.SDay21 + WPPQCT.SDay22 + WPPQCT.SDay23 + WPPQCT.SDay24 + WPPQCT.SDay25
                                + WPPQCT.SDay26 + WPPQCT.SDay27 + WPPQCT.SDay28 + WPPQCT.SDay29 + WPPQCT.SDay30
                                + WPPQCT.SDay31;

                            WPPQCT.TotalQty = sum;


                            //int sum = Convert.ToInt32(WPPQCT.SDay01) + Convert.ToInt32(WPPQCT.SDay02) + Convert.ToInt32(WPPQCT.SDay03)
                            //        + Convert.ToInt32(WPPQCT.SDay04) + Convert.ToInt32(WPPQCT.SDay05) + Convert.ToInt32(WPPQCT.SDay06)
                            //        + Convert.ToInt32(WPPQCT.SDay07) + Convert.ToInt32(WPPQCT.SDay08) + Convert.ToInt32(WPPQCT.SDay09)
                            //        + Convert.ToInt32(WPPQCT.SDay10) + Convert.ToInt32(WPPQCT.SDay11) + Convert.ToInt32(WPPQCT.SDay12)
                            //        + Convert.ToInt32(WPPQCT.SDay13) + Convert.ToInt32(WPPQCT.SDay14) + Convert.ToInt32(WPPQCT.SDay15)
                            //        + Convert.ToInt32(WPPQCT.SDay16) + Convert.ToInt32(WPPQCT.SDay17) + Convert.ToInt32(WPPQCT.SDay18)
                            //        + Convert.ToInt32(WPPQCT.SDay19) + Convert.ToInt32(WPPQCT.SDay20) + Convert.ToInt32(WPPQCT.SDay21)
                            //        + Convert.ToInt32(WPPQCT.SDay22) + Convert.ToInt32(WPPQCT.SDay23) + Convert.ToInt32(WPPQCT.SDay24)
                            //        + Convert.ToInt32(WPPQCT.SDay25) + Convert.ToInt32(WPPQCT.SDay26) + Convert.ToInt32(WPPQCT.SDay27)
                            //        + Convert.ToInt32(WPPQCT.SDay28) + Convert.ToInt32(WPPQCT.SDay29) + Convert.ToInt32(WPPQCT.SDay30)
                            //        + Convert.ToInt32(WPPQCT.SDay31);
                            //WPPQCT.TotalQty = lib.returnNumStringZero(Convert.ToString(sum));


                            //WPPQCT.SDay01 = lib.returnNumStringZero(WPPQCT.SDay01);
                            //WPPQCT.SDay02 = lib.returnNumStringZero(WPPQCT.SDay02);
                            //WPPQCT.SDay03 = lib.returnNumStringZero(WPPQCT.SDay03);
                            //WPPQCT.SDay04 = lib.returnNumStringZero(WPPQCT.SDay04);
                            //WPPQCT.SDay05 = lib.returnNumStringZero(WPPQCT.SDay05);
                            //WPPQCT.SDay06 = lib.returnNumStringZero(WPPQCT.SDay06);
                            //WPPQCT.SDay07 = lib.returnNumStringZero(WPPQCT.SDay07);
                            //WPPQCT.SDay08 = lib.returnNumStringZero(WPPQCT.SDay08);
                            //WPPQCT.SDay09 = lib.returnNumStringZero(WPPQCT.SDay09);
                            //WPPQCT.SDay10 = lib.returnNumStringZero(WPPQCT.SDay10);
                            //WPPQCT.SDay11 = lib.returnNumStringZero(WPPQCT.SDay11);
                            //WPPQCT.SDay12 = lib.returnNumStringZero(WPPQCT.SDay12);
                            //WPPQCT.SDay13 = lib.returnNumStringZero(WPPQCT.SDay13);
                            //WPPQCT.SDay14 = lib.returnNumStringZero(WPPQCT.SDay14);
                            //WPPQCT.SDay15 = lib.returnNumStringZero(WPPQCT.SDay15);
                            //WPPQCT.SDay16 = lib.returnNumStringZero(WPPQCT.SDay16);
                            //WPPQCT.SDay17 = lib.returnNumStringZero(WPPQCT.SDay17);
                            //WPPQCT.SDay18 = lib.returnNumStringZero(WPPQCT.SDay18);
                            //WPPQCT.SDay19 = lib.returnNumStringZero(WPPQCT.SDay19);
                            //WPPQCT.SDay20 = lib.returnNumStringZero(WPPQCT.SDay20);
                            //WPPQCT.SDay21 = lib.returnNumStringZero(WPPQCT.SDay21);
                            //WPPQCT.SDay22 = lib.returnNumStringZero(WPPQCT.SDay22);
                            //WPPQCT.SDay23 = lib.returnNumStringZero(WPPQCT.SDay23);
                            //WPPQCT.SDay24 = lib.returnNumStringZero(WPPQCT.SDay24);
                            //WPPQCT.SDay25 = lib.returnNumStringZero(WPPQCT.SDay25);
                            //WPPQCT.SDay26 = lib.returnNumStringZero(WPPQCT.SDay26);
                            //WPPQCT.SDay27 = lib.returnNumStringZero(WPPQCT.SDay27);
                            //WPPQCT.SDay28 = lib.returnNumStringZero(WPPQCT.SDay28);
                            //WPPQCT.SDay29 = lib.returnNumStringZero(WPPQCT.SDay29);
                            //WPPQCT.SDay30 = lib.returnNumStringZero(WPPQCT.SDay30);
                            //WPPQCT.SDay31 = lib.returnNumStringZero(WPPQCT.SDay31);

                            DataGridThisMonth.Items.Add(WPPQCT);
                        }

                        setGraphThisMonth(DataGridThisMonth);
                    }

                    DataGridThisMonthTotal.Items.Clear();

                    var tot = new Win_Prd_ProcessResultSum_Q_CodeView_ThisMonth()
                    {
                        Cnt = DataGridThisMonth.Items.Count.ToString(),
                        TotalQty = DataGridThisMonth.Items
                            .OfType<Win_Prd_ProcessResultSum_Q_CodeView_ThisMonth>()
                            .Sum(x => x.TotalQty)
                    };

                    DataGridThisMonthTotal.Items.Add(tot);


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
            if (str.Length > 3 && str.Length < 7)
            {
                string hour = str.Substring(0, 2);
                string min = str.Substring(2, 2);

                result = hour + ":" + min;
            }

            return result;
        }
        #endregion

        #region 텍스트박스 공용 키다운 이벤트
        private void txtBox_KeyDown_Search(object sender, KeyEventArgs e)
        {
            using (Loading lw = new Loading(re_Search))
            {
                lw.ShowDialog();
            }
        }
        #endregion

        private void LabelBuyerArticleNoSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (CheckBoxBuyerArticleNoSearch.IsChecked == true)
            {
                CheckBoxBuyerArticleNoSearch.IsChecked = false;
            }
            else
            {
                CheckBoxBuyerArticleNoSearch.IsChecked = true;
            }
        }

        private void CheckBoxBuyerArticleNoSearch_Checked(object sender, RoutedEventArgs e)
        {
            TextBoxBuyerArticleNoSearch.IsEnabled = true;
            ButtonBuyerArticleNoSearch.IsEnabled = true;
            TextBoxBuyerArticleNoSearch.Focus();
        }

        private void CheckBoxBuyerArticleNoSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            TextBoxBuyerArticleNoSearch.IsEnabled = false;
            ButtonBuyerArticleNoSearch.IsEnabled = false;
        }

        private void TextBoxBuyerArticleNoSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, "");
            }
        }

        private void ButtonBuyerArticleNoSearch_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, "");
        }

        private void tabconGrid_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            TabItem nowTab = tabconGrid.SelectedItem as TabItem;
            btnPrint.Visibility = nowTab.Name.Equals("tabWorker") == true ? Visibility.Visible : Visibility.Hidden;
        }

        public class ChartDTO
        {
            public SeriesCollection SeriesCollection { get; set; }
            public string[] Labels { get; set; }
            public Func<ChartPoint, string> Formatter { get; set; }
        }


        // 그래프들
        private void setGraphProcessMachine(DataGrid dataGrid)
        {
            try
            {
                if (lvcChartProcess.Series != null)
                {
                    lvcChartProcess.Series.Clear();
                }

                var list = new List<Win_Prd_ProcessResultSum_Q_ByProcessMachine>();

                for (int i = 0; i < dataGrid.Items.Count; i++)
                {
                    if (dataGrid.Items[i] is Win_Prd_ProcessResultSum_Q_ByProcessMachine row)
                    {
                        if (row.cls == "1" || string.IsNullOrWhiteSpace(row.cls))
                        {
                            list.Add(row);
                        }
                    }
                }

                var chart = new ChartDTO();
                chart.SeriesCollection = new SeriesCollection();
                chart.Labels = new string[list.Count];

                var qty = new ChartValues<double>();

                for (int i = 0; i < list.Count; i++)
                {
                    var row = list[i];

                    var p = (row.Process ?? "").Trim();
                    var m = (row.MachineNo ?? "").Trim();

                    chart.Labels[i] = string.IsNullOrEmpty(m) ? p : $"{p}-{m}";

                    if (!string.IsNullOrWhiteSpace(row.WorkQty))
                    {
                        double v;
                        if (double.TryParse(row.WorkQty.Replace(",", "").Trim(), out v))
                            qty.Add(v);
                        else
                            qty.Add(0);
                    }
                    else
                    {
                        qty.Add(0);
                    }
                }

                chart.Formatter = value => value.Y.ToString("N0") + "(개)";

                chart.SeriesCollection.Add(new LineSeries
                {
                    Values = qty,
                    DataLabels = true,
                    Title = "생산량",
                    LabelPoint = chart.Formatter
                });

                lvcChartProcess.DataContext = chart;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // 그래프들
        private void setGraphArticle(DataGrid dataGrid)
        {
            try
            {
                if (lvcChartArticle.Series != null)
                {
                    lvcChartArticle.Series.Clear();
                }

                var list = new List<Win_Prd_ProcessResultSum_Q_ByArticle>();

                for (int i = 0; i < dataGrid.Items.Count; i++)
                {
                    if (dataGrid.Items[i] is Win_Prd_ProcessResultSum_Q_ByArticle row)
                    {
                        if (row.cls == "1" || string.IsNullOrWhiteSpace(row.cls))
                        {
                            list.Add(row);
                        }
                    }
                }

                var chart = new ChartDTO();
                chart.SeriesCollection = new SeriesCollection();
                chart.Labels = new string[list.Count];

                var qty = new ChartValues<double>();

                for (int i = 0; i < list.Count; i++)
                {
                    var row = list[i];

                    chart.Labels[i] = (row.Article ?? "").Trim();

                    double v = 0;
                    if (!string.IsNullOrWhiteSpace(row.WorkQty))
                        double.TryParse(row.WorkQty.Replace(",", "").Trim(), out v);

                    qty.Add(v);
                }

                chart.Formatter = value => value.Y.ToString("N0") + "(개)";

                chart.SeriesCollection.Add(new LineSeries
                {
                    Values = qty,
                    DataLabels = true,
                    Title = "생산량",
                    LabelPoint = chart.Formatter
                });

                lvcChartArticle.DataContext = chart;

                if (list.Count == 1)
                {
                    lvcChartArticle.AxisX[0].MinValue = -0.5;
                    lvcChartArticle.AxisX[0].MaxValue = 0.5;
                }
                else
                {
                    lvcChartArticle.AxisX[0].MinValue = double.NaN;
                    lvcChartArticle.AxisX[0].MaxValue = double.NaN;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void setGraphWorker(DataGrid dataGrid)
        {
            try
            {
                if (lvcChartWorker.Series != null)
                {
                    lvcChartWorker.Series.Clear();
                }

                var list = new List<Win_Prd_ProcessResultSum_Q_ByWorker>();

                for (int i = 0; i < dataGrid.Items.Count; i++)
                {
                    if (dataGrid.Items[i] is Win_Prd_ProcessResultSum_Q_ByWorker row)
                    {
                        if (row.cls == "1" || string.IsNullOrWhiteSpace(row.cls))
                        {
                            list.Add(row);
                        }
                    }
                }

                var grp = list
                    .GroupBy(x => (x.Name ?? "").Trim())
                    .Select(g => new
                    {
                        Name = g.Key,
                        Qty = g.Sum(x =>
                        {
                            double v = 0;
                            if (!string.IsNullOrWhiteSpace(x.WorkQty))
                                double.TryParse(x.WorkQty.Replace(",", "").Trim(), out v);
                            return v;
                        })
                    })
                    .OrderByDescending(x => x.Qty)
                    .ToList();

                var chart = new ChartDTO();
                chart.SeriesCollection = new SeriesCollection();
                chart.Labels = new string[grp.Count];

                var qty = new ChartValues<double>();

                for (int i = 0; i < grp.Count; i++)
                {
                    chart.Labels[i] = grp[i].Name;
                    qty.Add(grp[i].Qty);
                }

                chart.Formatter = value => value.Y.ToString("N0") + "(개)";

                chart.SeriesCollection.Add(new LineSeries
                {
                    Values = qty,
                    DataLabels = true,
                    Title = "생산량",
                    LabelPoint = chart.Formatter
                });

                lvcChartWorker.DataContext = chart;

                if (grp.Count == 1)
                {
                    lvcChartWorker.AxisX[0].MinValue = -0.5;
                    lvcChartWorker.AxisX[0].MaxValue = 0.5;
                }
                else
                {
                    lvcChartWorker.AxisX[0].MinValue = double.NaN;
                    lvcChartWorker.AxisX[0].MaxValue = double.NaN;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void setGraphThisMonth(DataGrid dataGrid)
        {
            try
            {
                if (lvcChartThisMonth.Series != null)
                {
                    lvcChartThisMonth.Series.Clear();
                }

                double[] sums = new double[31];

                for (int i = 0; i < dataGrid.Items.Count; i++)
                {
                    if (dataGrid.Items[i] is Win_Prd_ProcessResultSum_Q_CodeView_ThisMonth row)
                    {
                        sums[0] += row.SDay01; sums[1] += row.SDay02; sums[2] += row.SDay03; sums[3] += row.SDay04; sums[4] += row.SDay05;
                        sums[5] += row.SDay06; sums[6] += row.SDay07; sums[7] += row.SDay08; sums[8] += row.SDay09; sums[9] += row.SDay10;
                        sums[10] += row.SDay11; sums[11] += row.SDay12; sums[12] += row.SDay13; sums[13] += row.SDay14; sums[14] += row.SDay15;
                        sums[15] += row.SDay16; sums[16] += row.SDay17; sums[17] += row.SDay18; sums[18] += row.SDay19; sums[19] += row.SDay20;
                        sums[20] += row.SDay21; sums[21] += row.SDay22; sums[22] += row.SDay23; sums[23] += row.SDay24; sums[24] += row.SDay25;
                        sums[25] += row.SDay26; sums[26] += row.SDay27; sums[27] += row.SDay28; sums[28] += row.SDay29; sums[29] += row.SDay30;
                        sums[30] += row.SDay31;
                    }
                }

                var chart = new ChartDTO();
                chart.SeriesCollection = new SeriesCollection();
                chart.Labels = new string[31];

                var qty = new ChartValues<double>();

                for (int d = 0; d < 31; d++)
                {
                    chart.Labels[d] = (d + 1).ToString();
                    qty.Add(sums[d]);
                }

                chart.Formatter = value => value.Y.ToString("N0");

                chart.SeriesCollection.Add(new LineSeries
                {
                    Values = qty,
                    DataLabels = true,
                    Title = "생산량",
                    LabelPoint = chart.Formatter
                });

                lvcChartThisMonth.DataContext = chart;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }




        #region 개인작업일보 인쇄
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ContextMenu menu = btnPrint.ContextMenu;
                menu.StaysOpen = true;
                menu.IsOpen = true;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnPrint_Click : " + ee.ToString());
            }
        }

        // 미리보기
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e) { menuPrint_Click(true); }
        // 바로인쇄
        private void menuRighPrint_Click(object sender, RoutedEventArgs e) { menuPrint_Click(false); }

        private void menuPrint_Click(bool Ahead)
        {
            try
            {
                TabItem nowTab = tabconGrid.SelectedItem as TabItem;
                if (nowTab.Name != "tabWorker")
                {
                    MessageBox.Show("작업자별 집계탭을 선택해주세요");
                    return;
                }

                if (chkCategory.IsChecked == false)
                {
                    MessageBox.Show("품명대분류를 선택하고 다시 검색해주세요");
                    return;
                }

                DateTime startTime = dtpSDate.SelectedDate.Value;
                DateTime EndTime = dtpEDate.SelectedDate.Value;

                if (startTime.CompareTo(EndTime) != 0)
                {
                    MessageBox.Show("시작일과 종료일이 같은 날짜만 가능합니다");
                    return;
                }

                if (dgdByWorker.Items.Count == 0)
                {
                    MessageBox.Show("인쇄할 내용이 없습니다");
                    return;
                }

                msg.Show();
                msg.Topmost = true;
                msg.Refresh();

                PrintWork(Ahead);
                msg.Visibility = Visibility.Hidden;
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - menuRighPrint_Click : " + ex.ToString());
            }
        }

        // 닫기
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ContextMenu menu = btnPrint.ContextMenu;
                menu.StaysOpen = false;
                menu.IsOpen = false;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - menuClose_Click : " + ee.ToString());
            }
        }

        private void PrintWork(bool Ahead)
        {
            Lib lib2 = new Lib();
            try
            {
                excelapp = new Microsoft.Office.Interop.Excel.Application();

                string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\개인작업일보.xlsx";
                workbook = excelapp.Workbooks.Add(MyBookPath);
                worksheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                // 일자
                workrange = worksheet.get_Range("J1", "K1");
                workrange.Value2 = dtpEDate.SelectedDate.Value.ToShortDateString().Replace("-", ".");

                // 제목
                string title = ((CodeView)cboCategory.SelectedItem).code_name.Contains("리니어모터") ? "리니어모터" : "슬라이드";

                workrange = worksheet.get_Range("A2", "G4");
                workrange.Value2 = "개인 작업 일보 (" + title + ")";

                int copyLine = 1;
                int copyRow = 34;

                int inputPossibleRowCnt = 27;           // 내역 입력 가능한 갯수
                int startRowNum = 6;                    // 내역 입력 시작점
                int endCnt = 0;                         // 엑셀 입력 종료 갯수

                int cnt = 0;
                int totCnt = 0;

                int pageCnt = 1;
                int totPageCnt = (dgdByWorker.Items.Count / inputPossibleRowCnt) + 1;

                List<Win_Prd_ProcessResultSum_Q_ByWorker> listWorker = new List<Win_Prd_ProcessResultSum_Q_ByWorker>();

                // 작업자 정보만 분류
                foreach (Win_Prd_ProcessResultSum_Q_ByWorker pair in dgdByWorker.Items)
                {
                    if (pair.cls != "1")
                        continue;

                    listWorker.Add(pair);
                    endCnt++;
                }

                int person_start_idx = startRowNum;
                string person_name = "";
                foreach (Win_Prd_ProcessResultSum_Q_ByWorker pair in listWorker)
                {
                    int rowNum = startRowNum + cnt;

                    // 성명
                    string name = pair.Name;
                    if (person_name != name)
                    {
                        workrange = worksheet.get_Range("A" + person_start_idx.ToString(), "A" + Math.Max((rowNum - 1), person_start_idx).ToString());
                        workrange.Merge();
                        workrange.Value2 = string.IsNullOrEmpty(person_name) ? name : person_name;

                        person_start_idx = rowNum;
                    }

                    person_name = name;

                    // 품명
                    workrange = worksheet.get_Range("B" + rowNum.ToString(), "E" + rowNum.ToString());
                    workrange.Value2 = pair.Article;

                    // 공정
                    workrange = worksheet.get_Range("F" + rowNum.ToString(), "H" + rowNum.ToString());
                    workrange.Value2 = pair.Process;

                    // 수량
                    workrange = worksheet.get_Range("I" + rowNum.ToString());
                    workrange.Value2 = pair.WorkQty;

                    // 시간 합계
                    workrange = worksheet.get_Range("J" + rowNum.ToString(), "K" + rowNum.ToString());
                    string[] workTimes = pair.WorkTime.Split(':');
                    if (workTimes.Length >= 2)
                        workrange.Value2 = workTimes[0] + "시간 " + workTimes[1] + "분";
                    else
                        workrange.Value2 = "";

                    cnt++;
                    totCnt++;

                    if (totCnt == endCnt || cnt == inputPossibleRowCnt)
                    {
                        // 마지막 작업자 이름 삽입
                        workrange = worksheet.get_Range("A" + person_start_idx.ToString(), "A" + rowNum.ToString());
                        workrange.Merge();
                        workrange.Value2 = person_name;

                        // 페이지수
                        workrange = worksheet.get_Range("A33");
                        workrange.NumberFormat = "@";
                        workrange.Value2 = pageCnt.ToString() + "/" + totPageCnt.ToString();

                        // 붙여넣기
                        worksheet.Select();
                        worksheet.UsedRange.EntireRow.Copy();
                        pastesheet.Select();
                        workrange = pastesheet.Rows[copyLine];
                        workrange.Select();
                        pastesheet.Paste();

                        // 내역 삭제
                        workrange = worksheet.get_Range("A06", "K32");
                        workrange.ClearContents();

                        // 병합 해제 후 테두리
                        workrange = worksheet.get_Range("A6", "A32");
                        workrange.UnMerge();

                        // 있는 모든 테두리들
                        workrange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        workrange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                        // 외곽만
                        workrange.BorderAround2(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin);
                        workrange.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

                        copyLine += copyRow;

                        cnt = 0;
                        person_start_idx = startRowNum;

                        pageCnt++;
                    }
                }

                excelapp.Visible = true;
                msg.Hide();

                if (Ahead == true)
                    pastesheet.PrintPreview();
                else
                    pastesheet.PrintOutEx(IgnorePrintAreas: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 = PrintWork : " + ex.ToString());
            }

            lib2.ReleaseExcelObject(workbook);
            lib2.ReleaseExcelObject(worksheet);
            lib2.ReleaseExcelObject(pastesheet);
            lib2.ReleaseExcelObject(excelapp);
            lib2 = null;
        }
        #endregion 개인작업일보 인쇄
    }



    class Win_Prd_ProcessResultSum_Q_ByProcessMachine : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string cls { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string BuyerModel { get; set; }
        public string MachineID { get; set; }
        public string MachineNo { get; set; }
        public string Article { get; set; }
        //public string ArticleID { get; set; }
        public string KCustom { get; set; }
        //public string ProdQtyPerBox { get; set; }
        public string WorkQty { get; set; }
        public string UnitPrice { get; set; }
        public string Amount { get; set; }
        public string WorkTime { get; set; }
        //public string OutQtyPerBox { get; set; }
        public string QtyPerBox { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Cnt { get; set; }

    }

    class Win_Prd_ProcessResultSum_Q_ByArticle : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string cls { get; set; }

        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }

        public string BuyerArticleNo { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string BuyerModelID { get; set; }
        public string Model { get; set; }

        public string WorkQty { get; set; }
        public string ProdQtyPerBox { get; set; }

        public int Num { get; set; }
        public string Cnt { get; set; }

    }

    class Win_Prd_ProcessResultSum_Q_ByWorker : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string cls { get; set; }
        public string WorkPersonID { get; set; }
        public string Name { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }

        public string MachineID { get; set; }
        public string Machine { get; set; }
        public string MachineNo { get; set; }
        public string BuyerModelID { get; set; }
        public string Model { get; set; }

        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }

        public string WorkQty { get; set; }
        public string ProdQtyPerBox { get; set; }
        public int Num { get; set; }

        public string BuyerModel { get; set; }

        public string WorkTime { get; set; }
        public string Cnt { get; set; }

    }

    class Win_Prd_ProcessResultSum_Q_CodeView_ThisMonth : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }

        public string Article { get; set; }
        public string ArticleID { get; set; }
        public string BuyerArticleNo { get; set; }
        public double TotalQty { get; set; }
        public double SDay01 { get; set; }
        public double SDay02 { get; set; }
        public double SDay03 { get; set; }
        public double SDay04 { get; set; }
        public double SDay05 { get; set; }
        public double SDay06 { get; set; }
        public double SDay07 { get; set; }
        public double SDay08 { get; set; }
        public double SDay09 { get; set; }
        public double SDay10 { get; set; }
        public double SDay11 { get; set; }
        public double SDay12 { get; set; }
        public double SDay13 { get; set; }
        public double SDay14 { get; set; }
        public double SDay15 { get; set; }
        public double SDay16 { get; set; }
        public double SDay17 { get; set; }
        public double SDay18 { get; set; }
        public double SDay19 { get; set; }
        public double SDay20 { get; set; }
        public double SDay21 { get; set; }
        public double SDay22 { get; set; }
        public double SDay23 { get; set; }
        public double SDay24 { get; set; }
        public double SDay25 { get; set; }
        public double SDay26 { get; set; }
        public double SDay27 { get; set; }
        public double SDay28 { get; set; }
        public double SDay29 { get; set; }
        public double SDay30 { get; set; }
        public double SDay31 { get; set; }
        public string Cnt { get; set; }

    }


}
