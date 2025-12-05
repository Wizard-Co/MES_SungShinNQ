using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WizMes_SungShinNQ.PopUP;

namespace WizMes_SungShinNQ
{
    /// <summary>
    /// Win_ord_OutWare_Scan.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_OutWare_Scan : UserControl
    {
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        // 인쇄 활용 용도 (프린트)
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        WizMes_SungShinNQ.PopUp.NoticeMessage msg = new WizMes_SungShinNQ.PopUp.NoticeMessage();

        List<Win_ord_OutWare_Scan_CodeView> lstOutwarePrint = new List<Win_ord_OutWare_Scan_CodeView>();


        int rowNum = 0;                          // 조회시 데이터 줄 번호 저장용도
        string strFlag = string.Empty;           // 추가, 수정 구분 
        string GetKey = "";

        List<string> LabelGroupList = new List<string>();         // packing ID 스캔에 따른 LabelID를 모아 담을 리스트 그릇입니다.
        bool EventStatus = false;        // 추가 / 수정 상태확인을 위한 이벤트 bool

        public Win_ord_OutWare_Scan()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                chkOutwareDay.IsChecked = true; //출고일자 IsCheked
                dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");   // 오늘 날짜 자동 반영

                CantBtnControl();
                SetComboBox();
     
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - UserControl_Loaded : " + ee.ToString());
            }
        }

        #region 콤보박스
        private void SetComboBox()
        {
            try
            {
                ObservableCollection<CodeView> cbOutClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "OCD", "Y", "", "PROD");
                this.cboOutClss.ItemsSource = cbOutClss;
                this.cboOutClss.DisplayMemberPath = "code_name";
                this.cboOutClss.SelectedValuePath = "code_id";
                this.cboOutClss.SelectedIndex = 0;

                ObservableCollection<CodeView> cbFromLoc = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "INSIDE");
                this.cboFromLoc.ItemsSource = cbFromLoc;
                this.cboFromLoc.DisplayMemberPath = "code_name";
                this.cboFromLoc.SelectedValuePath = "code_id";
                this.cboFromLoc.SelectedIndex = 0;

                ObservableCollection<CodeView> cbToLoc = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "NONE");
                this.cboToLoc.ItemsSource = cbToLoc;
                this.cboToLoc.DisplayMemberPath = "code_name";
                this.cboToLoc.SelectedValuePath = "code_id";
                this.cboToLoc.SelectedIndex = 0;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - SetComboBox : " + ee.ToString());
            }
        }
        #endregion 콤보박스

        #region 상단 레이아웃 조건 모음
        //출고일자 라벨 클릭시
        private void lblOutwareDay_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkOutwareDay.IsChecked == true)
            {
                chkOutwareDay.IsChecked = false;

                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkOutwareDay.IsChecked = true;

                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }

        //출고일자 체크 
        private void ChkOutwareDay_Checked(object sender, RoutedEventArgs e)
        {
            chkOutwareDay.IsChecked = true;

            dtpFromDate.IsEnabled = true;
            dtpToDate.IsEnabled = true;

        }

        //출고일자 체크해제
        private void ChkOutwareDay_Unchecked(object sender, RoutedEventArgs e)
        {
            chkOutwareDay.IsChecked = false;

            dtpFromDate.IsEnabled = false;
            dtpToDate.IsEnabled = false;
        }

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpToDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpFromDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpToDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        //거래처 라벨 클릭시
        private void lblCustomer_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomer.IsChecked == true)
            {
                chkCustomer.IsChecked = false;
                txtCustomer.IsEnabled = false;
                btnCustomer.IsEnabled = false;
            }
            else
            {
                chkCustomer.IsChecked = true;
                txtCustomer.IsEnabled = true;
                btnCustomer.IsEnabled = true;
                txtCustomer.Focus();
            }
        }

        //거래처 체크
        private void ChkCustomer_Checked(object sender, RoutedEventArgs e)
        {
            chkCustomer.IsChecked = true;
            txtCustomer.IsEnabled = true;
            btnCustomer.IsEnabled = true;
            txtCustomer.Focus();
        }

        //거래처 체크 해제
        private void ChkCustomer_Unchecked(object sender, RoutedEventArgs e)
        {
            chkCustomer.IsChecked = false;
            txtCustomer.IsEnabled = false;
            btnCustomer.IsEnabled = false;
        }



        //거래처-조건 텍스트박스 키다운 이벤트
        private void txtCustomer_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtCustomer, 0, "");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtCustomer_KeyDown : " + ee.ToString());
            }
        }

        //거래처-조건 플러스파인더 버튼
        private void btnCustomer_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtCustomer, 0, "");
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnCustomer_Click : " + ee.ToString());
            }
        }

        //품명 라벨 클릭시
        private void lblArticle_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                chkArticle.IsChecked = false;
                txtArticle.IsEnabled = false;
                btnArticle.IsEnabled = false;

            }
            else
            {
                chkArticle.IsChecked = true;
                txtArticle.IsEnabled = true;
                btnArticle.IsEnabled = true;
                txtArticle.Focus();
            }
        }

        //품명 체크
        private void ChkArticle_Checked(object sender, RoutedEventArgs e)
        {
            chkArticle.IsChecked = true;
            txtArticle.IsEnabled = true;
            btnArticle.IsEnabled = true;
            txtArticle.Focus();
        }

        //품명 체크 해제
        private void ChkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticle.IsChecked = false;
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
        }


        //품명 텍스트박스 키다운 이벤트(품번으로 변경요청, 2020.03.26, 장가빈)
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtArticle, 81, txtArticle.Text);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtArticle_KeyDown : " + ee.ToString());
            }
        }

        //품명 플러스파인더 버튼(품번으로 변경요청, 2020.03.26, 장가빈)
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtArticle, 81, txtArticle.Text);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnArticle_Click : " + ee.ToString());
            }
        }

        //관리번호 라벨 클릭시
        private void lblRadioOptionNum_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkRadioOptionNum.IsChecked == true)
            {
                chkRadioOptionNum.IsChecked = false;
                txtRadioOptionNum.IsEnabled = false;
            }
            else
            {
                chkRadioOptionNum.IsChecked = true;
                txtRadioOptionNum.IsEnabled = true;
                txtRadioOptionNum.Focus();
            }
        }

        //관리번호 체크
        private void ChkRadioOptionNum_Checked(object sender, RoutedEventArgs e)
        {
            chkRadioOptionNum.IsChecked = true;
            txtRadioOptionNum.IsEnabled = true;
            txtRadioOptionNum.Focus();
        }

        //관리번호 체크 해제
        private void ChkRadioOptionNum_Unchecked(object sender, RoutedEventArgs e)
        {
            chkRadioOptionNum.IsChecked = false;
            txtRadioOptionNum.IsEnabled = false;
        }

        //라디오버튼 OrderNo 버튼 클릭
        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {
            //hidden 2020.01.25 안씀
        }

        //라디오버튼 OrderID 버튼 클릭
        private void rbnManageNum_Click(object sender, RoutedEventArgs e)
        {
            //hidden 2020.01.25 안씀
        }
        #endregion

        #region 버튼 모음
        //추가버튼 클릭
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            //2021-06-02
            EventStatus = true;
            try
            {
                strFlag = "I";

                this.DataContext = new object();
                CanBtnControl();                             //버튼 컨트롤
                dtpOutDate.SelectedDate = DateTime.Today;

                cboOutClss.SelectedIndex = 0;
                cboFromLoc.SelectedIndex = 0; //사내제품창고가 기본값이 되게 설정
                cboToLoc.SelectedIndex = 0;

                dgdOutwareSub.Items.Clear();      
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnAdd_Click : " + ee.ToString());
            }
        }

        //수정버튼 클릭
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var OutwareItem = dgdOutware.SelectedItem as Win_ord_OutWare_Scan_CodeView;

                if (OutwareItem != null)
                {
                    string classname = OutwareItem.OutClssname;
                    if (!classname.Equals("예외출고"))
                    {
                        EventStatus = true;
                        strFlag = "U";

                        rowNum = dgdOutware.SelectedIndex;
                        CanBtnControl();
                    }
                    else
                    {
                        MessageBox.Show("예외출고 수정은 예외출고메뉴에서 해주시기 바랍니다.");
                        return;
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnUpdate_Click : " + ee.ToString());
            }
        }

        //삭제버튼 클릭
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                if (lstOutwarePrint.Count == 0)
                {
                    MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제 데이터를 지정하고 눌러주세요.");
                }
                else
                {
                    if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {

                        foreach (Win_ord_OutWare_Scan_CodeView RemoveData in lstOutwarePrint)
                        {
                            if (DeleteData(RemoveData.OutwareID))
                            {
                                rowNum = 0;
                                re_Search(rowNum);
                            }
                        }
                        lstOutwarePrint.Clear();
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnDelete_Click : " + ee.ToString());
            }
        }

        //닫기버튼 클릭
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Lib.Instance.ChildMenuClose(this.ToString());
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnClose_Click : " + ee.ToString());
            }
        }

        //검색버튼 클릭
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                try
                {
                    rowNum = 0;
                    re_Search(rowNum);

                }
                catch (Exception ee)
                {
                    MessageBox.Show("오류지점 - btnSearch_Click : " + ee.ToString());
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
            
           
        }

        //저장버튼 클릭
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CantBtnControl();           //버튼 컨트롤

                //for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                //{
                //    var OutwareSub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;
                //    if(OutwareSub.OutwareID == "" && OutwareSub.OutSubSeq == "" )
                //    {
                //        return;
                //    }
                //    if (!CheckStock(OutwareSub))
                //    {
                //        return;
                //    }
                //}

                //저장 전에 한번더 수량 계산 하도록 추가
                SumScanQty();

                if (SaveData(strFlag))
                {
                    if (strFlag.Equals("I"))
                    {
                        var outwareCount = dgdOutware.Items.Count;

                        rowNum = outwareCount;
                        //re_Search(rowNum);

                    }
                    //else if (strFlag.Equals("U"))
                    //{
                    //    re_Search(rowNum);
                    //}
                    //2021-06-02 
                    TextBoxClear(); // 저장했으면 클리어 해야지
                    //re_Search(rowNum);
                    strFlag = string.Empty;
                    //TextBoxClear(); //20200526 이거 때문에 거래처가 클리어 되서 수정할때 테그값이 없었음


                    re_Search(rowNum);
                    EventStatus = false;
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnSave_Click : " + ee.ToString());
            }
        }

        //취소버튼 클릭
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                EventStatus = false;
                CantBtnControl();           //버튼 컨트롤
                ClearGrdInput();

                if (strFlag.Equals("I"))
                {
                    re_Search(0);
                }
                else if (strFlag.Equals("U"))
                {
                    re_Search(rowNum);
                }


                strFlag = string.Empty;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnCancel_Click : " + ee.ToString());
            }
        }

        //엑셀버튼 클릭
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            Lib lib2 = new Lib();
            try
            {
                if (dgdOutware.Items.Count < 1)
                {
                    MessageBox.Show("먼저 검색해 주세요.");
                    return;
                }
                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[4];
                lst[0] = "메인그리드";
                lst[1] = "서브그리드";
                lst[2] = dgdOutware.Name;
                lst[3] = dgdOutwareSub.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdOutware.Name))
                    {
                        //MessageBox.Show("대분류");
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib2.DataGridToDTinHidden(dgdOutware);
                        else
                            dt = lib2.DataGirdToDataTable(dgdOutware);

                        Name = dgdOutware.Name;
                        if (lib2.GenerateExcel(dt, Name))
                        {
                            lib2.excel.Visible = true;
                            lib2.ReleaseExcelObject(lib2.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdOutwareSub.Name))
                    {
                        //MessageBox.Show("정성류");
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib2.DataGridToDTinHidden(dgdOutwareSub);
                        else
                            dt = lib2.DataGirdToDataTable(dgdOutwareSub);
                        Name = dgdOutwareSub.Name;
                        if (lib2.GenerateExcel(dt, Name))
                        {
                            lib2.excel.Visible = true;
                            lib2.ReleaseExcelObject(lib2.excel);
                        }
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
                MessageBox.Show("오류지점 - btnExcel_Click : " + ee.ToString());
            }
            finally
            {
                lib2 = null;
            }
        }

        //인쇄버튼 클릭
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

        //인쇄-미리보기 클릭
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgdOutware.Items.Count == 0)
                {
                    MessageBox.Show("먼저 검색해 주세요.");
                    return;
                }
                var OBJ = dgdOutware.SelectedItem as Win_ord_OutWare_Scan_CodeView;
                if (OBJ == null)
                {
                    MessageBox.Show("거래명세표 항목이 정확히 선택되지 않았습니다.");
                    return;
                }
                msg.Show();
                msg.Topmost = true;
                msg.Refresh();

                lib.Delay(1000);

                PrintWork(true);
                msg.Visibility = Visibility.Hidden;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - menuSeeAhead_Click : " + ee.ToString());
            }
        }

        //인쇄-바로인쇄 클릭
        private void menuRighPrint_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgdOutware.Items.Count == 0)
                {
                    MessageBox.Show("먼저 검색해 주세요.");
                    return;
                }
                var OBJ = dgdOutware.SelectedItem as Win_ord_OutWare_Scan_CodeView;
                if (OBJ == null)
                {
                    MessageBox.Show("거래명세표 항목이 정확히 선택되지 않았습니다.");
                    return;
                }
                msg.Show();
                msg.Topmost = true;
                msg.Refresh();

                lib.Delay(1000);

                PrintWork(false);
                msg.Visibility = Visibility.Hidden;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - menuRighPrint_Click : " + ee.ToString());
            }
        }

        //인쇄-닫기 클릭
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
        #endregion

        #region 키다운 이동 모음
        //관리번호 텍스트박스 키다운 이벤트
        private void txtOrderID_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtOrderID, 99, txtOrderID.Text);

                    if (txtOrderID.Text.Length > 0)
                    {
                        //관리번호 기반_ 항목 뿌리기 작업.
                        OrderID_OtherSearch(txtOrderID.Text);
                    }

                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtOrderID_KeyDown : " + ee.ToString());
            }
        }

   
   


        //수주거래처 키다운 이벤트
        private void TxtKCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtBuyerName.Focus();
            }
        }


        //납품거래처 텍스트박스 키다운 이벤트
        private void txtBuyerName_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtBuyerName, 0, "");

                    if (txtBuyerName.Text.Length > 0)
                    {
                        txtOutCustom.Text = txtBuyerName.Text;
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtBuyerName_KeyDown : " + ee.ToString());
            }
        }

   

        //출고처 키다운 이벤트
        private void TxtOutCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                cboFromLoc.IsDropDownOpen = true;
            }
        }
        #endregion

        #region 플러스파인더 및 데이터그리드 선택 변경

        //메인 데이터그리드 선택 변경
        private void dgdOutware_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var OutwareInfo = dgdOutware.SelectedItem as Win_ord_OutWare_Scan_CodeView;

                if (OutwareInfo != null)
                {
                    this.DataContext = OutwareInfo;
                    // 2021-06-02; 태그는 안넣어지니깐 클릭했는테그 넣어야지
                    txtKCustom.Tag = OutwareInfo.CustomID;
                    txtBuyerName.Tag = OutwareInfo.DvlyCustomID;
                    txtOutCustom.Tag = OutwareInfo.OutCustomID;
                    txtArticle_InGroupBox.Tag = OutwareInfo.ArticleID;

                    String OutwareID = OutwareInfo.OutwareID;
                    FillGridSub(OutwareID);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutware_SelectionChanged : " + ee.ToString());
            }
        }

        //관리번호 플러스파인더 버튼 클릭
        private void btnOrderID_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtOrderID, 99, txtOrderID.Text);

                if (txtOrderID.Text.Length > 0)
                {
                    //관리번호 기반_ 항목 뿌리기 작업.
                    OrderID_OtherSearch(txtOrderID.Text);
                }
                cboOutClss.IsDropDownOpen = true;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnOrderID_Click : " + ee.ToString());
            }
        }

        //납품거래처 플러스파인더 버튼
        private void btnBuyerName_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtBuyerName, 0, "");

                if (txtBuyerName.Text.Length > 0)
                {
                    txtOutCustom.Text = txtBuyerName.Text;
                }  
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnBuyerName_Click : " + ee.ToString());
            }
        }

        //라벨스캔 텍스트박스 키다운 이벤트
        private void txtScanData_KeyDown(object sender, KeyEventArgs e)
        {
            #region ....
            ////try
            ////{
            ////    if (e.Key == Key.Enter)
            ////    {
            ////        //1. 일반 케이스 (사내라벨 스캔시)
            ////        if (txtScanData.Text.Trim().Length != 11)   // 삼주테크 바코드 길이 13자리로 확정
            ////        {
            ////            MessageBox.Show("잘못된 바코드 입니다.");
            ////            txtScanData.Text = string.Empty;
            ////            return;
            ////        }

            ////        if (txtScanData.Text.Trim().Substring(0, 1) == "P")
            ////        {
            ////            //2018.07.05 PACKINGID SCAN 과정 추가._허윤구.
            ////            // 지금 스캔된 녀석은 PACKING이다.
            ////            // 성공적으로 Packing List를 가져왔을 때,
            ////            if (FindPackingLabelID(txtScanData.Text.Trim()) == true)
            ////            {
            ////                string InsideLabelID = string.Empty;

            ////                // 리스트 내 LabelID를 돌면서 박스 스캔. > SUBGRID 추가(여러개)
            ////                for (int j = 0; j < LabelGroupList.Count; j++)
            ////                {
            ////                    InsideLabelID = LabelGroupList[j].ToString();

            ////                    FindBoxScanData(InsideLabelID);
            ////                }
            ////            }
            ////        }
            ////        else
            ////        {
            ////            //부품식별표 박스ID 스캔 > SUBGRID 추가
            ////            FindBoxScanData(txtScanData.Text.Trim());
            ////        }
            ////        txtScanData.Text = string.Empty;
            ////    }

            ////    SumScanQty();
            ////}
            ////catch (Exception ee)
            ////{
            ////    MessageBox.Show("오류지점 - txtScanData_KeyDown : " + ee.ToString());
            ////}
            ///
            #endregion
            try
            {
                if (e.Key == Key.Enter)
                {
                    if (tgnMoveByID.IsChecked == true)
                    {
                        //1. 일반 케이스 (사내라벨 스캔시)
                        if (txtScanData.Text.Trim().Length != 11)   // 삼주테크 바코드 길이 13자리로 확정
                        {
                            MessageBox.Show("잘못된 바코드 입니다.");
                            txtScanData.Text = string.Empty;
                            return;
                        }

                        if (txtScanData.Text.Substring(0, 1) == "P")
                        {
                            //2018.07.05 PACKINGID SCAN 과정 추가._허윤구.
                            // 지금 스캔된 녀석은 PACKING이다.
                            // 성공적으로 Packing List를 가져왔을 때,
                            if (FindPackingLabelID(txtScanData.Text) == true)
                            {
                                string InsideLabelID = string.Empty;

                                // 리스트 내부 LabelID를 돌면서 박스 스캔. > SUBGRID 추가(여러개)
                                for (int j = 0; j < LabelGroupList.Count; j++)
                                {
                                    InsideLabelID = LabelGroupList[j].ToString();

                                    FindBoxScanData(InsideLabelID);
                                }
                            }
                        }
                        else
                        {
                            //부품식별표 박스ID 스캔 > SUBGRID 추가
                            FindBoxScanData(txtScanData.Text);
                        }
                        txtScanData.Text = string.Empty;
                    }

                    if (tgnMoveByQty.IsChecked == true && !string.IsNullOrEmpty(txtOrderID.Text))
                    {
                        
                        if (chkAutoPackingLoad.IsChecked == true && grdAutoPackingLoad.Visibility == Visibility.Visible)
                        {
                            List<Win_ord_OutWare_Scan_Sub_CodeView> LotIDsInSub = new List<Win_ord_OutWare_Scan_Sub_CodeView>();

                            if (dgdOutwareSub.Items.Count > 0)
                            foreach (Win_ord_OutWare_Scan_Sub_CodeView item in dgdOutwareSub.Items)
                            {
                                LotIDsInSub.Add(item);
                            }

                            // 바코드에 수량을 입력 → 숫자만 입력 가능하도록 유효성 검사
                            if (txtScanData.Text != "" && CheckConvertInt(txtScanData.Text))
                            {
                                List<Win_ord_OutWare_Scan_Sub_CodeView> Scan_Sub = FindBoxesByArticleID(LotIDsInSub, txtArticleID_InGroupBox.Text, txtScanData.Text);

                                if (Scan_Sub.Count > 0)
                                {
                                    foreach (Win_ord_OutWare_Scan_Sub_CodeView item in Scan_Sub)
                                    {
                                        dgdOutwareSub.Items.Add(item);
                                    }
                                }

                            }
                            else
                            {
                                MessageBox.Show("수량입력은 숫자만 가능합니다.", "확인");
                            }
                        }
                        else
                        {
                            if (txtScanData.Text != "" && CheckConvertInt(txtScanData.Text))
                            {
                                //수량 입력시 라벨 없이 입력됨
                                Win_ord_OutWare_Scan_Sub_CodeView label = new Win_ord_OutWare_Scan_Sub_CodeView();

                                int num = dgdOutwareSub.Items.Count + 1;
                                label.Num = num;
                                label.LabelID = "";
                                //label.Spec = "";
                                //label.Orderseq = orderSeq;
                                label.OutQty = stringFormatN0(txtScanData.Text);
                                label.UnitPrice = "0";
                                label.ArticleID = txtArticleID_InGroupBox.Text;
                                dgdOutwareSub.Items.Add(label);

                                // 데이터 그리드 등록 후 바코드 초기화
                            }
                            else
                            {
                                MessageBox.Show("수량입력은 숫자만 가능합니다.", "확인");
                            }
                        }


                        txtScanData.Text = "";

                    }
                    else
                    {
                        if (txtOrderID.Text == string.Empty || txtOrderID.Text == "")
                            MessageBox.Show("관리번호를 먼저 검색하여 주십시오.", "확인");
                    }

                }

                SumScanQty();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtScanData_KeyDown : " + ee.ToString());
            }
        }

        private List<Win_ord_OutWare_Scan_Sub_CodeView> FindBoxesByArticleID(List<Win_ord_OutWare_Scan_Sub_CodeView> SubItems,  string ArticleID, string OutQtyWant)
        {
            List<Win_ord_OutWare_Scan_Sub_CodeView> returnListCodeView = new List<Win_ord_OutWare_Scan_Sub_CodeView>();            

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("ArticleID", ArticleID);
            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sGetBoxes", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = null;
                dt = ds.Tables[0];

                if(dt.Rows.Count > 0) 
                {
                    int remainQty = Convert.ToInt32(OutQtyWant);
                    int totalBoxQty = 0;
                    int emptyItemsQty = 0;
                    if (SubItems != null)
                    {
                        foreach (var item in SubItems)
                        {
                            if (string.IsNullOrEmpty(item.LabelID))
                            {
                                emptyItemsQty += lib.RemoveComma(item.OutQty, 0);
                            }
                        }
                    }

                    if (emptyItemsQty > 0)
                    {
                        remainQty = Math.Max(0, remainQty - emptyItemsQty);
                   
                    }
                    foreach(DataRow dr in dt.Rows)
                    {
                        int BoxQty = lib.RemoveComma(dr["BoxQty"].ToString(), 0);
                        totalBoxQty += BoxQty;

                        if (remainQty <= 0)
                        {                   
                            break;
                        }

                        // 현재 LotID가 이미 있는지 확인
                        if (SubItems?.Any(x => x.LabelID == dr["LotID"].ToString()) == true)
                        {
                            remainQty -= BoxQty;
                            continue;
                        }


                        int OutQty = 0;
                        if(remainQty > BoxQty)
                        {
                            OutQty = BoxQty;
                            remainQty -= BoxQty;
                        }
                        else
                        {
                            OutQty = remainQty;
                            remainQty = 0;

                        }

                        var item = new Win_ord_OutWare_Scan_Sub_CodeView
                        {
                            ArticleID = dr["ArticleID"].ToString(),
                            LabelID = dr["LotID"].ToString(),
                            OutQty = stringFormatN0(OutQty),
                            UnitPrice = "0", //일단 0으로 해놓고 나중에 단가도 입력해달라하면...

                        };

                        returnListCodeView.Add(item);
          
                    }

                    if(remainQty > 0)
                    {
                        MessageBoxResult msgResult = MessageBox.Show($"출하희망량 : ({stringFormatN0(Convert.ToInt32(txtScanData.Text))})\n" +
                                                                     $"검사/포장수량 : ({stringFormatN0(totalBoxQty)})" +
                                                                     $"\n남은 수량은 바코드번호 없이 처리 하시겠습니까?","확인",MessageBoxButton.YesNo);
                        if(msgResult == MessageBoxResult.Yes)
                        {
                            var item = new Win_ord_OutWare_Scan_Sub_CodeView
                            {
                                ArticleID = txtArticleID_InGroupBox.Tag?.ToString() ?? string.Empty,
                                LabelID = string.Empty,
                                OutQty = stringFormatN0(remainQty)

                            };

                            returnListCodeView.Add(item);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("검사/포장 건이 없습니다.");
                    MessageBoxResult msgResult = MessageBox.Show("바코드번호 없이 처리 하시겠습니까?", "확인", MessageBoxButton.YesNo);
                    if (msgResult == MessageBoxResult.Yes)
                    {
                        var item = new Win_ord_OutWare_Scan_Sub_CodeView
                        {
                            ArticleID = txtArticleID_InGroupBox.Tag?.ToString() ?? string.Empty,
                            LabelID = string.Empty,
                            OutQty = stringFormatN0(Convert.ToInt32(txtScanData.Text))

                        };

                        returnListCodeView.Add(item);
                    }
                }
            }

            return returnListCodeView;
        }

        private bool CheckConvertInt(string str)
        {
            bool flag = false;
            int chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace(",", "");

                if (Int32.TryParse(str, out chkInt) == true)
                    flag = true;
            }

            return flag;
        }

        //PACKINGID SCAN 과정 > LABELID LIST 담기.
        private bool FindPackingLabelID(string PackingLabelID)
        {
            try
            {


                LabelGroupList.Clear();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("PackingLabelID", PackingLabelID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sPackingIDList", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("PackingID를 포함하고 있는 LabelID를 찾을 수 없습니다.");
                        return false;
                    }
                    else
                    {
                        LabelGroupList.Clear();
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            LabelGroupList.Add(dt.Rows[i]["InBoxID"].ToString());
                        }
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - FindPackingLabelID : " + ee.ToString());
                return false;
            }
        }
        private Boolean CheckStock(Win_ord_OutWare_Scan_Sub_CodeView scanData)
        {
            string outqty;
            DataSet ds;
            if (scanData.OutSubSeq == null && scanData.OutwareID == null)
            {
                outqty = "0";
            }
            else
            {
                String sql = "SELECT * FROM [OutwareSub] WHERE OutWareID = '" + scanData.OutwareID + "' AND OutSubSeq = " + scanData.OutSubSeq + "";
                 ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    DataRow row = dt.Rows[0];
                    outqty = row["OutQty"].ToString();
                }
                else
                {
                    outqty = "0";
                }
            }
            

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("BoxID", scanData.LabelID);

            ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sBoxIDOne_20200727_test", sqlParameter, false);
            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = null;
                dt = ds.Tables[0];

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("존재하지 않거나, 생산, 검사되지 않은 바코드 입니다.\r\n" +
                        "바코드 번호 :" + scanData.LabelID);
                    return false;
                }
                else
                {
                    DataRow DR = dt.Rows[0];
                    double availableQty = Convert.ToDouble(DR["qtyperbox"].ToString()) + Convert.ToDouble(outqty) ;

                    if (Convert.ToDouble(scanData.OutQty) > availableQty)
                    {
                        MessageBox.Show("입력한 수량이 실시간 현재고 보다 많습니다. 재고를 다시 확인해 주세요.");
                        return false;
                    }
                }
            }

            return true;
        }
        // 부품식별표 박스ID 스캔 > SUBGRID 추가
        private void FindBoxScanData(string ScanData)
        {
            try
            {
                LabelGroupList.Clear();
                ScanData = ScanData.ToUpper();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("BoxID", ScanData.ToUpper());

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sBoxIDOne", sqlParameter, false); ////// 2020.01.20 장가빈, wk_packing 테이블
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("존재하지 않거나, 생산, 검사되지 않은 바코드 입니다.\r\n" +
                            "바코드 번호 :" + ScanData);
                        return;
                    }
                    else
                    {
                        DataRow DR = dt.Rows[0];

                        //세부작업 1. 스캔값에 대한 각종검증작업. > 리턴처리

                        /* if (DR["OutDate"].ToString() != string.Empty) // OutDate 컬럼에 값이 들어가 있으면 
                         {
                             MessageBox.Show(ScanData + " : 이미 출고된 바코드 번호입니다.");
                             return;
                         }*/
                        double remainQty = (double)RemoveComma(DR["qtyperbox"].ToString(), true, typeof(double)); /*Convert.ToDouble(DR["qtyperbox"].ToString())*/;

                        if ((cboOutClss.SelectedValue.ToString() == "11" || cboOutClss.SelectedValue.ToString() == "17") && remainQty >= 0)  {}
                        else
                        {
                            if (DR["qtyperbox"].ToString() == null || DR["qtyperbox"].ToString() == string.Empty || remainQty <= 0)
                            {
                                MessageBox.Show("출고/반품 가능한 수량이 없습니다.");
                                return;
                            }
                        }
                        
                        if (DR["ScanDate"].ToString() == string.Empty) //ScanDate 컬럼에 값이 비어있으면 / ScanDate는 PackDate와 같다
                        {
                            MessageBox.Show("생산이력이 없는 바코드 번호입니다.");
                            return;
                        }
                        if (DR["inspectDate"].ToString() == string.Empty)   //wk_PackingCardList 테이블의 InspectDate / 검사일자가 비어있다면
                        {
                            MessageBox.Show("검사이력이 없는 바코드 번호입니다.");
                            return;
                        }
                        if ((lib.IsNumOrAnother(DR["GradeID"].ToString()) == true) && (lib.IsNumOrAnother(DR["DefectClss"].ToString()) == true)) //등급과 결함 구분에 값이 있으면
                        {
                            if (Convert.ToDouble(DR["GradeID"].ToString()) >= Convert.ToDouble(DR["DefectClss"].ToString())) //등급 >= 결함구분 값보다 크면
                            {
                                MessageBox.Show("불량등급이" + DR["GradeID"].ToString() + "이므로 출고가 불가능합니다.");
                                return;
                            }
                        }
                        if (txtArticle_InGroupBox.Tag != null) //품명 텍스트 박스에 값이 있고,
                        {
                            if (txtArticle_InGroupBox.Tag.ToString() != DR["ArticleID"].ToString()) //품명 텍스트 박스에 기재된 품명과 받아온 품명이 다르면
                            {
                                MessageBox.Show("서로 다른 품명을 동시에 출고처리 할 수 없습니다. \r\n" +
                                    "바코드 품명 :" + DR["Article"].ToString() + ". \r\n" +
                                    "출고 품명 :" + txtArticle_InGroupBox.Text + ".");
                                return;
                            }
                        }
                        if (txtKCustom.Tag != null) //거래처 텍스트 박스에 값이 있고, 
                        {
                            if (txtKCustom.Tag.ToString() != DR["CustomID"].ToString())         //거래처 텍스트 박스에 기재된 거래처와 받아온 거래처가 다르면
                            {
                                MessageBox.Show("서로 다른 거래처를 동시에 출고처리 할 수 없습니다. \r\n" +
                                    "바코드 거래처 :" + DR["CustomName"].ToString() + ". \r\n" +
                                    "출고 거래처 :" + txtKCustom.Text + ".");
                                return;
                            }
                        }

                        for (int i = 0; i < dgdOutwareSub.Items.Count; i++)     //이미 스캔한 바코드인지 체크, 
                        {
                            var OutSub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;

                            //DataGridRow dgr = lib.GetRow(i, dgdOutwareSub);
                            //var ViewReceiver = dgr.Item as Win_ord_OutWare_Scan_CodeView;

                            if (OutSub.LabelID.ToUpper() == ScanData.ToUpper())
                            {
                                MessageBox.Show((i + 1) + "줄에 이미 스캔된 바코드 입니다.");
                                return;
                            }
                        }

                        //세부작업 2. 관리번호가 비어있다면 > 스캔항목을 통해 관리번호 자동유추 > 관리번호 값 입력.
                        if (txtOrderID.Text == string.Empty)
                        {
                            txtOrderID.Tag = DR["OrderID"].ToString();
                            txtOrderID.Text = DR["OrderID"].ToString();

                            // 관리번호 기반_ 항목 뿌리기 작업.
                            OrderID_OtherSearch(txtOrderID.Text);
                        }
                        else
                        {
                            txtOrderID.Tag = DR["OrderID"].ToString();
                            txtOrderID.Text = DR["OrderID"].ToString();

                            OrderID_OtherSearch(txtOrderID.Text);
                        }

                        //세부작업 3. dgdOutwareSub에 ScanData Box DR 값 추가. (+ 1 Row)
                        var Win_ord_OutWare_Scan_Insert = new Win_ord_OutWare_Scan_Sub_CodeView()
                        {
                            LabelID = ScanData,                         //바코드 번호
                            OutQty = Lib.Instance.returnNumStringZero(DR["QtyPerBox"].ToString()),        //수량
                            OutRealQty = Lib.Instance.returnNumStringZero(DR["QtyPerBox"].ToString()),
                            UnitPrice = DR["UNITPRICE"].ToString(),     //단가
                            Orderseq = DR["OrderSeq"].ToString(),       //수주순서?
                            Amount = DR["Amount"].ToString(),           //금액
                            Vat_IND_YN = DR["VAT_IND_YN"].ToString(),    //부가세별도여부
                            LabelGubun = DR["labelGubun"].ToString(),    //라벨구분
                            Article = DR["Article"].ToString(),          //품명           
                            ArticleID = DR["ArticleID"].ToString(),

                            DeleteYN = "Y",
                        };

                        //dgdOutwareSub.Items.Add(Win_ord_OutWare_Scan_Insert);
                        dgdOutwareSub.Items.Insert(0, Win_ord_OutWare_Scan_Insert); //2021-05-21 최근에 스캔 한 것이 위로 오게 수정

                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - FindBoxScanData : " + ee.ToString());
            }
        }

        //서브 데이터 그리드 변경 이벤트
        private void dgdOutwareSub_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if ((btnSave.Visibility == Visibility.Visible) && (btnCancel.Visibility == Visibility.Visible))
                {
                    //추가 / 수정 이벤트가 진행중인 경우,
                    var deleteControl = dgdOutwareSub.SelectedItem as Win_ord_OutWare_Scan_Sub_CodeView;
                    if (deleteControl != null)
                    {
                        deleteControl.DeleteYN = "Y";
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutwareSub_SelectionChanged : " + ee.ToString());
            }
        }

        //서브 데이터 그리드 키다운 이벤트
        private void dgdOutwareSub_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Delete)
                {
                    //추가 / 수정 이벤트가 진행중인 경우,
                    if ((btnSave.Visibility == Visibility.Visible) && (btnCancel.Visibility == Visibility.Visible))
                    {
                        var OutwareSub = dgdOutwareSub.SelectedItem as Win_ord_OutWare_Scan_Sub_CodeView;
                        if (OutwareSub != null)
                        {
                            dgdOutwareSub.Items.Remove(OutwareSub);
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutwareSub_KeyDown : " + ee.ToString());
            }
        }

        #endregion

        #region Research
        private void re_Search(int rowNum)
        {
            try
            {
                //dgdOutware.Items.Clear();
                //dgdOutwareSub.Items.Clear();

                FillGrid();

                if (dgdOutware.Items.Count > 0)
                {
                    dgdOutware.SelectedIndex = rowNum;
                }
                else
                {
                    this.DataContext = new object();
                    return;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - re_Search : " + ee.ToString());
            }
        }

        #endregion

        #region 조회
        private void FillGrid()
        {
            if (dgdOutware.Items.Count > 0)
            {
                dgdOutware.Items.Clear();
                dgdOutwareSub.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                int i = 0;
                sqlParameter.Add("ChkDate", chkOutwareDay.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkOutwareDay.IsChecked == true ?
                                            dtpFromDate.ToString().Substring(0, 10).Replace("-", "") : "");
                sqlParameter.Add("EDate", chkOutwareDay.IsChecked == true ?
                                            dtpToDate.ToString().Substring(0, 10).Replace("-", "") : "");

                //sqlParameter.Add("ChkCustomID", chkCustomer.IsChecked == true ?
                //                                (txtCustomer.Tag.ToString() != null ? 1 : 2) : 0);

                sqlParameter.Add("ChkCustomID", chkCustomer.IsChecked == true ? (txtCustomer.Tag != null ? 1 : 2) : 0);

                //sqlParameter.Add("CustomID", chkCustomer.IsChecked == true ? (txtCustomer.Tag.ToString()) : "");

                sqlParameter.Add("CustomID", chkCustomer.IsChecked == true ? (txtCustomer.Tag == null ? "" : txtCustomer.Tag) : "");


                sqlParameter.Add("Custom", txtCustomer.Text == "" ? "" : txtCustomer.Text);

                //sqlParameter.Add("ChkArticleID", chkArticle.IsChecked == true ?
                //                                (txtArticle.Tag.ToString() != null ? 1 : 2) : 0);
                sqlParameter.Add("ChkArticleID", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true ? (txtArticle.Tag == null ? "" : txtArticle.Tag.ToString()) : "");
                //sqlParameter.Add("ArticleID", chkArticle.IsChecked == true ? (txtArticle.Tag.ToString()) : "");
                sqlParameter.Add("Article", txtArticle.Text);
                sqlParameter.Add("ChkOrder", chkRadioOptionNum.IsChecked == true ?
                                             (rbnManageNum.IsChecked == true ? 1 : 2) : 0);


                sqlParameter.Add("Order", chkRadioOptionNum.IsChecked == true ? (txtRadioOptionNum.Text) : "");
                sqlParameter.Add("OutFlag", 0);             // OutType조회, 0이면 구분없이 전체 조회
                sqlParameter.Add("OutClss", "");            // 출고구분 같은데, 빈값이면 전체 조회
                sqlParameter.Add("FromLocID", "");          // 무슨 일자인지 몰라서 빈값으로 전체조회
                sqlParameter.Add("ToLocID", "");            // 무슨 일자인지 몰라서 빈값으로 전체조회


                sqlParameter.Add("BuyerDirectYN", "Y");     //왜 Y만 검색하지?
                sqlParameter.Add("nBuyerArticleNo", "");      //모르겠어서 빈값으로 전체조회
                sqlParameter.Add("BuyerArticleNo", txtArticle.Text);

                sqlParameter.Add("ChkLabelID", chkLabelIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("LabelID", chkLabelIDSrh.IsChecked == true ? txtLabelIDSrh.Text : string.Empty);

                ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOrder", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다. 검색조건을 확인해 주세요.");
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow dr in drc)
                        {
                            i++;
                            double RemainQty = 0;   //잔여수량?
                            if ((lib.IsNumOrAnother(dr["OrderQty"].ToString()) == true) && (lib.IsNumOrAnother(dr["OutSumQty"].ToString()) == true))
                            {   //수주량 - 출고합계수량 = 잔여수량?
                                RemainQty = ConvertDouble(dr["OrderQty"].ToString()) - ConvertDouble(dr["OutSumQty"].ToString());
                            }

                            //double OutQty = 0;      //출고량
                            //OutQty = Convert.ToDouble(dr["OutQty"].ToString());

                            var Win_ord_OutWare_Scan_Insert = new Win_ord_OutWare_Scan_CodeView()
                            {
                                OutwareID = dr["OutwareID"].ToString(),       //출고번호
                                OrderID = dr["OrderID"].ToString(),           //관리번호
                                OutSeq = dr["OutSeq"].ToString(),             //순번
                                OrderNo = dr["OrderNo"].ToString(),           //OrderNo
                                CustomID = dr["CustomID"].ToString(),         //거래처코드

                                KCustom = dr["KCustom"].ToString(),           //수주거래처명
                                OutDate = dr["OutDate"].ToString(),           //출고일자
                                ArticleID = dr["ArticleID"].ToString(),       //품명코드
                                Article = dr["Article"].ToString(),           //품명

                                OutClss = dr["OutClss"].ToString(),           //출고구분
                                WorkID = dr["WorkID"].ToString(),             //가공구분코드?? 
                                OutRoll = dr["OutRoll"].ToString(),           //박스 수량
                                OutQty = dr["OutQty"].ToString(),             //개별 수량
                                OutRealQty = dr["OutRealQty"].ToString(),     //소요량??

                                ResultDate = dr["ResultDate"].ToString(),     //무슨날? outdate랑 같은 날인데??
                                RemainQty = RemainQty.ToString(),             //잔량
                                OrderQty = dr["OrderQty"].ToString(),         //수주량
                                UnitClss = dr["UnitClss"].ToString(),         //단위 
                                WorkName = dr["WorkName"].ToString(),         //작업명??

                                OutType = dr["OutType"].ToString(),           //출고구분(출고방식)
                                Remark = dr["Remark"].ToString(),             //비고
                                BuyerModel = dr["BuyerModel"].ToString(),     //차종
                                OutSumQty = dr["OutSumQty"].ToString(),       //누계출고
                                OutQtyY = dr["OutQtyY"].ToString(),           // ???

                                StuffinQty = dr["StuffinQty"].ToString(),     //입고 수량인가요?
                                OutWeight = dr["OutWeight"].ToString(),       //출고 중량??
                                OutRealWeight = dr["OutRealWeight"].ToString(), //출고 실중량??
                                BuyerDirectYN = dr["BuyerDirectYN"].ToString(), //납품처 직접출고

                                Vat_Ind_YN = dr["Vat_Ind_YN"].ToString(),         //부가세별도여부
                                InsStuffINYN = dr["InsStuffINYN"].ToString(),     //동시입고여부???
                                ExchRate = dr["ExchRate"].ToString(),             //환율
                                FromLocID = dr["FromLocID"].ToString(),           //?
                                TOLocID = dr["TOLocID"].ToString(),               // ??
                                UnitClssName = dr["UnitClssName"].ToString(),     //단위 EA, kg같은 거
                                FromLocName = dr["FromLocName"].ToString(),       //전 창고명
                                TOLocname = dr["TOLocname"].ToString(),           //후 창고명

                                OutClssname = dr["OutClssname"].ToString(),       //출고구분
                                //UnitPrice = dr["UnitPrice"].ToString(),           //단가
                                Amount = dr["Amount"].ToString(),                 //금액
                                VatAmount = dr["VatAmount"].ToString(),           //vat금액

                                DvlyCustomID = dr["DvlyCustomID"].ToString(),     //20210526
                                DvlyCustom = dr["DvlyCustom"].ToString(),         //20210526

                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(), //품번
                                OutCustomID = dr["OutCustomID"].ToString(),       //출고처코드
                                BuyerID = dr["BuyerID"].ToString(),               //납품거래처 코드
                                BuyerName = dr["BuyerName"].ToString(),           //납품거래처명
                                OutCustom = dr["OutCustom"].ToString(),           //출고처

                                //거래명세표에 필요한 데이터를 받아옴
                                Buyer_Chief = dr["Buyer_Chief"].ToString(),       //공급받는 자 대표자
                                Buyer_Address1 = dr["Buyer_Address1"].ToString(), //공급받는 자 주소
                                Buyer_Address2 = dr["Buyer_Address2"].ToString(), //공급받는 자 주소
                                Buyer_Address3 = dr["Buyer_Address3"].ToString(), //공급받는 자 주소
                                CustomNo = dr["CustomNo"].ToString(),             //사업자등록번호
                                Chief = dr["Chief"].ToString(),                   //공급하는 대표자명

                                Condition = dr["Condition"].ToString(),           //업테 2021-05-31
                                Category = dr["Category"].ToString(),             //종목 2021-05-31

                            };

                            //출고일자 데이트피커 포맷으로 변경
                            Win_ord_OutWare_Scan_Insert.OutDate = DatePickerFormat(Win_ord_OutWare_Scan_Insert.OutDate);
                            //잔량, 수주량, 소요량, 출고량, 누계출고, 단가 소숫점 두자리 변환
                            Win_ord_OutWare_Scan_Insert.RemainQty = Lib.Instance.returnNumStringZero(Win_ord_OutWare_Scan_Insert.RemainQty);
                            Win_ord_OutWare_Scan_Insert.OrderQty = Lib.Instance.returnNumStringZero(Win_ord_OutWare_Scan_Insert.OrderQty);
                            Win_ord_OutWare_Scan_Insert.OutRealQty = Lib.Instance.returnNumStringZero(Win_ord_OutWare_Scan_Insert.OutRealQty);
                            Win_ord_OutWare_Scan_Insert.OutQty = Lib.Instance.returnNumStringZero(Win_ord_OutWare_Scan_Insert.OutQty);
                            Win_ord_OutWare_Scan_Insert.OutSumQty = Lib.Instance.returnNumStringZero(Win_ord_OutWare_Scan_Insert.OutSumQty);
                            Win_ord_OutWare_Scan_Insert.UnitPrice = Lib.Instance.returnNumStringOne(Win_ord_OutWare_Scan_Insert.UnitPrice);

                            dgdOutware.Items.Add(Win_ord_OutWare_Scan_Insert);

                            //MessageBox.Show(Win_ord_OutWare_Scan_Insert.TOLocID);
                        }

                        tbkCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - FillGrid : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }
        #endregion

        #region Sub조회
        private void FillGridSub(string OutwareID)
        {
            try
            {
                if (dgdOutwareSub.Items.Count > 0)
                {
                    dgdOutwareSub.Items.Clear();
                }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("OutwareID", OutwareID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOutwareSubGroup_OFFICE", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow item in drc)
                        {
                            var Win_ord_OutWareSub_Scan_Insert = new Win_ord_OutWare_Scan_Sub_CodeView()
                            {
                                OutwareID = item["OutwareID"].ToString(),
                                OutSubSeq = item["OutSubSeq"].ToString(),
                                LabelID = item["LabelID"].ToString(),
                                LabelGubun = item["LabelGubun"].ToString(),
                                LabelGubunName = item["LabelGubunName"].ToString(),

                                OutQty = item["OutQty"].ToString(),
                                OutCnt = item["OutCnt"].ToString(),
                                OutRoll = item["OutRoll"].ToString(),
                                LotNo = item["LotNo"].ToString(),
                                Weight = item["Weight"].ToString(),

                                UnitPrice = item["UnitPrice"].ToString(),
                                Vat_IND_YN = item["Vat_IND_YN"].ToString(),
                                Orderseq = item["Orderseq"].ToString(),
                                Amount = item["Amount"].ToString(),
                                CustomBoxID = item["CustomBoxID"].ToString(),

                                FromLocID = item["FromLocID"].ToString(),
                                TOLocID = item["TOLocID"].ToString(),
                                UnitClss = item["UnitClss"].ToString(),
                                ArticleID = item["ArticleID"].ToString(),
                                Article = item["Article"].ToString(),

                                OutClss = item["OutClss"].ToString(),
                                Gubun = item["Gubun"].ToString(),
                                DefectID = item["DefectID"].ToString(),
                                DefectName = item["DefectName"].ToString(),

                                DeleteYN = "N",

                                OutRealQty = item["OutRealQty"].ToString()

                            };

                            Win_ord_OutWareSub_Scan_Insert.OutQty = lib.returnNumStringZero(Win_ord_OutWareSub_Scan_Insert.OutQty);
                            dgdOutwareSub.Items.Add(Win_ord_OutWareSub_Scan_Insert);
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - FillGridSub : " + ee.ToString());
            }
        }

        #endregion Sub조회

        #region 저장
        private bool SaveData(string strFlag)
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    string remarkTxt = "사무실에서 출고";

                    if ((cboOutClss.SelectedValue.ToString() == "11" || cboOutClss.SelectedValue.ToString() == "17"))
                    {
                        remarkTxt = "사무실에서 출고반품";
                    }
                    

                    #region 추가

                    if (strFlag == "I")
                    {
                        double cnt = 0;
                       

                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("OrderID", txtOrderID.Text);           //관리번호
                        sqlParameter.Add("CompanyID", "0001");                  //본인회사
                        sqlParameter.Add("OutSeq", "");
                        sqlParameter.Add("OutwareNo", "");
                        sqlParameter.Add("OutClss", cboOutClss.SelectedValue.ToString());

                        sqlParameter.Add("CustomID", txtKCustom.Tag != null ? txtKCustom.Tag.ToString() : "");
                        sqlParameter.Add("BuyerDirectYN", "Y");
                        sqlParameter.Add("WorkID", "0001");                 //지금은 샤프트가공 1개 뿐
                        sqlParameter.Add("ExchRate", 0);
                        sqlParameter.Add("UnitPriceClss", "0");

                        sqlParameter.Add("InsStuffInYN", "N");              //동시입고여부
                        //sqlParameter.Add("OutcustomID", txtBuyerName.Tag != null ? txtBuyerName.Tag.ToString() : "");                //납품거래처
                        sqlParameter.Add("OutcustomID", txtOutCustom.Tag != null ? txtOutCustom.Tag.ToString() : "");                //20210526
                        sqlParameter.Add("Outcustom", txtOutCustom.Text);
                        sqlParameter.Add("LossRate", 0);
                        sqlParameter.Add("LossQty", 0);

                        sqlParameter.Add("OutRoll", txtOutRoll.Text.Equals("") == true ? 0 : Convert.ToInt32(txtOutRoll.Text.Replace(",", "")));
                        sqlParameter.Add("OutQty", txtOutQty.Text.Equals("") == true ? 0 : ConvertDouble(txtOutQty.Text.Replace(",", "")));
                        sqlParameter.Add("OutRealQty", ConvertDouble(txtOutQty.Text.Replace(",", ""))); //실출고량인데, = outQty
                        sqlParameter.Add("OutDate", dtpOutDate.SelectedDate != null ?  dtpOutDate.SelectedDate.Value.ToString("yyyyMMdd") : DateTime.Today.ToString("yyyyMMdd"));
                        sqlParameter.Add("ResultDate", dtpOutDate.SelectedDate != null ? dtpOutDate.SelectedDate.Value.ToString("yyyyMMdd") : DateTime.Today.ToString("yyyyMMdd"));
                        sqlParameter.Add("Remark", txtRemark.Text.Equals("") ? remarkTxt : txtRemark.Text);
                        sqlParameter.Add("OutType", "3");                //스캔출고형태가 3번
                        sqlParameter.Add("OutSubType", "");              //안쓰니까 일단 빈값??
                        sqlParameter.Add("Amount", 0);                   //안쓰니까 일단 빈값??
                        sqlParameter.Add("VatAmount", 0);                //안쓰니까 일단 빈값??

                        sqlParameter.Add("VatINDYN", "Y");                //안쓰니까 일단 빈값??
                        sqlParameter.Add("FromLocID", cboFromLoc.SelectedValue != null ? cboFromLoc.SelectedValue.ToString() : "");
                        sqlParameter.Add("ToLocID", cboToLoc.SelectedValue != null ? cboToLoc.SelectedValue.ToString() : "");
                        sqlParameter.Add("UnitClss", 0);
                        sqlParameter.Add("ArticleID", txtArticleID_InGroupBox.Text != null ? txtArticleID_InGroupBox.Text : "");
                        sqlParameter.Add("DvlyCustomID", txtBuyerName.Tag == null ? "" : txtBuyerName.Tag.ToString()); //20210526

                        sqlParameter.Add("UserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Outware_iOutware";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "OutwareNo";      //OutwareNo = OutwareID
                        pro1.OutputLength = "12";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "OutwareNo")
                                {
                                    sGetID = kv.value;

                                    GetKey = kv.value;

                                    Prolist.RemoveAt(0);
                                    ListParameter.Clear();
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                        }


                        //sub그리드 아이템 수만큼 반복되어야 하므로
                        for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                        {
                            var OutwareSub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("OutwareID", GetKey);
                            sqlParameter.Add("OrderID", txtOrderID.Text);
                            sqlParameter.Add("OutSeq", "");
                            sqlParameter.Add("OutSubSeq", i + 1);
                            sqlParameter.Add("OrderSeq", tgnMoveByQty.IsChecked == true ? "1" : OutwareSub.Orderseq);

                            sqlParameter.Add("LineSeq", 0);
                            sqlParameter.Add("LineSubSeq", 0);
                            sqlParameter.Add("RollSeq", i);
                            sqlParameter.Add("LabelID", OutwareSub.LabelID);
                            sqlParameter.Add("LabelGubun", "2");        //박스라벨출고는 2번

                            sqlParameter.Add("LotNo", "0");
                            sqlParameter.Add("Gubun", "");              //용도를 몰라서 빈값
                            sqlParameter.Add("StuffQty", 0);
                            sqlParameter.Add("OutQty", lib.RemoveComma(OutwareSub.OutQty,0));
                            sqlParameter.Add("OutRoll", 1); // 하나당 박스 1개로 처리 하니, 1로 저장한다고 함

                            sqlParameter.Add("UnitPrice", lib.RemoveComma(OutwareSub.UnitPrice,0));
                            sqlParameter.Add("CustomBoxID", "");
                            sqlParameter.Add("DefectID", "");           //결함사유라는데.. 빈값으로 
                            sqlParameter.Add("BoxID", OutwareSub.LabelID);
                            sqlParameter.Add("ArticleID", OutwareSub.ArticleID);

                            sqlParameter.Add("UserID", MainWindow.CurrentUser);


                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_Outware_iOutwareSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "REQ_ID";
                            pro2.OutputLength = "10";

                            //cnt += (Double.Parse(OutwareSub.OutQty.Replace(",", "")) * Double.Parse(OutwareSub.UnitPrice.Replace(",", "")));

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);

                        }
                        //ListParameter[0]["Amount"] = cnt.ToString();
                        //ListParameter[0]["VatAmount"] = (cnt * 0.1).ToString();
                    }

                    #endregion   추가

                    #region 수정

                    else if (strFlag == "U")
                    {      // 1. outware 는 [xp_Outware_uOutware] : outware 수정 후 outwaresub, stuffin 도 같이 지우는 프로시저 
                           // 2. outwaresub 다시 등록
                           // 3. stuffin 다시 등록
                           // ssw 20210616 파라미터 값 넘기게 수정 (vatYN, Amount, va tAmount, UnitPrice, OutQty)
                        double cnt = 0;

                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("OutwareID", txtOutwareID.Text);
                        sqlParameter.Add("OrderID", txtOrderID.Text);
                        sqlParameter.Add("CompanyID", "0001");
                        sqlParameter.Add("OutClss", cboOutClss.SelectedValue.ToString());
                        sqlParameter.Add("CustomID", txtKCustom.Tag != null ? txtKCustom.Tag.ToString() : "");

                        sqlParameter.Add("BuyerDirectYN", "Y");
                        sqlParameter.Add("WorkID", "0001");
                        sqlParameter.Add("ExchRate", 0);
                        sqlParameter.Add("UnitPriceClss", "0");
                        sqlParameter.Add("InsStuffInYN", "N");

                        //sqlParameter.Add("OutcustomID", txtBuyerName.Tag != null ? txtBuyerName.Tag.ToString() : "");
                        sqlParameter.Add("OutcustomID", txtOutCustom.Tag != null ? txtOutCustom.Tag.ToString() : ""); //20210526
                        sqlParameter.Add("Outcustom", txtOutCustom.Text);
                        sqlParameter.Add("LossRate", 0);
                        sqlParameter.Add("LossQty", 0);
                        sqlParameter.Add("OutRoll", Convert.ToInt32(txtOutRoll.Text.Replace(",", "")));

                        sqlParameter.Add("OutQty", txtOutQty.Text.Replace(",", ""));
                        sqlParameter.Add("OutRealQty", txtOutQty.Text.Replace(",", ""));
                        sqlParameter.Add("OutDate", dtpOutDate.SelectedDate != null ? dtpOutDate.SelectedDate.Value.ToString("yyyyMMdd") : DateTime.Today.ToString("yyyyMMdd"));
                        sqlParameter.Add("ResultDate", dtpOutDate.SelectedDate != null ? dtpOutDate.SelectedDate.Value.ToString("yyyyMMdd") : DateTime.Today.ToString("yyyyMMdd"));
                        sqlParameter.Add("Remark", txtRemark.Text.Equals("") ? remarkTxt : txtRemark.Text);

                        sqlParameter.Add("OutType", "3");
                        sqlParameter.Add("OutSubType", "");
                        sqlParameter.Add("Amount", 0);
                        sqlParameter.Add("VatAmount", 0);
                        sqlParameter.Add("VatINDYN", 'Y');

                        sqlParameter.Add("FromLocID", cboFromLoc.SelectedValue.ToString());
                        sqlParameter.Add("ToLocID", cboToLoc.SelectedValue.ToString());
                        sqlParameter.Add("UnitClss", 0);
                        sqlParameter.Add("ArticleID", txtArticleID_InGroupBox.Text != null ? txtArticleID_InGroupBox.Text : "");
                        sqlParameter.Add("DvlyCustomID", txtBuyerName.Tag == null ? "" : txtBuyerName.Tag.ToString()); //20210526

                        sqlParameter.Add("UserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Outware_uOutware";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "";
                        pro1.OutputLength = "15";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        //sub그리드 아이템 수만큼 반복되어야 하므로 
                        for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                        {
                            var OutwareSub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("OutwareID", txtOutwareID.Text);
                            sqlParameter.Add("OrderID", txtOrderID.Text);
                            sqlParameter.Add("OutSeq", "");
                            sqlParameter.Add("OutSubSeq", i + 1);
                            sqlParameter.Add("OrderSeq", OutwareSub.Orderseq);

                            sqlParameter.Add("LineSeq", 0);
                            sqlParameter.Add("LineSubSeq", 0);
                            sqlParameter.Add("RollSeq", i);
                            sqlParameter.Add("LabelID", OutwareSub.LabelID);
                            sqlParameter.Add("LabelGubun", "2");        //박스라벨출고는 2번 3번은 로트아이디인 듯

                            sqlParameter.Add("LotNo", "0");
                            sqlParameter.Add("Gubun", "");              //용도를 몰라서 빈값
                            sqlParameter.Add("StuffQty", 0);
                            sqlParameter.Add("OutQty", OutwareSub.OutQty.Replace(",", ""));
                            sqlParameter.Add("OutRoll", 1); // 하나당 박스 1개로 처리 하니, 1로 저장한다고 함

                            sqlParameter.Add("UnitPrice", OutwareSub.UnitPrice.Replace(",", ""));
                            sqlParameter.Add("CustomBoxID", "");
                            sqlParameter.Add("DefectID", "");           //결함사유라는데.. 빈값으로 
                            sqlParameter.Add("BoxID", OutwareSub.LabelID);
                            sqlParameter.Add("ArticleID", OutwareSub.ArticleID);

                            sqlParameter.Add("UserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_Outware_iOutwareSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "REQ_ID";
                            pro2.OutputLength = "10";

                            cnt += (Double.Parse(OutwareSub.OutQty.Replace(",", "")) * Double.Parse(OutwareSub.UnitPrice.Replace(",", "")));

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        ListParameter[0]["Amount"] = cnt.ToString();
                        ListParameter[0]["VatAmount"] = (cnt * 0.1).ToString();
                    }

                    #endregion 수정

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                    if (Confirm[0] != "success")
                    {
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                        flag = false;
                        //return false;
                    }
                    else
                    {
                        //MessageBox.Show("성공");
                        flag = true;
                    }

                }
                else
                {
                    btnAdd_Click(null, null);
                    txtScanData.Focus();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - SaveData : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }

        #endregion 저장

        #region 데이터 체크
        // 그룹박스 데이터 기입체크
        private bool CheckData()
        {
            try
            {
                if (txtOrderID.Text == "")
                {
                    MessageBox.Show("관리번호를 반드시 입력하세요.");
                    return false;
                }

                if (txtKCustom.Text == "")
                //if (lib.IsNullOrWhiteSpace(txtKCustom.Text) == true)
                {
                    MessageBox.Show("거래처를 반드시 입력하세요.");
                    return false;
                }
                //if (lib.IsNumOrAnother(txtOutRoll.Text) == false)
                //{
                //    MessageBox.Show("출고박스 수량은 반드시 숫자로 입력하세요.");
                //    return false;
                //}
                //if (lib.IsNumOrAnother(txtOutQty.Text) == false)
                //{
                //    MessageBox.Show("출고 수량은 반드시 숫자로 입력하세요.");
                //    return false;
                //}
                if (cboOutClss.SelectedIndex < 0)
                {
                    MessageBox.Show("출고구분은 반드시 선택하세요.");
                    return false;
                }
                if (cboFromLoc.SelectedIndex < 0)
                {
                    MessageBox.Show("전 창고는 반드시 선택하세요.");
                    return false;
                }
                if (dgdOutwareSub.Items.Count == 0)
                {
                    MessageBox.Show("스캔된 라벨 정보가 없습니다.");
                    return false;
                }
                #region ...
                ////if (strFlag == "I" )
                ////{
                ////    if(cboOutClss.SelectedValue.ToString() != "11" && cboOutClss.SelectedValue.ToString() != "17")
                ////    {
                ////        for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                ////        {
                ////            var OutwareSub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;
                ////            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                ////            sqlParameter.Add("LabelID", OutwareSub.LabelID);
                ////            sqlParameter.Add("Qty", OutwareSub.OutQty.Replace(",", ""));
                ////            sqlParameter.Add("ArticleID", txtArticleID_InGroupBox.Text != null ? txtArticleID_InGroupBox.Text : "");
                ////            DataTable dt = DataStore.Instance.ProcedureToDataSet("xp_Outware_chkiOutware", sqlParameter, false).Tables[0];
                ////            if (dt.Rows[0][0].Equals("F"))
                ////            {
                ////                MessageBox.Show("재고에 있는 수량보다 많은 수량이 입력되었습니다.");
                ////                return false;
                ////            }
                ////        }
                ////    }
                ////}
                ////else
                ////{
                ////    for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                ////    {
                ////        var OutwareSub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;
                ////        CheckStock(OutwareSub);
                ////    }
                ////}
                #endregion

                if (strFlag == "I" && tgnMoveByID.IsChecked == true)
                {
                    for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                    {
                        var OutwareSub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Add("LabelID", OutwareSub.LabelID);
                        sqlParameter.Add("Qty", OutwareSub.OutQty.Replace(",", ""));
                        sqlParameter.Add("ArticleID", txtArticleID_InGroupBox.Text != null ? txtArticleID_InGroupBox.Text : "");
                        DataTable dt = DataStore.Instance.ProcedureToDataSet("xp_Outware_chkiOutware", sqlParameter, false).Tables[0];
                        if (dt.Rows[0][0].Equals("F"))
                        {
                            MessageBox.Show("재고에 있는 수량보다 많은 수량이 입력되었습니다.");
                            return false;
                        }
                    }
                }


                return true;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - CheckData : " + ee.ToString());
                return false;
            }
        }
        #endregion

        #region 삭제
        private bool DeleteData(string OutwareID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OutwareID", OutwareID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Outware_dOutware", sqlParameter, false);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - DeleteData : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }



            return flag;
        }


        #endregion 삭제

        //라벨스캔 토글버튼 클릭
        private void btnCustomerLabelScanYN_Click(object sender, RoutedEventArgs e)
        {
            //안쓸 듯
        }

        //서브 데이터 그리드 삭제컬럼 버튼 클릭
        private void dgdOutwareSub_btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var OutwareSub = dgdOutwareSub.SelectedItem as Win_ord_OutWare_Scan_Sub_CodeView;
                if (OutwareSub != null)
                {
                    dgdOutwareSub.Items.Remove(OutwareSub);
                }

                SumScanQty();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutwareSub_btnDelete_Click : " + ee.ToString());
            }
        }

        // 관리번호 기반_ 항목 뿌리기 작업.
        private void OrderID_OtherSearch(string OrderID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OrderID", OrderID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOrderOne", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        return;
                    }
                    else
                    {
                        DataRow DR = dt.Rows[0];
                        txtKCustom.Text = DR["KCustom"].ToString(); //20210526
                        txtKCustom.Tag = DR["CustomID"].ToString();
                        txtBuyerName.Text = DR["KCustom"].ToString();
                        txtBuyerName.Tag = DR["CustomID"].ToString();
                        txtOutCustom.Text = DR["KCustom"].ToString();
                        txtOutCustom.Tag = DR["CustomID"].ToString();
                        //if (txtKCustom.Text == string.Empty) { txtKCustom.Text = DR["KCustom"].ToString(); }
                        //if (txtKCustom.Tag == null) { txtKCustom.Tag = DR["CustomID"].ToString(); }
                        //if (txtBuyerName.Text == string.Empty) { txtBuyerName.Text = DR["KCustom"].ToString(); }
                        //if (txtBuyerName.Tag == null) { txtBuyerName.Tag = DR["CustomID"].ToString(); }
                        //if (txtOutCustom.Text == string.Empty) { txtOutCustom.Text = DR["KCustom"].ToString(); }
                        //if (txtOutCustom.Tag == null) { txtOutCustom.Tag = DR["CustomID"].ToString(); }

                        if (txtArticle_InGroupBox.Text == string.Empty) { txtArticle_InGroupBox.Text = DR["Article"].ToString(); }
                        if (txtArticle_InGroupBox.Tag == null)
                        {
                            txtArticle_InGroupBox.Tag = DR["ArticleID"].ToString();
                            txtArticleID_InGroupBox.Text = DR["ArticleID"].ToString();
                        }

                        if (txtArticleID_InGroupBox.Text == string.Empty)
                        {
                            txtArticleID_InGroupBox.Text = DR["ArticleID"].ToString();
                        }

                        if (txtOutQty.Text == string.Empty) { txtOutQty.Text = DR["OrderQty"].ToString(); }

                        txtBuyerModel.Text = DR["BuyerModel"].ToString();
                        txtBuyerModel.Tag = DR["BuyerModelID"].ToString();
                        txtBuyerArticleNo.Text = DR["BuyerArticleNo"].ToString();

                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - OrderID_OtherSearch : " + ee.ToString());
            }
        }

        // 실제 엑셀작업 스타트.
        // 2021-05-31
        private void PrintWork(bool previewYN)
        {
            Lib lib2 = new Lib();
            try
            {
                if (lstOutwarePrint.Count == 0 )
                {
                   MessageBox.Show("인쇄할 거래명세표를 선택하세요.");
                   lib2 = null;
                   return;
                }

                excelapp = new Microsoft.Office.Interop.Excel.Application();

                string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\거래명세서.xlsx";
                workbook = excelapp.Workbooks.Add(MyBookPath);
                worksheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                int page = (lstOutwarePrint.Count() / 14) + 1;

                if (lstOutwarePrint.Count() % 14 == 0)
                {
                    page = (lstOutwarePrint.Count() / 14);
                }
                int copyLine = 1;

                string str_OutDate = lstOutwarePrint[0].OutDate;                              // 거래일자.
                string str_BuyerName = lstOutwarePrint[0].BuyerName;                          // 공급받는 상호명.
                string str_Article = lstOutwarePrint[0].Article;                              // 품명
                string str_BuyerArticleNo = lstOutwarePrint[0].BuyerArticleNo;                // 품목
                //string str_UnitClssName = OutwareInfo.UnitClssName;                    // 단위
                //string str_OutQty = OutwareInfo.OutQty;                                // 수량
                //string str_UnitPrice = OutwareInfo.UnitPrice;                          // 단가
                //string str_SupPrice = OutwareInfo.Amount;                           // 공급가액
                //string str_Bugase = OutwareInfo.VatAmount;                              // 부가세
                //string str_Price = (Double.Parse(str_SupPrice) + Double.Parse(str_Bugase)).ToString();


                workrange = worksheet.get_Range("C4", "H4");    //거래일자
                workrange.Value2 = str_OutDate.Substring(0, 4) + "-" + str_OutDate.Substring(5, 2) + "-" + str_OutDate.Substring(8, 2);
                //workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                //workrange.Font.Size = 11;

                workrange = worksheet.get_Range("W7", "AB8");    //공급받는 상호명
                workrange.Value2 = str_BuyerName;
                //workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                //workrange.Font.Size = 10;

                workrange = worksheet.get_Range("W9", "AH10");    //공급받는 사업장 주소
                workrange.Value2 = lstOutwarePrint[0].Buyer_Address1 + lstOutwarePrint[0].Buyer_Address2 + lstOutwarePrint[0].Buyer_Address3;
                //workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                //workrange.Font.Size = 10;

                workrange = worksheet.get_Range("AE7", "AH8");    //공급받는 대표자 성명
                workrange.Value2 = lstOutwarePrint[0].Buyer_Chief;


                workrange = worksheet.get_Range("W5", "AH6");    //공급받는 사업자 등록번호
                workrange.Value2 = lstOutwarePrint[0].CustomNo;

                workrange = worksheet.get_Range("W11", "AB12");    // 업테 
                workrange.Value2 = lstOutwarePrint[0].Condition;

                workrange = worksheet.get_Range("AE11", "AH12");    // 종목
                workrange.Value2 = lstOutwarePrint[0].Category;
                //workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                //workrange.Font.Size = 10;
                int cnt = 0;

                Double totalUnitPrice = 0;
                Double totalAmount = 0;
                Double totalVatAmount = 0;


                foreach (Win_ord_OutWare_Scan_CodeView i in lstOutwarePrint)
                //for (int i = 0; i < lstOutwarePrint.Count; i++)
                {

                    string str_UnitClssName = i.UnitClssName;                    // 단위
                    string str_OutQty = i.OutQty;                                // 수량
                    string str_UnitPrice = i.UnitPrice;                          // 단가
                    string str_SupPrice = i.Amount;                           // 공급가액
                    string str_Bugase = i.VatAmount;                              // 부가세
                    string str_Price = (Double.Parse(str_SupPrice) + Double.Parse(str_Bugase)).ToString();

                    workrange = worksheet.get_Range("C" + (14 + cnt % 14).ToString(), "D" + (14 + cnt % 14).ToString());    //순번
                    workrange.Value2 = cnt + 1;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 9;


                    workrange = worksheet.get_Range("E" + (14 + cnt % 14).ToString(), "L" + (14 + cnt % 14).ToString());    //품명.
                    workrange.Value2 = str_BuyerArticleNo;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 9;

                    workrange = worksheet.get_Range("M" + (14 + cnt % 14).ToString(), "O" + (14 + cnt % 14).ToString());    //단위
                    workrange.Value2 = str_UnitClssName;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 9;

                    workrange = worksheet.get_Range("P" + (14 + cnt % 14).ToString(), "R" + (14 + cnt % 14).ToString());    //수량.
                    workrange.Value2 = str_OutQty;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 9;

                    workrange = worksheet.get_Range("S" + (14 + cnt % 14).ToString(), "V" + (14 + cnt % 14).ToString());    //단가.
                    workrange.Value2 = str_UnitPrice;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 9;

                    workrange = worksheet.get_Range("W" + (14 + cnt % 14).ToString(), "Z" + (14 + cnt % 14).ToString());    //공급가액.
                    workrange.Value2 = str_SupPrice;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 9;

                    workrange = worksheet.get_Range("AA" + (14 + cnt % 14).ToString(), "AD" + (14 + cnt % 14).ToString());    //부가세.
                    workrange.Value2 = str_Bugase;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 9;

                    workrange = worksheet.get_Range("AE" + (14 + cnt % 14).ToString(), "AH" + (14 + cnt % 14).ToString());    //금액.
                    workrange.Value2 = str_Price;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 9;

                    cnt += 1;

                    totalUnitPrice += Double.Parse(str_UnitPrice); //단가
                    totalAmount += Double.Parse(str_SupPrice); //공급가액
                    totalVatAmount += Double.Parse(str_Bugase); //부가세
                    //tatalPrice += Double.Parse(str_Price);

                    if (cnt % 14 == 0)
                    {
                        worksheet.Select();
                        worksheet.UsedRange.EntireRow.Copy();
                        pastesheet.Select();
                        workrange = pastesheet.Rows[copyLine];
                        workrange.Select();
                        pastesheet.Paste();
                        copyLine += 59;
                    }

                }
                if (cnt % 14 != 0)
                {
                    while (cnt > 14 && cnt % 14 != 0)
                    {

                        workrange = worksheet.get_Range("C" + (14 + cnt % 14).ToString(), "D" + (14 + cnt % 14).ToString());    //순번
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 9;


                        workrange = worksheet.get_Range("E" + (14 + cnt % 14).ToString(), "L" + (14 + cnt % 14).ToString());    //품명.
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 9;

                        workrange = worksheet.get_Range("M" + (14 + cnt % 14).ToString(), "O" + (14 + cnt % 14).ToString());    //단위
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 9;

                        workrange = worksheet.get_Range("P" + (14 + cnt % 14).ToString(), "R" + (14 + cnt % 14).ToString());    //수량.
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 9;

                        workrange = worksheet.get_Range("S" + (14 + cnt % 14).ToString(), "V" + (14 + cnt % 14).ToString());    //단가.
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 9;

                        workrange = worksheet.get_Range("W" + (14 + cnt % 14).ToString(), "Z" + (14 + cnt % 14).ToString());    //공급가액.
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 9;

                        workrange = worksheet.get_Range("AA" + (14 + cnt % 14).ToString(), "AD" + (14 + cnt % 14).ToString());    //부가세.
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 9;

                        workrange = worksheet.get_Range("AE" + (14 + cnt % 14).ToString(), "AH" + (14 + cnt % 14).ToString());    //금액.
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 9;

                        cnt += 1;
                    }

                    worksheet.Select();
                    worksheet.UsedRange.EntireRow.Copy();
                    pastesheet.Select();
                    //workrange = pastesheet.Cells[copyLine, 1];
                    workrange = pastesheet.Rows[copyLine];
                    workrange.Select();
                    pastesheet.Paste();
                }

                for (int i = 0; i < page; i++)
                {
                    workrange = pastesheet.get_Range("G" + (28 + i * 59).ToString(), "L" + (28 + i * 59).ToString());    //공급가액
                    workrange.Value2 = totalAmount.ToString();
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("Q" + (28 + i * 59).ToString(), "W" + (28 + i * 59).ToString());    //부가세
                    workrange.Value2 = totalVatAmount.ToString();
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AB" + (28 + i * 59).ToString(), "AH" + (28 + i * 59).ToString());    //합계
                    workrange.Value2 = (totalAmount + totalVatAmount).ToString();
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("G" + (57 + i * 59).ToString(), "L" + (57 + i * 59).ToString());    //공급가액
                    workrange.Value2 = totalAmount.ToString();
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("Q" + (57 + i * 59).ToString(), "W" + (57 + i * 59).ToString());    //부가세
                    workrange.Value2 = totalVatAmount.ToString();
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AB" + (57 + i * 59).ToString(), "AH" + (57 + i * 59).ToString());    //합계
                    workrange.Value2 = (totalAmount + totalVatAmount).ToString();
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;
                }

                excelapp.Visible = true;
                msg.Hide();


                if (previewYN == true)
                {
                    pastesheet.PrintPreview();
                }
                else
                {
                    pastesheet.PrintOutEx();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 = PrintWork : " + ee.ToString());
            }
            
            lib2.ReleaseExcelObject(workbook);
            lib2.ReleaseExcelObject(worksheet);
            lib2.ReleaseExcelObject(pastesheet);
            lib2.ReleaseExcelObject(excelapp);
            lib2 = null;
        }

        // 거래명세표 인쇄시 공급자 정보 구해오기(삼주테크)
        private DataTable Fill_DS_CompanyInfo()
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nChkCompany", 0);
                sqlParameter.Add("sCompanyID", "");
                sqlParameter.Add("sKCompany", "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Info_GetCompanyInfo", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable ddt = null;
                    ddt = ds.Tables[0];

                    if (ddt.Rows.Count == 0)
                    {
                        MessageBox.Show("공급자 정보를 구하지 못했습니다.");
                        return ddt;
                    }
                    else
                    {
                        return ddt;
                    }
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - Fill_DS_CompanyInfo : " + ee.ToString());
                return null;
            }
        }


        //추가, 수정일 때 
        private void CanBtnControl()
        {
            btnAdd.IsEnabled = false;               //추가
            btnUpdate.IsEnabled = false;            //수정
            btnDelete.IsEnabled = false;            //삭제
            btnClose.IsEnabled = true;              //닫기
            btnSearch.IsEnabled = false;            //검색
            btnSave.Visibility = Visibility.Visible;             //저장
            btnCancel.Visibility = Visibility.Visible;             //취소
            btnExcel.IsEnabled = false;             //엑셀
            btnPrint.IsEnabled = false;             //인쇄

            txtBuyerModel.IsHitTestVisible = false;  //차종은 땡겨오니까
            txtOutwareID.IsHitTestVisible = false;   //출고번호는 자동으로 생성되니까
            txtScanData.IsEnabled = true;           //바코드 스캔
            EventLabel.Visibility = Visibility.Visible; //자료입력중
            grbOutwareDetailBox.IsEnabled = true;       //DataContext Box
            dgdOutware.IsHitTestVisible = false;        //데이터그리드 클릭 안되게

            tgnMoveByID.IsHitTestVisible = true;
            tgnMoveByQty.IsHitTestVisible = true;

        }
        //저장, 취소일 때
        private void CantBtnControl()
        {
            btnAdd.IsEnabled = true;               //추가
            btnUpdate.IsEnabled = true;            //수정
            btnDelete.IsEnabled = true;            //삭제
            btnClose.IsEnabled = true;             //닫기
            btnSearch.IsEnabled = true;            //검색
            btnSave.Visibility = Visibility.Hidden;             //저장
            btnCancel.Visibility = Visibility.Hidden;             //취소
            btnExcel.IsEnabled = true;             //엑셀
            btnPrint.IsEnabled = true;             //인쇄

            txtBuyerModel.IsHitTestVisible = false;  //차종은 땡겨오니까
            txtScanData.IsEnabled = false;         //바코드 스캔
            EventLabel.Visibility = Visibility.Hidden; //자료입력중
            grbOutwareDetailBox.IsEnabled = false;       //DataContext Box
            dgdOutware.IsHitTestVisible = true;        //데이터그리드 클릭되게

            tgnMoveByID.IsHitTestVisible = false;
            tgnMoveByQty.IsHitTestVisible = false;

        }

        private void TextBoxClear()
        {
            txtOrderID.Text = string.Empty;
            txtArticleID_InGroupBox.Text = string.Empty;
            txtArticle_InGroupBox.Text = string.Empty;
            txtArticle_InGroupBox.Tag = null;
            cboOutClss.SelectedIndex = 0;
            txtBuyerModel.Text = string.Empty;
            txtOutwareID.Text = string.Empty;
            txtOutRoll.Text = string.Empty;
            txtOutQty.Text = string.Empty;
            cboFromLoc.SelectedIndex = 0;
            cboToLoc.SelectedIndex = 0;
            txtKCustom.Text = string.Empty;
            txtKCustom.Tag = null;
            txtBuyerName.Text = string.Empty;
            txtBuyerName.Tag = null;
            txtRemark.Text = string.Empty;
            txtOutCustom.Text = string.Empty;

        }

        private void SumScanQty()
        {
            try
            {
                int OutRoll = 0;
                double OutQty = 0;

                OutRoll = dgdOutwareSub.Items.Count;

                for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                {
                    var label = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;
                    if (label.OutQty != null)
                        OutQty += ConvertDouble(label.OutQty.ToString());
                }

                txtOutRoll.Text = stringFormatN0(OutRoll);
                txtOutQty.Text = stringFormatN0(OutQty);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - SumScanQty : " + ee.ToString());
            }
        }

        // 천자리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        //더블로 형식 변환
        private double ConvertDouble(string str)
        {
            double result = 0;
            double chkDouble = 0;

            try
            {
                if (!str.Trim().Equals(""))
                {
                    str = str.Trim().Replace(",", "");

                    if (double.TryParse(str, out chkDouble) == true)
                    {
                        result = double.Parse(str);
                    }
                }
                return result;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ConvertDouble : " + ee.ToString());
                return result;
            }
        }

        // 데이터피커 포맷으로 변경
        private string DatePickerFormat(string str)
        {
            string result = "";

            try
            {
                if (str.Length == 8)
                {
                    if (!str.Trim().Equals(""))
                    {
                        result = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
                    }
                }

                return result;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DatePickerFormat : " + ee.ToString());
                return result;
            }
        }

        //관리번호 숫자만 입력
        private void txtOrderID_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                lib.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtOrderID_PreviewTextInput : " + ee.ToString());
            }
        }

        //박스에 숫자만 입력
        private void txtOutRoll_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                lib.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtOutRoll_PreviewTextInput : " + ee.ToString());
            }
        }

        //수량에 숫자만 입력
        private void txtOutQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                lib.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtOutQty_PreviewTextInput : " + ee.ToString());
            }
        }

        //검색조건 - 관리번호에 숫자만 입력
        private void txtRadioOptionNum_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                lib.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtRadioOptionNum_PreviewTextInput : " + ee.ToString());
            }
        }


        private void chkReq_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            var Outware = chkSender.DataContext as Win_ord_OutWare_Scan_CodeView;

            if (Outware != null)
            {
                if (chkSender.IsChecked == true)
                {
                    Outware.Chk = true;

                    if (lstOutwarePrint.Contains(Outware) == false)
                    {
                        lstOutwarePrint.Add(Outware);
                    }
                }
                else
                {
                    Outware.Chk = false;

                    if (lstOutwarePrint.Contains(Outware) == true)
                    {
                        lstOutwarePrint.Remove(Outware);
                    }
                }

            }
        }

        private void txtQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (EventStatus == true)
            {
                //System.Windows.Controls.TextBox test = new TextBox();
                //test = (TextBox)sender;
                //string realQtyString = test.Text;

                var ViewReceiver = dgdOutwareSub.CurrentCell.Item as Win_ord_OutWare_Scan_Sub_CodeView;  //선택 줄.
                
                if (ViewReceiver != null)   // 널이 아니라면,
                {
                    try
                    {
                        if (e.Key == Key.Enter)
                        {
                            e.Handled = true;
                            int point = dgdOutwareSub.Items.IndexOf(ViewReceiver);

                            double realQty = Double.Parse(ViewReceiver.OutRealQty);
                            double beforeQty = Double.Parse(ViewReceiver.OutQty);

                            DataGridCell tempOutQtyCell = lib.GetCell(point, 4, dgdOutwareSub);
                            TextBox tempOutQtyTB = lib.GetVisualChild<TextBox>(tempOutQtyCell);

                            if ((cboOutClss.SelectedValue.ToString() == "11" || cboOutClss.SelectedValue.ToString() == "17"))
                            {
                                txtOutQty.Text = (Double.Parse(txtOutQty.Text) - beforeQty + Double.Parse(tempOutQtyTB.Text)).ToString();
                                ViewReceiver.OutQty = tempOutQtyTB.Text;
                            }
                            else if (Double.Parse(tempOutQtyTB.Text) > realQty)
                            {
                                MessageBox.Show("입력하신 수량이 재고수량보다 많습니다. 최대 입력가능 수량은 [ " + ViewReceiver.OutRealQty + " ]입니다.");
                                //tempOutQtyTB.Text = beforeQty.ToString();
                            }
                            else
                            {
                                txtOutQty.Text = (Double.Parse(txtOutQty.Text) - beforeQty + Double.Parse(tempOutQtyTB.Text)).ToString();
                                ViewReceiver.OutQty = tempOutQtyTB.Text;
                            }
                        }
                    }
                    catch (Exception ee)
                    {
                        MessageBox.Show("오류 시점 - 수량 입력후 엔터키" + ee.ToString());
                    }
                }
            }
        }

        private void dgdOutwareSubRequest_MouseClick(object sender, MouseButtonEventArgs e)
        {
            // 추가 상태로 들어와야 하고
            if (EventStatus == true)
            {
                var ViewReceiver = dgdOutwareSub.CurrentCell.Item as Win_ord_OutWare_Scan_Sub_CodeView;   //dgdOutRequest.SelectedItem as Win_out_OutwareReq_U_View;
                if (ViewReceiver != null)
                {
                    string eventer = ((DataGridCell)sender).Column.Header.ToString();

                    if (eventer == "수량")//(((eventer == "수량")) || (ButtonTag == "2") && (eventer == "Comments"))
                    {
                        List<TextBox> list = new List<TextBox>();
                        lib.FindChildGroup<TextBox>(dgdOutwareSub, "txtQty", ref list);
                        int target = dgdOutwareSub.Items.IndexOf(dgdOutwareSub.CurrentCell.Item);  //dgdOutRequest.SelectedIndex;
                        TextBox TextBoxComments = list[target];

                        TextBoxComments.IsReadOnly = false;
                        TextBoxComments.Focus();

                        Dispatcher.BeginInvoke((ThreadStart)delegate
                        {
                            TextBoxComments.Focus();
                        });
                    }
                }
            }
        }

        private void dgdOutwareSubRequest_KeyDown(object sender, KeyEventArgs e)
        {
            
        }

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

        private void ThisMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
                if (e.ClickCount == 2)
                {
                    try
                    {
                        UserControl userControl = Lib.Instance.GetParent<UserControl>(sender as DataGrid);
                        var ViewReceiver = dgdOutware.CurrentCell.Item as Win_ord_OutWare_Scan_CodeView;  //선택 줄.
                        string classname = ViewReceiver.OutClssname;
                        
                        if (!classname.Equals("예외출고"))
                        {
                            if (userControl != null)
                            {
                                object objUpdate = userControl.FindName("btnUpdate");
                                object objEdit = userControl.FindName("btnEdit");

                                if (objUpdate != null)
                                {
                                    if ((objUpdate as Button).IsEnabled == true)
                                    {
                                        (objUpdate as Button).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                                    }
                                }
                                else if (objEdit != null)
                                {
                                    if ((objEdit as Button).IsEnabled == true)
                                    {
                                        (objEdit as Button).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("예외출고 수정은 예외출고메뉴에서 해주시기 바랍니다.");
                            return;
                        }
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
        }

        private void tgnMoveByID_Click(object sender, RoutedEventArgs e)
        {
            tgnMoveByID.IsChecked = true;
            tgnMoveByQty.IsChecked = false;

            // 수량 입력 안되도록 → 수량기준이동 토글버튼이 활성화 됬을때만 입력 가능하도록
            txtOutRoll.IsHitTestVisible = false;
            txtOutQty.IsHitTestVisible = false;

            // 바코드 활성화
            txtScanData.IsHitTestVisible = true;

            // 그리드 변경
            dgdOutwareSub.Visibility = Visibility.Visible;

            grdAutoPackingLoad.Visibility = Visibility.Hidden;

            // OutRoll : 박스수, 서브그리드 갯수 / OutQty : 총 개수 - 구하기 
            //SetOutRollAndOutQty();
        }

        private void tgnMoveByQty_Click(object sender, RoutedEventArgs e)
        {
            tgnMoveByID.IsChecked = false;
            tgnMoveByQty.IsChecked = true;

            // 수량 입력 되도록 → 바코드로 입력하도록 막아놓자.
            txtOutRoll.IsHitTestVisible = false;
            txtOutQty.IsHitTestVisible = false;

            // 바코드 입력 안되도록 → 수량기준이동은 바코드가 아닌 수량으로 관리
            //txtBarCode.IsHitTestVisible = false;

            // 바코드 활성화
            txtScanData.IsHitTestVisible = true;

            // 그리드 변경
            dgdOutwareSub.Visibility = Visibility.Visible;

            grdAutoPackingLoad.Visibility = Visibility.Visible;


            // OutRoll : 박스수, 서브그리드 갯수 / OutQty : 총 개수 - 구하기 
            //SetOutRollAndOutQty();
        }

        private object RemoveComma(object obj, bool returnAsNumber = false, Type returnType = null)
        {
            //파라미터가 만약 null일때
            if (obj == null)
            {
                //숫자타입이 false면 string으로 내보내기
                if (!returnAsNumber) return "0";

                // 만약 숫자타입을 써야되면 returnType파라미터의 받은 형태로 전달
                // null일 때도 returnType에 따라 적절한 타입의 0 반환
                switch (returnType?.Name)
                {
                    case "Decimal": return (object)0m;  //monetary
                    case "Double": return (object)0d;   //double
                    case "Int64": return (object)0L;    //long
                    default: return (object)0;          //int
                }
            }

            string digits = obj.ToString()
                              .Trim()
                              .Replace(",", "");

            //만약 빈공백(blank)이더라도 0으로 내보내야한다.
            if (string.IsNullOrEmpty(digits))
            {
                if (!returnAsNumber) return "0";

                // returnType을 활용해서 적절한 타입으로 반환
                switch (returnType?.Name)
                {
                    case "Decimal": return (object)0m;
                    case "Double": return (object)0d;
                    case "Int64": return (object)0L;
                    default: return (object)0;
                }
            }


            try
            {
                Type targetType = returnType ?? typeof(int);

                //혹시나 하는 예외처리
                //입력 컨트롤간에 LostFocus나 TextChanged같은 걸로 계산을 할 때
                //처리 가능한 숫자 범위를 초과하면 오류가 발생하므로
                //초과하면 해당 자료형타입이 처리할 수 있는 최대 숫자를 표시해줌
                switch (targetType.Name)
                {
                    case "Int32":
                        if (decimal.TryParse(digits, out decimal intParsed))
                        {
                            if (intParsed > int.MaxValue) return int.MaxValue;
                            if (intParsed < int.MinValue) return int.MinValue;
                            return (int)intParsed;
                        }
                        return int.MaxValue;

                    case "Int64":
                        if (decimal.TryParse(digits, out decimal longParsed))
                        {
                            if (longParsed > long.MaxValue) return long.MaxValue;
                            if (longParsed < long.MinValue) return long.MinValue;
                            return (long)longParsed;
                        }
                        return long.MaxValue;

                    case "Double":
                        if (double.TryParse(digits, out double doubleParsed))
                        {
                            return doubleParsed;
                        }
                        return double.MaxValue;

                    case "Decimal":
                        if (decimal.TryParse(digits, out decimal decimalParsed))
                        {
                            return decimalParsed;
                        }
                        return decimal.MaxValue;

                    default:
                        return int.MaxValue;
                }
            }
            catch
            {

                if (returnType != null)
                {
                    switch (returnType.Name)
                    {
                        case "Int32":
                            return int.MaxValue;
                        case "Int64":
                            return long.MaxValue;
                        case "Double":
                            return double.MaxValue;
                        case "Decimal":
                            return decimal.MaxValue;
                        default:
                            return int.MaxValue;
                    }
                }
                return int.MaxValue;
            }
        }

        private void ClearGrdInput()
        {
            List<Grid> grids = new List<Grid> { grdInput};

            foreach (Grid grd in grids)
            {
                FindUiObject(grd, child =>
                {
                    if (child is TextBox textbox)
                    {
                        textbox.Text = string.Empty;
                        textbox.Tag = null;
                    }                 
                    else if (child is DatePicker dtp)
                    {
                        dtp.SelectedDate = null;
                    }

                    else if (child is DataGrid dgd)
                    {
                        dgd.Items.Clear();
                    }

                });
            }
        }


        //UI컨트롤 요소찾기
        private void FindUiObject(DependencyObject parent, Action<DependencyObject> action)
        {
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                action?.Invoke(child);

                FindUiObject(child, action);
            }
        }

        //컨트롤 안 특정 타입의 자식 컨트롤을 찾는 함수 (그리드내에서)
        //var parentContainer = VisualTreeHelper.GetParent(checkbox);
        //var datePicker = FindChild<DatePicker>(parentContainer);
        private T FindChild<T>(DependencyObject parent) where T : DependencyObject
        {
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                {
                    return typedChild;
                }

                // 재귀적으로 자식의 자식들도 검색
                var result = FindChild<T>(child);
                if (result != null)
                    return result;
            }
            return null;
        }


        // 자식요소 안에서 부모요소 찾기
        //DataGridRow row = FindVisualParent<DataGridRow>(checkBox); 데이터그리드안의 행속 체크박스의 부모행 찾기
        //DataGrid parentGrid = FindVisualParent<DataGrid>(row); 데이터그리드 행의 부모 데이터그리드 찾기
        private T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);

            if (parentObject == null)
                return null;

            T parent = parentObject as T;
            if (parent != null)
                return parent;
            else
                return FindVisualParent<T>(parentObject);
        }

        private void CommonControl_Click(object sender, MouseButtonEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void CommonControl_Click(object sender, RoutedEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

    }




    class Win_ord_OutWare_Scan_CodeView : BaseView
    {

        public bool Chk { get; set; }

        public string OutwareID { get; set; }
        public string OrderID { get; set; }
        public string OutSeq { get; set; }
        public string OrderNo { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string OutDate { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string OutClss { get; set; }
        public string WorkID { get; set; }
        public string OutRoll { get; set; }
        public string OutQty { get; set; }
        public string OutRealQty { get; set; }
        public string ResultDate { get; set; }
        public string OrderQty { get; set; }
        public string UnitClss { get; set; }
        public string WorkName { get; set; }
        public string OutType { get; set; }
        public string Remark { get; set; }
        public string BuyerModel { get; set; }
        public string OutSumQty { get; set; }
        public string OutQtyY { get; set; }
        public string StuffinQty { get; set; }
        public string OutWeight { get; set; }
        public string OutRealWeight { get; set; }
        public string UnitPriceClss { get; set; }
        public string BuyerDirectYN { get; set; }
        public string Vat_Ind_YN { get; set; }
        public string workID { get; set; }
        public string InsStuffINYN { get; set; }
        public string ExchRate { get; set; }
        public string FromLocID { get; set; }
        public string TOLocID { get; set; }
        public string UnitClssName { get; set; }
        public string FromLocName { get; set; }
        public string TOLocname { get; set; }
        public string OutClssname { get; set; }
        public string UnitPrice { get; set; }
        public string Amount { get; set; }
        public string VatAmount { get; set; }
        public string BuyerArticleNo { get; set; }
        public string OutCustomID { get; set; }
        public string BuyerID { get; set; }
        public string BuyerName { get; set; }
        public string Buyer_Chief { get; set; }
        public string Buyer_Address1 { get; set; }
        public string Buyer_Address2 { get; set; }
        public string Buyer_Address3 { get; set; }
        public string CustomNo { get; set; }
        public string Chief { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string OutCustom { get; set; }
        public string OutSubType { get; set; }

        public string RemainQty { get; set; }
        public string DvlyCustomID { get; set; }
        public string DvlyCustom { get; set; }

        //2021-05-31
        public string Category { get; set; }
        public string Condition { get; set; }

    }

    class Win_ord_OutWare_Scan_Sub_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public bool Chk { get; set; }
        public string OutwareID { get; set; }
        public string OutSubSeq { get; set; }
        public string LabelID { get; set; }
        public string LabelGubun { get; set; }
        public string LabelGubunName { get; set; }

        public string OutQty { get; set; }
        public string OutCnt { get; set; }
        public string OutRoll { get; set; }
        public string LotNo { get; set; }
        public string Weight { get; set; }

        public string UnitPrice { get; set; }
        public string Vat_IND_YN { get; set; }
        public string Orderseq { get; set; }
        public string Amount { get; set; }
        public string CustomBoxID { get; set; }

        public string FromLocID { get; set; }
        public string TOLocID { get; set; }
        public string UnitClss { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }

        public string OutClss { get; set; }
        public string Gubun { get; set; }
        public string DefectID { get; set; }
        public string DefectName { get; set; }

        public string DeleteYN { get; set; }

        public string OutRealQty { get; set; }
    }

}
