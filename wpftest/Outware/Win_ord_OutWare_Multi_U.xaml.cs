using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using WizMes_SungShinNQ.PopUP;
using Excel = Microsoft.Office.Interop.Excel;


namespace WizMes_SungShinNQ
{
    /// <summary>
    /// Win_ord_OutWare_Multi_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_OutWare_Multi_U : UserControl
    {
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();
        private IProgress<int> _progress;

        // 인쇄 활용 용도 (프린트)
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;

        bool preview_click = false;

        WizMes_SungShinNQ.PopUp.NoticeMessage msg = new WizMes_SungShinNQ.PopUp.NoticeMessage();

        private ToolTip currentToolTip;
        private System.Windows.Threading.DispatcherTimer currentTimer;

        private DispatcherTimer countdownTimer;
        private int countdownSeconds;

        //박스라벨이 될수있는 첫글자
        string[] LabelPrefix = { "C", "B" };

        Dictionary<string, int> ListTmp = new Dictionary<string, int>(); //우측 그리드에 잔여량 확인시 사용할 딕셔너리
        int cnt = 0;            //우측그리드 잔여량 체크용
        int rowNum = 0;                          // 조회시 데이터 줄 번호 저장용도
        int isQtyOver = 0;
        int isNumber = 0;
        string strFlag = string.Empty;           // 추가, 수정 구분 
        string GetKey = "";

        List<string> LabelGroupList = new List<string>();         // packing ID 스캔에 따른 LabelID를 모아 담을 리스트 그릇입니다.

        string strBoxID = string.Empty;
        string strLabelGbn = string.Empty;
        string StockQty = string.Empty;

        ObservableCollection<Win_ord_OutWare_Multi_U_dgdLEFT_CodeView> ovcDgdLeft = new ObservableCollection<Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>();
        ObservableCollection<Win_ord_OutWare_Multi_U_dgdRight_CodeView> ovcDgdRight = new ObservableCollection<Win_ord_OutWare_Multi_U_dgdRight_CodeView>();
        ObservableCollection<Win_ord_OutWare_Multi_U_CodeView> ovcdgdMain = new ObservableCollection<Win_ord_OutWare_Multi_U_CodeView>();
        List<Win_ord_OutWare_Multi_U_CodeView> lstMultiOutwarePrint = new List<Win_ord_OutWare_Multi_U_CodeView>();

        public Win_ord_OutWare_Multi_U()
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

                this.DataContext = new Win_ord_OutWare_Multi_U_CodeView();
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
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            //dtpFromDate.SelectedDate = DateTime.Today.AddDays(-1);
            //dtpToDate.SelectedDate = DateTime.Today.AddDays(-1);

            try
            {
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
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnYesterday_Click : " + ee.ToString());
            }
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                dtpFromDate.SelectedDate = DateTime.Today;
                dtpToDate.SelectedDate = DateTime.Today;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnToday_Click : " + ee.ToString());
            }
        }

        //전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //dtpFromDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[0];
            //dtpToDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[1];

            try
            {
                if (dtpFromDate.SelectedDate != null)
                {
                    DateTime ThatMonth1 = dtpFromDate.SelectedDate.Value.AddDays(-(dtpFromDate.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                    DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                    dtpFromDate.SelectedDate = LastMonth1;
                    dtpToDate.SelectedDate = LastMonth31;
                }
                else
                {
                    DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                    DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                    dtpFromDate.SelectedDate = LastMonth1;
                    dtpToDate.SelectedDate = LastMonth31;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnLastMonth_Click : " + ee.ToString());
            }
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                dtpFromDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
                dtpToDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnThisMonth_Click : " + ee.ToString());
            }
        }



        #endregion

        #region 버튼 모음

        //체크박스 클릭시 입력란 enabled/disabled 토글
        //그리드로 라벨,체크박스,입력란 묶을것
        private void CommonControl_Click(object sender, RoutedEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        //라벨 클릭시 입력란 enabled/disabled 토글
        //그리드로 라벨,체크박스,입력란 묶을것
        private void CommonControl_Click(object sender, MouseButtonEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        //입력란 키다운 이벤트
        private void CommonPlusfinder_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TextBox txtbox = (TextBox)sender;
                if (txtbox != null && !string.IsNullOrEmpty(txtbox.Name))
                {
                    if (txtbox.Name.Equals("txtCustomIDSrh") || txtbox.Name.Equals("txtInCustomIDSrh"))
                    {
                        pf.ReturnCode(txtbox, 0, "");
                    }
                    else if (txtbox.Name.Equals("txtArticleIDSrh"))
                    {
                        pf.ReturnCode(txtbox, 77, txtbox.Text);
                    }
                    else if (txtbox.Name.Equals("txtBuyerArticleNoSrh"))
                    {
                        pf.ReturnCode(txtbox, 76, txtbox.Text);
                    }
                    else if (txtbox.Name.Equals("txtOrderNo"))
                    {
                        pf.ReturnCode(txtbox, 99, txtbox.Text);
                        if (txtbox.Tag != null)
                        {
                            var (ordData, ordSubRight) = GetOrderData(txtbox.Text);
                            var item = this.DataContext as Win_ord_OutWare_Multi_U_CodeView;
                            if (item != null && ordData != null)
                            {
                                if (txtbox.Text != item.OriginOrderID && strFlag.Equals("U"))
                                {
                                    MessageBoxResult msgResult = MessageBox.Show("이미 저장된 데이터의 오더번호를 변경하면 기존에 입력된 출하 데이터는 초기화 됩니다. 계속 진행하시겠습니까?", "확인", MessageBoxButton.YesNo);
                                    if (msgResult == MessageBoxResult.Yes)
                                    {
                                        ovcDgdLeft.Clear();
                                        ovcDgdRight.Clear();
                                        dgdLeft.Items.Clear();
                                        dgdRight.Items.Clear();
                                    }
                                    else
                                    {
                                        return;
                                    }
                                }

                                item.OrderID = ordData.OrderID;
                                item.CustomID = ordData.CustomID;
                                item.KCustom = ordData.KCustom;
                                item.OutCustom = ordData.OutCustom;
                                item.OutCustomID = ordData.OutCustomID;
                                item.Article = ordData.Article;
                                item.ArticleID = ordData.ArticleID;
                                item.BuyerArticleNo = ordData.BuyerArticleNo;

                                txtScan.Focus();
                                lib.ShowTooltipMessage(txtScan, "박스라벨번호 스캔 또는 박스조회 버튼을 눌러 잔량을 조회 하세요", MessageBoxImage.Information, PlacementMode.Top, 1.3);

                            }
                        }

                    }
                }
            }
        }

        //입력란 버튼 클릭 이벤트
        private void CommonPlusfinder_Click(object sender, RoutedEventArgs e)
        {
            Button btn = (Button)sender;
            TextBox txtbox = lib.FindSiblingControl<TextBox>(btn);
            if (txtbox != null && !string.IsNullOrEmpty(txtbox.Name))
            {
                if (txtbox != null && !string.IsNullOrEmpty(txtbox.Name))
                {
                    if (txtbox.Name.Equals("txtCustomIDSrh") || txtbox.Name.Equals("txtInCustomIDSrh"))
                    {
                        pf.ReturnCode(txtbox, 0, "");
                    }
                    else if (txtbox.Name.Equals("txtArticleIDSrh"))
                    {
                        pf.ReturnCode(txtbox, 77, txtbox.Text);
                    }
                    else if (txtbox.Name.Equals("txtBuyerArticleNoSrh"))
                    {
                        pf.ReturnCode(txtbox, 76, txtbox.Text);
                    }
                    else if (txtbox.Name.Equals("txtOrderNo"))
                    {
                        pf.ReturnCode(txtbox, 99, txtbox.Text);
                        if (txtbox.Tag != null)
                        {
                            var (ordData, ordSubRight) = GetOrderData(txtbox.Text);
                            var item = this.DataContext as Win_ord_OutWare_Multi_U_CodeView;
                            if (item != null && ordData != null)
                            {
                                if (txtbox.Text != item.OriginOrderID && strFlag.Equals("U"))
                                {
                                    MessageBoxResult msgResult = MessageBox.Show("이미 저장된 데이터의 오더번호를 변경하면 기존에 입력된 출하 데이터는 초기화 됩니다. 계속 진행하시겠습니까?", "확인", MessageBoxButton.YesNo);
                                    if (msgResult == MessageBoxResult.Yes)
                                    {
                                        ovcDgdLeft.Clear();
                                        ovcDgdRight.Clear();
                                        dgdLeft.Items.Clear();
                                        dgdRight.Items.Clear();
                                    }
                                    else
                                    {
                                        return;
                                    }
                                }

                                item.OrderID = ordData.OrderID;
                                item.CustomID = ordData.CustomID;
                                item.KCustom = ordData.KCustom;
                                item.OutCustom = ordData.OutCustom;
                                item.OutCustomID = ordData.OutCustomID;
                                item.Article = ordData.Article;
                                item.ArticleID = ordData.ArticleID;
                                item.BuyerArticleNo = ordData.BuyerArticleNo;

                                txtScan.Focus();
                                lib.ShowTooltipMessage(txtScan, "박스라벨번호 스캔 또는 박스조회 버튼을 눌러 잔량을 조회 하세요", MessageBoxImage.Information, PlacementMode.Top, 1.3);
                            }
                        }

                    }
                }

            }
        }

        //추가버튼 클릭
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                strFlag = "I";

                this.DataContext = new Win_ord_OutWare_Multi_U_CodeView();
                if (dgdLeft.ItemsSource != null) dgdLeft.ItemsSource = null;
                if (dgdRight.ItemsSource != null) dgdRight.ItemsSource = null;

                ovcDgdLeft.Clear();
                ovcDgdRight.Clear();
                dgdLeft.Items.Clear();
                dgdRight.Items.Clear();       

                CanBtnControl();
                dtpOutDate.SelectedDate = DateTime.Today;

                cboOutClss.SelectedIndex = 0;
                cboFromLoc.SelectedIndex = 0;
                cboToLoc.SelectedIndex = 0;

                btnBoxSearch.IsEnabled = true;
                txtScan.Focus();


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
                if (this.DataContext == null || dgdOutware.SelectedItem == null) return;


                var OutwareItem = dgdOutware.SelectedItem as Win_ord_OutWare_Multi_U_CodeView;

                if (OutwareItem != null)
                {
                    strFlag = "U";

                    rowNum = dgdOutware.SelectedIndex;
                    CanBtnControl();
                    //dgdStuffQty.Visibility = Visibility.Visible;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnUpdate_Click : " + ee.ToString());
            }
        }

       

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                if (ovcdgdMain.Count == 0)
                {
                    MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 체크버튼을 이용해 삭제 데이터를 지정하고 눌러주세요.");
                }
                else
                {
                    if (MessageBox.Show("선택하신 출고 항목을 정말 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        bool flag = true;
                        List<Win_ord_OutWare_Multi_U_CodeView> itemsToRemove = new List<Win_ord_OutWare_Multi_U_CodeView>(ovcdgdMain);

                        foreach (var item in itemsToRemove)
                        {
                            var removeData = item as Win_ord_OutWare_Multi_U_CodeView;
                            if (removeData != null)
                            {
                                if (!DeleteData(removeData.OutWareID))
                                {
                                    flag = false;
                                    MessageBox.Show("삭제중 오류가 발생했습니다. 일부 데이터는 삭제 되지 않았을 수 있습니다.");
                                    break;
                                }
                                ovcdgdMain.Remove(item);
                            }
                        }

                        if (flag)
                        {
                            rowNum = 0;
                            re_Search(rowNum);
                        }
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
            if (lib.DatePickerCheck(dtpFromDate, dtpToDate, chkOutwareDay))
            {
                try
                {
                    rowNum = 0;
                    re_Search(rowNum);
                }
                catch (Exception ee)
                {
                    MessageBox.Show("오류지점 - btnSearch_Click : " + ee.ToString());
                }
            }

        }

        //저장버튼 클릭
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            Win_ord_OutWare_Multi_U_CodeView outData = this.DataContext as Win_ord_OutWare_Multi_U_CodeView;
            List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> rightData = ovcDgdRight.ToList();

            if (SaveData(strFlag, outData, rightData))
            {
                CantBtnControl();           //버튼 컨트롤
                if (strFlag.Equals("I"))
                {
                    var outwareCount = dgdOutware.Items.Count;

                    rowNum = outwareCount;
                    re_Search(rowNum);

                }
                else if (strFlag.Equals("U"))
                {
                    re_Search(rowNum);
                }

                MessageBox.Show("저장 되었습니다.", "확인");
                strFlag = string.Empty;
                re_Search(rowNum);
            }

        }

        //취소버튼 클릭
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CantBtnControl();           //버튼 컨트롤
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
            try
            {
                if (dgdOutware.Items.Count < 1)
                {
                    MessageBox.Show("먼저 검색해 주세요.");
                    return;
                }

                Lib lib = new Lib();
                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[4];
                lst[0] = "메인그리드";
                lst[1] = "서브그리드";
                lst[2] = dgdOutware.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdOutware.Name))
                    {
                        //MessageBox.Show("대분류");
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdOutware);
                        else
                            dt = lib.DataGirdToDataTable(dgdOutware);

                        Name = dgdOutware.Name;
                        lib.GenerateExcel(dt, Name);
                        lib.excel.Visible = true;
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
        }

        #region 거래명세서


        //인쇄버튼 클릭
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ContextMenu menu = new ContextMenu();
                Button btn = sender as Button;
                if (btn != null && btn.Tag != null)
                {
                    if (btn.Tag.Equals("Bill")) //메뉴 컨텍스트 돌려쓰려고
                        menu = btnPrint.ContextMenu;

                }
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
            MenuItem menu = sender as MenuItem;
            if (menu != null)
            {
                string menuTag = menu.Tag as string;
                menuPrint_Click(true, menuTag);

            }
        }

        //인쇄-바로인쇄 클릭
        private void menuRighPrint_Click(object sender, RoutedEventArgs e)
        {
            MenuItem menu = sender as MenuItem;
            if (menu != null)
            {
                string menuTag = menu.Tag as string;
                menuPrint_Click(false, menuTag);

            }
        }

        private async void menuPrint_Click(bool Ahead, string callFrom = null)
        {
            try
            {
                if (dgdOutware.Items.Count == 0)
                {
                    MessageBox.Show("먼저 검색해 주세요.","확인");
                    return;
                }
                else if(ovcdgdMain.Count == 0)
                {
                    MessageBox.Show("인쇄하실 출하건을 체크 하여 주십시오.", "확인");
                    return;
                }

                    preview_click = Ahead;

                DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");
                /*msg.Show();
                msg.Topmost = true;
                msg.Refresh();
                msg.Visibility = Visibility.Hidden;*/

                //using (Loading ld = new Loading("excel", ()=> PrintWork(preview_click)))
                //{
                //    ld.ShowDialog();
                //}
                //PrintWork(preview_click);
                this.IsHitTestVisible = false;
                EventLabel.Visibility = Visibility.Visible;

                _progress = new Progress<int>(percent =>
                {
                    tbkMsg.Text = $"준비중입니다... {percent}%";
                });

                await Task.Run(() =>
                {
                    PrintWork(preview_click, callFrom);
                });

                this.IsHitTestVisible = true;
                EventLabel.Visibility = Visibility.Hidden;
                tbkMsg.Text = "자료 입력 중";


            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - menuRighPrint_Click : " + ee.ToString());
                this.IsHitTestVisible = true;
                EventLabel.Visibility = Visibility.Hidden;
                tbkMsg.Text = "자료 입력 중";
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

        //엑셀 정품, 비정품 확인용
        private bool IsExcelActivated()
        {

            if (App._isExcelActivatedCache.HasValue)
            {
                return App._isExcelActivatedCache.Value;
            }

            try
            {
                Excel.Application testExcel = null;
                try
                {
                    testExcel = new Excel.Application();
                    testExcel.Visible = true;
                    bool isVisible = testExcel.Visible;
                    testExcel.Visible = false;

                    App._isExcelActivatedCache = isVisible;
                    return isVisible;
                }
                finally
                {
                    if (testExcel != null)
                    {
                        testExcel.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(testExcel);
                    }
                }
            }
            catch
            {
                App._isExcelActivatedCache = false;
                return false;
            }
        }

        //프린트메서드 수정판
        private void PrintWork(bool previewYN, string callFrom = null)
        {
            Excel.Application excelapp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Worksheet pastesheet = null;
            Excel.Range workrange = null;
            int excelProcessId = 0;

            string sheetName = "Org_거래명세표"; //내장 리소스 시트명


            try
            {
                _progress?.Report(0);


                List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> lstMultiOutWareSub = new List<Win_ord_OutWare_Multi_U_dgdRight_CodeView>();
                SetCompanyData setcompanyData = SetCompanyData.GetSetCompanyData();

                List<Win_ord_OutWare_Multi_U_CodeView> lstOutwarePrint = ovcdgdMain.ToList();

                lstOutwarePrint.ForEach(item =>
                {
                    List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> subItems = Win_ord_OutWare_Multi_U_dgdRight_CodeView.GetOutwareSubData(item.OutWareID);

                    lstMultiOutWareSub.AddRange(subItems);
                });

                //엑셀 생성
                excelapp = new Excel.Application();

                // 알림 및 화면 업데이트 비활성화
                excelapp.DisplayAlerts = false;
                excelapp.ScreenUpdating = false;

                //생성한 프로세스 아이디 저장(닫을때 EXCEL COM 정리용으로 사용함)
                excelProcessId = GetExcelProcessId();

                _progress?.Report(10);

                var assembly = Assembly.GetExecutingAssembly();
                string[] resourceNames = assembly.GetManifestResourceNames();
                string templateResourceName = resourceNames.FirstOrDefault(r => r.Contains(sheetName));

                // 내장 리소스 존재 확인
                if (string.IsNullOrEmpty(templateResourceName))
                {
                    throw new FileNotFoundException("시스템에 저장된 양식을 찾을 수 없습니다.\n관리자에게 문의해주세요");
                }

                // 임시 파일로 추출
                string templatePath = Path.Combine(Path.GetTempPath(), $"{sheetName}{Guid.NewGuid()}.xlsx");

                using (Stream stream = assembly.GetManifestResourceStream(templateResourceName))
                {
                    using (var fileStream = File.Create(templatePath))
                    {
                        stream.CopyTo(fileStream);
                    }
                }

                workbook = excelapp.Workbooks.Add(templatePath);
                worksheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                //먼저 원본 시트의 인쇄영역이 어디까지인지 구합니다.
                //여기서 오류 나면 원본시트에 인쇄영역(printArea)이 지정되어있는지 확인하세요
                Excel.Range printArea = worksheet.Range[worksheet.PageSetup.PrintArea];

                int columnCount = printArea.Columns.Count;          //인쇄영역으로 지정된 원본시트의 컬럼 합계수
                int rowsCount = printArea.Rows.Count;               //인쇄영역으로 지정된 원본시트의 로우 합계수
                int startRow = printArea.Row;                       //인쇄영역 설정 첫 시작지점

                string startColumnLetter = Regex.Match(printArea.Columns[1].Address[false, false], @"[A-Z]+").Value;
                string endColumnLetter = Regex.Match(printArea.Columns[printArea.Columns.Count].Address[false, false], @"[A-Z]+").Value;

                string workSheetStartRow = startRow.ToString();
                string workSheetRowsCount = rowsCount.ToString();

                //원본 조합
                string workSheetX = startColumnLetter + workSheetStartRow;      //위의 내용으로 원본시트의 시작지점과 끝지점을 구합니다.
                string workSheetY = endColumnLetter + workSheetRowsCount;

                _progress?.Report(15);

                //먼저 고정값을 원본시트에 적어놓습니다. 복사시트에 재활용 함   
                FillBaseInfo(worksheet, setcompanyData);

                //그 다음 원본시트를 복사시트에 복사하며 값을 넣습니다.
                FillDataIntoPasteSheet(worksheet, pastesheet, 10, 15, workSheetX, workSheetY, startColumnLetter, endColumnLetter, columnCount, rowsCount, startRow, lstOutwarePrint, lstMultiOutWareSub);

                //복사시트 선택
                pastesheet.Select();

                if (!IsPrinterAvailable())
                {
                    throw new Exception("윈도우에 연결된 기본 프린터가 없습니다.\n기본 프린터를 설정한 후 시도하여주세요.");
                }

                _progress?.Report(100);

                bool isActivated = IsExcelActivated();

                if (previewYN)
                {
                    if (isActivated)
                    {
                        //  정품 인증됨: 기존 방식 (Excel COM으로 직접 제어)
                        excelapp.ScreenUpdating = true;  // 화면 업데이트 활성화
                        excelapp.Visible = true;
                        excelapp.UserControl = true;
                        workbook.Saved = true;

                        ReleaseExcelObject(workrange);
                        ReleaseExcelObject(pastesheet);
                        ReleaseExcelObject(worksheet);
                        ReleaseExcelObject(workbook);
                        ReleaseExcelObject(excelapp);
                    }
                    else
                    {
                        //  정품 인증 안됨: 파일로 저장 후 열기(더블클릭 한다고 생각)
                        string tempFile = Path.Combine(Path.GetTempPath(), $"출하처리(스캔)_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                        workbook.SaveAs(tempFile);

                        workbook.Close(false);
                        excelapp.Quit();

                        if (excelProcessId != 0)
                        {
                            KillExcelProcess(excelProcessId);
                        }

                        ReleaseExcelObject(workrange);
                        ReleaseExcelObject(pastesheet);
                        ReleaseExcelObject(worksheet);
                        ReleaseExcelObject(workbook);
                        ReleaseExcelObject(excelapp);

                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        // 파일을 기본 Excel로 열기
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = tempFile,
                            UseShellExecute = true
                        });
                    }
                }
                else
                {
                    // 바로 인쇄는 정품 인증 관계없이 동일
                    pastesheet.PrintOut();

                    workbook.Close(false);
                    excelapp.Quit();

                    if (excelProcessId != 0)
                    {
                        KillExcelProcess(excelProcessId);
                    }
                    ReleaseExcelObject(workrange);
                    ReleaseExcelObject(pastesheet);
                    ReleaseExcelObject(worksheet);
                    ReleaseExcelObject(workbook);
                    ReleaseExcelObject(excelapp);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    if (workbook != null) workbook.Close(false);
                    if (excelapp != null) excelapp.Quit();
                }
                catch { }

                if (excelProcessId != 0)
                {
                    KillExcelProcess(excelProcessId);
                }
                ReleaseExcelObject(workrange);
                ReleaseExcelObject(pastesheet);
                ReleaseExcelObject(worksheet);
                ReleaseExcelObject(workbook);
                ReleaseExcelObject(excelapp);
                MessageBox.Show($"오류가 발생했습니다\n: {ex.Message}");
                throw;
            }
        }





        //원본시트의 서식, 행, 열 높이 등등 복사시트에 복사하는 메서드
        private void BaseCopySheet(Excel.Worksheet worksheet, Excel.Worksheet pastesheet, string worksheetX, string worksheetY, string pasteSheetX, string pasteSheetY)
        {
            //원본 시트의 범위를 소스로 잡습니다.
            Excel.Range sourceRange = worksheet.Range[$"{worksheetX}:{worksheetY}"];

            //복사할 위치 지정
            Excel.Range destination1 = pastesheet.Range[$"{pasteSheetX}"];

            // 클립보드를 사용하지 않고 직접 복사 (방법 1)
            sourceRange.Copy(destination1);

            //붙여넣고 나면 원본의 열, 행 높이넓이를 재지정
            int X = Convert.ToInt32(Regex.Replace(pasteSheetX, "[^0-9]", ""));

            X = X - 1;

            for (int i = 1; i <= sourceRange.Rows.Count; i++)
            {
                pastesheet.Rows[X + i].RowHeight = worksheet.Rows[i].RowHeight;
            }

            for (int j = 1; j <= sourceRange.Columns.Count; j++)
            {
                pastesheet.Columns[j].ColumnWidth = worksheet.Columns[j].ColumnWidth;
            }

            #region 용지 크기를 구해서 인쇄너비를 계산하기 버전
            //// 보통 A4
            //// A4 용지 크기 (포인트 단위, 1 inch = 72 points)
            //// A4 = 210mm x 297mm = 8.27 inch x 11.69 inch
            //const double A4_WIDTH_POINTS = 8.27 * 72;  // 약 595 points
            //const double A4_HEIGHT_POINTS = 11.69 * 72; // 약 842 points


            //pastesheet.PageSetup.PrintArea = $"{worksheetX}:{pasteSheetY}";
            //// 인쇄 영역 가져오기
            //Excel.Range printArea = pastesheet.Range[pastesheet.PageSetup.PrintArea];

            //// 인쇄 영역의 너비와 높이
            //double printAreaWidth = printArea.Width;
            //double printAreaHeight = printArea.Height;

            //// 여백 고려 (포인트 단위)
            //double availableWidth = A4_WIDTH_POINTS - (pastesheet.PageSetup.LeftMargin + pastesheet.PageSetup.RightMargin);
            //double availableHeight = A4_HEIGHT_POINTS - (pastesheet.PageSetup.TopMargin + pastesheet.PageSetup.BottomMargin);

            //// 배율 계산
            //double widthScale = (availableWidth / printAreaWidth) * 100;
            //double heightScale = (availableHeight / printAreaHeight) * 100;
            //int zoom = (int)Math.Min(widthScale, heightScale);

            //pastesheet.PageSetup.Zoom = zoom;
            #endregion

            //해보니까 배율을 자동
            //너비는 1페이지, 높이는 자동, 그리고 페이지브레이크만 넣으면 페이지 나누기,
            //그리고 인쇄영역을 인쇄할 부분 끝까지 지정하면
            //페이지 나누기 미리보기(실제 인쇄되면 나오는 부분)에서 딱 원본시트 복사한것 만큼 나온다
            //여러장에 적용 가능
            pastesheet.PageSetup.Zoom = false;
            pastesheet.PageSetup.FitToPagesWide = 1;
            pastesheet.PageSetup.FitToPagesTall = false;

            //페이지로 나눌 부분을 설정합니다.
            string pageBreakPointLetter_Row = Regex.Replace(pasteSheetY, "[^A-Z]", ""); //붙여넣는 부분 끝나는 지점이라 Y
            string pageBreakPointRows = Regex.Replace(pasteSheetY, "[^0-9]", "");
            int pageRowCount = Convert.ToInt32(pageBreakPointRows);

            //지정한 곳(인쇄영역으로 지정한 행 수 다음)으로 페이지 삽입을 합니다.
            Excel.Range nextPageRange = pastesheet.Range[$"{pageBreakPointLetter_Row}" + (pageRowCount + 1)];
            pastesheet.HPageBreaks.Add(nextPageRange);

            //인쇄영역을 처음부터 복사한 부분까지 지정합니다.
            pastesheet.PageSetup.PrintArea = $"{worksheetX}:{pasteSheetY}";

        }


        //고정부분 채우기
        private void FillBaseInfo(Excel.Worksheet sheet, SetCompanyData setCompanyData = null)
        {
            if (setCompanyData != null)
            {
                //공급자 부분
                workrange = sheet.Range["W6"];
                workrange.Value2 = setCompanyData.companyNo.Length.Equals(10) ? Regex.Replace(setCompanyData.companyNo, @"(\d{3})(\d{2})(\d{5})", "$1-$2-$3") : setCompanyData.companyNo;

                workrange = sheet.Range["W8"];
                workrange.Value2 = setCompanyData.kCompany;

                workrange = sheet.Range["AE8"];
                workrange.Value2 = setCompanyData.chief;

                workrange = sheet.Range["W10"];
                workrange.Value2 = setCompanyData.address1 + "\n" + setCompanyData.address2;

                workrange = sheet.Range["W12"];
                workrange.Value2 = setCompanyData.phone1;

                workrange = sheet.Range["AD12"];
                workrange.Value2 = setCompanyData.faxNo;
            }
        }


        private void FillDataIntoPasteSheet(Excel.Worksheet worksheet, Excel.Worksheet pastesheet,
                                            int perRow, int insertStartRow, string workSheetX, string workSheetY,
                                            string startColumnLetter, string endColumnLetter, int columnsCount, int RowsCount, int startRow,
                                            List<Win_ord_OutWare_Multi_U_CodeView> lstOutWarePrint, List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> lstOutWareSubPrint)
        {
            #region 파라미터 설명
            /*받는 파라미터 => (
                                    worksheet = 원본시트 
                                    pastesheet = 복사시트              
                                    perRow = 복사시트에 입력할 수 있는 행 수
                                    insertStartRow = 복사시트에 몇 줄부터 입력을 시작할지 정함
                                    workSheetX = 원본시트 시작행
                                    workSheetY = 원본시트 시작열
                                    startColumnLetter = 원본시트 시작열 문자
                                    endColumnLetter = 원본시트 끝나는 지점 문자
                                    columnsCount = 원본시트 컬럼 수
                                    RowsCount = 인쇄영역으로 지정된 원본시트의 행 수
                                    startRow = 원본시트 시작 행
                                    lstOutWarePrint 데이터그리드 체크한 항목
                                    lstOutWareSubPrint 메인데이터그리드 체크한 항목의 Sub값들
                                )*/
            #endregion

            // 1. 총 페이지 수 계산
            int totalPages = 0;
            List<int> pagesPerMain = new List<int>();  // 각 메인별 페이지 수 저장
            string fileName = ((Excel.Workbook)worksheet.Parent).Name;

            for (int i = 0; i < lstOutWarePrint.Count; i++)
            {
                var main = lstOutWarePrint[i];
                int subCount = lstOutWareSubPrint.Count(s => s.OutWareID == main.OutWareID);
                int pages = Math.Max(1, (int)Math.Ceiling((double)subCount / perRow));
                pagesPerMain.Add(pages);
                totalPages += pages;

                int percent = 10 + ((i + 1) * 10 / lstOutWarePrint.Count); // 10% ~ 20%
                _progress?.Report(percent);
            }



            // 2. 페이지 복사
            int row = RowsCount;
            string pasteSheetX = startColumnLetter + startRow;
            string pasteSheetY = endColumnLetter + row;
            int globalPageNum = 1;
            int mainIndex = 0;
            int pageInMain = 0;

            for (int i = 0; i < totalPages; i++)
            {
                BaseCopySheet(worksheet, pastesheet, workSheetX, workSheetY, pasteSheetX, pasteSheetY);

                int currentPageStartRow = startRow + (i * RowsCount);

                // 현재 어느 메인의 페이지인지 계산
                var mainItem = lstOutWarePrint[mainIndex];

                pasteSheetX = startColumnLetter + (startRow + row);
                pasteSheetY = endColumnLetter + (startRow + row + RowsCount - 1);


                //여러시트 사용할거면 수정하세요
                //거래명세표의 공급받는자부분
                /*if (fileName.Contains("파렛트"))
                //{
                //    // 메인 정보 입력
                //    pastesheet.Cells[currentPageStartRow + 2, 9] = mainItem.OutCustom;
                //    pastesheet.Cells[currentPageStartRow + 6, 9] = mainItem.KCustom;
                //    pastesheet.Cells[currentPageStartRow + 6, 36] = mainItem.OutDate?.ToString("yyyy-MM-dd");
                //}
                //else*/ if (fileName.Contains("거래명세표"))
                {
                    pastesheet.Cells[currentPageStartRow + 4, 3] = mainItem.OutDate?.ToString("yyyy-MM-dd");
                    pastesheet.Cells[currentPageStartRow + 5, 7] = mainItem.KCustom;
                    pastesheet.Cells[currentPageStartRow + 7, 7] = $"{mainItem.Address1}\n{mainItem.Address2}";
                    pastesheet.Cells[currentPageStartRow + 9, 7] = mainItem.Chief;

                }


                row += RowsCount;

                // 다음 메인으로 넘어가야 하는지 체크
                pageInMain++;
                if (pageInMain >= pagesPerMain[mainIndex])
                {
                    mainIndex++;
                    pageInMain = 0;
                }

                globalPageNum++;

                int percent = 20 + ((i + 1) * 40 / totalPages); // 20% ~ 60%
                _progress?.Report(percent);
            }

            // 3. 데이터 입력
            row = 0;
            mainIndex = 0;
            pageInMain = 0;

            for (int k = 0; k < totalPages; k++)
            {
                var mainItem = lstOutWarePrint[mainIndex];
                var subsForMain = lstOutWareSubPrint.Where(s => s.OutWareID == mainItem.OutWareID).ToList();


                // 이 페이지에 들어갈 서브 데이터
                int startIdx = pageInMain * perRow;
                int endIdx = Math.Min(startIdx + perRow, subsForMain.Count);

                //디버깅
                //System.Diagnostics.Debug.WriteLine($"=== 페이지 {k} ===");
                //System.Diagnostics.Debug.WriteLine($"OutWareID: {mainItem.OutWareID}");
                //System.Diagnostics.Debug.WriteLine($"subsForMain.Count: {subsForMain.Count}");
                //System.Diagnostics.Debug.WriteLine($"startIdx: {startIdx}, endIdx: {endIdx}");
                //System.Diagnostics.Debug.WriteLine($"반복 횟수: {endIdx - startIdx}");
                //System.Diagnostics.Debug.WriteLine($"row: {row}");


                //내역 넣기
                /*if (fileName.Contains("파렛트"))
                {
                    for (int j = 0; j < endIdx - startIdx; j++)
                    {
                        var subItem = subsForMain[startIdx + j];
                        int targetRow = insertStartRow + row + (j * 4);
                        //int targetRow = insertStartRow + row + j;
                        //if (j > 0) targetRow = targetRow + 3;

                        //System.Diagnostics.Debug.WriteLine($"j={j}, targetRow={targetRow}, Article={subItem.Article}");

                        pastesheet.Cells[targetRow, 2] = j + 1;           // 0 → 1로 변경
                        pastesheet.Cells[targetRow, 6] = subItem.Article;
                        pastesheet.Cells[targetRow, 18] = subItem.Spec;
                    }
                }
                else*/ if (fileName.Contains("거래명세표"))
                {
                    for (int j = 0; j < endIdx - startIdx; j++)
                    {
                        var subItem = subsForMain[startIdx + j];

                        pastesheet.Cells[insertStartRow + row + j, 3] = mainItem.OutDate?.ToString("MM");
                        pastesheet.Cells[insertStartRow + row + j, 4] = mainItem.OutDate?.ToString("dd");
                        pastesheet.Cells[insertStartRow + row + j, 5] = subItem.Article;
                        pastesheet.Cells[insertStartRow + row + j, 16] = subItem.OutQty;
                        pastesheet.Cells[insertStartRow + row + j, 18] = subItem.UnitPrice;
                        pastesheet.Cells[insertStartRow + row + j, 23] = subItem.OutQty * subItem.UnitPrice;
                        pastesheet.Cells[insertStartRow + row + j, 28] = Convert.ToInt32(lib.RemoveComma(subItem.OutQty * subItem.UnitPrice, 0) * 0.1);
                    }
                }


                row += RowsCount;

                // 다음 메인으로
                pageInMain++;
                if (pageInMain >= pagesPerMain[mainIndex])
                {
                    mainIndex++;
                    pageInMain = 0;
                }

                int percent = 60 + ((k + 1) * 30 / totalPages); // 60% ~ 90%
                _progress?.Report(percent);
            }
        }

        //기본 프린터가 하나라도 지정되었나요?
        private bool IsPrinterAvailable()
        {
            return System.Drawing.Printing.PrinterSettings.InstalledPrinters.Count > 0;
        }


        //프린트 핸들러
        private void HandlePrintPreview(Excel.Application app, Excel.Worksheet sheet, bool preview)
        {
            if (!IsPrinterAvailable())
            {
                throw new Exception("윈도우에 연결된 기본 프린터가 없습니다.\n기본 프린터를 설정한 후 시도하여주세요.");
            }

            app.Visible = true;
            if (preview)
            {
                sheet.PrintPreview();
            }
            else
            {
                sheet.PrintOut();
            }
        }

        //엑셀 리소스 정리
        private void ReleaseExcelObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
                catch
                {
                    obj = null;
                }
                finally
                {
                    GC.Collect();
                }
            }
        }

        // 실행 후 프로세스 아이디를 시간순 정렬해서 가져오기
        private int GetExcelProcessId()
        {
            var process = Process.GetProcessesByName("EXCEL")
                                .OrderByDescending(p => p.StartTime)
                                .FirstOrDefault();
            return process?.Id ?? 0;
        }

        //릴리즈해도 프로세스가 하나는 끝까지 살아남아서...
        private void KillExcelProcess(int processId)
        {
            try
            {
                Process process = Process.GetProcessById(processId);
                if (!process.HasExited)
                {
                    process.Kill();
                }
            }
            catch { }
        }



        #endregion

        #endregion

        #region 키다운 이동 모음

        private void TextBoxFocusInDataGrid(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        private void DataGridSubCell_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Left || e.Key == Key.Right)
            {
                DataGridSubCell_KeyDown(sender, e);
            }
        }

        private void DataGridSubCell_KeyDown(object sender, KeyEventArgs e)
        {
            if (EventLabel.Visibility == Visibility.Visible)
            {
                // 이 셀을 포함하는 DataGrid 가져오기
                DataGridCell cell = sender as DataGridCell;
                DataGrid dataGrid = lib.FindVisualParent<DataGrid>(cell);

                if (dataGrid == null) return;


                ComboBox comboBox = lib.FindChild<ComboBox>(cell);

                if (comboBox != null && comboBox.IsDropDownOpen)
                {
                    // 키 이벤트가 콤보박스로 전달되게 하기
                    return;
                }

                int rowCount = dataGrid.Items.IndexOf(dataGrid.CurrentItem);
                int colCount = dataGrid.Columns.IndexOf(dataGrid.CurrentCell.Column);
                int lastColcount = dataGrid.Columns.Count - 1;

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;

                    if (comboBox != null)
                    {
                        comboBox.IsDropDownOpen = true;
                        comboBox.Focus(); // 콤보박스에 포커스 주기
                        e.Handled = true; // 이벤트 처리 완료 표시
                    }
                }
                if (e.Key == Key.Space)
                {
                    e.Handled = true;

                
                    var rowData = dataGrid.Items[rowCount];

                    if (dataGrid.Name.Contains("Left"))
                    {
                        if (rowData is Win_ord_OutWare_Multi_U_dgdLEFT_CodeView leftData)
                        {
                            leftData.Chk = !leftData.Chk;
                        }
                    }
                    else if(dataGrid.Name.Contains("Right"))
                    {
                        if (rowData is Win_ord_OutWare_Multi_U_dgdRight_CodeView rightData)
                        {
                            rightData.Chk = !rightData.Chk;
                        }
                    }

                    return;
                }
                // Ctrl 키가 눌린 상태에서 방향키 처리
                if (Keyboard.Modifiers == ModifierKeys.Control)
                {
                    e.Handled = true;

                    // 2D 배열로 데이터그리드 배치 정의 (행과 열)
                    //dgdLeft의 row또는 cell이 선택된 상태에서 컨트롤 + 오른쪽 방향키 누르면 dgdRight로 이동
                    var gridLayout = new DataGrid[,] {
                                                        { dgdLeft, dgdRight },
                                                    
                                                     };

                    // 현재 데이터그리드의 위치 찾기
                    int currentRow = -1;
                    int currentCol = -1;

                    //GetLength, 0은 행의 갯수, 1은 열의 갯수입니다.
                    //행 숫자, 열 숫자만큼 반복해서 datagrid(이벤트를 발생시킨 데이터그리드)
                    //위치를 찾습니다.
                    for (int r = 0; r < gridLayout.GetLength(0); r++)
                    {
                        for (int c = 0; c < gridLayout.GetLength(1); c++)
                        {
                            if (gridLayout[r, c] == dataGrid)
                            {
                                currentRow = r;
                                currentCol = c;
                                break;
                            }
                        }
                        if (currentRow != -1) break;
                    }

                    if (currentRow != -1 && currentCol != -1) // 현재 데이터그리드가 배열에 있는 경우
                    {
                        int targetRow = currentRow;         //gridLayOut에 배치된 데이터그리드 행 위치
                        int targetCol = currentCol;         //gridLayOut에 배치된 데이터그리드 열 위치

                        // 방향키에 따라 다음 위치 계산
                        if (e.Key == Key.Down)
                        {
                            // 아래로 이동
                            //현재 데이터그리드 위치에서,  0 + 1 % 2  = 1 (첫번째 데이터그리드에서 바로 아래로)
                            // 1 + 1 % 2  = 0
                            targetRow = (currentRow + 1) % gridLayout.GetLength(0);
                        }
                        else if (e.Key == Key.Up)
                        {
                            // 위로 이동
                            targetRow = (currentRow - 1 + gridLayout.GetLength(0)) % gridLayout.GetLength(0);
                        }
                        else if (e.Key == Key.Right)
                        {
                            // 오른쪽으로 이동
                            targetCol = (currentCol + 1) % gridLayout.GetLength(1);
                        }
                        else if (e.Key == Key.Left)
                        {
                            // 왼쪽으로 이동
                            targetCol = (currentCol - 1 + gridLayout.GetLength(1)) % gridLayout.GetLength(1);
                        }

                        // 다른 데이터그리드로 이동하는 경우
                        if (targetRow != currentRow || targetCol != currentCol)
                        {
                            DataGrid targetGrid = gridLayout[targetRow, targetCol];
                            MoveToDataGrid(dataGrid, targetGrid);
                        }
                    }

                    return; // 처리 완료
                }

                else if (e.Key == Key.Down)
                {
                    e.Handled = true;
                    cell.IsEditing = false;
                    if (dataGrid.Items.Count - 1 > rowCount)
                    {
                        dataGrid.SelectedIndex = rowCount + 1;
                        dataGrid.CurrentCell = new DataGridCellInfo(dataGrid.Items[rowCount + 1], dataGrid.Columns[colCount]);
                    }
                    else if (dataGrid.Items.Count - 1 == rowCount)
                    {
                        if (lastColcount > colCount)
                        {
                            dataGrid.SelectedIndex = 0;
                            dataGrid.CurrentCell = new DataGridCellInfo(dataGrid.Items[0], dataGrid.Columns[colCount + 1]);
                        }
                    }
                }
                else if (e.Key == Key.Up)
                {
                    e.Handled = true;
                    cell.IsEditing = false;
                    if (rowCount > 0)
                    {
                        dataGrid.SelectedIndex = rowCount - 1;
                        dataGrid.CurrentCell = new DataGridCellInfo(dataGrid.Items[rowCount - 1], dataGrid.Columns[colCount]);
                    }
                }
                else if (e.Key == Key.Left)
                {
                    e.Handled = true;
                    cell.IsEditing = false;
                    if (colCount > 0)
                    {
                        dataGrid.CurrentCell = new DataGridCellInfo(dataGrid.Items[rowCount], dataGrid.Columns[colCount - 1]);
                    }
                }
                else if (e.Key == Key.Right)
                {
                    e.Handled = true;
                    cell.IsEditing = false;
                    if (lastColcount > colCount)
                    {
                        dataGrid.CurrentCell = new DataGridCellInfo(dataGrid.Items[rowCount], dataGrid.Columns[colCount + 1]);
                    }
                    else if (lastColcount == colCount)
                    {
                        if (dataGrid.Items.Count - 1 > rowCount)
                        {
                            dataGrid.SelectedIndex = rowCount + 1;
                  
                            int targetColIndex = 1; 
                            for (int i = 1; i < dataGrid.Columns.Count; i++)
                            {
                                if (!dataGrid.Columns[i].IsReadOnly &&
                                    dataGrid.Columns[i].Visibility == Visibility.Visible)
                                {
                                    targetColIndex = i;
                                    break;
                                }
                            }

                            dataGrid.CurrentCell = new DataGridCellInfo(dataGrid.Items[rowCount + 1], dataGrid.Columns[targetColIndex]);
                        }
                    }
                }
            }
        }

        // 한 데이터그리드에서 다른 데이터그리드로 이동하는 메서드
        private void MoveToDataGrid(DataGrid sourceGrid, DataGrid targetGrid)
        {
            targetGrid.Focus();
            if (targetGrid.Items.Count > 0)
            {
                int selectedIndex = sourceGrid.SelectedIndex;

                // 유효한 인덱스 계산
                if (selectedIndex >= 0 && selectedIndex < targetGrid.Items.Count)
                {
                    // 동일한 인덱스가 대상 그리드에 있는 경우
                    targetGrid.SelectedIndex = selectedIndex;
                    targetGrid.CurrentCell = new DataGridCellInfo(targetGrid.Items[selectedIndex], targetGrid.Columns[0]);
                }
                else
                {
                    // 인덱스가 범위를 벗어나면 마지막 항목으로 설정
                    int newIndex = Math.Min(targetGrid.Items.Count - 1, Math.Max(0, selectedIndex));
                    targetGrid.SelectedIndex = newIndex;
                    targetGrid.CurrentCell = new DataGridCellInfo(targetGrid.Items[newIndex], targetGrid.Columns[0]);
                }
            }
        }


        #endregion

        #region 플러스파인더 및 데이터그리드 선택 변경

        //메인 데이터그리드 선택 변경
        private void dgdOutware_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var OutwareInfo = dgdOutware.SelectedItem as Win_ord_OutWare_Multi_U_CodeView;

                if (OutwareInfo != null)
                {
                    this.DataContext = OutwareInfo;
                    FillGridRight(OutwareInfo.OutWareID);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutware_SelectionChanged : " + ee.ToString());
            }
        }


        //포장 재고량 검증(포커스 남은채로 저장했을때 대비)
        private void BoxQtyColumn_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox txtbox = sender as TextBox;
            if (txtbox == null || txtbox.DataContext == null)
                return;
            try
            {
                Win_ord_OutWare_Multi_U_CodeView MainData = this.DataContext as Win_ord_OutWare_Multi_U_CodeView;
                Win_ord_OutWare_Multi_U_dgdRight_CodeView rowData = txtbox.DataContext as Win_ord_OutWare_Multi_U_dgdRight_CodeView;
                var dgdLeftOriginInfo = new Dictionary<string, Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>();
                decimal boxQty = GetBoxQty(rowData.LabelID);
                //decimal boxQty = strFlag.Equals("I") ? GetBoxQty(rowData.LabelID) : GetBoxQty(rowData.LabelID) + GetoutwareIDQty(MainData.OutWareID,rowData.LabelID) ;
                //decimal boxQty = strFlag.Equals("I") ? GetBoxQty(rowData.LabelID) : GetBoxQty(rowData.LabelID) + rowData.OriginOutQty;


                decimal rollBackQty = 0;

                if (rowData != null)
                {
                    //조건 통합
                    if (!CalcuatePackQty(rowData, boxQty, out rollBackQty))
                    {
                        rowData.OutQty = rollBackQty;
                        txtbox.Focus();
                        txtbox.SelectAll();
                        return;
                    }


                    //왼쪽을 바꿀건데
                    if (dgdLeft.Items.Count > 0)
                    {
                        //왼쪽에 뭐 있었는지 저장
                        foreach (var leftItem in dgdLeft.Items.Cast<Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>())
                        {
                            dgdLeftOriginInfo[leftItem.LabelID] = leftItem;
                        }
                        //오른쪽에서 왼쪽과 똑같은게 뭐 있는지
                        if (dgdLeftOriginInfo.TryGetValue(rowData.LabelID, out Win_ord_OutWare_Multi_U_dgdLEFT_CodeView matchedLeftData))
                        {
                            //오른쪽에 있는게 만약 신규데이터(추가중이다)
                            bool isNewOut = string.IsNullOrEmpty(rowData.OutWareID);
                            decimal remainQty;

                            if (isNewOut)
                            {
                                // 새 출하는 남은 재고 - 출하량
                                remainQty = boxQty - rowData.OutQty;
                            }
                            else
                            {
                                // 기존 출하 수정은 남은 재고 + 기존 출하량 - 새 출하량
                                remainQty = boxQty + rowData.FixedOutQty - rowData.OutQty;
                            }

                            // 남은 수량이 0 이하면 왼쪽에서 제거
                            if (remainQty <= 0)
                            {
                                ovcDgdLeft.Remove(matchedLeftData);
                                dgdLeft.Items.Remove(matchedLeftData);
                            }
                            else
                            {
                                matchedLeftData.BoxQty = remainQty;
                                matchedLeftData.OutQty = remainQty;
                                matchedLeftData.OriginOutQty = remainQty;
                                rowData.BoxQty = rowData.OutQty;
                            }
                        }
                    }

                    SumScanQty();

                }
            }
            finally
            {
                e.Handled = true;  //루프 방지
            }
        }


        //포장재고량 검증(엔터키)
        private void BoxQtyColumn_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TextBox txtbox = sender as TextBox;
                if (txtbox == null || txtbox.DataContext == null)
                    return;
                try
                {
                    Win_ord_OutWare_Multi_U_CodeView MainData = this.DataContext as Win_ord_OutWare_Multi_U_CodeView;
                    Win_ord_OutWare_Multi_U_dgdRight_CodeView rowData = txtbox.DataContext as Win_ord_OutWare_Multi_U_dgdRight_CodeView;
                    var dgdLeftOriginInfo = new Dictionary<string, Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>();
                    decimal boxQty = GetBoxQty(rowData.LabelID);
                    //decimal boxQty = strFlag.Equals("I") ? GetBoxQty(rowData.LabelID) : GetBoxQty(rowData.LabelID) + GetoutwareIDQty(MainData.OutWareID, rowData.LabelID);

                    decimal rollBackQty = 0;

                    if (rowData != null)
                    {
                        //조건 통합
                        if (!CalcuatePackQty(rowData, boxQty, out rollBackQty))
                        {
                            rowData.OutQty = rollBackQty;
                            txtbox.Focus();
                            txtbox.SelectAll();
                            return;
                        }


                        //왼쪽을 바꿀건데
                        if (dgdLeft.Items.Count > 0)
                        {
                            //왼쪽에 뭐 있었는지 저장
                            foreach (var leftItem in dgdLeft.Items.Cast<Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>())
                            {
                                dgdLeftOriginInfo[leftItem.LabelID] = leftItem;
                            }
                            //오른쪽에서 왼쪽과 똑같은게 뭐 있는지
                            if (dgdLeftOriginInfo.TryGetValue(rowData.LabelID, out Win_ord_OutWare_Multi_U_dgdLEFT_CodeView matchedLeftData))
                            {
                                //오른쪽에 있는게 만약 신규데이터(추가중이다)
                                bool isNewOut = string.IsNullOrEmpty(rowData.OutWareID);
                                decimal remainQty;

                                if (isNewOut)
                                {
                                    // 새 출하는 남은 재고 - 출하량
                                    remainQty = boxQty - rowData.OutQty;
                                }
                                else
                                {
                                    // 기존 출하 수정은 남은 재고 + 기존 출하량 - 새 출하량
                                    remainQty = boxQty + rowData.FixedOutQty - rowData.OutQty;
                                }

                                // 남은 수량이 0 이하면 왼쪽에서 제거
                                if (remainQty <= 0)
                                {
                                    ovcDgdLeft.Remove(matchedLeftData);
                                    dgdLeft.Items.Remove(matchedLeftData);
                                }
                                else
                                {
                                    matchedLeftData.BoxQty = remainQty;
                                    matchedLeftData.OutQty = remainQty;
                                    matchedLeftData.OriginOutQty = remainQty;
                                    rowData.BoxQty = rowData.OutQty;
                                }
                            }
                        }

                        SumScanQty();

                    }
                }
                finally
                {
                    e.Handled = true; //LostFocus로 버블링 방지
                }
            }

        }

        //실시간 포장수량 계산
        //조건이 길어서 분리
        private bool CalcuatePackQty(Win_ord_OutWare_Multi_U_dgdRight_CodeView rowData, decimal boxQty, out decimal rollbackQty)
        {
            rollbackQty = 0;
            try
            {
                bool isNewOut = string.IsNullOrEmpty(rowData.OutWareID);

                // 잔량이 0인데 기존 출하량보다 더 출하하려는 경우
                if (boxQty.Equals(0) && rowData.OutQty > rowData.FixedOutQty)
                {
                    rollbackQty = rowData.FixedOutQty;
                    MessageBox.Show($"박스라벨:({rowData.LabelID})의 남은 잔량이 없어\n기존 출하량 ({stringFormatN0(rollbackQty)})을 초과할 수 없습니다.", "수량 초과");
                    return false;
                }
                // 기존 출하 데이터 수정: FixedOutQty + boxQty까지 허용
                else if (!isNewOut && rowData.OutQty > (rowData.FixedOutQty + boxQty))
                {
                    rollbackQty = rowData.FixedOutQty + boxQty;
                    MessageBox.Show($"박스라벨:({rowData.LabelID})의 허용 수량({stringFormatN0(rollbackQty)})을 초과할 수 없습니다.\n(기존 출하량 {stringFormatN0(rowData.FixedOutQty)} + 남은 잔량 {stringFormatN0(boxQty)})", "잔량 초과");
                    return false;
                }
                // 새 출하: boxQty까지만 허용
                else if (isNewOut && rowData.OutQty > boxQty)
                {
                    rollbackQty = boxQty;
                    MessageBox.Show($"박스라벨:({rowData.LabelID})의 남은 잔량({stringFormatN0(boxQty)})을 초과할 수 없습니다.", "잔량 초과");
                    return false;
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

        //박스라벨의 남은 수량 가져오기
        private decimal GetBoxQty(string BoxID)
        {
            decimal qty = 0;

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("LabelID", BoxID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sGetBoxQty", sqlParameter, false);


                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            qty = lib.RemoveComma(dr["BoxQty"].ToString(), 0m);
                        }
                    }
                }
            }
            catch
            {

            }

            return qty;
        }


        //수주와 ArticleID로 포장리스트 가져오기
        private void GetPackingList(Win_ord_OutWare_Multi_U_CodeView mainData)
        {
            try
            {

                if (!string.IsNullOrWhiteSpace(mainData.ArticleID) && !string.IsNullOrWhiteSpace(mainData.OrderID))
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("ArticleID", mainData.ArticleID);
                    sqlParameter.Add("OrderID", mainData.OrderID);

                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sPackingList", sqlParameter, false);

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];

                        if (dt.Rows.Count == 0)
                        {
                            MessageBox.Show("해당 수주로 검사/포장된 제품 또는 남은 포장재고가 없습니다.", "확인");
                            return;
                        }
                        else
                        {
                            DataRowCollection drc = dt.Rows; //프로시저에서 환원된 row를 컬렉션에 넣고

                            var dgdRightInfo = new Dictionary<string, Win_ord_OutWare_Multi_U_dgdRight_CodeView>();
                            var dgdLeftOriginInfo = new Dictionary<string, Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>();
                            if (dgdRight.Items.Count > 0)   //오른쪽에 데이터가 있으면
                            {
                                foreach (var rightItem in dgdRight.Items.Cast<Win_ord_OutWare_Multi_U_dgdRight_CodeView>()) //오른쪽 값을 캐스팅해서
                                {
                                    dgdRightInfo[rightItem.LabelID] = rightItem;                                            //기존데이터를 딕셔너리에 저장

                                }
                            }

                            //다시 그려주기전 비교를 위해 저장
                            if (dgdLeft.Items.Count > 0)    //왼쪽에 데이터가 있으면
                            {
                                foreach (var leftItem in dgdLeft.Items.Cast<Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>())    //마찬가지로 캐스팅해서 딕셔너리에 저장
                                {
                                    dgdLeftOriginInfo[leftItem.LabelID] = leftItem;

                                }
                            }

                            if (dgdLeft.ItemsSource != null) dgdLeft.ItemsSource = null;     //직접조작 오류 방지
                            ovcDgdLeft.Clear();
                            dgdLeft.Items.Clear();

                            int i = 0;

                            foreach (var leftItem in dgdLeftOriginInfo.Values)
                            {
                                // DB에서 dt로 받고 drc로 쪼갠 것에 해당 라벨이 없는 경우만 복원
                                bool existsInDB = drc.Cast<DataRow>().Any(dr => dr["LabelID"].ToString() == leftItem.LabelID);

                                if (!existsInDB)
                                {
                                    i++;
                                    leftItem.Num = i;
                                    ovcDgdLeft.Add(leftItem);
                                    dgdLeft.Items.Add(leftItem);
                                }
                            }

                            foreach (DataRow dr in drc)
                            {

                                decimal boxQty = lib.RemoveComma(dr["BoxQty"].ToString(), 0m);

                                var dgdLeftInfo = new Win_ord_OutWare_Multi_U_dgdLEFT_CodeView
                                {
                                    LabelID = dr["LabelID"].ToString(),
                                    OutQtySubulFromFn = lib.RemoveComma(dr["OutQty"].ToString(), 0m),    //재고함수에서 반환한 출하량(가장 정확)
                                    StuffinQty = lib.RemoveComma(dr["StuffINQty"].ToString(), 0m),
                                    UnitClssName = dr["UnitClssName"].ToString(),
                                    ArticleID = dr["ArticleID"].ToString() ?? string.Empty,
                                    Article = dr["Article"].ToString() ?? string.Empty,
                                    OrderID = dr["OrderID"].ToString() ?? string.Empty,
                                };

                                // 왼쪽에 기존 데이터가 있고 OutwareID가 있으면, 기존데이터 유지해서 계속 늘어나는거 방지
                                if (dgdLeftOriginInfo.TryGetValue(dr["LabelID"].ToString(), out Win_ord_OutWare_Multi_U_dgdLEFT_CodeView existsData)
                                 && !string.IsNullOrWhiteSpace(existsData.OutWareID))
                                {
                                    dgdLeftInfo.OutWareID = existsData.OutWareID;              //dgdLeftInfo는 DB에서 가져오는거니까 OutwareID 넣어주고 
                                    dgdLeftInfo.BoxQty = existsData.FixedOutQty + boxQty;      // 기존에 단건으로 내보낸 출하량 + 남은 재고량
                                    dgdLeftInfo.OutQty = existsData.FixedOutQty + boxQty;      // 동일 
                                    dgdLeftInfo.OriginOutQty = existsData.FixedOutQty + boxQty;
                                    dgdLeftInfo.FixedOutQty = existsData.FixedOutQty;
                                }
                                else
                                {
                                    dgdLeftInfo.OutWareID = string.Empty;
                                    dgdLeftInfo.BoxQty = boxQty;
                                    dgdLeftInfo.OutQty = boxQty;
                                    dgdLeftInfo.OriginOutQty = boxQty;
                                    dgdLeftInfo.FixedOutQty = 0;  // 새 출하는 0
                                }

                                //오른쪽에 데이터가 있을때                                
                                if (dgdRight.Items.Count > 0)
                                {
                                    //프로시저에서 불러온 라벨이 오른쪽에 이미 있으면
                                    if (dgdRightInfo.TryGetValue(dr["LabelID"].ToString(), out Win_ord_OutWare_Multi_U_dgdRight_CodeView matchedRightData))
                                    {
                                        //첫 출하이면(DB의 출하내역이 하나도 없을때)
                                        if (dgdLeftInfo.OutQtySubulFromFn.Equals(0))
                                        {
                                            dgdLeftInfo.BoxQty -= matchedRightData.OutQty; //그럼 재고에서 내보내고 싶은 양 빼기(오른쪽에서 수량 조정하면 OutQty 변경됨)
                                        }
                                        //첫 출하 이후
                                        else if (!dgdLeftInfo.OutQtySubulFromFn.Equals(0)) //출하한게 있으면
                                        {
                                            // DB에서 로드한 기존 출하 데이터를 수정하는 경우
                                            // 이때는 오른쪽에 OutwareID가 있으니까
                                            if (!string.IsNullOrEmpty(matchedRightData.OutWareID))
                                            {
                                                decimal totalQty = boxQty + matchedRightData.FixedOutQty; //재고 + 기존 출하량
                                                dgdLeftInfo.BoxQty = totalQty - matchedRightData.OutQty; //그중에서 내보내고 싶은 양 빼기
                                            }
                                            // 왼쪽에서 새로 옮겨온 데이터 (아직 저장 안 됨)
                                            // 이때는 재고에서 바로 빼기
                                            else
                                            {
                                                dgdLeftInfo.BoxQty = boxQty - matchedRightData.OutQty;
                                            }
                                        }

                                        // 공통 처리
                                        matchedRightData.BoxQty = matchedRightData.OutQty;

                                        // 0 이하면 왼쪽에 추가 안함
                                        if (dgdLeftInfo.BoxQty <= 0)
                                        {
                                            continue;
                                        }

                                        // 입고량만큼 출하하려는 경우도 패스
                                        if (matchedRightData.OutQty == dgdLeftInfo.StuffinQty)
                                        {
                                            continue;
                                        }
                                    }
                                }
                                //System.Diagnostics.Debug.WriteLine($"LabelID: {dgdLeftInfo.LabelID}, BoxQty: {dgdLeftInfo.BoxQty}, FixedOutQty: {dgdLeftInfo.FixedOutQty}");

                                i++;
                                dgdLeftInfo.Num = i;
                                ovcDgdLeft.Add(dgdLeftInfo);
                                dgdLeft.Items.Add(dgdLeftInfo);

                            }

                        }
                    }
                }
            }
            catch
            {

            }
        }


        #endregion

        #region Research
        private void re_Search(int rowNum)
        {
            try
            {
                dgdOutware.Items.Clear();

                ovcDgdLeft.Clear();
                ovcDgdRight.Clear();
                dgdLeft.Items.Clear();
                dgdRight.Items.Clear();
                ovcdgdMain.Clear();
                txtScan.Text = string.Empty;

                FillGrid();

                if (dgdOutware.Items.Count > 0)
                {
                    dgdOutware.SelectedIndex = rowNum;
                }
                else
                {
                    this.DataContext = null;
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

            dgdOutware.Items.Clear();
            dgdTotal.Items.Clear();

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", chkOutwareDay.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkOutwareDay.IsChecked == true ? dtpFromDate.SelectedDate?.ToString("yyyyMMdd") ?? string.Empty : string.Empty);
                sqlParameter.Add("EDate", chkOutwareDay.IsChecked == true ? dtpToDate.SelectedDate?.ToString("yyyyMMdd") ?? string.Empty : string.Empty);

                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkInCustomID", chkInCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("InCustomID", chkInCustomIDSrh.IsChecked == true ? txtInCustomIDSrh.Tag != null ? txtInCustomIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkOrder", chkOrderNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("Order", chkOrderNoSrh.IsChecked == true ? txtOrderNoSrh.Text : "");

                sqlParameter.Add("chkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : "" : "");

                ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOrder_Multi", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {

                        int i = 0;
                        decimal OutSumQty = 0;
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var OutwareInfo = new Win_ord_OutWare_Multi_U_CodeView()
                            {
                                Num = i,
                                OrderID = dr["OrderID"].ToString(),
                                OriginOrderID = dr["OrderID"].ToString(),
                                OutDate = lib.ToDateTime(dr["OutDate"].ToString()),
                                OutClss = dr["OutClss"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                OutCustomID = dr["OutCustomID"].ToString(),
                                OutCustom = dr["OutCustom"].ToString(),
                                OutWareID = dr["OutWareID"].ToString(),
                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                OutRoll = lib.RemoveComma(dr["OutRoll"].ToString(), 0),
                                OutQty = lib.RemoveComma(stringFormatN0(dr["OutQty"]), 0m),
                                FromLocID = dr["FromLocID"].ToString(),
                                Remark = dr["Remark"].ToString(),
                            };

                            dgdOutware.Items.Add(OutwareInfo);

                            OutSumQty += lib.RemoveComma(dr["OutQty"].ToString(), 0m);
                        }

                        if (dgdOutware.Items.Count > 0)
                        {
                            var total = new Win_ord_OutWare_Multi_U_dgdTotal_CodeView
                            {
                                OutRoll = i,
                                OutQtyTotal = OutSumQty,
                            };

                            dgdTotal.Items.Add(total);
                        }

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


        #region FillGridRight
        private void FillGridRight(string OutwareID)
        {
            try
            {
                if (dgdLeft.Items.Count > 0)
                {
                    ovcDgdLeft.Clear();
                    dgdLeft.Items.Clear();
                }
                if (dgdRight.Items.Count > 0)
                {
                    ovcDgdRight.Clear();
                    dgdRight.Items.Clear();
                }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("OutwareID", OutwareID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOutware_MultiSub_Right", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;
                        foreach (DataRow dr in drc)
                        {
                            decimal outQty = lib.RemoveComma(dr["OutQty"].ToString(), 0m);

                            var Search_Select = new Win_ord_OutWare_Multi_U_dgdRight_CodeView()
                            {
                                Num = i + 1,
                                OutWareID = dr["OutWareID"].ToString(),
                                LabelID = dr["LabelID"].ToString(),
                                OutQty = outQty,                                    //출하량 (사용자가 편집하면서 값 변경되는 것)
                                OriginOutQty = outQty,                              //데이터를 불러왔을 시점의 출하량
                                OutQtySubulFromFn = outQty,                     
                                UnitClssName = dr["UnitClssName"].ToString(),
                                CustomBoxID = dr["CustomBoxID"].ToString(),
                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                OrderID = dr["OrderID"].ToString(),
                                
                            };

                            ovcDgdRight.Add(Search_Select);
                            dgdRight.Items.Add(Search_Select);
                            i++;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("오류" + e.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }
        #endregion

        #region 저장
        private bool SaveData(string strFlag, Win_ord_OutWare_Multi_U_CodeView outData, List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> rightData)
        {    
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
            List<KeyValue> list_Result = new List<KeyValue>();
            string sGetID = string.Empty;

            try
            {
                if (CheckData(outData, rightData))
                {
                 
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        //화면에 있는 파라미터
                        sqlParameter.Add("OutWareID", outData.OutWareID ?? string.Empty);
                        sqlParameter.Add("OutSeq", outData.OutSeq);
                        sqlParameter.Add("OrderID", outData.OrderID ?? string.Empty);
                        sqlParameter.Add("OutClss", outData.OutClss);
                        sqlParameter.Add("OutDate", outData.OutDate?.ToString("yyyyMMdd"));
                        sqlParameter.Add("CustomID", outData.CustomID);
                        sqlParameter.Add("OutCustom", outData.OutCustom ?? string.Empty);
                        sqlParameter.Add("Memo", outData.Memo ?? string.Empty);
                        sqlParameter.Add("ToLocID", outData.TOLocID ?? "A0001");
                        sqlParameter.Add("FromLocID", outData.TOLocID ?? "A0001");
                        sqlParameter.Add("OutQty", outData.OutQty);
                        sqlParameter.Add("OutRealQty", outData.OutQty);
                        sqlParameter.Add("Amount", outData.Amount);
                        sqlParameter.Add("VatAmount", Math.Round(outData.VatAmount ?? outData.Amount * 0.1m, 0));
                        sqlParameter.Add("OutRoll", ovcDgdRight.Count);
                        sqlParameter.Add("UnitClss", outData.UnitClss ?? "0");


                        //널 비허용
                        sqlParameter.Add("WorkID", outData.WorkID ?? "0001");
                        sqlParameter.Add("ExchRate", outData.ExchRate);
                        sqlParameter.Add("ResultDate", outData.OutDate?.ToString("yyyyMMdd"));
                        sqlParameter.Add("Vat_Ind_YN", outData.Vat_Ind_YN ?? "N");
                        sqlParameter.Add("OutTime", DateTime.Now.ToString("HHmm"));
                        sqlParameter.Add("LoadTime", DateTime.Now.ToString("HHmm"));
                        sqlParameter.Add("TranNo", outData.TranNo ?? string.Empty);
                        sqlParameter.Add("TranSeq", outData.TranSeq);
                        sqlParameter.Add("OutType", outData.OutType ?? "3");
                        sqlParameter.Add("SetDate", outData.SetDate ?? DateTime.Today);
                        sqlParameter.Add("Remark", "사무실에서 출고");

                        //기본값

                        sqlParameter.Add("CompanyID", outData.CompanyID ?? "0001");
                        sqlParameter.Add("BuyerDirectYN", outData.BuyerDirectYN ?? "Y");
                        sqlParameter.Add("UnitPriceClss", outData.UnitPriceClss ?? "0");
                        sqlParameter.Add("InsStuffINYN", outData.InsStuffINYN ?? "N");
                        sqlParameter.Add("DvlyCustomID", outData.DvlyCustomID ?? outData.CustomID);
                        sqlParameter.Add("LossRate", outData.LossRate);
                        sqlParameter.Add("LossQty", outData.LossQty);
                        sqlParameter.Add("ArticleID", outData.ArticleID);

                        if (strFlag.Equals("I"))
                        {
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Outware_iOutware_Multi";
                            pro1.OutputUseYN = "Y";
                            pro1.OutputName = "OutWareID";
                            pro1.OutputLength = "12";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);

                            list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter, "C");


                            if (list_Result[0].key.ToLower() == "success")
                            {
                                list_Result.RemoveAt(0);
                                for (int i = 0; i < list_Result.Count; i++)
                                {
                                    KeyValue kv = list_Result[i];
                                    if (kv.key == "OutWareID")
                                    {
                                        sGetID = kv.value;

                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                                throw new Exception();
                            }

                            Prolist.Clear();
                            ListParameter.Clear();
                        }
                        else if (strFlag.Equals("U"))
                        {
                            sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Outware_uOutware_Multi";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "OutWareID";
                            pro1.OutputLength = "12";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);
                        }
             
                        foreach (var (subRow, index) in rightData.Select((item, idx) => (item, idx)))
                        {
                            sqlParameter = new Dictionary<string, object>();

                            sqlParameter.Add("OutWareID", string.IsNullOrEmpty(sGetID) ? outData.OutWareID : sGetID);
                            sqlParameter.Add("OrderID", outData.OrderID);
                            sqlParameter.Add("OrderSeq", subRow.OrderSeq);
                            sqlParameter.Add("OutSubSeq", index + 1);
                            sqlParameter.Add("LineSeq", 0);
                            sqlParameter.Add("LineSubSeq", 0);
                            sqlParameter.Add("RollSeq", index + 1);
                            sqlParameter.Add("LabelID", subRow.LabelID ?? string.Empty);
                            sqlParameter.Add("LabelGubun", "2");                    //박스출고가 2
                            sqlParameter.Add("LotNo", string.Empty);
                            sqlParameter.Add("StuffQty", 0);
                            sqlParameter.Add("OutRoll", subRow.OutRoll);
                            sqlParameter.Add("OutSeq", index + 1);
                            sqlParameter.Add("ArticleID", subRow.ArticleID);
                            //sqlParameter.Add("Spec", subRow.Spec ?? string.Empty);
                            sqlParameter.Add("UnitPrice", subRow.UnitPrice);
                            sqlParameter.Add("OutQty", subRow.OutQty);
                            sqlParameter.Add("OutRealQty", subRow.OutRealQty);
                            sqlParameter.Add("SetDate", DateTime.Today);
                            sqlParameter.Add("CustomBoxID", subRow.CustomBoxID ?? string.Empty);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            if (strFlag.Equals("U")) sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Outware_iOutwareSub_Multi";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "OutWareID";
                            pro1.OutputLength = "12";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);
                        }


                        list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter, "C");

                        if (list_Result[0].key.ToLower() != "success")
                        {
                            throw new Exception();
                        }

                        Prolist.Clear();
                        ListParameter.Clear();
                }
                else
                {
                    return false;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - SaveData : " + ex.ToString());
                return false;
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return true;
        }

        #endregion 저장

        #region 데이터 체크
        // 그룹박스 데이터 기입체크
        private bool CheckData(Win_ord_OutWare_Multi_U_CodeView mainData, List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> rightData)
        {
            try
            {
                string msg = string.Empty;

                if (string.IsNullOrWhiteSpace(mainData.OrderID))
                    msg += "오더번호가 입력되지 않았습니다. 오더번호를 검색 입력하세요";

                foreach (var item in rightData)
                {
                    if (item.OutQty == 0 || item.OutQty < 0)
                    {
                        msg += "출고 수량은 0 이상이어야 합니다. 수량을 확인하세요";
                        break;
                    }
                }

                if (!string.IsNullOrWhiteSpace(msg))
                {
                    MessageBox.Show(msg, "확인");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("저장 전 데이터 검증 중 오류\n" + ex.ToString());
                return false;
            }

            return true;

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




        #region LEFT RIGHT 그리드 좌우버튼 추가/삭제
        private void OVC_Remake_Select()
        {
            Win_ord_OutWare_Multi_U_dgdLEFT_CodeView RemainBoxList = null;
            Win_ord_OutWare_Multi_U_dgdRight_CodeView SelectedItems = null;
            int j = 0;
            bool overQty = false;

            List<Win_ord_OutWare_Multi_U_dgdLEFT_CodeView> itemsToRemove = new List<Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>();
            List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> dgdRightItems = dgdRight.Items.Cast<Win_ord_OutWare_Multi_U_dgdRight_CodeView>().ToList();

            List<string> duplicateLabels = new List<string>();
            List<Win_ord_OutWare_Multi_U_dgdLEFT_CodeView> checkedItems = new List<Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>();


            //오른쪽과 왼쪽과 같은 라벨이 존재할때
            for (int i = 0; i < dgdLeft.Items.Count; i++)
            {
                RemainBoxList = dgdLeft.Items[i] as Win_ord_OutWare_Multi_U_dgdLEFT_CodeView; //왼쪽 라벨 정보를 저장
                if (RemainBoxList.Chk == true)                                                //체크가 된거면
                {
                    checkedItems.Add(RemainBoxList);                                          //체크했다는 왼쪽 리스트 아이템에 넣음

                    if (dgdRightItems != null && dgdRightItems.Any(x => x.LabelID == RemainBoxList.LabelID)) //오른쪽 라벨정보가 왼쪽에도 있으면
                    {
                        if (!duplicateLabels.Contains(RemainBoxList.LabelID))                 //오른쪽과 왼쪽에 같이 있는 라벨이 있으면
                        {
                            duplicateLabels.Add(RemainBoxList.LabelID);                       //중복 라벨로 지정
                        }
                    }
                }
            }

            bool mergeDecision = false;
            if (duplicateLabels.Count > 0)
            {
                MessageBoxResult msgresult = MessageBox.Show($"우측 출하희망 목록에 등록된 라벨이 있습니다." +
                                                             $"\n({string.Join(", ", duplicateLabels)})\n남은 수량을 같은 라벨에 합치시겠습니까?"
                                                             , "확인", MessageBoxButton.YesNo);
                mergeDecision = (msgresult == MessageBoxResult.Yes);
            }

            //양쪽 같은 라벨이 있으면 라벨 수량 합치기
            foreach (var checkedItem in checkedItems) //왼쪽에 체크한 아이템들
            {
                if (duplicateLabels.Contains(checkedItem.LabelID)) //오른쪽에도 같은 라벨이 있으면
                {
                    //합치기로 했으면
                    if (mergeDecision)
                    {
                        //수량 셀 로스트포커스에서 만약 왼쪽 데이터그리드에 내용이 있으면 boxQty를 조정해주기로 했으므로
                        //기존 왼쪽의 남은 포장량(boxQty) + 내보내고 싶은량 (outQty)를 더하면 StuffinQty에 맞게 될 것
                        var existingItem = dgdRightItems.First(x => x.LabelID == checkedItem.LabelID);
                        decimal OutQtyWant = existingItem.OutQty + checkedItem.BoxQty;

                        existingItem.OutQty = OutQtyWant;
                        existingItem.OriginOutQty = OutQtyWant;
                        existingItem.BoxQty = OutQtyWant;


                        itemsToRemove.Add(checkedItem);
                    }
                }
                else
                {
                    // 새 라벨 추가
                    SelectedItems = new Win_ord_OutWare_Multi_U_dgdRight_CodeView()
                    {
                        Num = j + 1,
                        Chk = false,
                        OutWareID = !string.IsNullOrWhiteSpace(checkedItem.OutWareID) ? checkedItem.OutWareID : string.Empty,
                        LabelID = checkedItem.LabelID,
                        BoxQty = checkedItem.BoxQty,
                        OutQty = checkedItem.BoxQty,
                        OutRoll = checkedItem.OutRoll,
                        OriginOutQty = checkedItem.BoxQty,
                        FixedOutQty = checkedItem.FixedOutQty,
                        OutQtySubulFromFn = checkedItem.OutQtySubulFromFn,
                        Article = checkedItem.Article,
                        ArticleID = checkedItem.ArticleID,
                        UnitClssName = checkedItem.UnitClssName,
                        CustomBoxID = checkedItem.CustomBoxID ?? string.Empty,
                        OrderID = checkedItem.OrderID
                    };
                    ovcDgdRight.Add(SelectedItems);
                    itemsToRemove.Add(checkedItem);
                    j++;
                }
            }

            foreach (var item in itemsToRemove)
            {
                ovcDgdLeft.Remove(item);
            }
            dgdsubgrid_refill();

            if (overQty)
                lib.ShowTooltipMessage(dgdRight, "초과 입력된 수량은 자동 계산되었습니다.", MessageBoxImage.Information, PlacementMode.Top);

        }

        private void OVC_Remake_All()
        {
            Win_ord_OutWare_Multi_U_dgdLEFT_CodeView rightToLeft = null;
            Win_ord_OutWare_Multi_U_dgdRight_CodeView SelectedItems = null;
            int j = 0;
            List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> itemsToRemove = new List<Win_ord_OutWare_Multi_U_dgdRight_CodeView>();

            //기존에 왼쪽에 데이터가 있으면
            var dgdLeftOriginInfo = new Dictionary<string, Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>();
            if (dgdLeft.Items.Count > 0)
            {
                foreach (var leftItem in dgdLeft.Items.Cast<Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>())
                {
                    dgdLeftOriginInfo[leftItem.LabelID] = leftItem;

                }
            }

            for (int i = 0; i < dgdRight.Items.Count; i++)
            {

                SelectedItems = dgdRight.Items[i] as Win_ord_OutWare_Multi_U_dgdRight_CodeView;
                if (SelectedItems.Chk == true)
                {
                    rightToLeft = new Win_ord_OutWare_Multi_U_dgdLEFT_CodeView()
                    {
                        Num = j + 1,
                        Chk = false,
                        OutWareID = !string.IsNullOrWhiteSpace(SelectedItems.OutWareID) ? SelectedItems.OutWareID : string.Empty,
                        LabelID = SelectedItems.LabelID,
                        BoxQty = SelectedItems.OutQty,
                        OutQty = SelectedItems.OutQty,
                        OutRoll = SelectedItems.OutRoll,
                        OriginOutQty = SelectedItems.OriginOutQty,
                        FixedOutQty = SelectedItems.FixedOutQty,
                        Article = SelectedItems.Article,
                        ArticleID = SelectedItems.ArticleID,
                        OutQtySubulFromFn = SelectedItems.OutQtySubulFromFn,
                        UnitClssName = SelectedItems.UnitClssName,
                        CustomBoxID = SelectedItems.CustomBoxID,
                        OrderID = SelectedItems.OrderID
                    };

                    //왼쪽에 기존 데이터가 있고 오른쪽에서 넘겨줄 값이 매칭되는게 있으면
                    if (dgdLeftOriginInfo.TryGetValue(rightToLeft.LabelID, out Win_ord_OutWare_Multi_U_dgdLEFT_CodeView matchedLeftData))
                    {
                        //왼쪽에 있던거랑 오른쪽에 있던거 합쳐서 원복
                        matchedLeftData.BoxQty = rightToLeft.BoxQty + matchedLeftData.BoxQty;
                        matchedLeftData.OutQty = matchedLeftData.BoxQty;                       //윗줄에서 계산했으니까
                        matchedLeftData.OriginOutQty = matchedLeftData.BoxQty;

                        if (!string.IsNullOrEmpty(rightToLeft.OutWareID))
                        {
                            matchedLeftData.FixedOutQty = rightToLeft.FixedOutQty;
                            matchedLeftData.OutWareID = rightToLeft.OutWareID;
                        }

                    }
                    else
                    {
                        ovcDgdLeft.Add(rightToLeft);
                        j++;
                    }

                    itemsToRemove.Add(SelectedItems);

                }
            }

            foreach (var item in itemsToRemove)
            {
                ovcDgdRight.Remove(item);
            }

            dgdsubgrid_refill();
        }


        //다시그리기
        private void dgdsubgrid_refill()
        {
            int j = 0;
            int t = 0;

            if (dgdLeft.Items.Count > 0)
            {
                dgdLeft.Items.Clear();
            }
            if (dgdRight.Items.Count > 0)
            {
                dgdRight.Items.Clear();
            }
            for (j = 0; ovcDgdLeft.Count > j; j++)
            {
                var selectionItem = ovcDgdLeft[j] as Win_ord_OutWare_Multi_U_dgdLEFT_CodeView;
                selectionItem.Chk = false;
                selectionItem.Num = (j + 1);
                dgdLeft.Items.Add(selectionItem);
            }
            for (t = 0; ovcDgdRight.Count > t; t++)
            {
                var selectionItem = ovcDgdRight[t] as Win_ord_OutWare_Multi_U_dgdRight_CodeView;
                selectionItem.Chk = false;
                selectionItem.Num = (t + 1);
                dgdRight.Items.Add(selectionItem);
            }
        }
        #endregion

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

            //txtBuyerModel.IsHitTestVisible = false;  //차종은 땡겨오니까
            txtOutwareID.IsHitTestVisible = false;   //출고번호는 자동으로 생성되니까
            EventLabel.Visibility = Visibility.Visible; //자료입력중
            grbOutwareDetailBox.IsEnabled = true;       //DataContext Box
            dgdOutware.IsHitTestVisible = false;        //데이터그리드 클릭 안되게
            dgdLeft.IsHitTestVisible = true;
            dgdRight.IsHitTestVisible = true;
            btnAddSelectItem.IsEnabled = true;
            btnDelSelectItem.IsEnabled = true;

            btnSelectAllLeft.IsEnabled = true;
            btnSelectAllRight.IsEnabled = true;
            btnBoxSearch.IsEnabled = true;
            tbnCustomLabelID.IsEnabled = true;

            btnSelectAllLeft.Content = "전체선택";
            btnSelectAllRight.Content = "전체선택";
            txtScan.IsEnabled = true;


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

            //txtBuyerModel.IsHitTestVisible = false;  //차종은 땡겨오니까
            EventLabel.Visibility = Visibility.Hidden; //자료입력중
            grbOutwareDetailBox.IsEnabled = false;       //DataContext Box
            dgdOutware.IsHitTestVisible = true;        //데이터그리드 클릭되게
            dgdLeft.IsHitTestVisible = false;
            dgdRight.IsHitTestVisible = false;
            btnAddSelectItem.IsEnabled = false;
            btnDelSelectItem.IsEnabled = false;

            btnSelectAllLeft.IsEnabled = false;
            btnSelectAllRight.IsEnabled = false;
            btnBoxSearch.IsEnabled = false;
            tbnCustomLabelID.IsEnabled = false;

            btnSelectAllLeft.Content = "전체선택";
            btnSelectAllRight.Content = "전체선택";
            txtScan.IsEnabled = false;

        }



        private void SumScanQty()
        {
            try
            {
                if (this.DataContext == null) return;

                var mainItem = this.DataContext as Win_ord_OutWare_Multi_U_CodeView;
                decimal OutQty = 0;


                for (int i = 0; i < dgdRight.Items.Count; i++)
                {
                    var item = dgdRight.Items[i] as Win_ord_OutWare_Multi_U_dgdRight_CodeView;
                        OutQty += item.OutQty;
                }

                mainItem.OutQty = OutQty;
                mainItem.OutRoll = dgdRight.Items.Count;

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - SumScanQty : " + ee.ToString(), "경고");
            }
        }


      

        //수주 데이터 가져오기 
        private (Win_ord_OutWare_Multi_U_CodeView, List<Win_ord_OutWare_Multi_U_dgdRight_CodeView>) GetOrderData(string OrderID)
        {
            Win_ord_OutWare_Multi_U_CodeView outCodeView = new Win_ord_OutWare_Multi_U_CodeView();
            List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> outRightSub = new List<Win_ord_OutWare_Multi_U_dgdRight_CodeView>();

            try
            {
                try
                {
                    //이 쿼리에서 다 불러오자
                    //현재는 1:1의 경우만
                    string sql = "select " +
                                 "od.OrderID, mc.CustomID, mc.KCustom, InCustomID = mc1.CustomID, InCustom = mc1.KCustom, od.Remark " +
                                 ",ma.ArticleID, ma.Article, ma.BuyerArticleNo, ma.Spec " +
                                 "from [Order] od " +
                                 //"left join OrderColor oc on oc.OrderID = od.OrderID " +
                                 "left join mt_Custom mc on mc.CustomID = od.CustomID " +
                                 "left join mt_Custom mc1 on mc1.CustomID = od.InCustomID " +
                                 "left join mt_Article ma on ma.ArticleID = od.ArticleID " +
                                 //"left join mt_Article ma on ma.ArticleID = oc.ArticleID " +
                                 "where od.OrderID = @OrderID ";

                    var parameter = new Dictionary<string, object>
                    {
                        {"@OrderID", OrderID }
                    };

                    DataSet ds = DataStore.Instance.QueryToDataSetWithParam(sql, parameter);
                    if (ds != null)
                    {
                        DataTable dt = ds.Tables[0];
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow dr in drc)
                        {
                            var outware = new Win_ord_OutWare_Multi_U_CodeView
                            {
                                OrderID = OrderID,
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                OutCustomID = dr["InCustomID"].ToString(),
                                OutCustom = dr["InCustom"].ToString(),
                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Spec = dr["Spec"].ToString(),
                                OutDate = DateTime.Today,
                                TOLocID = "A0001",
                                Memo = dr["Remark"].ToString(),
                            };

                            outCodeView = outware;

                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("하위 정보 불러오기 실패했습니다" + ex.ToString());
                }


                //주석
                //출하선택에서는 자동으로 불러오기보다는 박스조회를 하고 수동으로 오른쪽으로 넘기는게 맞을듯
                //아래 코드는 오더번호(수주번호)를 불러오면 자동으로 우측데이터그리드에 포장데이터를 넘기려 했음
                //하지만 수정할때 출하스캔에서 넣은 수량기준과 뒤섞일것 같은 생각에 주석
                //try
                //{
                //    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                //    sqlParameter.Add("OrderID", OrderID);

                //    DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Outware_GetOrderData", sqlParameter, true, "R");


                //    if (ds != null && ds.Tables.Count > 0)
                //    {
                //        DataTable dt = ds.Tables[0];


                //        if (dt.Rows.Count > 0)
                //        {

                //            DataRowCollection drc = dt.Rows;
                //            foreach (var (dr, index) in drc.Cast<DataRow>().Select((row, idx) => (row, idx)))
                //            {
                //                var outwareSub = new Win_ord_OutWare_Multi_U_dgdRight_CodeView
                //                {
                //                    Num = index + 1,
                //                    OrderSeq = lib.RemoveComma(dr["OrderSeq"].ToString(), 0),
                //                    ArticleID = dr["ArticleID"].ToString(),
                //                    Article = dr["Article"].ToString(),
                //                    Spec = dr["Spec"].ToString(),
                //                    UnitPrice = lib.RemoveComma(dr["UnitPrice"].ToString(), 0m),
                //                    OutQty = lib.RemoveComma(dr["ColorQty"].ToString(), 0m),
                //                    OutRealQty = lib.RemoveComma(dr["ColorQty"].ToString(), 0m),
                //                };

                //                outwareSub.OutAmount = outwareSub.UnitPrice * outwareSub.OutQty;

                //                ovcDgdRight.Add(outwareSub);
                //            }

                //        }

                //    }
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show("OrderColor 데이터 불러오는 중 오류" + ex.ToString());
                //}        


            }
            catch
            {

            }

            return (outCodeView, outRightSub);
        }

        //스캔 텍스트박스 검증처리
        private void txtScan_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                try
                {
                    TextBox txtbox = sender as TextBox;
                    if (txtbox != null && !string.IsNullOrWhiteSpace(txtbox.Text) && tbnCustomLabelID.IsChecked == false)
                    {
                        if (!CheckScan(txtbox.Text ?? string.Empty))
                        {
                            txtbox.Text = string.Empty;
                            e.Handled = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("txtScan 입력 중 오류\n" + ex.ToString());
                }
            }
        }


        //검증을 통과하면 스캔값 처리
        private void txtScan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                try
                {
                    TextBox txtbox = (TextBox)sender;
                    if (tbnCustomLabelID.IsChecked == false)
                    {
                        var item = this.DataContext as Win_ord_OutWare_Multi_U_CodeView;

                        if (item != null && !string.IsNullOrWhiteSpace(txtbox.Text))
                        {
                            if (string.IsNullOrWhiteSpace(item.OrderID))
                            {
                                var ordData = GetOrderDataByLabel(txtbox.Text) as Win_ord_OutWare_Multi_U_CodeView;    //수주데이터를 가져옴
                                if (ordData != null)
                                {
                                    item.OrderID = ordData.OrderID;
                                    item.CustomID = ordData.CustomID;
                                    item.KCustom = ordData.KCustom;
                                    item.OutCustom = ordData.OutCustom;
                                    item.OutCustomID = ordData.OutCustomID;
                                    item.Article = ordData.Article;
                                    item.ArticleID = ordData.ArticleID;
                                    item.BuyerArticleNo = ordData.BuyerArticleNo;
                                }
                            }

                            BoxData boxData = new BoxData();
                            boxData = GetBoxData(txtbox.Text);      //라벨데이터 가져오기

                            var Subitem = new Win_ord_OutWare_Multi_U_dgdRight_CodeView         //서브그리드에 추가
                            {
                                Num = ovcDgdRight.Count + 1,
                                LabelID = boxData.LabelID,
                                OutQty = boxData.ColorQty < boxData.BoxQty ? boxData.ColorQty : boxData.BoxQty,
                                UnitPrice = boxData.UnitPrice,
                                OutRoll = 1,
                                ArticleID = boxData.ArticleID,
                                UnitClssName = boxData.UnitClssName,
                                OrderID = boxData.OrderID

                            };

                            ovcDgdRight.Add(Subitem);
                            dgdRight.Items.Add(Subitem);
                            txtbox.Text = string.Empty;
                            dgdRight.Items.Refresh();
                            txtbox.Focus();

                            SumScanQty();
                        }

                    }
                    else
                    {
                        List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> rightData = ovcDgdRight.ToList();
                        var existingItem = rightData.FirstOrDefault(x => string.IsNullOrWhiteSpace(x.CustomBoxID));
                        if (existingItem != null)
                        {
                            existingItem.CustomBoxID = txtbox.Text;
                            txtbox.Text = string.Empty;
                            txtbox.Focus();
                        }
                        else
                        {
                            lib.ShowTooltipMessage(txtbox, "우측 편 출하 데이터가 없습니다.",MessageBoxImage.Information, PlacementMode.Top, 1.3);
                        }

                    }

                }


                catch (Exception ex)
                {
                    MessageBox.Show("Label을 통한 생산정보, 수주정보를 가져오는 도중 오류 txtScan_KeyDown\n" + ex.ToString());
                }

            }
        }

        //라벨로 수주데이터 찾기(관리번호에서 입력 후 입력 컨트롤에 데이터 뿌리기)
        private Win_ord_OutWare_Multi_U_CodeView GetOrderDataByLabel(string LabelID)
        {
            Win_ord_OutWare_Multi_U_CodeView ordData = new Win_ord_OutWare_Multi_U_CodeView();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("BoxID", LabelID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sBoxIDOne_Multi", sqlParameter, false);


                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        return ordData = new Win_ord_OutWare_Multi_U_CodeView()
                        {
                            OrderID = dr["OrderID"].ToString(),
                            Article = dr["Article"].ToString(),
                            ArticleID = dr["ArticleID"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            CustomID = dr["CustomID"].ToString(),
                            KCustom = dr["CustomName"].ToString(),
                            OutCustomID = dr["OutCustomID"].ToString(),
                            OutCustom = dr["OutCustom"].ToString(),

                        };

                    }
                }
            }
            catch
            {

            }

            return ordData;
        }


        //스캔할때 검증
        private bool CheckScan(string LabelID)
        {
            var item = this.DataContext as Win_ord_OutWare_Multi_U_CodeView;
            BoxData boxdata = new BoxData();
            string msg = string.Empty;

           
            if (!LabelPrefix.Contains(LabelID.Substring(0, 1)))
            {
                msg += "올바른 바코드가 아닙니다.";
            }
            else if (item != null)
            {
                boxdata = GetBoxData(LabelID);

                if (!string.IsNullOrWhiteSpace(item.ArticleID) && boxdata.ArticleID != item.ArticleID)
                    msg += $"같은 품목의 데이터만 등록 할 수 있습니다.\n관리번호 품목명:({item.Article})\n스캔한 품목명:({boxdata.Article})";
                else if (boxdata.BoxQty <= 0) 
                {
                    if (!string.IsNullOrWhiteSpace(boxdata.RecentOutDate?.ToString("yyyy-MM-dd")))
                        msg += $"바코드번호 ({LabelID})는\n({boxdata.RecentOutDate?.ToString("yyyy-MM-dd")})에\n마지막으로 출하 후, 남은 잔량이 없습니다.";
                    else
                        msg += $"바코드번호 ({LabelID})는 생산되지 않거나 남은 잔량이 없습니다.";

                }

                else if (ovcDgdRight.Any(x => x.LabelID == LabelID))
                    msg += $"이미 추가된 바코드번호 입니다.";
                else if (!string.IsNullOrWhiteSpace(item.OrderID))
                {
                    MessageBoxResult result = MessageBoxResult.None;


                    /*if(tbnCustomLabelID.IsChecked == true && ovcDgdRight.Count == 0)
                    {
                        msg += "출하 데이터 추가 후 시도해 주세요";
                    }
                    else*/ if (ovcDgdRight.Count > 0 && ovcDgdRight.Any(x => string.IsNullOrWhiteSpace(x.LabelID)))
                    {
                        msg += "수량기준 출하 데이터가 있으므로 할 수 없습니다.";
                    }
                    else if (!string.IsNullOrWhiteSpace(item.OrderID) && !string.IsNullOrWhiteSpace(boxdata.OrderID) && boxdata.OrderID != item.OrderID)
                    {
                        string addMsg = $"[입력중인 오더번호]\n▷{item.OrderID}";

                        if (boxdata.CustomID != item.CustomID)
                        {
                            addMsg += $"\n▷거래처 : {item.KCustom}\n\n";
                            addMsg += $"[스캔한 바코드 관리번호]\n▷{boxdata.OrderID}\n▷거래처 : {boxdata.KCustom}\n\n";
                            addMsg += "다른 거래처의 바코드, 다른 수주입니다.";
                        }
                        else
                        {
                            addMsg += $"\n\n[스캔한 바코드 오더번호]\n▷{boxdata.OrderID}\n\n";
                            addMsg += "다른 수주입니다.";
                        }
                        addMsg += "\n진행하시겠습니까?";

                        result = MessageBox.Show(
                            addMsg,
                            "관리번호 불일치",
                            MessageBoxButton.YesNo);

                        if (result == MessageBoxResult.No)
                            msg += "작업이 취소되었습니다.";
                    }
                 


                }
            }       

            if (!string.IsNullOrWhiteSpace(msg))
            {
                MessageBox.Show(msg, "확인");
                return false;
            }

            return true;
        }


        //현재 라벨의 정보 가져오기
        private BoxData GetBoxData(string LabelID)
        {
            BoxData boxdata = new BoxData();
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("LabelID", LabelID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sGetBoxQty", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        return boxdata = new BoxData()
                        {
                            ArticleID = dr["ArticleID"].ToString(),
                            Article = dr["Article"].ToString(),
                            LabelID = dr["LotID"].ToString(),
                            StuffinQty = lib.RemoveComma(dr["StuffINQty"].ToString(), 0m),
                            OutQty = lib.RemoveComma(dr["OutQty"].ToString(), 0m),
                            BoxQty = lib.RemoveComma(dr["BoxQty"].ToString(), 0m),
                            OrderID = dr["OrderID"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            UnitClssName = dr["UnitClssName"].ToString(),
                            UnitPrice = lib.RemoveComma(dr["UnitPrice"].ToString(), 0m),
                            ColorQty = lib.RemoveComma(dr["ColorQty"].ToString(), 0m),
                            Spec = dr["Spec"].ToString(),
                            KCustom = dr["KCustom"].ToString(),
                            CustomID = dr["CustomID"].ToString(),
                            RecentOutDate = lib.ToDateTime(dr["OutDate"].ToString())
                        };

                    }
                }
            }
            catch
            {

            }

            return boxdata;

        }

        // 천자리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
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

        private void txtOutRoll_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.ImeProcessed) { e.Handled = true; }
        }

        private void txtOutQty_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.ImeProcessed) { e.Handled = true; }
        }


        private void DataGridCell_KeyUp(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        private void DataGridCell_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (EventLabel.Visibility == Visibility.Visible || tbkMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                if (cell.IsReadOnly != true)
                {
                    cell.IsEditing = true;
                }
                else
                {
                    cell.IsEditing = false;
                }
            }
        }


        private void chkReqDgdMain_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;

            var dgdAll = chkSender.DataContext as Win_ord_OutWare_Multi_U_CodeView;
            if (chkSender.IsChecked == true)
            {
                dgdAll.Chk = true;
                ovcdgdMain.Add(dgdAll);
                chkSender.IsChecked = true;
            }
            else
            {
                dgdAll.Chk = false;
                ovcdgdMain.Remove(dgdAll);
                chkSender.IsChecked = false;
            }

        }

        private void chkReqDgdLeft_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            if (tbkMsg.Visibility == Visibility.Visible)
            {
                var dgdAll = chkSender.DataContext as Win_ord_OutWare_Multi_U_dgdLEFT_CodeView;
                if (chkSender.IsChecked == true)
                {
                    dgdAll.Chk = true;
                }
                else
                {
                    dgdAll.Chk = false;
                }
            }
        }

        private void chkReqDgdRight_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            if (tbkMsg.Visibility == Visibility.Visible)
            {
                var dgdAll = chkSender.DataContext as Win_ord_OutWare_Multi_U_dgdRight_CodeView;
                if (chkSender.IsChecked == true)
                {
                    dgdAll.Chk = true;
                }
                else
                {
                    dgdAll.Chk = false;
                }
            }
            else
            {
                if (chkSender.IsChecked == true)
                {
                    chkSender.IsChecked = false;
                }
                else
                {
                    chkSender.IsChecked = true;
                }
                MessageBox.Show("체크박스를 사용하려면 먼저 추가나 수정을 누르고 진행해야 합니다.");
            }
        }

        //항목 오른쪽 넣기
        private void btnAddSelectItem_Click(object sender, RoutedEventArgs e)
        {
            bool flag = false;

            for (int i = 0; i < dgdLeft.Items.Count; i++)
            {
                var dgdLeftItem = dgdLeft.Items[i] as Win_ord_OutWare_Multi_U_dgdLEFT_CodeView;
                if (dgdLeftItem.Chk == true)
                {
                    flag = true;
                    break;
                }
            }

            if (!flag)
            {
                MessageBox.Show("체크된 항목이 없습니다.\n항목을 체크 후 눌러주세요");
                return;
            }


            OVC_Remake_Select();
            SumScanQty();

            if (dgdLeft.Items.Count == 0)
            {
                btnSelectAllLeft.Content = "전체선택";
            }

        }

        //항목 왼쪽으로 빼기
        private void btnDelSelectItem_Click(object sender, RoutedEventArgs e)
        {
            bool flag = true;
            string Errmsg = string.Empty;
            bool hasCheckedItem = false;  // 체크된 항목 존재 여부

            for (int i = 0; i < dgdRight.Items.Count; i++)
            {
                var dgdRightItem = dgdRight.Items[i] as Win_ord_OutWare_Multi_U_dgdRight_CodeView;
                if (dgdRightItem.Chk == true)
                {
                    hasCheckedItem = true;
                    if (string.IsNullOrEmpty(dgdRightItem.LabelID))
                    {
                        flag = false;
                        Errmsg = "선택 항목 중 수량출고로 처리한 항목이 있습니다.";
                        break;
                    }
                }
            }

            if (!hasCheckedItem)
            {
                flag = false;
                Errmsg = "체크된 항목이 없습니다.";
            }

            if (!flag)
            {
                MessageBox.Show($"{Errmsg}\n항목을 확인 후 눌러주세요", "확인");
                return;
            }

            OVC_Remake_All();
            SumScanQty();
            if (dgdRight.Items.Count == 0)
            {
                btnSelectAllRight.Content = "전체선택";
            }
        }

        private void btnBoxSearch_Click(object sender, RoutedEventArgs e)
        {
            var item = this.DataContext as Win_ord_OutWare_Multi_U_CodeView;
            if (this.DataContext == null || item == null) return;

            if (ovcDgdRight.Count > 0)
            {
                if (ovcDgdRight.Any(x => string.IsNullOrWhiteSpace(x.LabelID)))
                {
                    MessageBox.Show("수량기준 데이터로 저장된 출하건 입니다. 해당기능을 사용할 수 없습니다.","확인");
                    return;
                }            
            }
                StartCountdown();
                GetPackingList(item);
           

        }

        //박스조회 버튼 마구누르기 방지
        private void StartCountdown()
        {
            btnBoxSearch.IsEnabled = false;
            countdownSeconds = 2;

            if (currentTimer == null)
                currentTimer = new System.Windows.Threading.DispatcherTimer();

            currentTimer.Tick -= CountdownTimer_Tick;
            currentTimer.Tick += CountdownTimer_Tick;
            currentTimer.Interval = TimeSpan.FromSeconds(1);

            UpdateButtonText();
            currentTimer.Start();
        }

        private void CountdownTimer_Tick(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(strFlag))
            {
                currentTimer.Stop();
                if (txtBoxSearch != null)
                    txtBoxSearch.Text = "박스조회";
                if (btnBoxSearch != null)
                    btnBoxSearch.IsEnabled = false;
                return;
            }

            countdownSeconds--;

            if (countdownSeconds > 0)
            {
                if (txtBoxSearch != null)
                {
                    UpdateButtonText();
                }
                else
                {

                    currentTimer.Stop();
                }
            }
            else
            {
                currentTimer.Stop();
                if (txtBoxSearch != null)
                    txtBoxSearch.Text = "박스조회";
                if (btnBoxSearch != null)
                    btnBoxSearch.IsEnabled = true;
            }
        }


        private void UpdateButtonText()
        {
            if (countdownSeconds == 0)
            {
                txtBoxSearch.Text = $"박스조회";

            }
            else
            {
                txtBoxSearch.Text = $"박스조회(..{countdownSeconds})";
            }
        }




        //왼쪽 그리드용 전체선택 토글버튼
        private void btnSelectAllLeft_Click(object sender, RoutedEventArgs e)
        {
            if (btnSelectAllLeft.IsEnabled == true)
            {

                bool allChecked = true;

                foreach (var item in dgdLeft.Items)
                {
                    if (!(item as Win_ord_OutWare_Multi_U_dgdLEFT_CodeView).Chk)
                    {
                        allChecked = false;
                        break;
                    }
                }

                if (allChecked)
                {
                    foreach (var item in dgdLeft.Items)
                    {
                        (item as Win_ord_OutWare_Multi_U_dgdLEFT_CodeView).Chk = false;
                    }
                    btnSelectAllLeft.Content = "전체선택";
                }
                else
                {
                    foreach (var item in dgdLeft.Items)
                    {
                        (item as Win_ord_OutWare_Multi_U_dgdLEFT_CodeView).Chk = true;
                    }
                    btnSelectAllLeft.Content = "선택해제";
                }
            }
        }

        //오른쪽 그리드용 전체선택 토글버튼
        private void btnSelectAllRight_Click(object sender, RoutedEventArgs e)
        {

            if (btnSelectAllRight.IsEnabled == true)
            {
                bool allChecked = true;

                foreach (var item in dgdRight.Items)
                {
                    if (!(item as Win_ord_OutWare_Multi_U_dgdRight_CodeView).Chk)
                    {
                        allChecked = false;
                        break;
                    }
                }

                if (allChecked)
                {
                    foreach (var item in dgdRight.Items)
                    {
                        (item as Win_ord_OutWare_Multi_U_dgdRight_CodeView).Chk = false;
                    }
                    btnSelectAllRight.Content = "전체선택";
                }
                else
                {
                    foreach (var item in dgdRight.Items)
                    {
                        (item as Win_ord_OutWare_Multi_U_dgdRight_CodeView).Chk = true;
                    }
                    btnSelectAllRight.Content = "선택해제";
                }
            }
        }

        //왼쪽 그리드용 체크선택
        private void CheckBox_CheckedChangedLeft(object sender, RoutedEventArgs e)
        {
            bool allCheckedLeft = dgdLeft.Items.Cast<Win_ord_OutWare_Multi_U_dgdLEFT_CodeView>().All(item => item.Chk);

            if (allCheckedLeft)
            {
                btnSelectAllLeft.Content = "선택해제";
            }

        }

        //왼쪽 그리드용 체크선택 해제
        private void CheckBox_UnCheckedChangedLeft(object sender, RoutedEventArgs e)
        {
            btnSelectAllLeft.Content = "전체선택";
        }


        //오른쪽용 그리드 체크선택
        private void CheckBox_CheckedChangedRight(object sender, RoutedEventArgs e)
        {
            bool allCheckedRight = dgdRight.Items.Cast<Win_ord_OutWare_Multi_U_dgdRight_CodeView>().All(item => item.Chk);

            if (allCheckedRight)
            {
                btnSelectAllRight.Content = "선택해제";
            }
        }

        //오른쪽용 그리드 체크선택 해제
        private void CheckBox_UnCheckedChangedRight(object sender, RoutedEventArgs e)
        {
            btnSelectAllRight.Content = "전체선택";
        }   
        private void tbnCustomLabelID_Click(object sender, RoutedEventArgs e)
        {
            lib.ShowTooltipMessage(tbnCustomLabelID, "입력한 바코드가 고객라벨에 입력됩니다.", MessageBoxImage.Information, PlacementMode.Top);
            if (tbnCustomLabelID.IsChecked == true)
            {
                txtScan.Background = Brushes.AliceBlue;
            }
            else
            {
                txtScan.Background = Brushes.LightGreen;
            }
        }

        //다른 수주 포장건을 불러와서 출하시킨경우 셀 도움말
        private void DataGridRow_MouseEnter(object sender, MouseEventArgs e)
        {
            DataGridRow row = (DataGridRow)sender;
            var rowData = row.DataContext as Win_ord_OutWare_Multi_U_dgdRight_CodeView;
            var item = this.DataContext as Win_ord_OutWare_Multi_U_CodeView;

            if (rowData != null && item != null)
            {
                if (!string.IsNullOrWhiteSpace(item.OrderID) && !string.IsNullOrWhiteSpace(rowData.LabelID) && item.OrderID != rowData.OrderID)
                    lib.ShowTooltipMessage(row, $"현재 수주와 다른 수주번호\n({rowData.OrderID}로 생산된 라벨입니다.)", MessageBoxImage.Warning, PlacementMode.Top, 1.2);
            }

        }
        //돔말 닫기
        private void DataGridRow_MouseLeave(object sender, MouseEventArgs e)
        {
            lib.CloseToolTip();
        }


        private class Win_ord_OutWare_Multi_U_CodeView : BaseView
        {
            public int Num { get; set; }
            public bool Chk { get; set; }
            public string OutWareID { get; set; }
            public string CompanyID { get; set; }
            public string OriginOrderID { get; set; }
            public string OrderID { get; set; }
            public int OutSeq { get; set; }
            public int OrderSeq { get; set; }
            public string OrderNo { get; set; }
            public string CustomID { get; set; }

            private string _kcustom;
            public string KCustom
            {
                get => _kcustom;
                set
                {
                    _kcustom = value;
                    if (string.IsNullOrEmpty(value))
                    {
                        CustomID = null;
                    }
                }
            }

            public DateTime? OutDate { get; set; } = DateTime.Now;
            public string ArticleID { get; set; }
            public string Article { get; set; }
            public string Spec { get; set; }
            public string OutClss { get; set; }
            public string WorkID { get; set; }
            public decimal OutRoll { get; set; }
            public decimal OutQty { get; set; }
            public decimal OutRealQty { get; set; }
            public DateTime? ResultDate { get; set; } = DateTime.Now;
            public decimal OrderQty { get; set; }
            public string UnitClss { get; set; }
            public string WorkName { get; set; }
            public string OutType { get; set; }
            public string Remark { get; set; }
            public string BuyerModel { get; set; }
            public decimal OutSumQty { get; set; }
            public string OutQtyY { get; set; }
            public decimal StuffinQty { get; set; }
            public decimal OutWeight { get; set; }
            public decimal OutRealWeight { get; set; }
            public string UnitPriceClss { get; set; }
            public string BuyerDirectYN { get; set; }
            public string Vat_Ind_YN { get; set; }
            public string workID { get; set; }
            public string InsStuffINYN { get; set; }
            public double ExchRate { get; set; }
            public string FromLocID { get; set; }
            public string TOLocID { get; set; }
            public string UnitClssName { get; set; }
            public string FromLocName { get; set; }
            public string TOLocname { get; set; }
            public string OutClssname { get; set; }
            public decimal UnitPrice { get; set; }
            public decimal Amount { get; set; }
            public decimal OutAmount { get; set; }
            public decimal? VatAmount { get; set; }
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
            public string DvlyCustomID { get; set; }
            public double LossRate { get; set; }
            public decimal LossQty { get; set; }
            public string OutSubType { get; set; }
            public string OutTime { get; set; }
            public string LoadTime { get; set; }
            public string TranNo { get; set; }
            public int TranSeq { get; set; }
            public DateTime? SetDate { get; set; } = DateTime.Now;

            public decimal RemainQty { get; set; }
            public string Phone1 { get; set; }
            public string FaxNo { get; set; }
            public string CompanyNo { get; set; } //공급자 사업자 번호
            public string Memo { get; set; }

        }

        private class Win_ord_OutWare_Multi_U_dgdLEFT_CodeView : BaseView
        {
            public bool Chk { get; set; }
            public int Num { get; set; }
            public string OutWareID { get; set; }
            public string OrderID { get; set; }
            public int OrderSeq { get; set; }
            public int OutRoll { get; set; }
            public string LabelID { get; set; }
            public decimal BoxQty { get; set; }
            public string ArticleID { get; set; }
            public string Article { get; set; }
            public string Spec { get; set; }
            public decimal OutAmount { get; set; }
            public decimal UnitPrice { get; set; }
            public decimal OutRealQty { get; set; }
            public decimal OutQty { get; set; }                  //오른쪽 그리드가 쓰는 출고수량
            public decimal OutQtySubulFromFn { get; set; }        //재고함수로 나온 DB출고수량(한 labelID로 내보낸 전체 수량)
            public decimal StuffinQty { get; set; }
            //public decimal OriginPackedQty { get; set; }
            public decimal OriginOutQty { get; set; } //기존 출고 수량 근데 왼쪽 오른쪽 왔다갔다하며 변동될 수 있음
            public decimal FixedOutQty { get; set; }  //fillgridRight할때 조회된 수량
            public string UnitClssName { get; set; }
            public string CustomBoxID { get; set; }
        }

        private class Win_ord_OutWare_Multi_U_dgdRight_CodeView : BaseView
        {
            public bool Chk { get; set; }
            public int Num { get; set; }
            public string OutWareID { get; set; }
            public string OrderID { get; set; }
            public int OrderSeq { get; set; }
            public int OutRoll { get; set; }
            public string LabelID { get; set; }
            public decimal BoxQty { get; set; }
            public string ArticleID { get; set; }
            public string Article { get; set; }
            public string Spec { get; set; }
            public decimal OutAmount { get; set; }
            public decimal UnitPrice { get; set; }
            public decimal OutRealQty { get; set; }
            public decimal OutQty { get; set; }                  //오른쪽 그리드가 쓰는 출고수량
            public decimal OutQtySubulFromFn { get; set; }        //재고함수로 나온 DB출고수량(한 labelID로 내보낸 전체 수량)
            public decimal StuffinQty { get; set; }
            //public decimal OriginPackedQty { get; set; }
            public decimal OriginOutQty { get; set; } //기존 출고 수량 근데 왼쪽 오른쪽 왔다갔다하며 변동될 수 있음
            public decimal FixedOutQty { get; set; }  //fillgridRight할때 조회된 수량
            public string UnitClssName { get; set; }
            public string CustomBoxID { get; set; }

            //거래명표때 쓸
            public static List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> GetOutwareSubData(string outwareID)
            {
                List<Win_ord_OutWare_Multi_U_dgdRight_CodeView> lstOutwareSub = new List<Win_ord_OutWare_Multi_U_dgdRight_CodeView>();

                try
                {
                    string sql = "select  *                                               " +
                                 ",ma.Article, ma.Spec                                    " +
                                 "from OutwareSub ows                                     " +
                                 "left join mt_Article ma on ma.ArticleID = ows.ArticleID " +
                                 "where OutWareID like @OutWareID                         ";

                    var parameter = new Dictionary<string, object>
                    {
                        {"@OutWareID", outwareID }
                    };

                    DataSet ds = DataStore.Instance.QueryToDataSetWithParam(sql, parameter);
                    if (ds != null)
                    {
                        DataTable dt = ds.Tables[0];
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow dr in drc)
                        {
                            var outwareSub = new Win_ord_OutWare_Multi_U_dgdRight_CodeView
                            {
                                OutWareID = dr["OutWareID"].ToString(),
                                OutQty = Lib.Instance.RemoveComma(dr["OutQty"].ToString(), 0m),
                                UnitPrice = Lib.Instance.RemoveComma(dr["UnitPrice"].ToString(), 0m),
                                OutAmount = (Lib.Instance.RemoveComma(dr["OutQty"].ToString(), 0m) * Lib.Instance.RemoveComma(dr["UnitPrice"].ToString(), 0m)),
                                Article = dr["Article"].ToString(),
                                Spec = dr["Spec"].ToString(),

                            };

                            lstOutwareSub.Add(outwareSub);

                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("하위 정보 불러오기 실패했습니다" + ex.ToString());
                }
                finally
                {
                    DataStore.Instance.CloseConnection();
                }

                return lstOutwareSub;
            }
        }


        private class Win_ord_OutWare_Multi_U_dgdTotal_CodeView : BaseView
        {
            public int OutRoll { get; set; }
            public decimal OutQtyTotal { get; set; }
            public decimal OutPriceTotal { get; set; }
        }

        private class BoxData : BaseView
        {
            public string LabelID { get; set; }
            public decimal BoxQty { get; set; }
            public decimal StuffinQty { get; set; }
            public decimal OutQty { get; set; }
            public string ArticleID { get; set; }
            public string Article { get; set; }
            public string UnitClss { get; set; }
            public string UnitClssName { get; set; }
            public string OrderID { get; set; }
            public decimal UnitPrice { get; set; }
            public decimal ColorQty { get; set; }
            public string Spec { get; set; }
            public string KCustom { get; set; }
            public string CustomID { get; set; }
            public DateTime? RecentOutDate { get; set; }
        }

        private class SetCompanyData : BaseView
        {
            public string companyID { get; set; }
            public string chief { get; set; }
            public string kCompany { get; set; }
            public string companyNo { get; set; }
            public string address1 { get; set; }
            public string address2 { get; set; }
            public string phone1 { get; set; }
            public string faxNo { get; set; }

            //자사정보 불러오기
            public static SetCompanyData GetSetCompanyData()
            {
                SetCompanyData setCompanyData = new SetCompanyData();

                try
                {
                    //string sql = "select * from mt_setCompany where KCompany like '%' + @KCompany + '%' ";
                    string sql = "select top 1 * from mt_setCompany ";

                    //var parameter = new Dictionary<string, object>
                    //{
                    //    {"@KCompany", "신영" }
                    //};

                    DataSet ds = DataStore.Instance.QueryToDataSetWithParam(sql);
                    if (ds != null)
                    {
                        DataTable dt = ds.Tables[0];
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow dr in drc)
                        {
                            setCompanyData.kCompany = dr["KCompany"].ToString();
                            setCompanyData.companyNo = dr["CompanyNo"].ToString();
                            setCompanyData.chief = dr["chief"].ToString();
                            setCompanyData.address1 = dr["Address1"].ToString();
                            setCompanyData.address2 = dr["Address2"].ToString();
                            setCompanyData.phone1 = dr["Phone1"].ToString();
                            setCompanyData.faxNo = dr["FaxNo"].ToString();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("자사 정보를 불러오는 도중 오류\n" + ex.ToString());
                }
                finally
                {
                    DataStore.Instance.CloseConnection();
                }

                return setCompanyData;
            }

        }

   
    }


}


