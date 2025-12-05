/**
 * 
 * @details 작업 스케줄 관리 
 * @author 김수정
 * @date 2024-01-19
 * @version 1.0
 * 
 * @section MODIFYINFO 수정정보
 * - 수정일        - 수정자       : 수정내역
 * 
 * 
 * */

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


namespace WizMes_SungShinNQ
{

    public partial class Win_prd_ProcessRunningSchedule_U : UserControl
    {
        bool update = false;

        List<Win_prd_PrRunnig_U> lst = new List<Win_prd_PrRunnig_U>();

        public Win_prd_ProcessRunningSchedule_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            chkDate.IsChecked = true;
            btnToday_Click(null, null);

        }


        #region 일자변경

        private void lblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true) { chkDate.IsChecked = false; }
            else { chkDate.IsChecked = true; }
        }

        private void chkDate_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpSDate != null && dtpEDate != null)
            {
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        private void chkDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
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
        #endregion
        #region 체크 
        private void chkC_Checked(object sender, RoutedEventArgs e)
        {

            CheckBox chkSender = sender as CheckBox;
            var data = chkSender.DataContext as Win_prd_PrRunnig_U;

            if (data != null)
            {
                if (chkSender.IsChecked == true)
                {
                    data.IsCheck = true;

                    if (lst.Contains(data) == false)
                        lst.Add(data);
                }
            }
        }

        private void chkC_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            var data = chkSender.DataContext as Win_prd_PrRunnig_U;
            if (data != null)
            {
                data.IsCheck = false;
                if (lst.Contains(data) == true)
                    lst.Remove(data);
            }
        }

        // 전체 선택 체크
        private void AllCheck_Checked(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                var Main = dgdMain.Items[i] as Win_prd_PrRunnig_U;
                if (Main != null)
                {
                    Main.IsCheck = true;
                }
            }
        }
        // 전체 선택 체크 해제
        private void AllCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                var Main = dgdMain.Items[i] as Win_prd_PrRunnig_U;
                if (Main != null)
                {
                    Main.IsCheck = false;
                }
            }
        }
        #endregion

        #region 우측 상단 버튼
        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            FillGrid();
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            controlU();
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            update = true;

            //이미 테이블에 기간 내 데이터가 있으면 
            if (checkData())
            {
                MessageBox.Show("해당 날짜에 기존 데이터가 있으므로 없는 공정만 생성됩니다");
            }
            AddFill();

        }
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            controlU();
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            update = true;
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            for(int i = 0;i < dgdMain.Items.Count; i++)
            {
                Win_prd_PrRunnig_U data = dgdMain.Items[i] as Win_prd_PrRunnig_U;
                if (SaveData(data))
                {
                    controlD();
                };
            }
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            FillGrid();

        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            for(int i = 0; i < lst.Count; i++)
            {
                DeleteData(lst[i]);
            }
            lst.Clear();
            FillGrid();
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            controlD();
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            FillGrid();
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "작업지시 주간계획";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

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

        // 수정, 추가시 버튼 활성화 
        public void controlU()
        {
            chkDate.IsEnabled = false;
            chkDate.IsChecked = false;
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //취소시 
        public void controlD()
        {
            chkDate.IsEnabled = true;
            chkDate.IsChecked = true;
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }
        #endregion

        #region
        private void AddFill()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("SDate", dtpSDate.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("EDate", dtpEDate.SelectedDate.Value.ToString("yyyyMMdd"));

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_mcProcessRunSch_I", sqlParameter, false);

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
                            var WinWeekPlan = new Win_prd_PrRunnig_U()
                            {
                                Num = i,
                                yyyymmdd = Lib.Instance.StrDateTimeBar(dr["yyyymmdd"].ToString()),
                                DayName = dr["DayName"].ToString(),
                                processid = dr["processid"].ToString(),
                                process = dr["process"].ToString(),
                                planWorkTime = dr["PlanWorkTime"].ToString(),
                                Comments = dr["Comments"].ToString(),
                            };

                            dgdMain.Items.Add(WinWeekPlan);
                        }
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
#endregion

        private bool checkData()
        {
            try
            {
                string sData = Lib.Instance.DateFormat(dtpSDate.SelectedDate.ToString());
                string EData = Lib.Instance.DateFormat(dtpEDate.SelectedDate.ToString());
                string sql = "select * from mt_ProcessRunningSchedule where yyyymmdd between" + "'" +sData + "'" + " and " + "'" + EData + "'";

                DataSet ds = DataStore.Instance.QueryToDataSet(sql);

                
                if(ds.Tables[0].Rows.Count > 0)
                {
                    return true;
                } else
                {
                    return false;
                }
                
             }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
            return false;
        }

        #region 조회
        //
        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("SDate", dtpSDate.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("EDate", dtpEDate.SelectedDate.Value.ToString("yyyyMMdd"));

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_mcProcessRunSch_S", sqlParameter, false);

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
                            var WinWeekPlan = new Win_prd_PrRunnig_U()
                            {
                                Num = i,
                                yyyymmdd = Lib.Instance.StrDateTimeBar(dr["yyyymmdd"].ToString()),
                                DayName = dr["DayName"].ToString(),
                                processid = dr["processid"].ToString(),
                                process = dr["process"].ToString(),
                                planWorkTime = dr["PlanWorkTime"].ToString(),
                                Comments = dr["Comments"].ToString(),
                            };

                            dgdMain.Items.Add(WinWeekPlan);
                        }
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
        #endregion

        #region 저장
        private bool SaveData(Win_prd_PrRunnig_U data)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Clear();
                sqlParameter.Add("sYYYYmmdd", Lib.Instance.DateFormat(data.yyyymmdd));
                sqlParameter.Add("sProcessID", data.processid);
                sqlParameter.Add("nPlanWorkTime", data.planWorkTime);
                sqlParameter.Add("sComments", data.Comments);
                sqlParameter.Add("UserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_mcProcessRunSch_U", sqlParameter, false);

                if (result[0].Equals("success"))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
                return false;
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
            return false;
        }

        #endregion

        #region 삭제 
        private void DeleteData(Win_prd_PrRunnig_U data)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Clear();
                sqlParameter.Add("sYYYYmmdd", data.yyyymmdd.Replace("-",""));
                sqlParameter.Add("sProcessID", data.processid);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_mcProcessRunSch_D", sqlParameter, false);

                if (result[0].Equals("success"))
                {
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
        #endregion

        #region 데이터 그리드 

        // KeyDown 이벤트
        private void DataGird_KeyDown(object sender, KeyEventArgs e)
        {
            int currRow = dgdMain.Items.IndexOf(dgdMain.CurrentItem);
            int currCol = dgdMain.Columns.IndexOf(dgdMain.CurrentCell.Column);
            int startCol = 1;
            int endCol = 7;

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 열, 마지막 행 아님
                if (endCol == currCol && dgdMain.Items.Count - 1 > currRow)
                {
                    dgdMain.SelectedIndex = currRow + 1; // 이건 한줄 파란색으로 활성화 된 걸 조정하는 것입니다.
                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[currRow + 1], dgdMain.Columns[startCol]);

                } // 마지막 열 아님
                else if (endCol > currCol && dgdMain.Items.Count - 1 >= currRow)
                {
                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[currRow], dgdMain.Columns[currCol + 1]);
                } // 마지막 열, 마지막 행

            }
            else if (e.Key == Key.Down)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 행 아님
                if (dgdMain.Items.Count - 1 > currRow)
                {
                    dgdMain.SelectedIndex = currRow + 1;
                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[currRow + 1], dgdMain.Columns[currCol]);
                } // 마지막 행일때
                else if (dgdMain.Items.Count - 1 == currRow)
                {
                    if (endCol > currCol) // 마지막 열이 아닌 경우, 열을 오른쪽으로 이동
                    {
                        //dgdSub.SelectedIndex = 0;
                        dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[currRow], dgdMain.Columns[currCol + 1]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Up)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 첫행 아님
                if (currRow > 0)
                {
                    dgdMain.SelectedIndex = currRow - 1;
                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[currRow - 1], dgdMain.Columns[currCol]);
                } // 첫 행
                else if (dgdMain.Items.Count - 1 == currRow)
                {
                    if (0 < currCol) // 첫 열이 아닌 경우, 열을 왼쪽으로 이동
                    {
                        //dgdSub.SelectedIndex = 0;
                        dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[currRow], dgdMain.Columns[currCol - 1]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Left)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (startCol < currCol)
                {
                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[currRow], dgdMain.Columns[currCol - 1]);
                }
                else if (startCol == currCol)
                {
                    if (0 < currRow)
                    {
                        dgdMain.SelectedIndex = currRow - 1;
                        dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[currRow - 1], dgdMain.Columns[endCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Right)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (endCol > currCol)
                {

                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[currRow], dgdMain.Columns[currCol + 1]);
                }
                else if (endCol == currCol)
                {
                    if (dgdMain.Items.Count - 1 > currRow)
                    {
                        dgdMain.SelectedIndex = currRow + 1;
                        dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[currRow + 1], dgdMain.Columns[startCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
        }
        private void DataGridIn_TextFocus(object sender, KeyEventArgs e)
        {
            try
            {
                Lib.Instance.DataGridINControlFocus(sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_TextFocus " + ee.ToString());
            }
        }
        // GotFocus 이벤트
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible && update == true)
            {
                int currCol = dgdMain.Columns.IndexOf(dgdMain.CurrentCell.Column);

                DataGridCell cell = sender as DataGridCell;
                if (currCol == 6 || currCol == 7)
                {
                    cell.IsEditing = true;
                }
            }
        }
        // 2019.08.27 MouseUp 이벤트
        private void DataGridCell_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Lib.Instance.DataGridINBothByMouseUP(sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_MouseUp " + ee.ToString());
            }
        }

        #endregion

        #region 기타 메서드 
        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumericOnly((TextBox)sender, e);
        }

        #endregion

    }

    #region CodeView
    class Win_prd_PrRunnig_U
    {
        public int Num { get; set; }
        public bool IsCheck { get; set; }
        public string yyyymmdd { get; set; }
        public string DayName { get; set; }
        public string processid { get; set; }

        public string process { get; set; }
        public string planWorkTime { get; set; }
        public string Comments { get; set; }
    }

    #endregion
}
