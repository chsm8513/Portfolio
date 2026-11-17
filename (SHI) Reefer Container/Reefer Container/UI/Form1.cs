using Aucotec.EngineeringBase.Interop;
using DevExpress.Internal.WinApi.Windows.UI.Notifications;
using DevExpress.Spreadsheet;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraSpreadsheet;
using DevExpress.XtraSpreadsheet.Menu;
using DevExpress.XtraSpreadsheet.Model;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Drawing;
using Reefer_Container;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EBApp = Aucotec.EngineeringBase.Interop.Application;
using Gem = GemBox.Spreadsheet;
using R_EBApp = Aucotec.EngineeringBase.Client.Runtime.Application;
using Runtime = Aucotec.EngineeringBase.Client.Runtime;

namespace Reefer_Container
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public static EBApp ebapp = null;
        public EBHelper ebHelper;
        public static R_EBApp r_ebapp = null;

        [DllImport("f_nxldr.dll", EntryPoint = "EnableDRM")]
        private static extern bool EnableDRM();

        List<ReeferItem> raw_List = new List<ReeferItem>(); // 1시트 원본
        List<BayGroupDto> grouped_List = new List<BayGroupDto>(); // 출력용 그룹
        Dictionary<string, string> RDB_dict = new Dictionary<string, string>();
        Dictionary<string, string> Gangbox_dict = new Dictionary<string, string>();
        Dictionary<string, string> Reefer_dict = new Dictionary<string, string>();
        private readonly Dictionary<string, IAucObject> _rdbObjectCache = new Dictionary<string, IAucObject>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, IAucObject> _gangboxObjectCache = new Dictionary<string, IAucObject>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, IAucObject> _reeferObjectCache = new Dictionary<string, IAucObject>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, IAucObject> _ebObjectCacheById = new Dictionary<string, IAucObject>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _pendingOriginalReeferColors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _workbookReeferEbIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public string RDB_folder_name = "46:ELEC. DISTRIBUTION SYSTEM(SWBD)";
        public string Gangbox_folder_name = "75:REEFER CONTAINER SOCKET";
        public string Transformer_folder_name = "16:H/V REEFER TRANSFORMER";
        
        #region 사용자 설정 변수       
        public static int Max_AF = 1250; // 최대 AF 수량 (Bay당 Reefer 수량)
        public static int Max_rdb_reefer_cnt = 0; // 1RDB gang 수 제한
        public int Max_gang_cnt = 0; // gangbox에 들어갈 최대 gang 수       

        public int deck_hold_type = 0; // 0: ON DECK  1: ON DECK + IN HOLD 우선
        private bool _processDeckBeforeHold = true;
        private bool IsDeckHoldProcessingEnabled => deck_hold_type == 0 && _processDeckBeforeHold;
        public double Power_factor = 0.8; // 0.8

        public double Reefer_basis = 0; // kw     
        public double distboard_basis = 0; // kw   
        public double transformer_basis = 0; // kw       
        public double srd_basis = 0; // %
        public double generator_basis = 0; // kw


        public IAucObject Reefer_folder;
        public IAucObject RDB_folder;
        public IAucObject Gangbox_folder;
        public IAucObject Transformer_folder;
        public IAucObject Origin_eb_reefer;
        public IAucObject RDB_catalog;
        public IAucObject Gangbox_catalog;
        public IAucProject Project;

        private int _rightClickRowIndex = -1;
        private int _rightClickColumnIndex = -1;
        private DevExpress.Spreadsheet.Cell _rightClickCell;
        private Point _rightClickPoint;
        private System.Windows.Forms.ComboBox _reeferColorCombo;
        private DevExpress.Spreadsheet.Cell _reeferColorComboCell;
        private List<DevExpress.Spreadsheet.Cell> _reeferColorTargetCells = new List<DevExpress.Spreadsheet.Cell>();
        private string _spreadsheetRangeFillInput = string.Empty;
        private bool _spreadsheetRangeFillInputActive;

        private bool _additionalSettingControlsInitialized;

        //string downloadPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),"Downloads");
        string downloadPath = $@"C:\Users\{Environment.UserName}\AppData\Local\Temp";
        string templatePath = ""; // 템플릿 파일 경로 지정
        string savePath = ""; // 저장할 파일 경로 지정
        string template_FileName = "";
        private string settingsPath = "";
        private const string settingsFileName = "Reefer Container Settings.json";
        private string outputPath = "";
        private const string outputFileName = "Reefer Container Output.xlsx";
        private const string DefaultRevisionName = "Rev0";
        private const string EbSyncStateWorksheetName = "__EB_SYNC_STATE";
        private const string EbSyncStateMarker = "REEFER_EB_SYNC_STATE_V1";
        private string _savedRevisionName = DefaultRevisionName;
        private string _validatedEbOidSessionKey;

        private readonly GroupingService _groupingService = new GroupingService();
        private readonly EvenGroupingService _evenGroupingService = new EvenGroupingService();
        private readonly OddGroupingService _oddGroupingService = new OddGroupingService();
        private readonly UpdateService _updateService = new UpdateService();
        private readonly ExcelService _excelService = new ExcelService();
        private SettingService _settingService;

        private int _renderedBayAsc = 0;   // 0: FWD->AFT, 1: AFT->FWD
        private int _renderedTierAsc = 0;  // 0: Top->Bottom, 1: Bottom->Top
        private int _renderedRowAsc = 0;   // 0/1 기존 로직 기준 그대로 사용
        private bool _hasRenderedOrderOptions = false;

        #endregion


        public Form1()
        {
            InitializeComponent();
            settingsPath = Path.Combine(downloadPath, settingsFileName);
            outputPath = Path.Combine(downloadPath, outputFileName);
            InitializeAdditionalSettingControls();

            KeyPreview = true;
            KeyDown += spreadsheetControl1_KeyDown;
            spreadsheetControl1.MouseDown += spreadsheetControl1_MouseDown;
            spreadsheetControl1.KeyDown += spreadsheetControl1_KeyDown;
            spreadsheetControl1.PopupMenuShowing += spreadsheetControl1_PopupMenuShowing;
            InitializeReeferColorCombo();
            // 정렬값 변경 시 기존 Import 결과는 유지한다.
            // 변경된 설정은 사용자가 다시 Import할 때만 반영한다.
            //sortPriorityPanel.PriorityChanged += SortPriorityPanel_PriorityChanged;
        }

        private void InitializeAdditionalSettingControls()
        {
            if (_additionalSettingControlsInitialized)
                return;

            _additionalSettingControlsInitialized = true;

            UpdateBasisCalculationLabels();
        }

        private void UpdateBasisCalculationLabels()
        {
            if (num_srd_basis == null || num_generator_basis == null ||
                lbl_rdb_result == null || lbl_srd_result == null || lbl_tr_result == null)
                return;

            double reeferBasisKw = (double)num_Reefer_basis.Value;
            double rdbKw = reeferBasisKw * (double)num_distboard_basis.Value / 100.0;
            double srdKw = reeferBasisKw * (double)num_srd_basis.Value / 100.0;
            double trKw = reeferBasisKw * (double)num_transformer_basis.Value / 100.0;

            num_RDB.Value = (decimal)rdbKw;
            num_SRD.Value = (decimal)srdKw;
            num_TR.Value = (decimal)trKw;

            lbl_rdb_result.Text = $"{CalculateRatedCurrent(rdbKw):0.0}A";
            lbl_srd_result.Text = $"{CalculateRatedCurrent(srdKw):0.0}A";
            lbl_tr_result.Text = $"{CalculateRatedCurrent(trKw):0.0}A";
        }

        private double CalculateRatedCurrent(double powerKw)
        {
            const double ratedVoltage = 440.0;
            const double electricalPhaseCount = 3.0;
            return powerKw * 1000.0 /
                (Math.Sqrt(electricalPhaseCount) * ratedVoltage * Power_factor);
        }

        private void UpdateTransformerCount()
        {
            int count = 0;

            try
            {
                if (Transformer_folder != null && Transformer_folder.Children != null)
                    count = Transformer_folder.Children.Count;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[UpdateTransformerCount] fail: {ex.Message}");
            }

            if (lbl_tr_count != null)
                lbl_tr_count.Text = $"{count} EA";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //version
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"Reefer Container Grouping v{version}";

            //gebox license
            Gem.SpreadsheetInfo.SetLicense("SN-2022Oct17-DnYYtaAe7mJpwiwjXviXn7Sqdvj2ZGU1bFD2Qp72UhXkrPYDLPopj99indaxL4R/1daNVuUSUdJnOTyOH6p3Z8yRsEg==A");

            //DRM
            try
            {
                EnableDRM();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

            ebHelper = new EBHelper();

            if (ebapp == null)
            {
                MessageBox.Show($"EB App을 찾을 수 없습니다.");
                this.Close();
                return;
            }                

            if (ebapp.Selection.Count == 0)
            {
                MessageBox.Show("EB Selection을 찾지 못했습니다.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Project = ebapp.Selection.Item(1).Project;
            if(Project == null)
            {
                MessageBox.Show("EB Project을 찾지 못했습니다.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            lb_projectname.Text = Project.Name;
            _settingService = new SettingService(Project, ebHelper);

            IAucObjects finds = Project.EquipmentFolder.FindObjectsEx(ebHelper.Get_filter("Equipment"), AucVbSearchType.aucSearchFlat, true);
            if(finds.Count == 0)
            {
                MessageBox.Show("EB Project.EquipmentFolder -> Equipment folder를 찾지 못했습니다.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            IAucObject equipment_folder = Project.EquipmentFolder.FindObjectsEx(ebHelper.Get_filter("Equipment"), AucVbSearchType.aucSearchFlat, true).Item(1);
                        
            Reefer_folder = ebHelper.Get_object(equipment_folder, $"# Reefer", AucObjectKind.aucObjUnit, AucType.aucUnit); // 112, 12는 찾음 근데 1764 이건 못찾음
            Debug.WriteLine($"Reefer: {Reefer_folder.ID}");

            IAucObject EP_folder = ebHelper.Get_object(equipment_folder, $"EP", AucObjectKind.aucObjUnit, AucType.aucUnitCommon);
            RDB_folder = ebHelper.Get_object(EP_folder, $"{RDB_folder_name}", AucObjectKind.aucObjUnit, AucType.aucUnitCommon); 
            Debug.WriteLine($"RDB: {RDB_folder.ID}");

            Gangbox_folder = ebHelper.Get_object(EP_folder, $"{Gangbox_folder_name}", AucObjectKind.aucObjUnit, AucType.aucUnitCommon); 
            Debug.WriteLine($"Gangbox: {Gangbox_folder.ID}");

            Transformer_folder = ebHelper.Get_object(
                EP_folder,
                Transformer_folder_name,
                AucObjectKind.aucObjUnit,
                AucType.aucUnitCommon,
                bCreate: false);

            if (Transformer_folder == null)
            {
                Transformer_folder = ebHelper.Get_object(
                    equipment_folder,
                    Transformer_folder_name,
                    AucObjectKind.aucObjUnit,
                    AucType.aucUnitCommon,
                    bCreate: false);
            }

            UpdateTransformerCount();

            this.Resize += UpdateWaitFormLocation;
            this.LocationChanged += UpdateWaitFormLocation;

            btn_create_data.Visible = false;
            simpleButton1.Visible = false;
            btn_callworksheet.Visible = false;
            btn_rebuild.Visible = false;
            btn_change_color.BringToFront();
            spreadsheetControl1.BringToFront();

            gridView1.OptionsBehavior.Editable = false;
            gridView1.RowCellStyle += gridView1_RowCellStyle;
            gridControl1.ViewRegistered += gridControl1_ViewRegistered;
        }

        private void UpdateWaitFormLocation(object sender, EventArgs e)
        {
            // SplashScreenManager가 현재 WaitForm을 화면에 보여주고 있을 때만 실행
            if (SplashScreenManager.Default != null && SplashScreenManager.Default.IsSplashFormVisible)
            {
                // SendCommand를 사용하여 WaitForm으로 현재 Form의 전체 영역 정보(Bounds)를 보냅니다.
                // this.Bounds는 Form의 X, Y 좌표, Width, Height를 모두 포함하는 Rectangle 객체입니다.
                SplashScreenManager.Default.SendCommand(null, this.Bounds);
            }
        }

        #region Cell color combo box of Reefer
        private sealed class ReeferColorComboItem
        {
            public int Number { get; set; }
            public string Hex { get; set; }
            public Color Color { get; set; }

            public override string ToString()
            {
                return Number.ToString();
            }
        }

        private void InitializeReeferColorCombo()
        {
            _reeferColorCombo = new System.Windows.Forms.ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 22,
                IntegralHeight = true,
                Visible = false,
                Width = 92,
                Height = 24
            };

            for (int i = 0; i < _groupColorPalette.Length; i++)
            {
                string hex = _excelService.NormalizeHtmlColor(_groupColorPalette[i]);
                _reeferColorCombo.Items.Add(new ReeferColorComboItem
                {
                    Number = i + 1,
                    Hex = hex,
                    Color = ColorTranslator.FromHtml(hex)
                });
            }

            _reeferColorCombo.DrawItem += ReeferColorCombo_DrawItem;
            _reeferColorCombo.SelectionChangeCommitted += ReeferColorCombo_SelectionChangeCommitted;
            _reeferColorCombo.DropDownClosed += (s, e) => _reeferColorCombo.Visible = false;

            Controls.Add(_reeferColorCombo);
            _reeferColorCombo.BringToFront();
        }

        private void spreadsheetControl1_MouseDown(object sender, MouseEventArgs e)
        {
            HideReeferColorCombo();
            ClearSpreadsheetRangeFillInput();

            if (e.Button != MouseButtons.Right)
                return;

            try
            {
                DevExpress.Spreadsheet.Cell cell = spreadsheetControl1.GetCellFromPoint(new PointF(e.X, e.Y));

                if (cell == null)
                {
                    _rightClickRowIndex = -1;
                    _rightClickColumnIndex = -1;
                    _rightClickCell = null;
                    _rightClickPoint = Point.Empty;
                    return;
                }

                _rightClickRowIndex = cell.RowIndex;
                _rightClickColumnIndex = cell.ColumnIndex;
                _rightClickCell = cell;
                _rightClickPoint = e.Location;
            }
            catch
            {
                _rightClickRowIndex = -1;
                _rightClickColumnIndex = -1;
                _rightClickCell = null;
                _rightClickPoint = Point.Empty;
            }
        }

        private void spreadsheetControl1_KeyDown(object sender, KeyEventArgs e)
        {
            if (spreadsheetControl1 == null || !spreadsheetControl1.ContainsFocus)
                return;

            if (e.KeyCode == Keys.Escape)
            {
                ClearSpreadsheetRangeFillInput();
                return;
            }

            if (e.KeyCode == Keys.Enter && _spreadsheetRangeFillInputActive)
            {
                if (FillSelectedSpreadsheetRanges(_spreadsheetRangeFillInput))
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }

                ClearSpreadsheetRangeFillInput();
                return;
            }

            if (!HasMultiCellSpreadsheetSelection())
            {
                ClearSpreadsheetRangeFillInput();
                return;
            }

            if (e.KeyCode == Keys.Back && _spreadsheetRangeFillInputActive)
            {
                if (_spreadsheetRangeFillInput.Length > 0)
                    _spreadsheetRangeFillInput = _spreadsheetRangeFillInput.Substring(0, _spreadsheetRangeFillInput.Length - 1);

                UpdateSpreadsheetRangeFillPreview(_spreadsheetRangeFillInput);
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (TryGetSpreadsheetRangeFillCharacter(e, out char inputCharacter))
            {
                if (inputCharacter == '-' && _spreadsheetRangeFillInput.Length > 0)
                    return;

                char decimalSeparator = Convert.ToChar(System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                if (inputCharacter == decimalSeparator && _spreadsheetRangeFillInput.IndexOf(decimalSeparator) >= 0)
                    return;

                _spreadsheetRangeFillInput += inputCharacter;
                _spreadsheetRangeFillInputActive = true;
                UpdateSpreadsheetRangeFillPreview(_spreadsheetRangeFillInput);
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (e.KeyCode != Keys.ShiftKey && e.KeyCode != Keys.ControlKey && e.KeyCode != Keys.Menu)
                ClearSpreadsheetRangeFillInput();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (spreadsheetControl1 != null &&
                spreadsheetControl1.ContainsFocus &&
                (keyData == Keys.Enter || keyData == Keys.Return) &&
                _spreadsheetRangeFillInputActive)
            {
                bool filled = FillSelectedSpreadsheetRanges(_spreadsheetRangeFillInput);
                ClearSpreadsheetRangeFillInput();

                if (filled)
                    return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private bool HasMultiCellSpreadsheetSelection()
        {
            try
            {
                var ranges = spreadsheetControl1.GetSelectedRanges();
                if (ranges == null)
                    return false;

                foreach (DevExpress.Spreadsheet.CellRange range in ranges)
                {
                    if (range != null &&
                        (range.TopRowIndex != range.BottomRowIndex ||
                         range.LeftColumnIndex != range.RightColumnIndex))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[HasMultiCellSpreadsheetSelection] fail: {ex.Message}");
            }

            return false;
        }

        private bool TryGetSpreadsheetRangeFillCharacter(KeyEventArgs e, out char inputCharacter)
        {
            inputCharacter = '\0';

            if (e.Control || e.Alt)
                return false;

            if (e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9)
            {
                inputCharacter = (char)('0' + (e.KeyCode - Keys.D0));
                return true;
            }

            if (e.KeyCode >= Keys.NumPad0 && e.KeyCode <= Keys.NumPad9)
            {
                inputCharacter = (char)('0' + (e.KeyCode - Keys.NumPad0));
                return true;
            }

            if (e.KeyCode == Keys.OemMinus || e.KeyCode == Keys.Subtract)
            {
                inputCharacter = '-';
                return true;
            }

            if (e.KeyCode == Keys.Decimal || e.KeyCode == Keys.OemPeriod || e.KeyCode == Keys.Oemcomma)
            {
                inputCharacter = Convert.ToChar(System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                return true;
            }

            return false;
        }

        private void UpdateSpreadsheetRangeFillPreview(string input)
        {
            try
            {
                var ranges = spreadsheetControl1.GetSelectedRanges();
                if (ranges == null)
                    return;

                foreach (DevExpress.Spreadsheet.CellRange range in ranges)
                {
                    if (range == null ||
                        (range.TopRowIndex == range.BottomRowIndex &&
                         range.LeftColumnIndex == range.RightColumnIndex))
                    {
                        continue;
                    }

                    DevExpress.Spreadsheet.Cell previewCell =
                        range.Worksheet.Cells[range.TopRowIndex, range.LeftColumnIndex];

                    if (string.IsNullOrEmpty(input))
                    {
                        previewCell.Value = string.Empty;
                    }
                    else if (double.TryParse(
                        input,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.CurrentCulture,
                        out double numericValue))
                    {
                        previewCell.Value = numericValue;
                    }
                    else
                    {
                        previewCell.Value = input;
                    }

                    break;
                }

                spreadsheetControl1.Refresh();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[UpdateSpreadsheetRangeFillPreview] fail: {ex.Message}");
            }
        }

        private bool FillSelectedSpreadsheetRanges(string input)
        {
            if (string.IsNullOrWhiteSpace(input) || !HasMultiCellSpreadsheetSelection())
                return false;

            try
            {
                if (!double.TryParse(
                    input,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.CurrentCulture,
                    out double numericValue))
                {
                    return false;
                }

                var ranges = spreadsheetControl1.GetSelectedRanges();
                if (ranges == null)
                    return false;

                int changedCount = 0;
                foreach (DevExpress.Spreadsheet.CellRange range in ranges)
                {
                    if (range == null)
                        continue;

                    for (int rowIndex = range.TopRowIndex; rowIndex <= range.BottomRowIndex; rowIndex++)
                    {
                        for (int columnIndex = range.LeftColumnIndex; columnIndex <= range.RightColumnIndex; columnIndex++)
                        {
                            range.Worksheet.Cells[rowIndex, columnIndex].Value = numericValue;
                            changedCount++;
                        }
                    }
                }

                if (changedCount == 0)
                    return false;

                spreadsheetControl1.Refresh();
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[FillSelectedSpreadsheetRanges] fail: {ex}");
                MessageBox.Show($"선택 범위에 값을 입력하는 중 오류가 발생했습니다: {ex.Message}");
                return false;
            }
        }

        private void ClearSpreadsheetRangeFillInput()
        {
            _spreadsheetRangeFillInput = string.Empty;
            _spreadsheetRangeFillInputActive = false;
        }

        private void ShowReeferColorCombo(Point clickPoint)
        {
            try
            {
                DevExpress.Spreadsheet.Cell cell = spreadsheetControl1.GetCellFromPoint(new PointF(clickPoint.X, clickPoint.Y));
                ShowReeferColorCombo(cell, clickPoint);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ShowReeferColorCombo] fail: {ex}");
                HideReeferColorCombo();
            }
        }

        private void ShowReeferColorCombo(DevExpress.Spreadsheet.Cell cell, Point clickPoint)
        {
            try
            {
                if (cell == null)
                {
                    HideReeferColorCombo();
                    return;
                }

                var selectedCells = GetSelectedReeferCells();
                var clickedReefer = FindReeferByCell(cell.RowIndex, cell.ColumnIndex);
                bool clickedCellIsReefer = IsReeferSpreadsheetCell(cell);
                bool clickedInSelection = selectedCells.Any(x =>
                    x.RowIndex == cell.RowIndex &&
                    x.ColumnIndex == cell.ColumnIndex);

                var targetCells = new List<DevExpress.Spreadsheet.Cell>();
                if (selectedCells.Count > 0 && clickedInSelection)
                    targetCells.AddRange(selectedCells);
                else if (clickedCellIsReefer)
                    targetCells.Add(cell);
                else if (selectedCells.Count > 0)
                    targetCells.AddRange(selectedCells);

                if (targetCells.Count == 0)
                {
                    HideReeferColorCombo();
                    return;
                }

                DevExpress.Spreadsheet.Cell sourceCell = clickedCellIsReefer
                    ? cell
                    : targetCells.FirstOrDefault();
                var sourceReefer = clickedReefer ?? FindReeferByCell(sourceCell.RowIndex, sourceCell.ColumnIndex);

                string cellColor = _excelService.NormalizeHtmlColor(GetSpreadsheetCellColorHex(sourceCell));
                string reeferColor = _excelService.NormalizeHtmlColor(sourceReefer?.GangBox_ID);
                string selectedColor = !string.IsNullOrWhiteSpace(cellColor) ? cellColor : reeferColor;
                if (string.IsNullOrWhiteSpace(selectedColor))
                {
                    HideReeferColorCombo();
                    return;
                }

                _reeferColorComboCell = sourceCell;
                _reeferColorTargetCells = targetCells;

                int selectedIndex = FindPaletteIndex(selectedColor);
                _reeferColorCombo.SelectedIndex = selectedIndex;
                _reeferColorCombo.Location = GetColorComboLocation(clickPoint);
                _reeferColorCombo.Visible = true;
                _reeferColorCombo.BringToFront();
                _reeferColorCombo.Focus();
                _reeferColorCombo.DroppedDown = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ShowReeferColorCombo] fail: {ex}");
                HideReeferColorCombo();
            }
        }

        private void HideReeferColorCombo()
        {
            if (_reeferColorCombo != null)
                _reeferColorCombo.Visible = false;

            _reeferColorComboCell = null;
            _reeferColorTargetCells.Clear();
        }

        private Point GetColorComboLocation(Point clickPoint)
        {
            int x = clickPoint.X;
            int y = clickPoint.Y + 2;

            if (_reeferColorCombo != null)
            {
                x = Math.Max(0, Math.Min(x, ClientSize.Width - _reeferColorCombo.Width));
                y = Math.Max(0, Math.Min(y, ClientSize.Height - _reeferColorCombo.Height));
            }

            return new Point(x, y);
        }

        private ReeferItem FindReeferByCell(int rowIndex, int columnIndex)
        {
            return grouped_List
                .SelectMany(bay => bay.Tiers ?? Enumerable.Empty<TierGroupDto>())
                .SelectMany(tier => tier.Rows ?? Enumerable.Empty<ReeferItem>())
                .FirstOrDefault(x =>
                    x != null &&
                    x.Reefer &&
                    Convert.ToInt32(x.X) == rowIndex &&
                    Convert.ToInt32(x.Y) == columnIndex);
        }

        private List<DevExpress.Spreadsheet.Cell> GetSelectedReeferCells()
        {
            var result = new List<DevExpress.Spreadsheet.Cell>();
            var seen = new HashSet<string>();

            try
            {
                var ranges = spreadsheetControl1.GetSelectedRanges();
                if (ranges == null)
                    return result;

                foreach (DevExpress.Spreadsheet.CellRange range in ranges)
                {
                    if (range == null)
                        continue;

                    foreach (DevExpress.Spreadsheet.Cell cell in range.ExistingCells)
                    {
                        if (cell == null)
                            continue;

                        string key = cell.RowIndex + "|" + cell.ColumnIndex;
                        if (!seen.Add(key))
                            continue;

                        if (IsReeferSpreadsheetCell(cell))
                            result.Add(cell);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GetSelectedReeferCells] fail: {ex}");
            }

            return result;
        }

        private int FindPaletteIndex(string hex)
        {
            string normalized = _excelService.NormalizeHtmlColor(hex);
            for (int i = 0; i < _groupColorPalette.Length; i++)
            {
                if (string.Equals(_excelService.NormalizeHtmlColor(_groupColorPalette[i]), normalized, StringComparison.OrdinalIgnoreCase))
                    return i;
            }

            return -1;
        }

        private string GetSpreadsheetCellColorHex(DevExpress.Spreadsheet.Cell cell)
        {
            if (cell == null)
                return null;

            Color color = cell.FillColor;
            if (color.IsEmpty || color.ToArgb() == Color.White.ToArgb())
                return null;

            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        /// <summary>
        /// Change Color는 EB Reefer 객체가 생성되기 전에도 사용한다.
        /// 따라서 EB ID나 메모리 X/Y 매칭이 아닌, Excel에 표시된
        /// GangBox 팔레트 색상을 주 판별 기준으로 사용한다.
        /// </summary>
        private bool IsReeferSpreadsheetCell(DevExpress.Spreadsheet.Cell cell)
        {
            if (cell == null)
                return false;

            string cellColor = _excelService.NormalizeHtmlColor(GetSpreadsheetCellColorHex(cell));
            if (FindPaletteIndex(cellColor) >= 0)
                return true;

            // 기존 메모리 매칭은 이전 파일/ReBuild 흐름을 위한 보조 판별로만 유지한다.
            return FindReeferByCell(cell.RowIndex, cell.ColumnIndex) != null;
        }

        private int ApplyColorToSelectedReeferCells(Color color)
        {
            var cells = _reeferColorTargetCells
                .Where(x => x != null)
                .GroupBy(x => x.RowIndex + "|" + x.ColumnIndex)
                .Select(g => g.First())
                .ToList();

            if (cells.Count == 0 &&
                _reeferColorComboCell != null &&
                IsReeferSpreadsheetCell(_reeferColorComboCell))
            {
                cells.Add(_reeferColorComboCell);
            }

            string worksheetName = spreadsheetControl1.ActiveWorksheet?.Name ?? string.Empty;
            string newColor = _excelService.NormalizeHtmlColor(
                $"#{color.R:X2}{color.G:X2}{color.B:X2}");
            int changedCount = 0;

            foreach (var cell in cells)
            {
                string currentColor = _excelService.NormalizeHtmlColor(GetSpreadsheetCellColorHex(cell)) ?? string.Empty;
                if (string.Equals(currentColor, newColor, StringComparison.OrdinalIgnoreCase))
                    continue;

                string cellKey = worksheetName + "|" + cell.RowIndex + "|" + cell.ColumnIndex;
                if (!_pendingOriginalReeferColors.ContainsKey(cellKey))
                    _pendingOriginalReeferColors[cellKey] = currentColor;

                cell.FillColor = color;
                changedCount++;

                string originalColor = _pendingOriginalReeferColors[cellKey];
                if (string.Equals(originalColor, newColor, StringComparison.OrdinalIgnoreCase))
                    _pendingOriginalReeferColors.Remove(cellKey);
            }

            return changedCount;
        }

        private void ReeferColorCombo_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();

            if (e.Index < 0 || e.Index >= _reeferColorCombo.Items.Count)
                return;

            var item = _reeferColorCombo.Items[e.Index] as ReeferColorComboItem;
            if (item == null)
                return;

            Rectangle numberBounds = new Rectangle(e.Bounds.Left + 4, e.Bounds.Top + 2, 24, e.Bounds.Height - 4);
            Rectangle swatchBounds = new Rectangle(e.Bounds.Left + 32, e.Bounds.Top + 3, e.Bounds.Width - 38, e.Bounds.Height - 6);

            using (Brush textBrush = new SolidBrush(e.ForeColor))
            using (Brush swatchBrush = new SolidBrush(item.Color))
            {
                e.Graphics.DrawString(item.Number.ToString(), e.Font, textBrush, numberBounds);
                e.Graphics.FillRectangle(swatchBrush, swatchBounds);
            }

            e.Graphics.DrawRectangle(Pens.DimGray, swatchBounds);
            e.DrawFocusRectangle();
        }

        private void ReeferColorCombo_SelectionChangeCommitted(object sender, EventArgs e)
        {
            var item = _reeferColorCombo.SelectedItem as ReeferColorComboItem;
            if (item == null || _reeferColorComboCell == null)
                return;

            try
            {
                int changedCount = ApplyColorToSelectedReeferCells(item.Color);
                if (changedCount > 0 && !string.IsNullOrWhiteSpace(savePath))
                    spreadsheetControl1.SaveDocument(savePath);

                spreadsheetControl1.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"색상 변경 중 오류 발생: {ex.Message}");
            }
            finally
            {
                HideReeferColorCombo();
            }
        }

        #endregion

        #region Navigate menu
        private void spreadsheetControl1_PopupMenuShowing(object sender, PopupMenuShowingEventArgs e)
        {
            if (e.Menu == null)
                return;

            if (e.MenuType != SpreadsheetMenuType.Cell)
                return;
            
            // 1
            SpreadsheetMenuItem navigateItem = new SpreadsheetMenuItem("Navigate", NavigateMenu_Click);
            navigateItem.BeginGroup = true;

            e.Menu.Items.Add(navigateItem);

            // 2
            //SpreadsheetMenuItem changeColorItem = new SpreadsheetMenuItem("Change Color", ChangeColorMenu_Click);
            //changeColorItem.BeginGroup = true;

            //e.Menu.Items.Add(changeColorItem);

            //3

        }

        private void ChangeColorMenu_Click(object sender, EventArgs e)
        {
            try
            {
                if (_rightClickCell == null || _rightClickRowIndex < 0 || _rightClickColumnIndex < 0)
                {
                    MessageBox.Show("우클릭한 셀 좌표를 찾을 수 없습니다.");
                    return;
                }

                bool hasClickedReefer = IsReeferSpreadsheetCell(_rightClickCell);
                bool hasSelectedReefers = GetSelectedReeferCells().Count > 0;
                if (!hasClickedReefer && !hasSelectedReefers)
                {
                    MessageBox.Show($"해당 셀에 Reefer가 없습니다. Cell=({_rightClickRowIndex},{_rightClickColumnIndex})");
                    return;
                }

                Point formPoint = PointToClient(spreadsheetControl1.PointToScreen(_rightClickPoint));
                ShowReeferColorCombo(_rightClickCell, formPoint);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"색상 메뉴 처리 중 오류 발생: {ex.Message}");
            }
        }

        private void btn_change_color_Click(object sender, EventArgs e)
        {
            try
            {
                var selectedCells = GetSelectedReeferCells();
                DevExpress.Spreadsheet.Cell cell = selectedCells.FirstOrDefault() ?? _rightClickCell;

                if (!IsReeferSpreadsheetCell(cell))
                {
                    MessageBox.Show("먼저 Reefer 셀을 선택해 주세요.");
                    return;
                }

                Point comboPoint = new Point(btn_change_color.Left, btn_change_color.Bottom + 1);
                ShowReeferColorCombo(cell, comboPoint);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"색상 버튼 처리 중 오류 발생: {ex.Message}");
            }
        }

        private void NavigateMenu_Click(object sender, EventArgs e)
        {
            try
            {
                if (_rightClickRowIndex < 0 || _rightClickColumnIndex < 0)
                {
                    MessageBox.Show("우클릭한 셀 좌표를 찾을 수 없습니다.");
                    return;
                }

                int rowIndex = _rightClickRowIndex;
                int columnIndex = _rightClickColumnIndex;
                
                // Search
                var reefer = grouped_List
                    .SelectMany(bay => bay.Tiers ?? Enumerable.Empty<TierGroupDto>())
                    .SelectMany(tier => tier.Rows ?? Enumerable.Empty<ReeferItem>())
                    .FirstOrDefault(x =>
                        Convert.ToInt32(x.X) == rowIndex &&
                        Convert.ToInt32(x.Y) == columnIndex);

                if (reefer == null)
                {
                    MessageBox.Show($"해당 셀에 Reefer가 없습니다. Cell=({rowIndex},{columnIndex})");
                    return;
                }

                string reeferEbId = reefer.Reefer_EB_ID;
                if (string.IsNullOrWhiteSpace(reeferEbId))
                {
                    string reeferDesignation = Get_Reefer_Name(reefer);
                    _workbookReeferEbIds.TryGetValue(reeferDesignation, out reeferEbId);
                }

                if (string.IsNullOrWhiteSpace(reeferEbId))
                {
                    MessageBox.Show(
                        "선택한 Reefer의 EB OID가 Excel 동기화 시트에 없습니다.\r\n" +
                        "EB Update가 완료된 Excel인지 확인하세요.",
                        "Navigate",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                try
                {
                    ebapp.Automation.Parameters.Clear();
                    //ebapp.Automation.Parameters.Add(57112);
                    //ebapp.Utils.RunMacro("Toolbar_Assistants.m99_Functions.sbExecuteCommand", item);
                    ebapp.Automation.Parameters.Add(reeferEbId);
                    ebapp.Utils.RunMacro("PhaseBalance.9.Navigation", ebapp.ActiveProject);
                }
                catch { MessageBox.Show("Navigate할 수 없습니다.\r\n객체를 다시 확인하세요.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Navigate 처리 중 오류 발생: {ex.Message}");
            }
        }
        #endregion

        private void gridControl1_ViewRegistered(object sender, DevExpress.XtraGrid.ViewOperationEventArgs e)
        {
            try
            {
                var gv = e.View as DevExpress.XtraGrid.Views.Grid.GridView;
                if (gv == null)
                    return;

                gv.RowCellStyle -= DetailRowsView_RowCellStyle;
                gv.RowCellStyle += DetailRowsView_RowCellStyle;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"gridControl1_ViewRegistered 오류: {ex}");
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (!LoadSettingsFromEB())
            {
                MessageBox.Show(
                    $"EB에서 '{settingsFileName}'을 불러오지 못했습니다.\r\n프로그램을 종료합니다.",
                    "Setting Load 오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                Close();
                return;
            }

            Value_Settings();
            LoadOutputAndShowIfAvailable();
        }

        private void LoadOutputAndShowIfAvailable()
        {
            if (!LoadOutputFromEB())
            {
                if (tb_revision != null)
                    tb_revision.Text = DefaultRevisionName;
                return;
            }

            // 최초 실행 시 EB에서 내려받은 Output Excel도 Rebuild 경로로 읽어야
            // 화면 Cell과 ReeferItem이 연결되고 __EB_SYNC_STATE의 OID로 Navigate할 수 있다.
            if (RebuildFromFile(outputPath, false))
                return;

            // 이전 형식 등으로 Rebuild할 수 없는 파일도 기존처럼 화면에는 표시한다.
            sheet8grid_tfn = true;
            spreadsheetControl1.LoadDocument(outputPath);
            spreadsheetControl1.BringToFront();
        }
        

        public void Value_Settings()
        {
            try
            {
                SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);

                Max_AF = (int)num_Max_AF.Value;

                transformer_basis = (double)num_transformer_basis.Value;
                Debug.WriteLine($"transformer_basis: {transformer_basis}");

                distboard_basis = (double)num_distboard_basis.Value;
                Debug.WriteLine($"distboard_basis: {distboard_basis}");

                Reefer_basis = (double)num_Reefer_basis.Value;
                Debug.WriteLine($"Reefer_basis: {Reefer_basis}");

                srd_basis = (double)num_srd_basis.Value;
                generator_basis = (double)num_generator_basis.Value;
                UpdateBasisCalculationLabels();

                Max_rdb_reefer_cnt = (int)num_Max_rdb_reefer_cnt.Value;

                Max_gang_cnt = (int)num_Max_gang_cnt.Value;

                deck_hold_type = cb_deck_hold_type.SelectedIndex;
                _processDeckBeforeHold = chk_deck_first.Checked;

                SplashScreenManager.CloseForm();
            }
            catch
            {

            }
        }

        private ReeferSettingsDto CaptureSettings()
        {
            var sortPriority = sortPriorityPanel == null
                ? new List<ReeferSortSettingDto>()
                : sortPriorityPanel.GetPriorityItems()
                    .Select(x => new ReeferSortSettingDto
                    {
                        Label = x.Label,
                        Key = x.Key,
                        SelectedOrder = x.SelectedOrder
                    })
                    .ToList();

            return new ReeferSettingsDto
            {
                MaxAf = (int)num_Max_AF.Value,
                MaxRdbReeferCount = (int)num_Max_rdb_reefer_cnt.Value,
                MaxGangCount = (int)num_Max_gang_cnt.Value,
                ReeferBasis = (double)num_Reefer_basis.Value,
                RdbPercent = (double)num_distboard_basis.Value,
                SrdPercent = (double)num_srd_basis.Value,
                TrPercent = (double)num_transformer_basis.Value,
                GeneratorBasis = (double)num_generator_basis.Value,
                DeckHoldType = cb_deck_hold_type.SelectedIndex,
                ProcessDeckBeforeHold = chk_deck_first.Checked,
                SortPriority = sortPriority,
                Revision = _settingService.NormalizeRevisionName(tb_revision?.Text)
            };
        }

        private void ApplySettings(ReeferSettingsDto settings)
        {
            if (settings == null)
                return;

            num_Max_AF.Value = ClampDecimal(settings.MaxAf, num_Max_AF.Minimum, num_Max_AF.Maximum);
            num_Max_rdb_reefer_cnt.Value = ClampDecimal(settings.MaxRdbReeferCount, num_Max_rdb_reefer_cnt.Minimum, num_Max_rdb_reefer_cnt.Maximum);
            num_Max_gang_cnt.Value = ClampDecimal(settings.MaxGangCount, num_Max_gang_cnt.Minimum, num_Max_gang_cnt.Maximum);
            num_Reefer_basis.Value = ClampDecimal(settings.ReeferBasis, num_Reefer_basis.Minimum, num_Reefer_basis.Maximum);
            num_distboard_basis.Value = ClampDecimal(settings.RdbPercent, num_distboard_basis.Minimum, num_distboard_basis.Maximum);
            num_srd_basis.Value = ClampDecimal(settings.SrdPercent, num_srd_basis.Minimum, num_srd_basis.Maximum);
            num_transformer_basis.Value = ClampDecimal(settings.TrPercent, num_transformer_basis.Minimum, num_transformer_basis.Maximum);
            num_generator_basis.Value = ClampDecimal(settings.GeneratorBasis, num_generator_basis.Minimum, num_generator_basis.Maximum);

            if (settings.DeckHoldType >= 0 && settings.DeckHoldType < cb_deck_hold_type.Items.Count)
                cb_deck_hold_type.SelectedIndex = settings.DeckHoldType;

            // 기존 JSON에 항목이 없으면 새 기본값인 Checked를 사용한다.
            chk_deck_first.Checked = settings.ProcessDeckBeforeHold ?? true;
            tb_revision.Text = _settingService.TryNormalizeRevisionName(settings.Revision, out string revision)
                ? revision
                : DefaultRevisionName;
            _savedRevisionName = tb_revision.Text;

            sortPriorityPanel?.ApplyPriorityItems(settings.SortPriority?.Select(x => new SortOrderItem
            {
                Label = x.Label,
                Key = x.Key,
                SelectedOrder = x.SelectedOrder
            }));

            UpdateBasisCalculationLabels();
        }

        private static decimal ClampDecimal(double value, decimal minimum, decimal maximum)
        {
            decimal converted = (decimal)value;
            return Math.Min(maximum, Math.Max(minimum, converted));
        }

        private void ApplyDefaultSettings()
        {
            num_Max_AF.Value = 1250m;
            num_Max_rdb_reefer_cnt.Value = 80m;
            num_Max_gang_cnt.Value = 10m;
            num_Reefer_basis.Value = 10.3m;
            num_distboard_basis.Value = 100m;
            num_srd_basis.Value = 90m;
            num_transformer_basis.Value = 70m;
            num_generator_basis.Value = 5m;
            cb_deck_hold_type.SelectedIndex = 0;
            chk_deck_first.Checked = true;
            tb_revision.Text = DefaultRevisionName;
            sortPriorityPanel?.ResetToDefaults();
            UpdateBasisCalculationLabels();
        }

        private void SaveSettingsToEB()
        {
            ReeferSettingsDto settings = CaptureSettings();
            _settingService.SaveSettings(settings, settingsPath, settingsFileName);
            _savedRevisionName = settings.Revision;
        }

        private bool LoadSettingsFromEB()
        {
            try
            {
                ReeferSettingsDto settings = _settingService.LoadSettings(settingsPath, settingsFileName);
                if (settings == null)
                {
                    ApplyDefaultSettings();
                    SaveSettingsToEB();
                    return true;
                }
                ApplySettings(settings);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LoadSettingsFromEB] fail: {ex.Message}");
                return false;
            }
        }

        private bool LoadOutputFromEB()
        {
            try
            {
                if (!_settingService.DownloadExternalDocument(outputFileName, outputPath))
                    return false;

                savePath = outputPath;
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LoadOutputFromEB] fail: {ex.Message}");
                return false;
            }
        }

        private void SaveOutputToEB()
        {
            if (spreadsheetControl1.Document == null)
                throw new InvalidOperationException("저장할 Excel 문서가 없습니다.");

            Directory.CreateDirectory(downloadPath);
            spreadsheetControl1.SaveDocument(outputPath);
            savePath = outputPath;

            // 로컬 저장이 끝난 최신 Output을 EB에도 업로드한다.
            SaveToEB_UpLoad();
                        
        }

        private void gridView1_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            try
            {
                if (e.RowHandle < 0) return;

                // Warning / HasChangedItems 는 전체 행에 적용
                var view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
                if (view == null) return;

                var row = view.GetRow(e.RowHandle) as BayGroupDto;
                if (row == null) return;

                if (row.Warning)
                {
                    e.Appearance.BackColor = Color.DarkGray;
                    e.Appearance.ForeColor = Color.Yellow;
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
                }
                else if (row.HasChangedItems)
                {
                    e.Appearance.BackColor = Color.LightSkyBlue;
                    e.Appearance.ForeColor = Color.DarkBlue;
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"gridView1_RowCellStyle 오류: {ex}");
            }
        }

        private void DetailRowsView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            try
            {
                if (e.RowHandle < 0)
                    return;

                var view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
                if (view == null)
                    return;

                // Rows detail view 의 실제 row 타입
                var item = view.GetRow(e.RowHandle) as ReeferItem;
                if (item == null)
                    return;

                // Before_RDB_Color 값이 없으면 변경 안 된 것
                if (string.IsNullOrWhiteSpace(item.Before_RDB_Color) && !item.IsChanged)
                    return;

                // IsChanged 항목은 행 전체를 연한 노란색으로 하이라이트
                if (item.IsChanged &&
                    e.Column.FieldName != nameof(ReeferItem.GangBox_ID) &&
                    e.Column.FieldName != nameof(ReeferItem.GangBox_Color) &&
                    e.Column.FieldName != nameof(ReeferItem.Before_RDB_Color))
                {
                    e.Appearance.BackColor = Color.LightYellow;
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
                    return;
                }
                // RDB_Color 컬럼만 칠하기
                if (e.Column.FieldName == nameof(ReeferItem.GangBox_ID) || e.Column.FieldName == nameof(ReeferItem.GangBox_Color))
                {
                    string colorText = !string.IsNullOrWhiteSpace(item.GangBox_ID)
                        ? item.GangBox_ID
                        : item.GangBox_Color;

                    Color backColor;
                    try
                    {
                        backColor = ColorTranslator.FromHtml(colorText);
                    }
                    catch
                    {
                        return;
                    }

                    e.Appearance.BackColor = backColor;
                    e.Appearance.ForeColor = GetContrastColor(backColor);
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
                }
                else if (e.Column.FieldName == nameof(ReeferItem.Before_RDB_Color))
                {
                    Color backColor;
                    try
                    {
                        backColor = ColorTranslator.FromHtml(item.Before_RDB_Color);
                    }
                    catch
                    {
                        return;
                    }

                    e.Appearance.BackColor = backColor;
                    e.Appearance.ForeColor = GetContrastColor(backColor);
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
                }
                else
                {
                    return;
                }


            }
            catch (Exception ex)
            {
                Debug.WriteLine($"DetailRowsView_RowCellStyle 오류: {ex}");
            }
        }

        private Color GetContrastColor(Color bg)
        {
            int brightness = (bg.R * 299 + bg.G * 587 + bg.B * 114) / 1000;
            return brightness >= 128 ? Color.Black : Color.White;
        }

        private string Get_RDB_Name(string rdbid)
        {
            if (string.IsNullOrWhiteSpace(rdbid))
                return string.Empty;

            string rdbNoText = _excelService.ExtractRdbNumber(rdbid.Trim());
            if (!int.TryParse(rdbNoText, out int rdbNo))
                return string.Empty;

            return $"EP46-2{rdbNo.ToString("D2")}";
        }

        private string Get_Gangbox_Name(string gangboxid, string rdbid)
        {
            if (string.IsNullOrWhiteSpace(gangboxid) || string.IsNullOrWhiteSpace(rdbid))
                return string.Empty;

            string rdbNoText = _excelService.ExtractRdbNumber(rdbid.Trim());
            string gangboxNoText = _excelService.ExtractRdbNumber(gangboxid.Trim());

            if (!int.TryParse(rdbNoText, out int rdbNo) ||
                !int.TryParse(gangboxNoText, out int gangboxNo))
            {
                return string.Empty;
            }

            return $"EP75-{rdbNo.ToString("D2")}{gangboxNo.ToString("D2")}";
        }

        private string Get_Reefer_Name(ReeferItem reefer)
        {
            if (reefer == null)
                return string.Empty;

            return $"{reefer.Place}-{reefer.BayNo}-{reefer.RowNo}-{reefer.TierNo}-{reefer.CHNo}";           
        }

        private void CacheEbObject(IAucObject ebObject,string designation,Dictionary<string, IAucObject> designationCache)
        {
            // IAucObject COM 참조는 보관하지 않는다.
            // 실제 객체가 필요할 때 GetSnglObjectByID()로 새 참조를 가져온다.
        }

        private void TryAddExistingObjectToMap(IAucObject folder,string designation,Dictionary<string, string> idMap)
        {
            if (folder == null || string.IsNullOrWhiteSpace(designation) ||
                idMap == null || idMap.ContainsKey(designation))
            {
                return;
            }

            IAucObjects candidates = folder.FindObjectsEx(
                ebHelper.Get_filter(designation),
                AucVbSearchType.aucSearchDeep,
                true);
            for (int index = 1; index <= candidates.Count; index++)
            {
                IAucObject candidate = candidates.Item(index);
                if (candidate != null && string.Equals(
                    candidate.Name,
                    designation,
                    StringComparison.OrdinalIgnoreCase))
                {
                    idMap[designation] = candidate.ID;
                    return;
                }
            }
        }

        private IAucObject GetCachedObjectById(string objectId)
        {
            if (string.IsNullOrWhiteSpace(objectId) || ebapp == null)
                return null;

            // EB COM 객체는 Store/MoveTo/Association 이후 기존 참조가 stale 상태가 될 수 있다.
            // ID는 캐시하되, 실제 작업에 사용할 IAucObject는 매번 EB에서 새로 조회한다.
            IAucObject ebObject = ebapp.Utils.GetSnglObjectByID(objectId);
            return ebObject;
        }

        private IAucObject GetCachedObjectByDesignation(string designation,Dictionary<string, string> idMap,
            Dictionary<string, IAucObject> designationCache)
        {
            if (string.IsNullOrWhiteSpace(designation))
                return null;

            if (!idMap.TryGetValue(designation, out string objectId))
                return null;

            // designationCache에는 COM 객체를 재사용하지 않고 ID 맵의 객체를 매번 재조회한다.
            return GetCachedObjectById(objectId);
        }

        private int GetGangBoxIndexFromColor(string color)
        {
            string normalizedColor = _excelService.NormalizeHtmlColor(color);
            for (int colorIndex = 0; colorIndex < _groupColorPalette.Length; colorIndex++)
            {
                if (string.Equals(
                    normalizedColor,
                    _excelService.NormalizeHtmlColor(_groupColorPalette[colorIndex]),
                    StringComparison.OrdinalIgnoreCase))
                {
                    return colorIndex + 1;
                }
            }

            return -1;
        }

        /// <summary>
        /// Excel Import는 현재 화면의 데이터를 이어 붙이는 작업이 아니라
        /// 선택한 파일을 기준으로 새 작업을 시작한다.
        ///
        /// 특히 프로그램 시작 시 Output을 ReBuild한 뒤 Import하는 경우에는
        /// raw_List에 ReBuild 결과가 이미 들어 있으므로, 이를 남겨 둔 채
        /// 새 Excel을 읽으면 동일한 슬롯이 중복되어 Gang 수가 부풀려진다.
        /// </summary>
        private void ResetImportSessionState()
        {
            raw_List.Clear();
            grouped_List.Clear();
            _pendingOriginalReeferColors.Clear();

            RDB_dict.Clear();
            Gangbox_dict.Clear();
            Reefer_dict.Clear();

            _rdbObjectCache.Clear();
            _gangboxObjectCache.Clear();
            _reeferObjectCache.Clear();
            _ebObjectCacheById.Clear();

            _hasRenderedOrderOptions = false;
            _renderedBayAsc = 0;
            _renderedTierAsc = 0;
            _renderedRowAsc = 0;
        }

        private void Create_Dict_All(bool splashscreen)
        {
            // 기존 EB 객체 ID Dictionary 다시 구성
            RDB_dict.Clear();
            Gangbox_dict.Clear();
            Reefer_dict.Clear();
            _rdbObjectCache.Clear();
            _gangboxObjectCache.Clear();
            _reeferObjectCache.Clear();
            _ebObjectCacheById.Clear();

            // 이전 생성된 EB 객체 탐색
            if (splashscreen)
                SplashScreenManager.Default.SetWaitFormDescription($"Loading... EB WorkSheet... Reefer");
            object attrs = AucAttribute.aucAttrDesignation;
            IAucWorksheet ebReefer_ws = Reefer_folder.OpenWorksheetWithFilter(ebHelper.Get_filter("*"), AucWorksheetOpenType.aucWorksheetOpenStructDeep, ref attrs);
            for (int i = 1; i <= ebReefer_ws.RowCount; i++)
            {
                IAucObject obj = ebReefer_ws.GetObjectItem(i);
                if (!Reefer_dict.ContainsKey(obj.Name))
                {
                    Reefer_dict[obj.Name] = obj.ID;
                    CacheEbObject(obj, obj.Name, _reeferObjectCache);
                }
            }
            if (splashscreen)
                SplashScreenManager.Default.SetWaitFormDescription($"Loading... EB WorkSheet... RDB");

            IAucWorksheet ebRdb_ws = RDB_folder.OpenWorksheetWithFilter(ebHelper.Get_filter("EP46-2*"), AucWorksheetOpenType.aucWorksheetOpenStructDeep, ref attrs);
            for (int i = 1; i <= ebRdb_ws.RowCount; i++)
            {
                IAucObject obj = ebRdb_ws.GetObjectItem(i);
                if (!RDB_dict.ContainsKey(obj.Name))
                {
                    RDB_dict[obj.Name] = obj.ID;
                    CacheEbObject(obj, obj.Name, _rdbObjectCache);
                }
            }
            if (splashscreen)
                SplashScreenManager.Default.SetWaitFormDescription($"Loading... EB WorkSheet... Gangbox");

            IAucWorksheet ebGangbox_ws = Gangbox_folder.OpenWorksheetWithFilter(ebHelper.Get_filter("EP75-*"), AucWorksheetOpenType.aucWorksheetOpenStructDeep, ref attrs);
            for (int i = 1; i <= ebGangbox_ws.RowCount; i++)
            {
                IAucObject obj = ebGangbox_ws.GetObjectItem(i);
                if (!Gangbox_dict.ContainsKey(obj.Name))
                {
                    Gangbox_dict[obj.Name] = obj.ID;
                    CacheEbObject(obj, obj.Name, _gangboxObjectCache);
                }
            }
        }

        private void AssignGangBoxCandidates(List<GangBoxAssignmentCandidate> candidates,Dictionary<string, RdbAccumulator> rdbStateByKey, 
            ref int nextRdbNo,ref int groupColorIndex)
        {
            var options = new GroupingOptions
            {
                MaxGangCount = Max_gang_cnt,
                MaxRdbReeferCount = Max_rdb_reefer_cnt,
                DeckHoldType = deck_hold_type,
                ProcessDeckBeforeHold = _processDeckBeforeHold,
                GangBoxColorPalette = _groupColorPalette,
                RdbColorPalette = _rdbColorPalette
            };
            var ordered = _groupingService.OrderByPlacePolicy(OrderGangBoxCandidatesByPriority(candidates),options);
            foreach (var candidate in ordered)
                _groupingService.AssignCandidate(candidate,rdbStateByKey,options,OrderReeferItemsForPosition,Get_RDB_Name,ref nextRdbNo);

            // RDB 할당 후 색상과 GangBox 번호를 서로 독립적으로 부여한다.
            // 색상은 PORT/STBD 대칭, 번호는 UI Order By 진행 순서를 따른다.
            AssignGangBoxNumbersAndSymmetricColors(candidates);
        }



        private void bnt_setting_Click(object sender, EventArgs e)
        {
            panel_setting.Visible = !panel_setting.Visible;
            panel_setting.BringToFront();
        }

        private void btn_excel_import_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx;*.xls";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string filePath = ofd.FileName;

                    ExcelFile import_excel = Gem.ExcelFile.Load(filePath);
                    if (import_excel.Worksheets.Count < 1)
                        throw new InvalidOperationException("엑셀에 최소 1개의 시트가 필요합니다. (1시트: 원본)");

                    // ReBuild 이후 Import하는 경우에도 이전 실행의 raw/group 데이터가
                    // 새 Import 데이터에 섞이지 않도록 Import 세션을 새로 시작한다.
                    ResetImportSessionState();

                    SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);

                    DateTime start_time = DateTime.Now;
                    Debug.WriteLine($"[START] {start_time:HH:mm:ss.fff}");
                    Stopwatch st = Stopwatch.StartNew();

                    // 1시트 원본
                    SplashScreenManager.Default.SetWaitFormDescription($"Loading... Excel Import...");
                    ExcelWorksheet ws = import_excel.Worksheets[0];
                    for (int row = 2; row <= ws.Rows.Count(); row++)
                    {
                        if (string.IsNullOrEmpty(GemBoxHelper.GetCellText(ws, row, 0)))
                            continue;

                        var item = new ReeferItem
                        {
                            TagNo = GemBoxHelper.GetCellText(ws, row, 0),
                            Reefer = GemBoxHelper.GetCellText(ws, row, 1) == "TRUE" ? true : false,
                            BayNo = int.Parse(GemBoxHelper.GetCellText(ws, row, 2)),
                            TierNo = int.Parse(GemBoxHelper.GetCellText(ws, row, 3)),
                            RowNo = int.Parse(GemBoxHelper.GetCellText(ws, row, 4)),
                            CHNo = GemBoxHelper.GetCellText(ws, row, 5),
                            Place = GemBoxHelper.GetCellText(ws, row, 6)
                        };

                        raw_List.Add(item);
                        
                    }
                    // SortPriorityPanel에서 정렬 방향 읽기
                    var sortItems = sortPriorityPanel.GetPriorityItems();
                    //int Bay_asc = sortItems.FirstOrDefault(x => x.Key == "Bay")?.SelectedOrder == "Fwd->Aft" ? 0 : 1;
                    int Bay_asc = GetBayAscFromPanel();
                    int Tier_asc = sortItems.FirstOrDefault(x => x.Key == "Tier")?.SelectedOrder == "Top->Bottom" ? 0 : 1;
                    int Row_asc = sortItems.FirstOrDefault(x => x.Key == "Row")?.SelectedOrder == "PORT->ST'BD" ? 1 : 0;
                    // 정렬기준 저장
                    SaveRenderedOrderOptions(Bay_asc, Tier_asc, Row_asc);

                    grouped_List = (deck_hold_type == 0
                                    ? raw_List
                                    .GroupBy(x => new
                                    {
                                        x.BayNo,
                                        x.CHNo,
                                        x.Place
                                    })
                                    .Select(g => new BayGroupDto
                                    {
                                        BayNo = g.Key.BayNo,
                                        CHNo = g.Key.CHNo,
                                        Place = g.Key.Place,
                                        Tiers = (Tier_asc == 0
                                                 ? g.GroupBy(x => x.TierNo)
                                                  .Select(t => new TierGroupDto
                                                  {
                                                      TierNo = t.Key,
                                                      Rows = _excelService.BuildRowList(t, Row_asc)
                                                  })
                                                 .OrderByDescending(x => x.TierNo)
                                                 .ToList()
                                                 : g.GroupBy(x => x.TierNo)
                                                  .Select(t => new TierGroupDto
                                                  {
                                                      TierNo = t.Key,
                                                      Rows = _excelService.BuildRowList(t, Row_asc)
                                                  })
                                                 .OrderBy(x => x.TierNo)
                                                 .ToList())
                                    }).ToList()
                                    : raw_List
                                    .GroupBy(x => new
                                    {
                                        x.BayNo,
                                        x.CHNo,
                                        x.Place
                                    })
                                    .Select(g => new BayGroupDto
                                    {
                                        BayNo = g.Key.BayNo,
                                        CHNo = g.Key.CHNo,
                                        Place = g.Key.Place,
                                        Tiers = (Tier_asc == 0
                                                 ? g.GroupBy(x => x.TierNo)
                                                  .Select(t => new TierGroupDto
                                                  {
                                                      TierNo = t.Key,
                                                      Rows = _excelService.BuildRowList(t, Row_asc)
                                                  })
                                                 .OrderByDescending(x => x.TierNo)
                                                 .ToList()
                                                 : g.GroupBy(x => x.TierNo)
                                                  .Select(t => new TierGroupDto
                                                  {
                                                      TierNo = t.Key,
                                                      Rows = _excelService.BuildRowList(t, Row_asc)
                                                  })
                                                 .OrderBy(x => x.TierNo)
                                                 .ToList())
                                    })                                    
                                    .ToList()
                                    );

                    grouped_List = (Bay_asc == 0 ? grouped_List.OrderBy(x => x.BayNo) : grouped_List.OrderByDescending(x => x.BayNo))
                                    .ThenBy(g => g.CHNo)                                  
                                    .ToList();

                    grouped_List = ApplyDeckFirstProcessingOrder(grouped_List);

                    gridControl1.DataSource = grouped_List;
                    if (deck_hold_type == 1)
                        gridView1.Columns["Place"].Visible = false;
                    gridView1.Columns["Warning"].Visible = false;

                    TemplateMode mode = _excelService.DetectTemplateMode(grouped_List);

                    switch (mode)
                    {
                        case TemplateMode.Even:
                            _evenGroupingService.Group(grouped_List, Max_gang_cnt, AssignGangBoxCandidates);
                            break;
                        case TemplateMode.Odd:
                            _oddGroupingService.Group(grouped_List, Max_gang_cnt, Max_rdb_reefer_cnt, _evenGroupingService, AssignGangBoxCandidates);
                            break;
                        case TemplateMode.Mixed:
                            throw new InvalidOperationException("짝수/홀수 데이터가 혼합되어 있습니다.");
                        default:
                            throw new InvalidOperationException("템플릿 모드를 판별할 수 없습니다.");
                    }


                    //            
                    if (mode == TemplateMode.Even)
                    {
                        templatePath = Path.Combine(downloadPath, "(짝수)Reefer Container Template.xlsx");
                        template_FileName = "(짝수)Reefer Container Template.xlsx";
                        savePath = outputPath;
                    }
                    else if (mode == TemplateMode.Odd)
                    {
                        templatePath = Path.Combine(downloadPath, "(홀수)Reefer Container Template.xlsx");
                        template_FileName = "(홀수)Reefer Container Template.xlsx";
                        savePath = outputPath;
                    }
                    else
                    {
                        throw new InvalidOperationException("짝수/홀수 파일탐색에 실패했습니다.");
                    }

                    if (File.Exists(templatePath))
                    {
                        File.Delete(templatePath);
                    }

                    if (true)
                    {
                        IAucObject template_folder = Project.TemplatesFolder
                                               .FindObjectsEx(ebHelper.Get_filter("Settings"), AucVbSearchType.aucSearchFlat, true).Item(1)
                                               .FindObjectsEx(ebHelper.Get_filter("template"), AucVbSearchType.aucSearchFlat, true).Item(1);

                        IAucObjects template_files = template_folder.FindObjectsEx(ebHelper.Get_filter(template_FileName), AucVbSearchType.aucSearchFlat, true);
                        if (template_files.Count == 0)
                        {
                            MessageBox.Show($"아래 경로의 파일을 확인해주세요.\r\n1. Project -> Templates -> Settings -> template -> Template Excel \r\n2. {templatePath}");
                            throw new InvalidOperationException("짝수/홀수 템플릿 판별에 실패했습니다.");
                        }

                        ExternalDocument temp = (ExternalDocument)template_files.Item(1);

                        temp.SaveAs(templatePath);
                        if (!File.Exists(templatePath))
                        {
                            MessageBox.Show($"아래 경로의 파일을 확인해주세요.\r\n1. Project -> Templates -> Settings -> template -> Template Excel \r\n2. {templatePath}");
                            throw new InvalidOperationException("짝수/홀수 템플릿 판별에 실패했습니다.");
                        }
                    }

                    var workbook = ExcelFile.Load(templatePath);
                    if (workbook.Worksheets.Count == 0)
                    {
                        MessageBox.Show("워크시트가 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var ws2 = _updateService.FindWorksheetByName(workbook, DefaultRevisionName) ?? workbook.Worksheets[0];
                    if (!string.Equals(ws2.Name, DefaultRevisionName, StringComparison.OrdinalIgnoreCase))
                        ws2.Name = DefaultRevisionName;

                    workbook.Worksheets.ActiveWorksheet = ws2;
                    tb_revision.Text = DefaultRevisionName;
                    _excelService.FillGroupedDataToTemplate(
                        ws2,
                        grouped_List,
                        mode,
                        _hasRenderedOrderOptions ? _renderedBayAsc : GetBayAscFromPanel());

                    workbook.Save(savePath);

                    var wb_last = ExcelFile.Load(savePath);
                    var ws_last = _updateService.FindWorksheetByName(wb_last, DefaultRevisionName) ?? wb_last.Worksheets[0];
                    ReplaceEBTextboxValue(ws_last, wb_last);

                    sheet8grid_tfn = true;
                    spreadsheetControl1.LoadDocument(savePath);
                    spreadsheetControl1.BringToFront();

                    // save to EB
                    savePath = outputPath;
                    SaveToEB_UpLoad();

                    st.Stop();
                    DateTime endTime = DateTime.Now;
                    Debug.WriteLine($"[END]   {endTime:HH:mm:ss.fff} / {st.Elapsed.TotalSeconds:N3} sec == {st.Elapsed.TotalMilliseconds:N3} ms");

                    SplashScreenManager.CloseForm();
                    MessageBox.Show("완료되었습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        private void SaveRenderedOrderOptions(int bayAsc,int tierAsc,int rowAsc)
        {
            _renderedBayAsc = bayAsc;
            _renderedTierAsc = tierAsc;
            _renderedRowAsc = rowAsc;
            _hasRenderedOrderOptions = true;

            Debug.WriteLine(
                $"[ORDER SAVE] BayAsc={_renderedBayAsc}, TierAsc={_renderedTierAsc}, " +
                $"RowAsc={_renderedRowAsc}");
        }
        private int GetBayAscFromPanel()
        {
            var sortItems = sortPriorityPanel.GetPriorityItems();

            string selected = sortItems
                .FirstOrDefault(x => x.Key == "Bay")
                ?.SelectedOrder ?? "";

            selected = selected.Replace(" ", "").Trim().ToUpperInvariant();

            if (selected == "AFT->FWD")
                return 1;

            return 0;
        }

        /// <summary>
        /// 그룹 형상은 Order By와 무관하게 PORT/STBD를 항상 분리해 만든다.
        /// 생성 순서는 고정하고, 완성된 그룹의 numbering만 사용자 Order By를 따른다.
        /// </summary>
        private static IEnumerable<string> GetSymmetricGroupingSideOrder()
        {
            return new[] { "PORT", "STBD" };
        }

        private List<GangBoxAssignmentCandidate> OrderGangBoxCandidatesByPriority(List<GangBoxAssignmentCandidate> candidates, 
            bool firstPriorityMovesFirst = false)
        {
            if (candidates == null || candidates.Count < 2)
                return candidates ?? new List<GangBoxAssignmentCandidate>();

            Func<GangBoxAssignmentCandidate, int> getBayOrder = candidate =>
                _renderedBayAsc == 0
                    ? candidate.Bay.BayNo
                    : -candidate.Bay.BayNo;

            Func<GangBoxAssignmentCandidate, int> getTierOrder = candidate =>
            {
                int fallbackTierNo = _renderedTierAsc == 0 ? 0 : int.MaxValue;
                int tierNo = candidate.TierBand == null
                    ? fallbackTierNo
                    : candidate.TierBand
                        .Where(x => x != null)
                        .Select(x => x.TierNo)
                        .DefaultIfEmpty(fallbackTierNo)
                        .First();

                return _renderedTierAsc == 0 ? -tierNo : tierNo;
            };

            Func<GangBoxAssignmentCandidate, int> getRowOrder = candidate => candidate.Items == null
                ? int.MaxValue
                : candidate.Items
                    .Where(x => x != null)
                    .Select(x => GetTraversalRowOrderKey(
                        x.RowNo,
                        _renderedRowAsc,
                        candidate.SideName))
                    .DefaultIfEmpty(int.MaxValue)
                    .Min();

            var keySelectors = new List<Func<GangBoxAssignmentCandidate, int>>();
            var priorityItems = sortPriorityPanel?.GetPriorityItems();

            if (priorityItems != null)
            {
                foreach (var priorityItem in priorityItems)
                {
                    Func<GangBoxAssignmentCandidate, int> keySelector = null;

                    if (string.Equals(priorityItem?.Key, "Bay", StringComparison.OrdinalIgnoreCase))
                        keySelector = getBayOrder;
                    else if (string.Equals(priorityItem?.Key, "Tier", StringComparison.OrdinalIgnoreCase))
                        keySelector = getTierOrder;
                    else if (string.Equals(priorityItem?.Key, "Row", StringComparison.OrdinalIgnoreCase))
                        keySelector = getRowOrder;

                    if (keySelector == null)
                        continue;

                    keySelectors.Add(keySelector);
                }
            }

            // GangBox 번호만 다시 부여할 때는 UI의 1순위가 가장 먼저 변해야 한다.
            // RDB 할당에서는 false를 사용하므로 기존 정렬과 번호가 유지된다.
            if (firstPriorityMovesFirst)
                keySelectors.Reverse();

            var result = candidates.ToList();
            result.Sort((left, right) =>
            {
                foreach (var keySelector in keySelectors)
                {
                    int comparison = keySelector(left).CompareTo(keySelector(right));
                    if (comparison != 0)
                        return comparison;
                }

                return left.CreationOrder.CompareTo(right.CreationOrder);
            });

            return result;
        }

        private void AssignGangBoxNumbersAndSymmetricColors(
            List<GangBoxAssignmentCandidate> candidates)
        {
            if (candidates == null || candidates.Count == 0 ||
                _groupColorPalette == null || _groupColorPalette.Length == 0)
            {
                return;
            }

            var validCandidates = candidates
                .Where(candidate =>
                    candidate?.Items != null &&
                    candidate.Items.Any(item => item != null && !string.IsNullOrWhiteSpace(item.RDB_ID)))
                .OrderBy(candidate => candidate.CreationOrder)
                .ToList();

            // GangBox 번호는 전체 그룹의 전역 번호가 아니라 RDB 내부 번호다.
            // 예: 07 RDB의 첫 번째/두 번째 GangBox는 EP75-0701/EP75-0702.
            var nextGangBoxNoByRdb =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var bayBlocks = validCandidates.GroupBy(candidate => new
            {
                candidate.Bay.BayNo,
                CHNo = (candidate.Bay.CHNo ?? string.Empty).Trim().ToUpperInvariant(),
                Place = _excelService.NormalizePlace(candidate.Bay.Place)
            });

            foreach (var bayBlock in bayBlocks)
            {
                // 양쪽을 모두 외측 -> 중심 순서로 놓고 같은 palette index를 부여한다.
                // 그 결과 실제 PORT -> STBD 화면에서는
                // PORT: 빨강,주황 / STBD: 주황,빨강처럼 완전 대칭이 된다.
                foreach (string side in GetSymmetricGroupingSideOrder())
                {
                    int sideColorIndex = 0;
                    foreach (var candidate in bayBlock
                        .Where(x => string.Equals(x.SideName, side, StringComparison.OrdinalIgnoreCase))
                        .OrderBy(x => x.CreationOrder))
                    {
                        string gangBoxColor =
                            _groupColorPalette[sideColorIndex % _groupColorPalette.Length];

                        foreach (var item in candidate.Items)
                        {
                            if (item == null)
                                continue;

                            item.GangBox_ID = gangBoxColor;
                            item.GangBox_Color = gangBoxColor;
                        }

                        sideColorIndex++;
                    }
                }

                // PORT/STBD가 아닌 예외 candidate도 안정적으로 색상을 받도록 한다.
                int fallbackColorIndex = 0;
                foreach (var candidate in bayBlock
                    .Where(x =>
                        !string.Equals(x.SideName, "PORT", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(x.SideName, "STBD", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(x => x.CreationOrder))
                {
                    string gangBoxColor =
                        _groupColorPalette[fallbackColorIndex % _groupColorPalette.Length];
                    foreach (var item in candidate.Items.Where(x => x != null))
                    {
                        item.GangBox_ID = gangBoxColor;
                        item.GangBox_Color = gangBoxColor;
                    }
                    fallbackColorIndex++;
                }

                // GangBox 번호는 색상 index와 무관하게 Order By 순서로 증가한다.
                foreach (var candidate in OrderGangBoxCandidatesByPriority(
                    bayBlock.ToList(),
                    firstPriorityMovesFirst: true))
                {
                    string rdbDesignation = candidate.Items
                        .Where(x => x != null && !string.IsNullOrWhiteSpace(x.RDB_ID))
                        .Select(x => Get_RDB_Name(x.RDB_ID))
                        .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
                    if (string.IsNullOrWhiteSpace(rdbDesignation))
                        continue;

                    nextGangBoxNoByRdb.TryGetValue(rdbDesignation, out int currentGangBoxNo);
                    int gangBoxNo = currentGangBoxNo + 1;

                    foreach (var item in candidate.Items.Where(x => x != null))
                        item.GangBox_No = gangBoxNo;

                    nextGangBoxNoByRdb[rdbDesignation] = gangBoxNo;
                }
            }
        }

        /// <summary>
        /// Excel은 색상은 보관하지만 GangBox 번호를 셀에 직접 표시하지 않는다.
        /// 대칭 색상으로 인해 같은 색이 여러 GangBox에 재사용될 수 있으므로,
        /// RDB/색상이 같은 인접 셀을 하나의 논리 그룹으로 복원한 뒤 Order By로 재채번한다.
        /// </summary>
        private void AssignGangBoxNumbersFromExcel(IEnumerable<ReeferItem> items)
        {
            var sourceItems = (items ?? Enumerable.Empty<ReeferItem>())
                .Where(item =>
                    item != null &&
                    item.Reefer &&
                    !string.IsNullOrWhiteSpace(item.RDB_ID) &&
                    GetGangBoxIndexFromColor(item.GangBox_Color) > 0)
                .ToList();

            foreach (var item in sourceItems)
                item.GangBox_No = 0;

            var componentCandidates = new List<GangBoxAssignmentCandidate>();
            var colorGroups = sourceItems.GroupBy(item => new
            {
                item.BayNo,
                CHNo = (item.CHNo ?? string.Empty).Trim().ToUpperInvariant(),
                Place = _excelService.NormalizePlace(item.Place),
                RdbId = Get_RDB_Name(item.RDB_ID),
                Color = _excelService.NormalizeHtmlColor(item.GangBox_Color)
            });

            foreach (var colorGroup in colorGroups)
            {
                var itemByCell = colorGroup
                    .GroupBy(item => item.X + "|" + item.Y)
                    .ToDictionary(group => group.Key, group => group.First());
                var unvisited = new HashSet<ReeferItem>(itemByCell.Values);

                while (unvisited.Count > 0)
                {
                    ReeferItem seed = unvisited.First();
                    unvisited.Remove(seed);

                    var componentItems = new List<ReeferItem>();
                    var queue = new Queue<ReeferItem>();
                    queue.Enqueue(seed);

                    while (queue.Count > 0)
                    {
                        ReeferItem current = queue.Dequeue();
                        componentItems.Add(current);

                        string[] neighborKeys =
                        {
                            (current.X - 1) + "|" + current.Y,
                            (current.X + 1) + "|" + current.Y,
                            current.X + "|" + (current.Y - 1),
                            current.X + "|" + (current.Y + 1)
                        };

                        foreach (string neighborKey in neighborKeys)
                        {
                            if (!itemByCell.TryGetValue(neighborKey, out ReeferItem neighbor) ||
                                !unvisited.Remove(neighbor))
                            {
                                continue;
                            }

                            queue.Enqueue(neighbor);
                        }
                    }

                    var tierBandQuery = componentItems
                        .GroupBy(item => item.TierNo)
                        .Select(group => new TierGroupDto
                        {
                            TierNo = group.Key,
                            Rows = group.ToList()
                        });
                    var tierBand = (_renderedTierAsc == 0
                            ? tierBandQuery.OrderByDescending(tier => tier.TierNo)
                            : tierBandQuery.OrderBy(tier => tier.TierNo))
                        .ToList();

                    componentCandidates.Add(new GangBoxAssignmentCandidate
                    {
                        Bay = new BayGroupDto
                        {
                            BayNo = colorGroup.Key.BayNo,
                            CHNo = colorGroup.Key.CHNo,
                            Place = colorGroup.Key.Place
                        },
                        TierBand = tierBand,
                        Items = componentItems,
                        SideName = GetComponentSideName(componentItems, sourceItems),
                        CreationOrder = componentCandidates.Count
                    });
                }
            }

            // 숨김 번호가 없는 Excel을 복원할 때도 RDB마다 01부터 다시 번호를 부여한다.
            var nextGangBoxNoByRdb =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var bayBlocks = componentCandidates
                .OrderBy(candidate => candidate.CreationOrder)
                .GroupBy(candidate => new
                {
                    candidate.Bay.BayNo,
                    CHNo = (candidate.Bay.CHNo ?? string.Empty).Trim().ToUpperInvariant(),
                    Place = _excelService.NormalizePlace(candidate.Bay.Place)
                });

            foreach (var bayBlock in bayBlocks)
            {
                foreach (var candidate in OrderGangBoxCandidatesByPriority(
                    bayBlock.ToList(),
                    firstPriorityMovesFirst: true))
                {
                    string rdbDesignation = candidate.Items
                        .Where(item => item != null && !string.IsNullOrWhiteSpace(item.RDB_ID))
                        .Select(item => Get_RDB_Name(item.RDB_ID))
                        .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
                    if (string.IsNullOrWhiteSpace(rdbDesignation))
                        continue;

                    nextGangBoxNoByRdb.TryGetValue(rdbDesignation, out int currentGangBoxNo);
                    int gangBoxNo = currentGangBoxNo + 1;

                    foreach (var item in candidate.Items.Where(item => item != null))
                        item.GangBox_No = gangBoxNo;

                    nextGangBoxNoByRdb[rdbDesignation] = gangBoxNo;
                }
            }
        }
        
        private static string GetComponentSideName(IEnumerable<ReeferItem> componentItems,IEnumerable<ReeferItem> allItems)
        {
            var component = (componentItems ?? Enumerable.Empty<ReeferItem>())
                .Where(item => item != null)
                .ToList();
            bool hasPort = component.Any(item => item.RowNo != 0 && item.RowNo % 2 == 0);
            bool hasStbd = component.Any(item => item.RowNo != 0 && item.RowNo % 2 != 0);

            if (hasPort != hasStbd)
                return hasPort ? "PORT" : "STBD";

            ReeferItem centerItem = component.FirstOrDefault(item => item.RowNo == 0);
            if (centerItem == null)
                return "ALL";

            var centerItems = (allItems ?? Enumerable.Empty<ReeferItem>())
                .Where(item =>
                    item != null &&
                    item.RowNo == 0 &&
                    item.BayNo == centerItem.BayNo &&
                    string.Equals(item.CHNo, centerItem.CHNo, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(item.Place, centerItem.Place, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(item => item.TierNo)
                .ThenBy(item => item.X)
                .ToList();
            int centerIndex = centerItems.IndexOf(centerItem);
            return centerIndex >= 0 && centerIndex % 2 != 0 ? "STBD" : "PORT";
        }

        private List<ReeferItem> OrderReeferItemsForPosition(IEnumerable<ReeferItem> items)
        {
            var indexedItems = (items ?? Enumerable.Empty<ReeferItem>())
                .Where(item => item != null && item.Reefer)
                .Select((item, index) => new { Item = item, Index = index })
                .ToList();

            if (indexedItems.Count < 2)
                return indexedItems.Select(x => x.Item).ToList();

            Func<ReeferItem, int> getBayOrder = item =>
                _renderedBayAsc == 0 ? item.BayNo : -item.BayNo;

            Func<ReeferItem, int> getTierOrder = item =>
                _renderedTierAsc == 0 ? -item.TierNo : item.TierNo;

            Func<ReeferItem, int> getRowOrder = item =>
                GetTraversalRowOrderKey(item.RowNo, _renderedRowAsc);

            var keySelectors = new List<Func<ReeferItem, int>>();
            var priorityItems = sortPriorityPanel?.GetPriorityItems();

            if (priorityItems != null)
            {
                foreach (var priorityItem in priorityItems)
                {
                    Func<ReeferItem, int> keySelector = null;

                    if (string.Equals(priorityItem?.Key, "Bay", StringComparison.OrdinalIgnoreCase))
                        keySelector = getBayOrder;
                    else if (string.Equals(priorityItem?.Key, "Tier", StringComparison.OrdinalIgnoreCase))
                        keySelector = getTierOrder;
                    else if (string.Equals(priorityItem?.Key, "Row", StringComparison.OrdinalIgnoreCase))
                        keySelector = getRowOrder;

                    if (keySelector != null)
                        keySelectors.Add(keySelector);
                }
            }

            // UI 1순위가 Position에서 가장 먼저 변하는 축이다.
            // 예: Tier > Row이면 같은 Row에서 Tier 방향으로 1,2,3... 진행한 뒤
            // Row 방향의 다음 열로 이동한다.
            keySelectors.Reverse();

            indexedItems.Sort((left, right) =>
            {
                foreach (var keySelector in keySelectors)
                {
                    int comparison = keySelector(left.Item).CompareTo(keySelector(right.Item));
                    if (comparison != 0)
                        return comparison;
                }

                return left.Index.CompareTo(right.Index);
            });

            return indexedItems.Select(x => x.Item).ToList();
        }

        /// <summary>
        /// PORT/STBD 대칭 그룹 형상을 만들기 위한 side 내부 Row 순서다.
        /// numbering용 Order By 방향과 무관하게 양쪽 모두 외측에서 중심으로 진행한다.
        /// </summary>
        private static List<int> OrderRowsForSymmetricGrouping(IEnumerable<int> rowNumbers)
        {
            if (rowNumbers == null)
                return new List<int>();

            return rowNumbers
                .Distinct()
                // PORT 20,18,...,2 / STBD 19,17,...,1.
                // Center 0은 해당 side의 가장 안쪽으로 보낸다.
                .OrderByDescending(rowNo => rowNo)
                .ToList();
        }

        /// <summary>
        /// 완성된 그룹과 그룹 내부 Reefer numbering에만 사용하는
        /// 선박 폭 방향의 실제 이동 순서다.
        ///   PORT -> STBD: 20,18,...,2,0,1,3,...,19
        ///   STBD -> PORT: 19,17,...,1,0,2,4,...,20
        /// </summary>
        private static int GetTraversalRowOrderKey(int rowNo,int rowAsc,string centerSideName = null)
        {
            if (rowNo == 0)
            {
                if (string.Equals(centerSideName, "PORT", StringComparison.OrdinalIgnoreCase))
                    return rowAsc == 1 ? 0 : 2000;

                if (string.Equals(centerSideName, "STBD", StringComparison.OrdinalIgnoreCase))
                    return rowAsc == 1 ? 2000 : 0;

                return 1000;
            }

            bool isPort = rowNo % 2 == 0;
            int section;
            int value;

            if (rowAsc == 1) // PORT -> STBD
            {
                section = isPort ? 0 : 2;
                value = isPort ? -rowNo : rowNo;
            }
            else // STBD -> PORT
            {
                section = isPort ? 2 : 0;
                value = isPort ? rowNo : -rowNo;
            }

            return (section * 1000) + value;
        }

        private List<BayGroupDto> ApplyDeckFirstProcessingOrder(List<BayGroupDto> groups)
        {
            if (!IsDeckHoldProcessingEnabled || groups == null || groups.Count < 2)
                return groups;

            // 체크 시 기존 Bay/CH 정렬 결과의 순서는 유지하고,
            // Place만 최우선으로 정렬하여 ON DECK 전체를 먼저 처리한다.
            return groups
                .Select((group, index) => new { Group = group, Index = index })
                .OrderBy(x => GetPlaceProcessingPriority(x.Group?.Place))
                .ThenBy(x => x.Index)
                .Select(x => x.Group)
                .ToList();
        }

        private static int GetPlaceProcessingPriority(string place)
        {
            string normalized = (place ?? string.Empty).Trim().ToUpperInvariant();

            if (normalized == "ON DECK")
                return 0;

            if (normalized == "IN HOLD")
                return 1;

            return 2;
        }

        private List<BayLayoutTemplateDto> BuildDrawSlotsFromGroupedList(List<BayGroupDto> groupedList)
        {
            if (groupedList == null || groupedList.Count == 0)
                return new List<BayLayoutTemplateDto>();

            //int bayAsc = GetBayAscFromPanel();
            int bayAsc = _hasRenderedOrderOptions ? _renderedBayAsc : GetBayAscFromPanel();

            var slots = groupedList
                .Where(x => x != null)
                .GroupBy(x => new
                {
                    x.BayNo,
                    CHNo = (x.CHNo ?? "").Trim()
                })
                .Select(g => new BayLayoutTemplateDto
                {
                    BayNo = g.Key.BayNo,
                    CHNo = g.Key.CHNo
                });

            var sorted = bayAsc == 0
                ? slots.OrderBy(x => x.BayNo).ThenBy(x => x.CHNo).ToList()
                : slots.OrderByDescending(x => x.BayNo).ThenBy(x => x.CHNo).ToList();

            for (int i = 0; i < sorted.Count; i++)
                sorted[i].SlotOrder = i + 1;

            return sorted;
        }

        private int GetTagNoNumeric(string tagNo)
        {
            if (string.IsNullOrWhiteSpace(tagNo))
                return int.MaxValue;

            var digits = new string(tagNo.Where(char.IsDigit).ToArray());

            return int.TryParse(digits, out int value) ? value : int.MaxValue;
        }

        private readonly string[] _groupColorPalette =
        {
            "#FF0000", // 빨강
            "#FFA500", // 주황
            "#FFFF00", // 노랑
            "#00B050", // 초록
            "#0070C0", // 파랑
            //"#002060", // 남색
            "#7030A0", // 보라
            //"#000000", // 검정
            "#00B0F0", // 청록
            "#808080", // 회색
            "#FFC0CB", // 분홍
            "#FF00FF", // 자홍
            "#92D050", // 연두
            "#8B4513", // 갈색
            "#C0C0C0"  // 은색
        };

        private readonly string[] _rdbColorPalette =
        {
            //"#000000", // 검정
            "#FF0000", // 빨강
            "#0070C0", // 파랑
            "#00B050", // 초록
            "#7030A0", // 보라
            "#FFA500", // 주황
            "#8B4513", // 갈색
            "#FF00FF", // 자홍
            "#00B0F0", // 청록
            "#808080", // 회색
            "#C00000", // 진빨강
            "#548235", // 진초록
            "#1F4E79", // 진파랑
            "#806000", // 진노랑/갈색
            "#660066"  // 진보라
        };

        private void btn_create_data_Click(object sender, EventArgs e)
        {
            try
            {
                if (grouped_List == null || grouped_List.Count == 0)
                {
                    MessageBox.Show("그룹 데이터가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (OpenFileDialog dlg = new OpenFileDialog())
                {
                    dlg.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls";
                    dlg.Title = "템플릿 파일 선택";

                    if (dlg.ShowDialog() != DialogResult.OK)
                        return;

                    string templatePath = dlg.FileName;

                    using (SaveFileDialog saveDlg = new SaveFileDialog())
                    {
                        saveDlg.Filter = "Excel Files (*.xlsx)|*.xlsx";
                        saveDlg.Title = "저장할 파일명 선택";
                        saveDlg.FileName = $"{Path.GetFileNameWithoutExtension(templatePath)}_output.xlsx";

                        if (saveDlg.ShowDialog() != DialogResult.OK)
                            return;

                        savePath = saveDlg.FileName;

                        var workbook = ExcelFile.Load(templatePath);
                        if (workbook.Worksheets.Count == 0)
                        {
                            MessageBox.Show("워크시트가 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        var ws = workbook.Worksheets[0];

                        TemplateMode mode = _excelService.DetectTemplateMode(grouped_List);

                        var allItems = grouped_List
                                    .Where(b => b?.Tiers != null)
                                    .SelectMany(b => b.Tiers)
                                    .Where(t => t?.Rows != null)
                                    .SelectMany(t => t.Rows)
                                    .Where(x => x != null && x.Reefer)
                                    .ToList();

                        int withGangBox = allItems.Count(x => !string.IsNullOrEmpty(x.GangBox_ID));
                        int withRdb = allItems.Count(x => !string.IsNullOrEmpty(x.RDB_ID));

                        MessageBox.Show(
                            $"전체 Reefer: {allItems.Count}\n" +
                            $"GangBox_ID 있음: {withGangBox}\n" +
                            $"RDB_ID 있음: {withRdb}\n" +
                            $"GangBox_ID 예시: {allItems.FirstOrDefault(x => !string.IsNullOrEmpty(x.GangBox_ID))?.GangBox_ID}");


                        _excelService.FillGroupedDataToTemplate(
                            ws,
                            grouped_List,
                            mode,
                            _hasRenderedOrderOptions ? _renderedBayAsc : GetBayAscFromPanel());

                        workbook.Save(savePath);

                        var wb_last = ExcelFile.Load(savePath);
                        var ws_last = wb_last.Worksheets[0];
                        ReplaceEBTextboxValue(ws_last, wb_last);

                        MessageBox.Show("완료되었습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void ReplaceEBTextboxValue(ExcelWorksheet ws, ExcelFile wb)
        {
            const int totalRowIndex = 156; // Excel 157행 (zero-based)
            wb.Calculate();
            TemplateMode mode = _excelService.DetectTemplateMode(grouped_List);
            foreach (var shape in ws.Shapes)
            {
                string text = shape.Text.ToString();
                Debug.WriteLine($"{text}");
                if(text.StartsWith("ON DECK"))
                {
                    Debug.WriteLine("ON DECK");
                    Debug.WriteLine($"A157  = {ws.Cells[totalRowIndex, 0].Value}, Formula = {ws.Cells[totalRowIndex, 0].Formula}");
                    if (mode is TemplateMode.Even)
                    {
                        Debug.WriteLine($"AD157 = {ws.Cells[totalRowIndex, 29].Value}, Formula = {ws.Cells[totalRowIndex, 29].Formula}");
                    }
                    else
                    {
                        Debug.WriteLine($"AE157 = {ws.Cells[totalRowIndex, 30].Value}, Formula = {ws.Cells[totalRowIndex, 30].Formula}");
                    }
                    int left = mode is TemplateMode.Even ? int.Parse(ws.Cells[totalRowIndex, 0].Value?.ToString()) : int.Parse(ws.Cells[totalRowIndex, 0].Value?.ToString());
                    int right = mode is TemplateMode.Even ? int.Parse(ws.Cells[totalRowIndex, 29].Value?.ToString()) : int.Parse(ws.Cells[totalRowIndex, 30].Value?.ToString());
                    text = text.Replace("#EB:TOTAL", $"{left + right}");
                    Debug.WriteLine($"text: {text}");
                    shape.Text.Paragraphs.Clear();
                    Gem.Drawing.TextRun run = shape.Text.Paragraphs.Add().Elements.AddRun(text);
                    run.Format.Font = "Arial";
                    run.Format.Size = Length.From(13, LengthUnit.Point);
                    run.Format.Fill.SetSolid(Gem.Drawing.DrawingColor.FromRgb(255, 0, 0));
                }
                else if (text.StartsWith("IN HOLD"))
                {
                    Debug.WriteLine("IN HOLD");                    
                    if (mode is TemplateMode.Even)
                    {                       
                        Debug.WriteLine($"AF157 = {ws.Cells[totalRowIndex, 31].Value}, Formula = {ws.Cells[totalRowIndex, 31].Formula}");
                        Debug.WriteLine($"BG157 = {ws.Cells[totalRowIndex, 58].Value}, Formula = {ws.Cells[totalRowIndex, 58].Formula}");
                    }
                    else
                    {                        
                        Debug.WriteLine($"AG157 = {ws.Cells[totalRowIndex, 32].Value}, Formula = {ws.Cells[totalRowIndex, 32].Formula}");
                        Debug.WriteLine($"BH157 = {ws.Cells[totalRowIndex, 60].Value}, Formula = {ws.Cells[totalRowIndex, 60].Formula}");
                    }
                    int left = mode is TemplateMode.Even ? int.Parse(ws.Cells[totalRowIndex, 31].Value?.ToString()) : int.Parse(ws.Cells[totalRowIndex, 32].Value?.ToString());
                    int right = mode is TemplateMode.Even ? int.Parse(ws.Cells[totalRowIndex, 58].Value?.ToString()) : int.Parse(ws.Cells[totalRowIndex, 60].Value?.ToString());
                    text = text.Replace("#EB:TOTAL", $"{left + right}");
                    Debug.WriteLine($"text: {text}");
                    shape.Text.Paragraphs.Clear();
                    Gem.Drawing.TextRun  run = shape.Text.Paragraphs.Add().Elements.AddRun(text);
                    run.Format.Font = "Arial";
                    run.Format.Size = Length.From(13, LengthUnit.Point);
                    run.Format.Fill.SetSolid(Gem.Drawing.DrawingColor.FromRgb(255, 0, 0));
                }
            }

            wb.Save(savePath);
        }

        private void EnsureOutputWorkbookReadyForUpdate()
        {
            // 현재 Project에서 Excel Import가 한 번이라도 완료되었다면
            // EB Template 폴더에 Output 문서가 존재해야 한다.
            if (_settingService.FindExternalDocument(outputFileName) == null)
            {
                throw new InvalidOperationException(
                    $"EB에 '{outputFileName}' 파일이 없습니다.\r\n" +
                    "Excel Import를 먼저 실행해주세요.");
            }

            // 로컬 임시 파일이 삭제된 경우에는 EB의 Output을 다시 내려받는다.
            if (!File.Exists(outputPath) && !LoadOutputFromEB())
            {
                throw new InvalidOperationException(
                    $"EB의 '{outputFileName}' 파일을 로컬로 내려받지 못했습니다.");
            }

            ExcelFile workbook;
            try
            {
                workbook = ExcelFile.Load(outputPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"'{outputFileName}' 파일을 읽을 수 없습니다.",
                    ex);
            }

            if (workbook.Worksheets.Count == 0)
                throw new InvalidOperationException("Output Excel에 워크시트가 없습니다.");

            // 최초 Import에서 기준 시트를 무조건 Rev0으로 생성하므로,
            // Rev0가 없다면 이 Project의 정상적인 Revision Output으로 볼 수 없다.
            if (_updateService.FindWorksheetByName(workbook, DefaultRevisionName) == null)
            {
                throw new InvalidOperationException(
                    $"'{outputFileName}'에 기준 Revision 시트 '{DefaultRevisionName}'이 없습니다.\r\n" +
                    "Excel Import를 다시 실행해 정상적인 Output을 생성해주세요.");
            }

            ExcelWorksheet activeWorksheet =
                workbook.Worksheets.ActiveWorksheet ?? workbook.Worksheets[0];
            if (!_settingService.TryNormalizeRevisionName(activeWorksheet.Name, out _))
            {
                throw new InvalidOperationException(
                    $"현재 활성 시트 '{activeWorksheet.Name}'은 Revision 시트가 아닙니다.\r\n" +
                    "Rev0 이상의 Revision 시트를 활성화한 후 Update해주세요.");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SaveOutputToEB();
                MessageBox.Show(
                    $"저장 완료되었습니다.\r\n{outputPath}",
                    "완료",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Output 저장 실패: {ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btn_update_data_Click(object sender, EventArgs e)
        {
            try
            {
                EnsureOutputWorkbookReadyForUpdate();
                string updateWorksheetName = _updateService.GetActiveRevisionWorksheetName(outputPath);

                // 이번 Update 시트에서 직접 수정한 고유 Reefer만 센다.
                // GangBox 재채번/연결 변경 수로 대체하지 않으며, 원복한 셀은 제외한다.
                int directlyEditedReeferCount = _updateService.CountDirectlyEditedReefers(
                    _updateService.FindWorksheetByName(ExcelFile.Load(outputPath), updateWorksheetName),
                    _pendingOriginalReeferColors,
                    cell => _excelService.NormalizeHtmlColor(_excelService.GetCellColorHex(cell)));

                // Update는 기존 메모리의 변경분만 반영하지 않고,
                // 사용자가 저장한 최종 Excel을 정답으로 다시 읽어 전체 EB 관계를 생성/동기화한다.
                if (!RebuildFromFile(outputPath, false, true, updateWorksheetName))
                    return;

                SaveToEB_UpLoad();

                int reeferCount = raw_List.Count(x => x != null && x.Reefer);
                int rdbCount = raw_List
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.RDB_EB_ID))
                    .Select(x => x.RDB_EB_ID)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count();
                int gangBoxCount = raw_List
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.GangBox_EB_ID))
                    .Select(x => x.GangBox_EB_ID)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count();

                string revisionResult = $"{updateWorksheetName} 활성 Revision 기준 EB Update";

                foreach (string cellKey in _pendingOriginalReeferColors.Keys
                    .Where(key => key.StartsWith(updateWorksheetName + "|", StringComparison.OrdinalIgnoreCase))
                    .ToList())
                {
                    _pendingOriginalReeferColors.Remove(cellKey);
                }

                MessageBox.Show(
                    $"최종 Excel 기준 EB 객체 생성/동기화 완료\r\n" +
                    $"Revision: {revisionResult}\r\n" +
                    $"RDB: {rdbCount}\r\n" +
                    $"GangBox: {gangBoxCount}\r\n" +
                    $"Reefer: {reeferCount}\r\n" +
                    $"변경 건수: {directlyEditedReeferCount}",
                    "완료",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                    
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "최종 Excel 기준 EB 객체 동기화 실패: " + ex.Message,
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        


        #region event handlers
        private void btn_save_setting_Click(object sender, EventArgs e)
        {
            try
            {
                string revisionName = _settingService.NormalizeRevisionName(tb_revision.Text);
                tb_revision.Text = revisionName;
                bool revisionChanged = !string.Equals(revisionName, _savedRevisionName, StringComparison.OrdinalIgnoreCase);

                // 최초 Import 전에는 Output Excel이 없으므로 Setting만 저장한다.
                // Output이 존재하는 기존 Project에서 Rev1 이상일 때만 Revision 시트를 생성하거나 갱신한다.
                bool outputWorkbookExists = _settingService.FindExternalDocument(outputFileName) != null;
                bool revisionWorksheetRequired = !string.Equals(revisionName, DefaultRevisionName, StringComparison.OrdinalIgnoreCase);
                if (outputWorkbookExists && revisionWorksheetRequired)
                {
                    EnsureOutputWorkbookReadyForUpdate();
                    ExcelFile revisionWorkbook = ExcelFile.Load(outputPath);
                    bool revisionWorksheetExists = _updateService.FindWorksheetByName(revisionWorkbook, revisionName) != null;

                    if (revisionChanged || !revisionWorksheetExists)
                    {
                        _updateService.PrepareRevisionWorksheet(outputPath, revisionName, out _);

                        // GemBox로 변경한 Output을 EB에 반영한 후,
                        // SpreadsheetControl에도 다시 로드하여 새 Revision 탭을 표시한다.
                        SaveToEB_UpLoad();
                        savePath = outputPath;
                        sheet8grid_tfn = true;
                        spreadsheetControl1.LoadDocument(outputPath);
                        spreadsheetControl1.BringToFront();
                    }
                }

                SaveSettingsToEB();
                panel_setting.Visible = !panel_setting.Visible;
                panel_setting.BringToFront();
                MessageBox.Show("Setting이 JSON으로 저장되어 EB에 업로드되었습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Setting 저장 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_default_setting_Click(object sender, EventArgs e)
        {
            try
            {
                ApplyDefaultSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Default Setting 적용 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void num_Max_AF_ValueChanged(object sender, EventArgs e)
        {
            Max_AF = (int)num_Max_AF.Value;
            Debug.WriteLine($"Max_AF updated: {Max_AF}");
        }

        private void num_Max_rdb_reefer_cnt_ValueChanged(object sender, EventArgs e)
        {
            Max_rdb_reefer_cnt = (int)num_Max_rdb_reefer_cnt.Value;            
            Debug.WriteLine($"Max_rdb_reefer_cnt updated: {Max_rdb_reefer_cnt}");
        }

        private void num_Max_gangbox_cnt_ValueChanged(object sender, EventArgs e)
        {
            Max_gang_cnt = (int)num_Max_gang_cnt.Value;           
            Debug.WriteLine($"Max_gangbox_cnt updated: {Max_gang_cnt}");
        }

        private void num_distboard_basis_ValueChanged(object sender, EventArgs e)
        {
            distboard_basis = (double)num_distboard_basis.Value;
            num_RDB.Value = (decimal)num_Reefer_basis.Value * ((decimal)num_distboard_basis.Value / 100);
            Debug.WriteLine($"DistBoard_basis updated: {distboard_basis}");
            UpdateBasisCalculationLabels();
        }
              
        private void num_transformer_basis_ValueChanged(object sender, EventArgs e)
        {
            transformer_basis = (double)num_transformer_basis.Value;
            num_TR.Value = (decimal)num_Reefer_basis.Value * ((decimal)num_transformer_basis.Value / 100);
            Debug.WriteLine($"Transformer_basis updated: {transformer_basis}");
            UpdateBasisCalculationLabels();
        }
             
        private void num_Reefer_basis_ValueChanged(object sender, EventArgs e)
        {
            Reefer_basis = (double)num_Reefer_basis.Value;
            num_RDB.Value = (decimal)num_Reefer_basis.Value * ((decimal)num_distboard_basis.Value / 100);
            num_SRD.Value = (decimal)num_Reefer_basis.Value * ((decimal)num_srd_basis.Value / 100);
            num_TR.Value = (decimal)num_Reefer_basis.Value * ((decimal)num_transformer_basis.Value / 100);
            Debug.WriteLine($"Reefer_basis updated: {Reefer_basis}");
            UpdateBasisCalculationLabels();
        }

        private void num_srd_basis_ValueChanged(object sender, EventArgs e)
        {
            srd_basis = (double)num_srd_basis.Value;
            num_SRD.Value = (decimal)num_Reefer_basis.Value * ((decimal)num_srd_basis.Value / 100);
            Debug.WriteLine($"SRD_basis updated: {srd_basis}");
            UpdateBasisCalculationLabels();
        }

        private void num_generator_basis_ValueChanged(object sender, EventArgs e)
        {
            generator_basis = (double)num_generator_basis.Value;
            Debug.WriteLine($"Generator_basis updated: {generator_basis}");
        }
               
        private void cb_deck_hold_type_SelectedIndexChanged(object sender, EventArgs e)
        {
            deck_hold_type = cb_deck_hold_type.SelectedIndex; // 0: ON DECK / IN HOLD,  1: ON DECK + IN HOLD
            Debug.WriteLine($"Deck/Hold type updated: {(deck_hold_type == 0 ? "ON DECK / IN HOLD" : "ON DECK + IN HOLD")}");
        }

        private void chk_deck_first_CheckedChanged(object sender, EventArgs e)
        {
            _processDeckBeforeHold = chk_deck_first.Checked;
            Debug.WriteLine($"Deck first processing updated: {_processDeckBeforeHold}");
        }
        
        #endregion


        private void SortPriorityPanel_PriorityChanged(object sender, EventArgs e)
        {
            // Setting의 정렬값 변경은 기존 Import 결과에 반영하지 않는다.
            // 변경된 값은 사용자가 다시 Import할 때 적용된다.
            //RebuildGroupedList();
        }
               
        /// <summary>
        /// IGrouping 으로부터 Tier 정렬 + Row 정렬이 반영된 TierGroupDto 리스트를 생성한다.
        /// </summary>
        private bool sheet8grid_tfn = false;
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (!sheet8grid_tfn)
            {
                spreadsheetControl1.BringToFront();
                sheet8grid_tfn = true;
            }
            else
            {
                gridControl1.BringToFront();
                sheet8grid_tfn = false;
            }
        }

        private IAucObject First_Create_EB_Reefer(AucObjectKind kind = AucObjectKind.aucObjDevice, AucType type = AucType.aucElectricalComponent)
        {
            try
            {
                IAucObject eb_reefer = null;
                IAucObjects temp_eb_reefer = Reefer_folder.FindObjectsEx(ebHelper.Get_filter("Origin_Reefer"), AucVbSearchType.aucSearchFlat, true);
                if (temp_eb_reefer.Count > 0)
                {
                    eb_reefer = temp_eb_reefer.Item(1);
                    return eb_reefer;
                }
                else
                {
                    //eb_reefer = Parent.Children.AddNew(kind, type);
                    IAucObject catalog = ebapp.Folders.Catalogs
                        .FindObjectsEx(ebHelper.Get_filter("SHI Shipbuilding Catalog"), AucVbSearchType.aucSearchFlat, true).Item(1)
                        .FindObjectsEx(ebHelper.Get_filter("Materials"), AucVbSearchType.aucSearchFlat, true).Item(1)
                        .FindObjectsEx(ebHelper.Get_filter("ELEC_EQUIP"), AucVbSearchType.aucSearchFlat, true).Item(1)
                        .FindObjectsEx(ebHelper.Get_filter("REF"), AucVbSearchType.aucSearchFlat, true).Item(1).Children.Item(1);

                    eb_reefer = ebHelper.TryCopyTo(Reefer_folder, catalog);
                    if (eb_reefer == null)
                    {
                        Debug.WriteLine($"{Parent.Name}에 Children.AddNew {"Origin_Reefer"} 생성 중 에러 발생\r");
                        return null;
                    }
                    else
                    {                     
                        ebHelper.SetAttributes(eb_reefer, new[]
                        {
                            new KeyValuePair<int, string>(5, "Origin_Reefer"),
                            new KeyValuePair<int, string>(25, "Reefer"),
                            new KeyValuePair<int, string>(100236, "10.3"),
                            new KeyValuePair<int, string>(102193, ""), //Bay
                            new KeyValuePair<int, string>(102194, ""), //Tier
                            new KeyValuePair<int, string>(102195, ""), //Row
                            new KeyValuePair<int, string>(102196, "")  //C/H no
                        });
                       
                        return eb_reefer;
                    }
                        

                    
                }
                    
            }
            catch
            {
                return null;
            }            
              
           
        }

        private IAucObject Create_EB_Reefer(ReeferItem reefer,int position,AucObjectKind kind = AucObjectKind.aucObjDevice,
            AucType type = AucType.aucElectricalComponent,bool updateExistingAttributes = true)
        {
            IAucObject eb_reefer = null;
            string reefer_name = Get_Reefer_Name(reefer);
            bool isNewReefer = false;

            TryAddExistingObjectToMap(Reefer_folder, reefer_name, Reefer_dict);

            if (Reefer_dict.ContainsKey(reefer_name))
            {
                eb_reefer = GetCachedObjectByDesignation(reefer_name, Reefer_dict, _reeferObjectCache);
                if (eb_reefer == null)
                {
                    Debug.WriteLine($"EB_reefer Find Fail\rName: <{reefer_name}>");
                    return null;
                }

                if (!string.Equals(
                    eb_reefer.Parent?.ID,
                    Reefer_folder.ID,
                    StringComparison.OrdinalIgnoreCase))
                {
                    Debug.WriteLine($"{eb_reefer.Parent.ID}");
                    Debug.WriteLine($"{Reefer_folder.ID}");
                    bool result = ebHelper.TryMoveTo(Reefer_folder, eb_reefer);
                    if (!result)
                        Debug.WriteLine($"{eb_reefer.Name}->{Reefer_folder.Name} MoveTo Fail");
                }

            }
            else
            {                              
                eb_reefer = ebHelper.TryCopyTo(Reefer_folder, Origin_eb_reefer);
                if (eb_reefer == null)
                {
                    Debug.WriteLine($"EB_reefer Find Fail");
                    return null;
                }

                isNewReefer = true;
                // 중복생성 컨트롤
                if (!Reefer_dict.ContainsKey(reefer_name))
                    Reefer_dict[reefer_name] = eb_reefer.ID;
            }

            CacheEbObject(eb_reefer, reefer_name, _reeferObjectCache);

            if (!isNewReefer && !updateExistingAttributes)
                return eb_reefer;

            // attribute
            var reeferAttributes = new List<KeyValuePair<int, string>>
            {
                new KeyValuePair<int, string>(102193, reefer.BayNo.ToString("00")), //Bay
                new KeyValuePair<int, string>(102194, reefer.TierNo.ToString()), //Tier
                new KeyValuePair<int, string>(102195, reefer.RowNo.ToString("00")), //Row
                new KeyValuePair<int, string>(102196, reefer.CHNo ?? string.Empty), //C/H no
                new KeyValuePair<int, string>(261, position.ToString()) // Position in GangBox
            };

            if (isNewReefer)
                reeferAttributes.Insert(0, new KeyValuePair<int, string>(5, reefer_name));

            ebHelper.SetAttributes(eb_reefer, reeferAttributes);

            return eb_reefer;
        }

        private IAucObject Create_EB_RDB(string rdbDesignation, AucObjectKind kind = AucObjectKind.aucObjDevice, AucType type = AucType.aucElectricalComponent)
        {
            if (RDB_folder == null)
                throw new InvalidOperationException("RDB_folder 값이 null 입니다.");

            if (string.IsNullOrWhiteSpace(rdbDesignation))
                throw new ArgumentException("rdbDesignation 값이 비어 있습니다.", nameof(rdbDesignation));

            TryAddExistingObjectToMap(RDB_folder, rdbDesignation, RDB_dict);
            
            IAucObject eb_rdb = null;           
            if (RDB_dict.ContainsKey(rdbDesignation))
            {
                eb_rdb = GetCachedObjectByDesignation(rdbDesignation, RDB_dict, _rdbObjectCache);
                if (eb_rdb == null)
                {
                    Debug.WriteLine($"EB_rdb Find Fail");
                    return null;
                }

                if (!string.Equals(
                    eb_rdb.Parent?.ID,
                    RDB_folder.ID,
                    StringComparison.OrdinalIgnoreCase))
                {
                    Debug.WriteLine($"{eb_rdb.Parent.ID}");
                    Debug.WriteLine($"{RDB_folder.ID}");
                    bool result = ebHelper.TryMoveTo(RDB_folder, eb_rdb);
                    if (!result)
                        Debug.WriteLine($"{eb_rdb.Name}->{RDB_folder.Name} MoveTo Fail");                    
                }
            }
            else
            {
                //eb_rdb = Reefer_Container_folder.Children.AddNew(kind, type);
                if (RDB_catalog == null)
                    RDB_catalog = ebapp.Folders.Catalogs
                        .FindObjectsEx(ebHelper.Get_filter("SHI Shipbuilding Catalog"), AucVbSearchType.aucSearchFlat, true).Item(1)
                        .FindObjectsEx(ebHelper.Get_filter("Materials"), AucVbSearchType.aucSearchFlat, true).Item(1)
                        .FindObjectsEx(ebHelper.Get_filter("ELEC_EQUIP"), AucVbSearchType.aucSearchFlat, true).Item(1)
                        .FindObjectsEx(ebHelper.Get_filter("DIST_SYS_CLS"), AucVbSearchType.aucSearchFlat, true).Item(1).Children.Item(1);

                eb_rdb = ebHelper.TryCopyTo(RDB_folder, RDB_catalog);
                if (eb_rdb == null)
                {
                    Debug.WriteLine($"{Parent.Name}에 Children.AddNew {"eb_rdb"} 생성 중 에러 발생\r");
                    return null;
                }
                else
                {
                    ebHelper.SetAttributes(eb_rdb, new[]
                    {
                        new KeyValuePair<int, string>(5, rdbDesignation),
                        new KeyValuePair<int, string>(25, "RDB"),
                        new KeyValuePair<int, string>(100236, "7.2"),
                        new KeyValuePair<int, string>(100237, "440V"),
                        new KeyValuePair<int, string>(100239, "60Hz"),
                        new KeyValuePair<int, string>(100240, "3")
                    });


                    // 중복생성 컨트롤
                    if (!RDB_dict.ContainsKey(rdbDesignation))
                        RDB_dict[rdbDesignation] = eb_rdb.ID;
                }              
            }

            CacheEbObject(eb_rdb, rdbDesignation, _rdbObjectCache);

            return eb_rdb;
        }

        private IAucObject Create_EB_GangBox(string gangBoxDesignation, AucObjectKind kind = AucObjectKind.aucObjDevice, AucType type = AucType.aucElectricalComponent)
        {
            TryAddExistingObjectToMap(Gangbox_folder, gangBoxDesignation, Gangbox_dict);

            if (Gangbox_folder == null)
                throw new ArgumentNullException(nameof(Gangbox_folder));

            if (string.IsNullOrWhiteSpace(gangBoxDesignation))
                throw new ArgumentException("gangBoxDesignation 값이 비어 있습니다.", nameof(gangBoxDesignation));


            IAucObject eb_gangbox = null;           
            if (Gangbox_dict.ContainsKey(gangBoxDesignation))
            {
                eb_gangbox = GetCachedObjectByDesignation(gangBoxDesignation, Gangbox_dict, _gangboxObjectCache);
                if (eb_gangbox == null)
                {
                    Debug.WriteLine($"EB_gangbox Find Fail");
                    return null;
                }

                if (!string.Equals(
                    eb_gangbox.Parent?.ID,
                    Gangbox_folder.ID,
                    StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        Debug.WriteLine($"{eb_gangbox.Parent.ID} / {eb_gangbox.Parent.Name} / {eb_gangbox.Parent.Parent.Name} / {eb_gangbox.Parent.Parent.Name} / {eb_gangbox.Parent.Parent.Parent.Name} / {eb_gangbox.Parent.Parent.Parent.Parent.Name}");
                        Debug.WriteLine($"{Gangbox_folder.ID} / {Gangbox_folder.Name} / {Gangbox_folder.Parent.Name} / {Gangbox_folder.Parent.Parent.Name} / {Gangbox_folder.Parent.Parent.Parent.Name} / {Gangbox_folder.Parent.Parent.Parent.Parent.Name}");
                        bool result = ebHelper.TryMoveTo(Gangbox_folder, eb_gangbox);
                        if (!result)
                            Debug.WriteLine($"{eb_gangbox.Parent.Name}->{Gangbox_folder.Name} MoveTo Fail");
                    }
                    catch
                    {

                    }

                }

            }
            else           
            {
                //eb_gangbox = rdbObj.Children.AddNew(kind, type);
                if (Gangbox_catalog == null)
                    Gangbox_catalog = ebapp.Folders.Catalogs
                        .FindObjectsEx(ebHelper.Get_filter("SHI Shipbuilding Catalog"), AucVbSearchType.aucSearchFlat, true).Item(1)
                        .FindObjectsEx(ebHelper.Get_filter("Materials"), AucVbSearchType.aucSearchFlat, true).Item(1)
                        .FindObjectsEx(ebHelper.Get_filter("ELEC_EQUIP"), AucVbSearchType.aucSearchFlat, true).Item(1)
                        .FindObjectsEx(ebHelper.Get_filter("REF_SCK_CLS"), AucVbSearchType.aucSearchFlat, true).Item(1).Children.Item(1);

                eb_gangbox = ebHelper.TryCopyTo(Gangbox_folder, Gangbox_catalog);
                if (eb_gangbox == null)
                {
                    Debug.WriteLine($"{Parent.Name}에 Children.AddNew {"eb_rdb"} 생성 중 에러 발생\r");
                    return null;
                }
                else               
                {
                    ebHelper.SetAttributes(eb_gangbox, new[]
                    {
                        new KeyValuePair<int, string>(5, gangBoxDesignation),
                        new KeyValuePair<int, string>(25, "Gangbox"),
                        new KeyValuePair<int, string>(100236, "9.27"),
                        new KeyValuePair<int, string>(100237, "440V"),
                        new KeyValuePair<int, string>(100239, "60Hz"),
                        new KeyValuePair<int, string>(100240, "3")
                    });


                    // 중복생성 컨트롤
                    if (!Gangbox_dict.ContainsKey(gangBoxDesignation))
                        Gangbox_dict[gangBoxDesignation] = eb_gangbox.ID;

                }           
            }

            CacheEbObject(eb_gangbox, gangBoxDesignation, _gangboxObjectCache);

            return eb_gangbox;
        }

        private void ClearDetachedReeferPositions(string gangBoxName,IEnumerable<IAucObject> detachedReefers)
        {
            if (detachedReefers == null)
                return;

            foreach (IAucObject detachedReefer in detachedReefers.Where(x => x != null))
            {
                bool positionCleared = ebHelper.SetAttributes(
                    detachedReefer,
                    new[] { new KeyValuePair<int, string>(261, string.Empty) });

                if (!positionCleared)
                {
                    throw new InvalidOperationException(
                        $"GangBox에서 제외된 Reefer의 Position 초기화에 실패했습니다. " +
                        $"GangBox={gangBoxName}, Reefer={detachedReefer.Name}");
                }

                Debug.WriteLine(
                    $"[REEFER POSITION CLEAR] GangBox={gangBoxName}, " +
                    $"Reefer={detachedReefer.Name}, Attribute=261");
            }
        }

        private Dictionary<string, string> GetAssociatedReeferIds(IAucObject gangbox)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (gangbox == null)
                return result;

            foreach (IAucAssociation association in gangbox.SourceAssociations.Filter((AucRole)119))
            {
                IAucObject reefer = association.RelatedObject;
                if (reefer == null || string.IsNullOrWhiteSpace(reefer.Name))
                    continue;

                if (!result.ContainsKey(reefer.Name))
                    result[reefer.Name] = reefer.ID;
            }

            return result;
        }

        private void UpdateRdbAttributes(IEnumerable<ReeferItem> sourceItems)
        {
            if (sourceItems == null)
                return;

            var items = sourceItems
                .Where(x => x != null && x.Reefer && !string.IsNullOrWhiteSpace(x.RDB_ID))
                .ToList();

            foreach (var rdbGroup in items.GroupBy(x => Get_RDB_Name(x.RDB_ID)))
            {
                string rdbDesignation = rdbGroup.Key;
                if (string.IsNullOrWhiteSpace(rdbDesignation))
                    continue;

                IAucObject rdb = GetCachedObjectByDesignation(
                    rdbDesignation,
                    RDB_dict,
                    _rdbObjectCache);

                if (rdb == null)
                    continue;

                var distinctReefers = rdbGroup
                    .GroupBy(Get_Reefer_Name)
                    .Select(x => x.First())
                    .ToList();

                int rdbNo = int.TryParse(_excelService.GetRdbDisplayNumber(rdbDesignation), out int parsedRdbNo)
                    ? parsedRdbNo
                    : 0;

                Debug.WriteLine($"[RDB ATTR START] {rdbDesignation}");

                double ratedPowerKw = Reefer_basis * (srd_basis / 100.0) * distinctReefers.Count;
                double ratedCurrentA = CalculateRatedCurrent(Reefer_basis) * (srd_basis / 100.0) * distinctReefers.Count;

                ebHelper.SetAttributes(rdb, new[]
                {
                    new KeyValuePair<int, string>(100035, $"REEFER DIST. BOARD NO.{rdbNo}(RD{rdbNo})"),
                    new KeyValuePair<int, string>(100236, ratedPowerKw.ToString("0.0")),
                    new KeyValuePair<int, string>(100233, ratedCurrentA.ToString("0.0")),
                    new KeyValuePair<int, string>(100237, "440V"),
                    new KeyValuePair<int, string>(100239, "60Hz"),
                    new KeyValuePair<int, string>(100240, "3")
                });

                Debug.WriteLine($"[RDB ATTR END] {rdbDesignation}");
            }
        }

        private void UpdateGangBoxAttributes(IEnumerable<ReeferItem> sourceItems)
        {
            if (sourceItems == null)
                return;

            var gangBoxGroups = sourceItems
                .Where(x => x != null && x.Reefer && !string.IsNullOrWhiteSpace(x.RDB_ID))
                .Select(item =>
                {
                    int gangBoxNo = item.GangBox_No;
                    string rdbDesignation = Get_RDB_Name(item.RDB_ID);

                    return new
                    {
                        Item = item,
                        RdbDesignation = rdbDesignation,
                        GangBoxNo = gangBoxNo,
                        GangBoxDesignation = gangBoxNo > 0
                            ? Get_Gangbox_Name(gangBoxNo.ToString(), rdbDesignation)
                            : string.Empty
                    };
                })
                .Where(x => x.GangBoxNo > 0 && !string.IsNullOrWhiteSpace(x.GangBoxDesignation))
                .GroupBy(x => new { x.RdbDesignation, x.GangBoxDesignation });

            foreach (var gangBoxGroup in gangBoxGroups)
            {
                IAucObject gangbox = GetCachedObjectByDesignation(
                    gangBoxGroup.Key.GangBoxDesignation,
                    Gangbox_dict,
                    _gangboxObjectCache);

                if (gangbox == null)
                    continue;

                var distinctReefers = gangBoxGroup
                    .Select(x => x.Item)
                    .GroupBy(Get_Reefer_Name)
                    .Select(x => x.First())
                    .OrderBy(x => x.BayNo)
                    .ThenBy(x => x.RowNo)
                    .ThenByDescending(x => x.TierNo)
                    .ThenBy(x => x.CHNo ?? string.Empty)
                    .Take(10)
                    .ToList();

                int gangBoxNo = gangBoxGroup.First().GangBoxNo;
                int socketCount = gangBoxGroup
                    .Select(x => x.Item)
                    .GroupBy(Get_Reefer_Name)
                    .Count();

                double ratedPowerKw = Reefer_basis * (distboard_basis / 100.0) * socketCount;
                double ratedCurrentA = CalculateRatedCurrent(Reefer_basis) * (distboard_basis / 100.0) * socketCount;

                Debug.WriteLine(
                    $"[GANGBOX ATTR START] {gangBoxGroup.Key.GangBoxDesignation}, " +
                    $"socketCount={socketCount}");

                var gangBoxAttributes = new List<KeyValuePair<int, string>>
                {
                    new KeyValuePair<int, string>(100035, $"REEFER CONTAINER SOCKET NO.{gangBoxNo}({socketCount}G)"),
                    new KeyValuePair<int, string>(102196, gangBoxGroup.First().Item.CHNo),
                    new KeyValuePair<int, string>(101125, gangBoxGroup.First().Item.CHNo),
                    new KeyValuePair<int, string>(100236, ratedPowerKw.ToString("0.0")),
                    new KeyValuePair<int, string>(100233, ratedCurrentA.ToString("0.0")),
                    new KeyValuePair<int, string>(102192, socketCount.ToString()),
                    new KeyValuePair<int, string>(100237, "440V"),
                    new KeyValuePair<int, string>(100239, "60Hz"),
                    new KeyValuePair<int, string>(100240, "3")
                };

                ebHelper.SetAttributes(gangbox, gangBoxAttributes);

                Debug.WriteLine($"[GANGBOX ATTR END] {gangBoxGroup.Key.GangBoxDesignation}");
            }
        }

        private void btn_rebuild_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx;*.xls";
                ofd.Title = "완성된 Reefer Container Excel 선택";

                if (ofd.ShowDialog() != DialogResult.OK)
                    return;

                RebuildFromFile(ofd.FileName, true);
            }
        }

        private bool RebuildFromFile(string filePath,bool showCompletionMessage,bool synchronizeEbObjects = false,string worksheetName = null)
        {
            bool waitFormShown = false;

            try
            {
                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                    return false;

                ExcelFile import_excel = Gem.ExcelFile.Load(filePath);
                if (import_excel.Worksheets.Count == 0)
                    throw new InvalidOperationException("워크시트가 없습니다.");

                ExcelWorksheet selectedWorksheet = string.IsNullOrWhiteSpace(worksheetName)
                    ? import_excel.Worksheets.ActiveWorksheet ?? import_excel.Worksheets[0]
                    : _updateService.FindWorksheetByName(import_excel, worksheetName);

                if (selectedWorksheet == null)
                    throw new InvalidOperationException($"'{worksheetName}' Revision 시트를 찾을 수 없습니다.");

                var persistedSyncStates = new List<EbSyncReeferState>();
                // Navigation은 EB 객체 생성 여부와 무관하게 Excel의 숨김 OID 맵을 사용한다.
                bool hasPersistedSyncState =
                    _updateService.TryReadEbSyncState(
                        import_excel,
                        EbSyncStateWorksheetName,
                        EbSyncStateMarker,
                        out persistedSyncStates);
                string ebOidValidationSessionKey =
                    _updateService.GetOidValidationSessionKey(Project, filePath, selectedWorksheet.Name);

                SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);
                waitFormShown = true;

                if (hasPersistedSyncState && !string.Equals( _validatedEbOidSessionKey, ebOidValidationSessionKey, StringComparison.OrdinalIgnoreCase))
                {
                    SplashScreenManager.Default.SetWaitFormDescription("Loading... EB OID Validation...");
                    _updateService.RefreshPersistedObjectIds(
                        persistedSyncStates,
                        Project,
                        id => ebapp?.Utils.GetSnglObjectByID(id),
                        () => Create_Dict_All(false),
                        Reefer_dict,
                        Gangbox_dict,
                        RDB_dict);
                }
                else if (hasPersistedSyncState)
                {
                    Debug.WriteLine($"[EB OID VALIDATION SKIP] 이미 검증된 프로그램 세션: {ebOidValidationSessionKey}");
                }

                _workbookReeferEbIds.Clear();
                if (hasPersistedSyncState)
                {
                    foreach (EbSyncReeferState state in persistedSyncStates)
                    {
                        if (state == null ||
                            string.IsNullOrWhiteSpace(state.ReeferDesignation) ||
                            string.IsNullOrWhiteSpace(state.ReeferEbId))
                        {
                            continue;
                        }

                        _workbookReeferEbIds[state.ReeferDesignation] = state.ReeferEbId;
                    }
                }

                DateTime start_time = DateTime.Now;
                Debug.WriteLine($"[REBUILD START] {start_time:HH:mm:ss.fff}");
                Stopwatch st = Stopwatch.StartNew();

                    raw_List.Clear();
                    grouped_List.Clear();

                    SplashScreenManager.Default.SetWaitFormDescription("Loading... Completed Excel Read...");
                    ExcelWorksheet ws = selectedWorksheet;

                    int evenStartCount = _excelService.FindBayStartRowsFromTierColumn(ws, 25).Count + _excelService.FindBayStartRowsFromTierColumn(ws, 54).Count;
                    int oddStartCount = _excelService.FindBayStartRowsFromTierColumn(ws, 26).Count + _excelService.FindBayStartRowsFromTierColumn(ws, 56).Count;

                    if (evenStartCount == 0 && oddStartCount == 0)
                        throw new InvalidOperationException("완성된 짝수/홀수 템플릿 구조를 찾지 못했습니다. Tier 기준 셀을 확인하세요.");

                    TemplateMode mode = oddStartCount > evenStartCount
                        ? TemplateMode.Odd
                        : TemplateMode.Even;

                    var usedRange = ws.GetUsedCellRange(true);
                    int usedLastRow = usedRange != null ? usedRange.LastRowIndex : ws.Rows.Count - 1;

                    int importedCount = 0;

                    Action<string, IReadOnlyDictionary<int, int>, int, int, int> readPlace =
                        (place, rowNoToColumnMap, bayCol, chCol, tierCol) =>
                        {
                            List<int> bayStartRows = _excelService.FindBayStartRowsFromTierColumn(ws, tierCol)
                                .OrderBy(x => x)
                                .ToList();
                            var allowedTierNumbers = new HashSet<int>(_excelService.GetAllowedTierNumbers(place));

                            if (bayStartRows.Count == 0)
                                return;

                            for (int slotIndex = 0; slotIndex < bayStartRows.Count; slotIndex++)
                            {
                                int bayStartRow = bayStartRows[slotIndex];

                                int nextStartRow = slotIndex + 1 < bayStartRows.Count
                                    ? bayStartRows[slotIndex + 1]
                                    : usedLastRow + 1;

                                string bayText = GemBoxHelper.GetCellTextResolveMerge(ws, bayStartRow, bayCol);
                                string chNo = GemBoxHelper.GetCellTextResolveMerge(ws, bayStartRow, chCol);

                                int bayNo;
                                if (!int.TryParse(_excelService.ExtractRdbNumber(bayText), out bayNo))
                                    continue;

                                if (string.IsNullOrWhiteSpace(chNo))
                                    continue;

                                chNo = chNo.Trim();

                                for (int row = bayStartRow; row < nextStartRow; row++)
                                {
                                    string tierText = GemBoxHelper.GetCellTextResolveMerge(ws, row, tierCol);

                                    int tierNo;
                                    if (!int.TryParse(_excelService.ExtractRdbNumber(tierText), out tierNo))
                                        continue;

                                    if (!allowedTierNumbers.Contains(tierNo))
                                        continue;

                                    foreach (var pair in rowNoToColumnMap)
                                    {
                                        int rowNo = pair.Key;
                                        int colIndex = pair.Value;

                                        var cell = ws.Cells[row, colIndex];

                                        string cellText = GemBoxHelper.GetCellText(ws, row, colIndex);
                                        string rdbNoText = _excelService.ExtractRdbNumber(cellText);

                                        string cellColor = _excelService.NormalizeHtmlColor(_excelService.GetCellColorHex(cell));

                                        // 완성 Excel 기준:
                                        // 숫자 = RDB
                                        // 색상 = GangBox
                                        if (string.IsNullOrWhiteSpace(rdbNoText) &&
                                            string.IsNullOrWhiteSpace(cellColor))
                                        {
                                            continue;
                                        }

                                        if (string.IsNullOrWhiteSpace(cellColor))
                                            cellColor = "#C0C0C0";

                                        int rdbNo;
                                        string rdbId = "";
                                        string rdbColor = null;

                                        if (int.TryParse(rdbNoText, out rdbNo))
                                        {
                                            rdbId = Get_RDB_Name(rdbNo.ToString()); // 1 -> EP46-201
                                            rdbColor = GroupingService.GetRdbColor(rdbNo, _rdbColorPalette);
                                        }

                                        raw_List.Add(new ReeferItem
                                        {
                                            // 아래에서 R1~RMax 재채번
                                            TagNo = "",

                                            Reefer = true,
                                            BayNo = bayNo,
                                            TierNo = tierNo,
                                            RowNo = rowNo,
                                            CHNo = chNo,
                                            Place = place,

                                            // Excel 색상 = GangBox
                                            GangBox_ID = cellColor,
                                            GangBox_Color = cellColor,

                                            // Excel 숫자 = RDB
                                            RDB_ID = rdbId,
                                            RDB_Color = rdbColor,

                                            // Navigate용 위치
                                            X = row,
                                            Y = colIndex
                                        });

                                        importedCount++;
                                    }
                                }
                            }
                        };

                    if (mode == TemplateMode.Even)
                    {
                        readPlace("ON DECK", ReeferTemplateLayout.RowNoToColumnMapOnDeck, 2, 24, 25);
                        readPlace("IN HOLD", ReeferTemplateLayout.RowNoToColumnMapInHold, 33, 53, 54);
                    }
                    else if (mode == TemplateMode.Odd)
                    {
                        readPlace("ON DECK", ReeferTemplateLayout.RowNoToColumnMapOnDeckOdd, 2, 25, 26);
                        readPlace("IN HOLD", ReeferTemplateLayout.RowNoToColumnMapInHoldOdd, 34, 55, 56);
                    }
                    else
                    {
                        throw new InvalidOperationException("템플릿 모드를 판별할 수 없습니다.");
                    }

                    if (raw_List.Count == 0)
                        throw new InvalidOperationException("완성된 Excel에서 Reefer 셀을 찾지 못했습니다.");

                    SplashScreenManager.Default.SetWaitFormDescription("Loading... TagNo Rebuild...");
                    
                    var orderedItems = raw_List
                        .Where(x => x != null && x.Reefer)
                        .OrderBy(x => x.BayNo)
                        .ThenByDescending(x => x.TierNo)
                        .ThenBy(x =>
                        {
                            int chNoNumeric;
                            return int.TryParse(_excelService.ExtractRdbNumber(x.CHNo), out chNoNumeric)
                                ? chNoNumeric
                                : int.MaxValue;
                        })
                        .ThenBy(x => x.CHNo ?? "")
                        .ThenBy(x =>
                        {
                            string place = _excelService.NormalizePlace(x.Place);

                            if (place == "ON DECK")
                                return 0;

                            if (place == "IN HOLD")
                                return 1;

                            return 9;
                        })
                        .ThenBy(x =>
                        {
                            if (x.RowNo == 0)
                                return 1;

                            if (x.RowNo % 2 != 0)
                                return 0; // ST'BD

                            return 2;     // PORT
                        })
                        .ThenBy(x =>
                        {
                            if (x.RowNo % 2 != 0)
                                return -x.RowNo; // 19,17,15...1

                            return x.RowNo;      // 0,2,4...20
                        })
                        .ThenBy(x => x.X)
                        .ThenBy(x => x.Y)
                        .ToList();

                    for (int i = 0; i < orderedItems.Count; i++)
                        orderedItems[i].TagNo = $"R{i}";

                    SplashScreenManager.Default.SetWaitFormDescription("Loading... Group Build...");
                    
                    var sortItems = sortPriorityPanel.GetPriorityItems();
                    
                    int Bay_asc = GetBayAscFromPanel();
                    int Tier_asc = sortItems.FirstOrDefault(x => x.Key == "Tier")?.SelectedOrder == "Top->Bottom" ? 0 : 1;
                    int Row_asc = sortItems.FirstOrDefault(x => x.Key == "Row")?.SelectedOrder == "PORT->ST'BD" ? 1 : 0;

                    SaveRenderedOrderOptions(Bay_asc, Tier_asc, Row_asc);

                    // 대칭 색상은 GangBox 번호와 1:1로 결합되지 않는다.
                    // Excel의 인접 색상 그룹을 복원한 뒤 Order By 순서로 번호를 다시 부여한다.
                    AssignGangBoxNumbersFromExcel(raw_List);

                    grouped_List = raw_List
                        .Where(x => x != null && x.Reefer)
                        .GroupBy(x => new
                        {
                            x.BayNo,
                            x.CHNo,
                            x.Place
                        })
                        .Select(g => new BayGroupDto
                        {
                            BayNo = g.Key.BayNo,
                            CHNo = g.Key.CHNo,
                            Place = g.Key.Place,
                            Tiers = _excelService.BuildTierList(g, Tier_asc, Row_asc)
                        })
                        .ToList();

                    grouped_List = (Bay_asc == 0
                            ? grouped_List.OrderBy(x => x.BayNo)
                            : grouped_List.OrderByDescending(x => x.BayNo))
                        .ThenBy(x =>
                        {
                            int chNoNumeric;
                            return int.TryParse(_excelService.ExtractRdbNumber(x.CHNo), out chNoNumeric)
                                ? chNoNumeric
                                : int.MaxValue;
                        })
                        .ThenBy(x => x.CHNo ?? "")
                        .ThenBy(x =>
                        {
                            string place = _excelService.NormalizePlace(x.Place);

                            if (place == "ON DECK")
                                return 0;

                            if (place == "IN HOLD")
                                return 1;

                            return 9;
                        })
                        .ToList();

                    grouped_List = ApplyDeckFirstProcessingOrder(grouped_List);

                    var importedItems = raw_List
                        .Where(x => x != null && x.Reefer)
                        .ToList();

                    // Excel-only Rebuild에서도 선택 Cell이 바로 Navigate될 수 있도록
                    // 숨김 시트의 Reefer OID를 현재 메모리 항목에 연결한다.
                    foreach (ReeferItem item in importedItems)
                    {
                        string reeferDesignation = Get_Reefer_Name(item);
                        if (_workbookReeferEbIds.TryGetValue(
                            reeferDesignation,
                            out string reeferEbId))
                        {
                            item.Reefer_EB_ID = reeferEbId;
                        }
                    }

                    if (synchronizeEbObjects)
                    {
                    // Update 시 사용자가 수정한 셀의 실제 RDB 번호를 기준으로
                    // PORT/STBD 노란색 요약 영역을 다시 계산한다.
                    _excelService.UpdateRdbSummaryCells(selectedWorksheet, mode, importedItems);

                    SplashScreenManager.Default.SetWaitFormDescription("Loading... EB WorkSheet...");

                    if (hasPersistedSyncState)
                    {
                        RDB_dict.Clear();
                        Gangbox_dict.Clear();
                        Reefer_dict.Clear();
                        _rdbObjectCache.Clear();
                        _gangboxObjectCache.Clear();
                        _reeferObjectCache.Clear();
                        _ebObjectCacheById.Clear();

                        foreach (EbSyncReeferState state in persistedSyncStates)
                        {
                            if (!string.IsNullOrWhiteSpace(state.ReeferDesignation) &&
                                !string.IsNullOrWhiteSpace(state.ReeferEbId))
                            {
                                Reefer_dict[state.ReeferDesignation] = state.ReeferEbId;
                            }

                            if (!string.IsNullOrWhiteSpace(state.GangBoxDesignation) &&
                                !string.IsNullOrWhiteSpace(state.GangBoxEbId))
                            {
                                Gangbox_dict[state.GangBoxDesignation] = state.GangBoxEbId;
                            }

                            if (!string.IsNullOrWhiteSpace(state.RdbDesignation) &&
                                !string.IsNullOrWhiteSpace(state.RdbEbId))
                            {
                                RDB_dict[state.RdbDesignation] = state.RdbEbId;
                            }
                        }

                        Debug.WriteLine(
                            $"[EB FAST DIFF] persisted state loaded: {persistedSyncStates.Count}");
                    }
                    else
                    {
                        // 기존 Output처럼 동기화 기준값이 없는 파일은 최초 1회만
                        // EB 전체 맵을 구성한 후 숨김 기준 시트를 생성한다.
                        Create_Dict_All(waitFormShown);
                    }

                    Origin_eb_reefer = null;

                    List<EbSyncReeferState> desiredSyncStates =
                        _updateService.BuildEbSyncStates(
                            importedItems,
                            Get_RDB_Name,
                            Get_Gangbox_Name,
                            Get_Reefer_Name,
                            OrderReeferItemsForPosition);
                    var persistedSyncStateByReefer = persistedSyncStates
                        .Where(x => x != null && !string.IsNullOrWhiteSpace(x.ReeferDesignation))
                        .GroupBy(x => x.ReeferDesignation, StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);
                    HashSet<string> rdbsRequiringGangBoxRenumbering = hasPersistedSyncState
                        ? _updateService.GetRdbDesignationsRequiringGangBoxRenumbering(
                            persistedSyncStates,
                            desiredSyncStates)
                        : new HashSet<string>(
                            desiredSyncStates
                                .Where(x => !string.IsNullOrWhiteSpace(x.RdbDesignation))
                                .Select(x => x.RdbDesignation),
                            StringComparer.OrdinalIgnoreCase);
                    if (hasPersistedSyncState)
                    {
                        Debug.WriteLine(
                            $"[GANGBOX RENUMBER RDB] " +
                            $"{string.Join(", ", rdbsRequiringGangBoxRenumbering.OrderBy(x => x))}");
                    }

                    SplashScreenManager.Default.SetWaitFormDescription("Loading... EB ID Mapping...");

                    foreach (var item in importedItems)
                    {
                        // 기존 EB ID 매핑은 한 번만 수행한다.
                        if (!string.IsNullOrWhiteSpace(item.RDB_ID) && RDB_dict.TryGetValue(item.RDB_ID, out string rdbId))
                            item.RDB_EB_ID = rdbId;

                        int gangBoxNo = item.GangBox_No;
                        if (!string.IsNullOrWhiteSpace(item.RDB_ID) && gangBoxNo > 0)
                        {
                            string gangBoxName = Get_Gangbox_Name(gangBoxNo.ToString(), item.RDB_ID);
                            if (Gangbox_dict.TryGetValue(gangBoxName, out string gangBoxId))
                                item.GangBox_EB_ID = gangBoxId;
                        }

                        string reeferName = Get_Reefer_Name(item);
                        if (Reefer_dict.TryGetValue(reeferName, out string reeferId))
                            item.Reefer_EB_ID = reeferId;
                    }

                    var invalidAssociationItem = importedItems.FirstOrDefault(item =>
                        string.IsNullOrWhiteSpace(item.RDB_ID) ||
                        GetGangBoxIndexFromColor(item.GangBox_Color) <= 0 ||
                        item.GangBox_No <= 0);

                    if (invalidAssociationItem != null)
                    {
                        throw new InvalidOperationException(
                            $"최종 Excel에 RDB 번호 또는 GangBox 색상이 올바르지 않은 Reefer가 있습니다. " +
                            $"Bay={invalidAssociationItem.BayNo}, Tier={invalidAssociationItem.TierNo}, " +
                            $"Row={invalidAssociationItem.RowNo}, CH={invalidAssociationItem.CHNo}, " +
                            $"RDB={invalidAssociationItem.RDB_ID}, Color={invalidAssociationItem.GangBox_Color}");
                    }

                    // 같은 GangBox의 Reefer들을 한 그룹으로 묶어 객체 생성/association을 처리한다.
                    var importedAssociationGroups = importedItems
                        .Select(item => new
                        {
                            Item = item,
                            GangBoxNo = item.GangBox_No
                        })
                        .Where(x => !string.IsNullOrWhiteSpace(x.Item.RDB_ID) && x.GangBoxNo > 0)
                        .Select(x => new
                        {
                            x.Item,
                            x.GangBoxNo,
                            RdbDesignation = Get_RDB_Name(x.Item.RDB_ID),
                            GangBoxDesignation = Get_Gangbox_Name(x.GangBoxNo.ToString(), Get_RDB_Name(x.Item.RDB_ID))
                        })
                        .Where(x => !string.IsNullOrWhiteSpace(x.RdbDesignation) && !string.IsNullOrWhiteSpace(x.GangBoxDesignation))
                        .GroupBy(x => new { x.RdbDesignation, x.GangBoxDesignation })
                        .ToList();

                    var desiredGangBoxDesignations = new HashSet<string>(
                        importedAssociationGroups.Select(x => x.Key.GangBoxDesignation),
                        StringComparer.OrdinalIgnoreCase);
                    var desiredGangBoxByReefer = importedAssociationGroups
                        .SelectMany(group => group.Select(entry => new
                        {
                            ReeferName = Get_Reefer_Name(entry.Item),
                            GangBox = group.Key.GangBoxDesignation
                        }))
                        .GroupBy(x => x.ReeferName, StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(
                            x => x.Key,
                            x => x.First().GangBox,
                            StringComparer.OrdinalIgnoreCase);
                    var affectedRdbDesignations = new HashSet<string>(
                        rdbsRequiringGangBoxRenumbering,
                        StringComparer.OrdinalIgnoreCase);
                    var affectedGangBoxDesignations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    // __EB_SYNC_STATE는 OID 조회 힌트로만 사용한다.
                    // Association의 현재 상태는 대상 프로젝트의 EB 객체에서 직접 조회한다.

                    int importedGroupIndex = 0;
                    foreach (var group in importedAssociationGroups)
                    {
                        importedGroupIndex++;
                        List<ReeferItem> orderedGroupItems =
                            OrderReeferItemsForPosition(group.Select(x => x.Item));

                        Debug.WriteLine(
                            $"[EB MAP GROUP START] {importedGroupIndex}/{importedAssociationGroups.Count}: " +
                            $"{group.Key.GangBoxDesignation} -> {group.Key.RdbDesignation}");

                        if (SplashScreenManager.Default != null && SplashScreenManager.Default.IsSplashFormVisible)
                        {
                            SplashScreenManager.Default.SetWaitFormDescription(
                                $"Loading... EB ID Mapping - GangBox {importedGroupIndex}/{importedAssociationGroups.Count}\n" +
                                $"Preparing: {group.Key.GangBoxDesignation} -> {group.Key.RdbDesignation}");
                        }

                        var desiredReeferNames = new HashSet<string>(
                            group.Select(x => Get_Reefer_Name(x.Item)),
                            StringComparer.OrdinalIgnoreCase);

                        bool gangBoxExists = Gangbox_dict.ContainsKey(group.Key.GangBoxDesignation);
                        IAucObject existingGangBox = gangBoxExists
                            ? GetCachedObjectByDesignation(
                                group.Key.GangBoxDesignation,
                                Gangbox_dict,
                                _gangboxObjectCache)
                            : null;
                        Dictionary<string, string> currentReeferIds =
                            GetAssociatedReeferIds(existingGangBox);

                        // 한 GangBox의 추가/삭제로 뒤쪽 번호가 밀릴 수 있으므로,
                        // 변경 RDB에 속한 GangBox는 모두 최종 번호/내용을 다시 동기화한다.
                        bool groupChanged =
                            rdbsRequiringGangBoxRenumbering.Contains(group.Key.RdbDesignation) ||
                            !gangBoxExists ||
                            !RDB_dict.ContainsKey(group.Key.RdbDesignation) ||
                            !desiredReeferNames.SetEquals(currentReeferIds.Keys);

                        if (!groupChanged && hasPersistedSyncState)
                        {
                            List<EbSyncReeferState> previousGroupStates = persistedSyncStates
                                .Where(x => string.Equals(
                                    x.GangBoxDesignation,
                                    group.Key.GangBoxDesignation,
                                    StringComparison.OrdinalIgnoreCase))
                                .OrderBy(x => x.Position)
                                .ToList();
                            var desiredGroupEntries = orderedGroupItems;

                            groupChanged = previousGroupStates.Count != desiredGroupEntries.Count;
                            for (int entryIndex = 0;
                                !groupChanged && entryIndex < desiredGroupEntries.Count;
                                entryIndex++)
                            {
                                groupChanged =
                                    previousGroupStates[entryIndex].Position != entryIndex + 1 ||
                                    !string.Equals(
                                        previousGroupStates[entryIndex].ReeferDesignation,
                                        Get_Reefer_Name(desiredGroupEntries[entryIndex]),
                                        StringComparison.OrdinalIgnoreCase) ||
                                    !string.Equals(
                                        previousGroupStates[entryIndex].RdbDesignation,
                                        group.Key.RdbDesignation,
                                        StringComparison.OrdinalIgnoreCase);
                            }
                        }

                        if (!groupChanged)
                        {
                            string existingRdbId = RDB_dict[group.Key.RdbDesignation];
                            string existingGangBoxId = Gangbox_dict[group.Key.GangBoxDesignation];
                            var existingDesiredReefers = new List<IAucObject>();

                            foreach (var entry in group)
                            {
                                string reeferName = Get_Reefer_Name(entry.Item);
                                entry.Item.RDB_EB_ID = existingRdbId;
                                entry.Item.GangBox_EB_ID = existingGangBoxId;

                                if (currentReeferIds.TryGetValue(reeferName, out string currentReeferId))
                                {
                                    entry.Item.Reefer_EB_ID = currentReeferId;
                                    IAucObject currentReefer = GetCachedObjectById(currentReeferId);
                                    if (currentReefer != null)
                                        existingDesiredReefers.Add(currentReefer);
                                }
                                else if (Reefer_dict.TryGetValue(reeferName, out string reeferId))
                                {
                                    entry.Item.Reefer_EB_ID = reeferId;
                                    IAucObject currentReefer = GetCachedObjectById(reeferId);
                                    if (currentReefer != null)
                                        existingDesiredReefers.Add(currentReefer);
                                }
                            }

                            if (existingDesiredReefers.Count != group.Count())
                            {
                                throw new InvalidOperationException(
                                    $"기존 GangBox의 Reefer 객체 조회에 실패했습니다. " +
                                    $"GangBox={group.Key.GangBoxDesignation}");
                            }

                            bool existingReeferAssociationSynchronized =
                                _updateService.SynchronizeGangBoxReefers(
                                    ebHelper,
                                    existingGangBox,
                                    existingDesiredReefers,
                                    out int existingAddedCount,
                                    out int existingRemovedCount);
                            IAucObject existingRdb = GetCachedObjectById(existingRdbId);
                            bool existingRdbAssociationSynchronized =
                                _updateService.SynchronizeRdbGangBox(
                                    ebHelper,
                                    existingRdb,
                                    existingGangBox);

                            if (!existingReeferAssociationSynchronized ||
                                !existingRdbAssociationSynchronized)
                            {
                                throw new InvalidOperationException(
                                    $"기존 그룹의 실제 EB Association 확인에 실패했습니다. " +
                                    $"GangBox={group.Key.GangBoxDesignation}");
                            }

                            Debug.WriteLine(
                                $"[EB MAP GROUP ASSOCIATION CHECK] {group.Key.GangBoxDesignation}, " +
                                $"unchanged={group.Count()}, +{existingAddedCount}, -{existingRemovedCount}");
                            continue;
                        }

                        affectedRdbDesignations.Add(group.Key.RdbDesignation);
                        affectedGangBoxDesignations.Add(group.Key.GangBoxDesignation);

                        if (group.Any(x => !Reefer_dict.ContainsKey(Get_Reefer_Name(x.Item))) &&
                            Origin_eb_reefer == null)
                        {
                            Origin_eb_reefer = First_Create_EB_Reefer();
                            if (Origin_eb_reefer == null)
                                throw new InvalidOperationException("Origin Reefer 객체 생성에 실패했습니다.");
                        }

                        Debug.WriteLine($"[EB MAP CREATE RDB START] {group.Key.RdbDesignation}");
                        IAucObject rdb = Create_EB_RDB(group.Key.RdbDesignation);
                        Debug.WriteLine($"[EB MAP CREATE RDB END] {group.Key.RdbDesignation}");

                        Debug.WriteLine($"[EB MAP CREATE GANGBOX START] {group.Key.GangBoxDesignation}");
                        IAucObject gangbox = Create_EB_GangBox(group.Key.GangBoxDesignation);
                        Debug.WriteLine($"[EB MAP CREATE GANGBOX END] {group.Key.GangBoxDesignation}");
                        if (rdb == null || gangbox == null)
                        {
                            throw new InvalidOperationException(
                                $"EB 객체 생성에 실패했습니다. " +
                                $"RDB={group.Key.RdbDesignation}, GangBox={group.Key.GangBoxDesignation}");
                        }

                        var desiredReefers = new List<IAucObject>();
                        bool canSynchronizeReeferAssociations = gangbox != null;
                        int reeferPosition = 0;

                        foreach (var item in orderedGroupItems)
                        {
                            int position = reeferPosition + 1;
                            string reeferDesignation = Get_Reefer_Name(item);
                            bool updateReeferAttributes = true;
                            if (hasPersistedSyncState &&
                                persistedSyncStateByReefer.TryGetValue(
                                    reeferDesignation,
                                    out EbSyncReeferState previousReeferState))
                            {
                                updateReeferAttributes =
                                    previousReeferState.Position != position ||
                                    !string.Equals(
                                        previousReeferState.GangBoxDesignation,
                                        group.Key.GangBoxDesignation,
                                        StringComparison.OrdinalIgnoreCase) ||
                                    !string.Equals(
                                        previousReeferState.RdbDesignation,
                                        group.Key.RdbDesignation,
                                        StringComparison.OrdinalIgnoreCase);
                            }

                            IAucObject reefer = Create_EB_Reefer(
                                item,
                                position,
                                updateExistingAttributes: updateReeferAttributes);
                            if (reefer == null)
                            {
                                throw new InvalidOperationException(
                                    $"Reefer EB 객체 생성에 실패했습니다. " +
                                    $"Reefer={Get_Reefer_Name(item)}, GangBox={group.Key.GangBoxDesignation}");
                            }

                            reeferPosition = position;

                            item.RDB_EB_ID = rdb?.ID;
                            item.GangBox_EB_ID = gangbox.ID;
                            item.Reefer_EB_ID = reefer.ID;

                            desiredReefers.Add(reefer);
                        }

                        int addedReeferAssociationCount = 0;
                        int removedReeferAssociationCount = 0;
                        var detachedReefers = new List<IAucObject>();
                        bool reeferAssociationSynchronized = false;
                        if (canSynchronizeReeferAssociations)
                        {
                            Debug.WriteLine($"[EB MAP REEFER SYNC START] {gangbox.Name}");
                            reeferAssociationSynchronized =
                                _updateService.SynchronizeGangBoxReefers(
                                    ebHelper,
                                    gangbox,
                                    desiredReefers,
                                    out addedReeferAssociationCount,
                                    out removedReeferAssociationCount,
                                    out detachedReefers);
                            Debug.WriteLine($"[EB MAP REEFER SYNC END] {gangbox.Name}");
                        }

                        if (!reeferAssociationSynchronized)
                        {
                            throw new InvalidOperationException(
                                $"GangBox-Reefer Association 동기화에 실패했습니다. GangBox={gangbox.Name}");
                        }

                        // 다른 GangBox로 이동한 Reefer는 새 그룹에서 계산한 Position을 유지해야 한다.
                        // 최종 Excel의 어떤 GangBox에도 속하지 않는 Reefer만 Position을 초기화한다.
                        ClearDetachedReeferPositions(
                            gangbox.Name,
                            detachedReefers.Where(x =>
                                x != null &&
                                !desiredGangBoxByReefer.ContainsKey(x.Name)));

                        Debug.WriteLine(
                            $"[EB MAP RDB ASSOCIATION START] {rdb?.Name} -> {gangbox?.Name}");
                        bool rdbAssociationSynchronized =
                            _updateService.SynchronizeRdbGangBox(ebHelper, rdb, gangbox);
                        Debug.WriteLine(
                            $"[EB MAP RDB ASSOCIATION END] {rdb?.Name} -> {gangbox?.Name}");

                        if (SplashScreenManager.Default != null && SplashScreenManager.Default.IsSplashFormVisible)
                        {
                            SplashScreenManager.Default.SetWaitFormDescription(
                                $"Loading... EB ID Mapping - GangBox {importedGroupIndex}/{importedAssociationGroups.Count}\n" +
                                $"{(gangbox != null ? gangbox.Name : "")} -> {(rdb != null ? rdb.Name : "")} : {rdbAssociationSynchronized}\n" +
                                $"Reefer association: +{addedReeferAssociationCount}, -{removedReeferAssociationCount}");
                        }

                        Debug.WriteLine(
                            $"==> {group.Key.GangBoxDesignation} -> {group.Key.RdbDesignation}, " +
                            $"Reefer synchronized: {reeferAssociationSynchronized}, " +
                            $"added={addedReeferAssociationCount}, removed={removedReeferAssociationCount}, " +
                            $"positionCleared={detachedReefers.Count}");

                    }

                    // 최종 Excel에서 빠진 기존 GangBox는 객체를 삭제하지 않고
                    // 알려진 Association과 Reefer Position만 최종 상태에 맞춘다.
                    var obsoleteGangBoxes =
                        new Dictionary<string, IAucObject>(StringComparer.OrdinalIgnoreCase);

                    foreach (string designation in Gangbox_dict.Keys
                        .Where(x => !desiredGangBoxDesignations.Contains(x))
                        .ToList())
                    {
                        IAucObject gangBox = GetCachedObjectByDesignation(
                            designation,
                            Gangbox_dict,
                            _gangboxObjectCache);
                        if (gangBox != null)
                            obsoleteGangBoxes[designation] = gangBox;
                    }

                    foreach (var obsoleteEntry in obsoleteGangBoxes)
                    {
                        string obsoleteGangBoxDesignation = obsoleteEntry.Key;
                        IAucObject obsoleteGangBox = obsoleteEntry.Value;
                        affectedGangBoxDesignations.Add(obsoleteGangBoxDesignation);

                        int obsoleteAddedCount;
                        int obsoleteRemovedCount;
                        List<IAucObject> obsoleteDetachedReefers;
                        bool obsoleteGangBoxSynchronized =
                            _updateService.SynchronizeGangBoxReefers(
                                ebHelper,
                                obsoleteGangBox,
                                Enumerable.Empty<IAucObject>(),
                                out obsoleteAddedCount,
                                out obsoleteRemovedCount,
                                out obsoleteDetachedReefers);

                        if (!obsoleteGangBoxSynchronized)
                        {
                            throw new InvalidOperationException(
                                $"최종 Excel에서 제외된 GangBox Association 정리에 실패했습니다. " +
                                $"GangBox={obsoleteGangBoxDesignation}");
                        }

                        ClearDetachedReeferPositions(
                            obsoleteGangBoxDesignation,
                            obsoleteDetachedReefers.Where(x =>
                                x != null &&
                                !desiredGangBoxByReefer.ContainsKey(x.Name)));

                        Debug.WriteLine(
                            $"[OBSOLETE GANGBOX RETAINED] GangBox={obsoleteGangBoxDesignation}, " +
                            $"removedAssociation={obsoleteRemovedCount}");
                    }

                    Debug.WriteLine("[REBUILD RDB ATTRIBUTES START]");
                    if (SplashScreenManager.Default != null && SplashScreenManager.Default.IsSplashFormVisible)
                        SplashScreenManager.Default.SetWaitFormDescription("Loading... RDB attributes...");
                    // Reefer 구성이 변하지 않아도 Reefer/SRD 기준값과 역률이 바뀌면
                    // 정격 전력/전류가 달라지므로 모든 RDB 속성을 다시 계산한다.
                    /*UpdateRdbAttributes(importedItems.Where(item => affectedRdbDesignations.Contains(Get_RDB_Name(item.RDB_ID))));*/
                    UpdateRdbAttributes(importedItems);
                    Debug.WriteLine("[REBUILD RDB ATTRIBUTES END]");

                    Debug.WriteLine("[REBUILD GANGBOX ATTRIBUTES START]");
                    if (SplashScreenManager.Default != null && SplashScreenManager.Default.IsSplashFormVisible)
                        SplashScreenManager.Default.SetWaitFormDescription("Loading... GangBox attributes...");
                    // Reefer 구성이 변하지 않아도 Reefer/DistBoard 기준값과 역률이 바뀌면
                    // 정격 전력/전류가 달라지므로 모든 GangBox 속성을 다시 계산한다.                    
                    /*UpdateGangBoxAttributes(importedItems.Where(item =>
                    {string gangBoxDesignation = Get_Gangbox_Name(item.GangBox_No.ToString(),Get_RDB_Name(item.RDB_ID)); 
                        return affectedGangBoxDesignations.Contains(gangBoxDesignation);}));*/
                    UpdateGangBoxAttributes(importedItems);
                    Debug.WriteLine("[REBUILD GANGBOX ATTRIBUTES END]");

                    _updateService.WriteEbSyncState(
                        import_excel,
                        selectedWorksheet,
                        _updateService.BuildEbSyncStates(
                            importedItems,
                            Get_RDB_Name,
                            Get_Gangbox_Name,
                            Get_Reefer_Name,
                            OrderReeferItemsForPosition),
                        EbSyncStateWorksheetName,
                        EbSyncStateMarker);
                    import_excel.Save(filePath);

                    }

                    gridView1.BeginUpdate();

                    try
                    {
                        gridControl1.DataSource = null;
                        gridControl1.DataSource = grouped_List;

                        var placeCol = gridView1.Columns[nameof(BayGroupDto.Place)];
                        if (placeCol != null)
                            placeCol.Visible = deck_hold_type != 1;

                        var warningCol = gridView1.Columns[nameof(BayGroupDto.Warning)];
                        if (warningCol != null)
                            warningCol.Visible = false;
                    }
                    finally
                    {
                        gridView1.EndUpdate();
                    }

                    gridControl1.RefreshDataSource();
                    gridView1.RefreshData();

                    // ReBuild는 입력 Excel을 정답으로 사용한다.
                    // grouping 재계산이나 별도 Output 저장 없이 입력 파일을 그대로 유지한다.
                    savePath = filePath;
                    sheet8grid_tfn = true;

                    spreadsheetControl1.LoadDocument(savePath);
                    spreadsheetControl1.BringToFront();

                    st.Stop();

                    int rdbMapped = raw_List.Count(x => !string.IsNullOrWhiteSpace(x.RDB_EB_ID));
                    int gangMapped = raw_List.Count(x => !string.IsNullOrWhiteSpace(x.GangBox_EB_ID));
                    int reeferMapped = raw_List.Count(x => !string.IsNullOrWhiteSpace(x.Reefer_EB_ID));

                    Debug.WriteLine($"[REBUILD END] {DateTime.Now:HH:mm:ss.fff} / {st.Elapsed.TotalSeconds:N3} sec");
                    if (synchronizeEbObjects)
                        Debug.WriteLine($"[EB MAP] RDB={rdbMapped}, GangBox={gangMapped}, Reefer={reeferMapped}, Total={raw_List.Count}");

                    if (showCompletionMessage)
                    {
                        string completionMessage = synchronizeEbObjects
                            ? $"완성 Excel 기준 EB 객체 동기화 완료\r\n" +
                              $"Mode: {mode}\r\n" +
                              $"raw_List: {raw_List.Count}\r\n" +
                              $"grouped_List: {grouped_List.Count}\r\n" +
                              $"RDB EB 매핑: {rdbMapped}/{raw_List.Count}\r\n" +
                              $"GangBox EB 매핑: {gangMapped}/{raw_List.Count}\r\n" +
                              $"Reefer EB 매핑: {reeferMapped}/{raw_List.Count}"
                            : $"완성 Excel 불러오기 완료 (EB 객체 생성 없음)\r\n" +
                              $"Mode: {mode}\r\n" +
                              $"raw_List: {raw_List.Count}\r\n" +
                              $"grouped_List: {grouped_List.Count}\r\n" +
                              $"Navigate OID 매핑: {reeferMapped}/{raw_List.Count}";

                        MessageBox.Show(completionMessage, "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    if (synchronizeEbObjects)
                        _validatedEbOidSessionKey = ebOidValidationSessionKey;

                    return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "완성 Excel 역변환 실패: " + ex.Message,
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                if (waitFormShown &&
                    SplashScreenManager.Default != null &&
                    SplashScreenManager.Default.IsSplashFormVisible)
                {
                    SplashScreenManager.CloseForm();
                }
            }
        }
         
        private void btn_callworksheet_Click(object sender, EventArgs e)
        {
            SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);

            //ebapp.Utils.RunMacro("Edu_Test01.Edu_Test04", ebapp.Selection.Item(1).Project);

            SplashScreenManager.CloseForm();

            MessageBox.Show($"Worksheet Open!!");
        }

        private void SaveToEB_UpLoad()
        {
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("EB에 업로드할 Output 파일을 찾지 못했습니다.", outputPath);
            _settingService.UploadAllExternalDocuments(outputPath, outputFileName);
        }


    }
}
