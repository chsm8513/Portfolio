using Aucotec.EngineeringBase.Interop;
using Auto_SLD.Models;
using Auto_SLD.Controls;
using Auto_SLD.Service;
using EngineeringBase;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;

namespace Auto_SLD;

public partial class Form1 : DevExpress.XtraEditors.XtraForm
{
    private const string EngineeringBaseProgId = "EngineeringBase.Application.32";
    private const string TestSelectionColumnName = "TestSelection";
    private bool _Test = false;
    private readonly IEbAccessService _accessService;
    private readonly IEbRepository _repository;
    private readonly DataService _dataService;
    private string? _currentProjectId;
    private bool _busy;

    public Form1() : this(new EbAccessService(EngineeringBaseProgId))
    {
    }
    private void Form1_Load(object sender, EventArgs e)
    {
        string version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "0.0.0.0";
        this.Text = $"Auto SLD(Single Line Diagram) v{version}";        
    }

    internal Form1(IEbAccessService accessService)
    {
        _accessService = accessService ?? throw new ArgumentNullException(nameof(accessService));
        _repository = new EbRepository(_accessService);
        _dataService = new DataService(_repository);
        InitializeComponent();
        UpdateActionState();
    }

    private async void btnFindProject_Click(object? sender, EventArgs e)
    {
        if (_busy) return;

        try
        {
            SetBusy(true, "EB 프로젝트 조회 중", "현재 EB 프로젝트를 찾는 중...");
            ClearLoadedData();

            var progress = new Progress<ProgressInfo>(info => worksheetProgressPanel.UpdateProgress(info.Description, info.Caption));

            ProjectDataResult result = await RunStaThreadAsync(() => LoadProjectData(progress));

            _currentProjectId = result.ProjectId;
            txtProject.Text = result.ProjectName;
            txtProject.Tag = result.ProjectId;
            dataGridView.DataSource = result.LoadDevices;
            ConfigureTestSelectionColumn();
            SetStatus($"EB Project '{result.ProjectName}' → Load Device {result.LoadDevices.Count:N0}개 조회 완료");
        }
        catch (Exception exception)
        {
            ClearLoadedData();
            ShowError("EB 프로젝트 데이터 조회 실패", exception);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private ProjectDataResult LoadProjectData(IProgress<ProgressInfo> progress)
    {
        IAucProject? project = null;
        IAucObject? equipmentFolder = null;

        try
        {
            if (!_accessService.IsEngineeringBaseRunning())
                throw new InvalidOperationException("Engineering Base를 실행하고 프로젝트를 활성화하세요.");

            project = _accessService.GetActiveProject() ?? throw new InvalidOperationException("현재 활성화된 EB 프로젝트가 없습니다.");
            equipmentFolder = _repository.GetEquipmentRoot(project);

            IReadOnlyList<LoadDeviceDto> loadDevices = _dataService.GetLoadDevices(project, equipmentFolder, progress);
            return new ProjectDataResult(project.ID, project.Name, loadDevices);
        }
        finally
        {
            EbComHelper.ReleaseComObject(equipmentFolder);
            EbComHelper.ReleaseComObject(project);
        }
    }

    private void ClearLoadedData()
    {
        _currentProjectId = null;
        _dataService.loadDevices.Clear();
        txtProject.Clear();
        txtProject.Tag = null;
        dataGridView.DataSource = null;
    }

    private async void btnAutoDrawing_Click(object? sender, EventArgs e)
    {
        ClearFailureHighlight();
        await RunAutoDrawingAsync();
    }

    private async Task RunAutoDrawingAsync()
    {
        if (_busy) return;
        if (_dataService.loadDevices.Count == 0)
        {
            MessageBox.Show(this, "Project에서 DataSet이 없습니다.", "No DataSet", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (string.IsNullOrWhiteSpace(_currentProjectId)) return;

        if (dataGridView.IsCurrentCellDirty) dataGridView.CommitEdit(DataGridViewDataErrorContexts.Commit);
        dataGridView.EndEdit();
        string projectId = _currentProjectId!;
        LoadDeviceDto[] devices = GetDrawingDevices();
        if (devices.Length == 0)
        {
            MessageBox.Show(this, "도면을 그릴 항목을 체크하세요.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        try
        {
            _accessService.GetApplication().AlertResponse = AucAlertResponse.aucResponseYesOK;
            SetBusy(true, "도면 작성 중", "Typical 도면을 복사하는 중...");
            var progress = new Progress<ProgressInfo>(info => worksheetProgressPanel.UpdateProgress(info.Description, info.Caption));
            AutoDrawingResult result = await RunStaThreadAsync(() =>
            {
                // 조회 작업에서 사용한 RCW를 재사용하지 않고 이 STA에서 EB 객체를 다시 취득한다.
                var accessService = new EbAccessService(EngineeringBaseProgId);
                var repository = new EbRepository(accessService);
                using var mouseInputGuard = new MouseInputGuard();
                var service = new AutoDrawingService(accessService, repository);
                return service.New_draw(projectId, devices, progress);
            });
            ClearFailureHighlight();
            SetBusy(true, "도면 완료", $"도면 작성 완료: Sheet {result.SheetCount:N0}장 / Load {result.FeederCount:N0}개 / Equipment 이동 {result.MovedEquipmentCount:N0}개 / association {result.ReassociatedSymbolCount:N0}개 / 임시 객체 삭제 {result.RemovedDummyCount:N0}개");
            SetBusy(false);
            MessageBox.Show($"도면 작성 완료: Sheet {result.SheetCount:N0}장 / Load {result.FeederCount:N0}개 / Equipment 이동 {result.MovedEquipmentCount:N0}개 / association {result.ReassociatedSymbolCount:N0}개 / 임시 객체 삭제 {result.RemovedDummyCount:N0}개");
        }
        catch (Exception exception)
        {
            if (exception.Data["FailedDeviceIndex"] is int failedDeviceIndex && failedDeviceIndex >= 0 && failedDeviceIndex < devices.Length) HighlightFailureRow(devices[failedDeviceIndex]);
            ShowError("Auto Drawing 실패", exception);
        }
        finally
        {
            SetBusy(false);
            _accessService.GetApplication().AlertResponse = AucAlertResponse.aucResponseNone;
        }
    }
    
    private void btnExit_Click(object? sender, EventArgs e) => Close();

    private static Task<T> RunStaThreadAsync<T>(Func<T> action)
    {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            try
            {
                completion.SetResult(action());
            }
            catch (Exception exception)
            {
                completion.SetException(exception);
            }
        })
        {
            IsBackground = true,
            Name = "Engineering Base Worker"
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }

    private void SetBusy(bool busy,string? caption = null,string? description = null)
    {
        _busy = busy;
        btnFindProject.Enabled = !busy;
        btnExit.Enabled = !busy;
        dataGridView.Enabled = !busy;

        if (busy)
        {
            worksheetProgressPanel.ShowProgress(caption ?? "처리 중", description ?? string.Empty);

            if (!string.IsNullOrWhiteSpace(description))
                SetStatus(description!);
        }
        else
        {
            worksheetProgressPanel.HideProgress();
        }

        UpdateActionState();
    }

    private void UpdateActionState()
    {
        btnAutoDrawing.Enabled = !_busy && !string.IsNullOrWhiteSpace(_currentProjectId) && dataGridView.Rows.Count > 0;
    }

    private void ClearFailureHighlight()
    {
        foreach (DataGridViewRow row in dataGridView.Rows)
        {
            row.DefaultCellStyle.BackColor = Color.Empty;
            row.DefaultCellStyle.SelectionBackColor = Color.Empty;
        }
    }

    private void ConfigureTestSelectionColumn()
    {
        if (dataGridView.Columns.Contains(TestSelectionColumnName)) dataGridView.Columns.Remove(TestSelectionColumnName);
        dataGridView.ReadOnly = !_Test;
        if (!_Test) return;

        var selectionColumn = new DataGridViewCheckBoxColumn
        {
            Name = TestSelectionColumnName,
            HeaderText = "선택",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
            Width = 50,
            Frozen = true,
            ReadOnly = false,
            SortMode = DataGridViewColumnSortMode.NotSortable,
            TrueValue = true,
            FalseValue = false
        };
        dataGridView.Columns.Insert(0, selectionColumn);
        foreach (DataGridViewColumn column in dataGridView.Columns) column.ReadOnly = column.Name != TestSelectionColumnName;
    }

    private LoadDeviceDto[] GetDrawingDevices()
    {
        if (!_Test) return _dataService.loadDevices.ToArray();
        return dataGridView.Rows.Cast<DataGridViewRow>().Where(row => Convert.ToBoolean(row.Cells[TestSelectionColumnName].Value)).Select(row => row.DataBoundItem).OfType<LoadDeviceDto>().ToArray();
    }

    private void HighlightFailureRow(LoadDeviceDto failedDevice)
    {
        ClearFailureHighlight();
        DataGridViewRow? failedRow = dataGridView.Rows.Cast<DataGridViewRow>().FirstOrDefault(row => ReferenceEquals(row.DataBoundItem, failedDevice));
        if (failedRow == null) return;
        failedRow.DefaultCellStyle.BackColor = Color.LightYellow;
        failedRow.DefaultCellStyle.SelectionBackColor = Color.Gold;
        dataGridView.ClearSelection();
        failedRow.Selected = true;
        dataGridView.FirstDisplayedScrollingRowIndex = failedRow.Index;
    }

    private void SetStatus(string message) => lblStatus.Text = message;

    private void ShowError(string title, Exception exception)
    {
        SetStatus(title);
        MessageBox.Show(this, exception.Message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    
}
