using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;


namespace Reefer_Container 
{
    public partial class SortPriorityPanel : UserControl
    {
        private List<SortOrderItem> _items;
        private Panel[] _rows;
        private Label[] _priorityLabels;
        private Label[] _itemLabels;
        private ComboBox[] _orderCombos;
        private Button[] _upButtons;
        private Button[] _downButtons;
        private bool _isRebuildingUi;
        public event EventHandler PriorityChanged;
        private static readonly Dictionary<string, string[]> OrderOptions = new Dictionary<string, string[]>
        {
            { "Bay",  new[] { "Fwd->Aft", "Aft->Fwd" } },
            { "Tier", new[] { "Top->Bottom", "Bottom->Top" } },
            { "Row",  new[] { "PORT->ST'BD", "ST'BD->PORT" } },
        };

        public SortPriorityPanel()
        {
            InitializeComponent();
            _items = CreateDefaultItems();
            InitializeItemControls();
            RebuildUI();
        }

        /// <summary>현재 우선순위 순서대로 SortOrderItem 목록 반환</summary>
        public IReadOnlyList<SortOrderItem> GetPriorityItems() => _items.AsReadOnly();

        public void ResetToDefaults()
        {
            _items = CreateDefaultItems();

            RebuildUI();
            PriorityChanged?.Invoke(this, EventArgs.Empty);
        }

        public void ApplyPriorityItems(IEnumerable<SortOrderItem> items)
        {
            if (items == null)
                return;

            var result = new List<SortOrderItem>();
            var seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (SortOrderItem item in items)
            {
                if (item == null || string.IsNullOrWhiteSpace(item.Key) ||
                    !OrderOptions.ContainsKey(item.Key) || !seenKeys.Add(item.Key))
                    continue;

                string selectedOrder = item.SelectedOrder;

                if (!OrderOptions[item.Key].Contains(selectedOrder))
                    selectedOrder = OrderOptions[item.Key][0];

                result.Add(new SortOrderItem
                {
                    Label = string.IsNullOrWhiteSpace(item.Label) ? item.Key + " Order by" : item.Label,
                    Key = item.Key,
                    SelectedOrder = selectedOrder
                });
            }

            foreach (SortOrderItem defaultItem in CreateDefaultItems())
            {
                if (seenKeys.Add(defaultItem.Key))
                    result.Add(defaultItem);
            }

            _items = result;
            RebuildUI();
        }

        private static List<SortOrderItem> CreateDefaultItems()
        {
            return new List<SortOrderItem>
            {
                new SortOrderItem { Label = "Bay Order by", Key = "Bay", SelectedOrder = "Fwd->Aft" },
                new SortOrderItem { Label = "Tier Order by", Key = "Tier", SelectedOrder = "Top->Bottom" },
                new SortOrderItem { Label = "Row Order by", Key = "Row", SelectedOrder = "PORT->ST'BD" },
            };
        }

        private void InitializeItemControls()
        {
            _rows = new[] { _row1, _row2, _row3 };
            _priorityLabels = new[] { _priorityLabel1, _priorityLabel2, _priorityLabel3 };
            _itemLabels = new[] { _itemLabel1, _itemLabel2, _itemLabel3 };
            _orderCombos = new[] { _orderCombo1, _orderCombo2, _orderCombo3 };
            _upButtons = new[] { _upButton1, _upButton2, _upButton3 };
            _downButtons = new[] { _downButton1, _downButton2, _downButton3 };

            for (int i = 0; i < _rows.Length; i++)
            {
                int index = i;
                _orderCombos[i].SelectedIndexChanged += OrderCombo_SelectedIndexChanged;
                _upButtons[i].Click += (s, e) => MoveItem(index, -1);
                _downButtons[i].Click += (s, e) => MoveItem(index, +1);
            }
        }

        private void OrderCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isRebuildingUi)
                return;

            var combo = sender as ComboBox;
            var item = combo?.Tag as SortOrderItem;
            if (combo?.SelectedItem == null || item == null)
                return;

            item.SelectedOrder = combo.SelectedItem.ToString();
            PriorityChanged?.Invoke(this, EventArgs.Empty);
        }

        private void RebuildUI()
        {
            _isRebuildingUi = true;
            try
            {
                for (int i = 0; i < _rows.Length; i++)
                {
                    bool hasItem = i < _items.Count;
                    _rows[i].Visible = hasItem;
                    if (!hasItem)
                        continue;

                    SortOrderItem item = _items[i];
                    _rows[i].Tag = item;
                    _priorityLabels[i].Text = $"{i + 1}.";
                    _itemLabels[i].Text = item.Label;

                    ComboBox combo = _orderCombos[i];
                    combo.BeginUpdate();
                    try
                    {
                        combo.Items.Clear();
                        combo.Items.AddRange(OrderOptions[item.Key]);
                        combo.SelectedItem = item.SelectedOrder;
                    }
                    finally
                    {
                        combo.EndUpdate();
                    }

                    combo.Tag = item;
                    combo.Enabled = true;
                    combo.TabStop = true;
                    _upButtons[i].Enabled = i > 0;
                    _downButtons[i].Enabled = i < _items.Count - 1;
                }
            }
            finally
            {
                _isRebuildingUi = false;
            }
        }

        private void MoveItem(int index, int direction)
        {
            int newIndex = index + direction;
            if (newIndex < 0 || newIndex >= _items.Count) return;

            var temp = _items[index];
            _items[index] = _items[newIndex];
            _items[newIndex] = temp;

            RebuildUI();
            PriorityChanged?.Invoke(this, EventArgs.Empty);
        }
    }

}

