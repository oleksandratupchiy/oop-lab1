using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetApp11.Core;

namespace SpreadsheetApp11.UI
{
    public class MainForm : Form
    {
        private readonly DataGridView _grid = new DataGridView();
        private readonly SpreadsheetService _service = new SpreadsheetService();
        private readonly ToolStrip _toolbar = new ToolStrip();
        private readonly ToolStripButton _btnToggleMode = new ToolStripButton("Відображення: Значення");
        private readonly ToolStripButton _btnInsRow = new ToolStripButton("+Рядок");
        private readonly ToolStripButton _btnDelRow = new ToolStripButton("-Рядок");
        private readonly ToolStripButton _btnInsCol = new ToolStripButton("+Стовпець");
        private readonly ToolStripButton _btnDelCol = new ToolStripButton("-Стовпець");
        private readonly ToolStripButton _btnSave = new ToolStripButton("Зберегти");
        private readonly ToolStripButton _btnLoad = new ToolStripButton("Відкрити");
        private readonly ToolStripButton _btnSaveDrive = new ToolStripButton("Зберегти в Drive");
        private readonly ToolStripButton _btnOpenDrive = new ToolStripButton("Відкрити з Drive");
        private readonly ToolStripButton _btnHelp = new ToolStripButton("Довідка");

        private Core.DisplayMode _mode = Core.DisplayMode.Values;

        // Поле для відстеження змін
        private bool _isModified = false;

        private const int InitialRows = 20;
        private const int InitialCols = 10;

        public MainForm()
        {
            Text = "Електронні таблиці Lab 11";
            Width = 1100; Height = 700;

            _toolbar.Items.AddRange(new ToolStripItem[]
            {
                _btnToggleMode, new ToolStripSeparator(),
                _btnInsRow, _btnDelRow, _btnInsCol, _btnDelCol,
                new ToolStripSeparator(),
                _btnSave, _btnLoad, new ToolStripSeparator(),
                _btnSaveDrive, _btnOpenDrive, new ToolStripSeparator(),
                _btnHelp
            });
            _toolbar.GripStyle = ToolStripGripStyle.Hidden;

            _grid.Dock = DockStyle.Fill;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.RowHeadersWidth = 60;
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _grid.Font = new Font("Segoe UI", 10);

            Controls.Add(_grid);
            Controls.Add(_toolbar);
            _toolbar.Dock = DockStyle.Top;

            _service.CreateEmpty(InitialRows, InitialCols);
            BuildGrid();

            _grid.CellEndEdit += OnCellEndEdit;
            _grid.CellBeginEdit += OnCellBeginEdit;
            _grid.CellFormatting += OnCellFormatting;

            _btnToggleMode.Click += (s, e) => ToggleMode();
            _btnInsRow.Click += (s, e) => InsertRowAtSelection();
            _btnDelRow.Click += (s, e) => DeleteRowAtSelection();
            _btnInsCol.Click += (s, e) => InsertColAtSelection();
            _btnDelCol.Click += (s, e) => DeleteColAtSelection();
            _btnSave.Click += (s, e) => SaveFile();
            _btnLoad.Click += (s, e) => LoadFile();
            _btnSaveDrive.Click += (s, e) => SaveToDrive();
            _btnOpenDrive.Click += (s, e) => LoadFromDrive();
            _btnHelp.Click += (s, e) => ShowHelpInfo();

            this.FormClosing += MainForm_FormClosing;
        }

        private void SetModified(bool modified)
        {
            _isModified = modified;
            string title = "Електронні таблиці Lab 11";
            Text = modified ? $"{title} *" : title;
        }

        private bool PromptSaveIfModified()
        {
            if (!_isModified) return true;

            var result = MessageBox.Show(
                "У вас є незбережені зміни. Бажаєте зберегти їх перед продовженням?",
                "Незбережені зміни",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Cancel)
            {
                return false;
            }

            if (result == DialogResult.Yes)
            {
                return SaveFile();
            }

            return true;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!PromptSaveIfModified())
            {
                e.Cancel = true;
            }
        }

        private void BuildGrid()
        {
            _grid.Columns.Clear();
            _grid.Rows.Clear();

            for (int c = 0; c < _service.ColCount; c++)
            {
                _grid.Columns.Add($"Col{c}", GetColumnName(c));
            }

            _grid.Rows.Add(_service.RowCount);

            for (int r = 0; r < _service.RowCount; r++)
            {
                _grid.Rows[r].HeaderCell.Value = (r + 1).ToString();
            }

            for (int r = 0; r < _service.RowCount; r++)
            {
                for (int c = 0; c < _service.ColCount; c++)
                {
                    _grid[c, r].Value = _service.GetCellDisplay(r, c, _mode);
                }
            }
        }

        private string GetColumnName(int colIndex)
        {
            string name = "";
            int index = colIndex;
            while (index >= 0)
            {
                name = (char)('A' + (index % 26)) + name;
                index = (index / 26) - 1;
            }
            return name;
        }

        private void OnCellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int r = e.RowIndex;
            int c = e.ColumnIndex;
            string newValue = _grid[c, r].Value?.ToString() ?? "";

            var result = _service.SetCell(r, c, newValue);
            if (!result.Ok)
            {
                MessageBox.Show(result.Error, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _grid[c, r].Value = newValue;
            }
            else
            {
                SetModified(true);
                RefreshGrid();
            }
        }

        private void OnCellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            int r = e.RowIndex;
            int c = e.ColumnIndex;
            _grid[c, r].Value = _service.GetCellRaw(r, c);
        }

        private void OnCellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (_mode == Core.DisplayMode.Formulas)
            {
                e.CellStyle.BackColor = Color.AliceBlue;
                e.CellStyle.ForeColor = Color.DarkSlateBlue;
                e.CellStyle.SelectionBackColor = Color.SteelBlue;
                e.CellStyle.SelectionForeColor = Color.White;
                e.CellStyle.Font = new Font(_grid.Font, FontStyle.Italic);
            }
            else
            {
                e.CellStyle.BackColor = _grid.DefaultCellStyle.BackColor;
                e.CellStyle.ForeColor = _grid.DefaultCellStyle.ForeColor;
                e.CellStyle.SelectionBackColor = _grid.DefaultCellStyle.SelectionBackColor;
                e.CellStyle.SelectionForeColor = _grid.DefaultCellStyle.SelectionForeColor;
                e.CellStyle.Font = _grid.Font;
            }
        }

        private void ToggleMode()
        {
            _mode = _mode == Core.DisplayMode.Values ? Core.DisplayMode.Formulas : Core.DisplayMode.Values;
            _btnToggleMode.Text = _mode == Core.DisplayMode.Values ? "Відображення: Значення" : "Відображення: Формули";
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            try
            {
                for (int r = 0; r < _grid.Rows.Count; r++)
                {
                    for (int c = 0; c < _grid.Columns.Count; c++)
                    {
                        string displayValue = _service.GetCellDisplay(r, c, _mode) ?? "";
                        _grid[c, r].Value = displayValue;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing grid: {ex.Message}");
            }
        }

        private void InsertRowAtSelection()
        {
            int rowIndex = _grid.CurrentCell?.RowIndex ?? _service.RowCount;
            _service.InsertRows(rowIndex, 1);
            SetModified(true);
            BuildGrid();
        }

        private void DeleteRowAtSelection()
        {
            int rowIndex = _grid.CurrentCell?.RowIndex ?? -1;
            if (rowIndex >= 0 && rowIndex < _service.RowCount)
            {
                _service.DeleteRows(rowIndex, 1);
                SetModified(true);
                BuildGrid();
            }
        }

        private void InsertColAtSelection()
        {
            int colIndex = _grid.CurrentCell?.ColumnIndex ?? _service.ColCount;
            _service.InsertColumns(colIndex, 1);
            SetModified(true);
            BuildGrid();
        }

        private void DeleteColAtSelection()
        {
            int colIndex = _grid.CurrentCell?.ColumnIndex ?? -1;
            if (colIndex >= 0 && colIndex < _service.ColCount)
            {
                _service.DeleteColumns(colIndex, 1);
                SetModified(true);
                BuildGrid();
            }
        }

        private bool SaveFile()
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "JSON файли (*.json)|*.json|Всі файли (*.*)|*.*";
                dialog.DefaultExt = "json";
                dialog.Title = "Зберегти таблицю";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _service.SaveLocal(dialog.FileName);
                        SetModified(false);
                        MessageBox.Show($"Файл успішно збережено!\n{dialog.FileName}",
                            "Успіх", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Помилка збереження: {ex.Message}",
                            "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }
            return false;
        }

        private void LoadFile()
        {
            if (!PromptSaveIfModified()) return;

            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "JSON файли (*.json)|*.json|Всі файли (*.*)|*.*";
                dialog.DefaultExt = "json";
                dialog.Title = "Відкрити таблицю";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _service.LoadLocal(dialog.FileName);
                        BuildGrid();
                        SetModified(false);
                        MessageBox.Show($"Файл успішно завантажено!\n{dialog.FileName}",
                            "Успіх", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Помилка завантаження: {ex.Message}",
                            "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private async void SaveToDrive()
        {
            string fileName = Microsoft.VisualBasic.Interaction.InputBox(
                "Введіть назву файлу для Google Drive:",
                "Зберегти в Google Drive",
                $"spreadsheet_{DateTime.Now:yyyyMMdd_HHmmss}.json");

            if (!string.IsNullOrEmpty(fileName))
            {
                try
                {
                    _btnSaveDrive.Enabled = false;
                    _btnSaveDrive.Text = "Збереження...";

                    await Task.Run(() => _service.SaveToDrive(fileName));

                    SetModified(false);

                    MessageBox.Show($"Файл '{fileName}' успішно збережено в Google Drive!",
                        "Успіх", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Помилка збереження в Google Drive: {ex.Message}",
                        "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    _btnSaveDrive.Enabled = true;
                    _btnSaveDrive.Text = "Зберегти в Drive";
                }
            }
        }

        private async void LoadFromDrive()
        {
            if (!PromptSaveIfModified()) return;

            string fileName = Microsoft.VisualBasic.Interaction.InputBox(
                "Введіть назву файлу або ID з Google Drive:",
                "Відкрити з Google Drive",
                "spreadsheet.json");

            if (!string.IsNullOrEmpty(fileName))
            {
                try
                {
                    _btnOpenDrive.Enabled = false;
                    _btnOpenDrive.Text = "Завантаження...";

                    await Task.Run(() => _service.LoadFromDrive(fileName));

                    BuildGrid();
                    SetModified(false);
                    MessageBox.Show($"Файл '{fileName}' успішно завантажено з Google Drive!",
                        "Успіх", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Помилка завантаження з Google Drive: {ex.Message}",
                        "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    _btnOpenDrive.Enabled = true;
                    _btnOpenDrive.Text = "Відкрити з Drive";
                }
            }
        }

        private void ShowHelpInfo()
        {
            string helpText = $@"Електронні таблиці - Лабораторна робота 1*

Розробник: Тупчій Олександра
Група: K-25
Варіант 11";

            MessageBox.Show(helpText, "Довідка",
                           MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}