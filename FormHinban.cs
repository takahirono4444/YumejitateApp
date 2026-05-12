using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows.Forms;

namespace YumejitateApp
{
    /// <summary>
    /// AB品番検索フォーム (VB6: form_hinban.frm の移植)
    /// </summary>
    public class FormHinban : Form
    {
        // ----------------------------------------------------------------
        // コントロール
        // ----------------------------------------------------------------
        private ComboBox _cmbItem;
        private ComboBox _cmbStoneShape;
        private ComboBox _cmbSetStyle;
        private ComboBox _cmbJigane;

        private ComboBox _cmbHinban1;
        private ComboBox _cmbHinban2;
        private ComboBox _cmbHinban3;
        private ComboBox _cmbHinban4;
        private ComboBox _cmbHinban5;
        private ComboBox _cmbHinban6;
        private ComboBox _cmbHinban7;

        private ComboBox _cmbSizeUeJu;
        private ComboBox _cmbSizeUeIchi;
        private ComboBox _cmbSizeUeKo;
        private ComboBox _cmbSizeShitaJu;
        private ComboBox _cmbSizeShitaIchi;
        private ComboBox _cmbSizeShitaKo;

        private ComboBox _cmbRingsizeJu;
        private ComboBox _cmbRingsizeIchi;
        private ComboBox _cmbGrade;

        private Button _btnSearch;
        private Button _btnBack;

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------
        public FormHinban()
        {
            InitializeComponent();
        }

        // ----------------------------------------------------------------
        // InitializeComponent
        // ----------------------------------------------------------------
        private void InitializeComponent()
        {
            this.SuspendLayout();

            this.Text = "夢仕立て - 品番検索";
            this.BackColor = System.Drawing.Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.AutoScroll = true;
            this.MinimumSize = new System.Drawing.Size(1000, 650);

            var fntNormal = new System.Drawing.Font("ＭＳ Ｐゴシック", 14.25f, System.Drawing.FontStyle.Bold);
            var fntTitle = new System.Drawing.Font("ＭＳ Ｐゴシック", 18f, System.Drawing.FontStyle.Italic);
            var fntBtn = new System.Drawing.Font("ＭＳ Ｐゴシック", 15.75f, System.Drawing.FontStyle.Bold);
            var cmbSz = new System.Drawing.Size(52, 27);

            // タイトル
            var lblTitle = MkLbl("品番検索", fntTitle, new System.Drawing.Point(432, 16), true);

            // ================================================================
            // GroupBox
            // ================================================================
            var grp = new GroupBox
            {
                Text = "検索条件",
                Font = fntNormal,
                Location = new System.Drawing.Point(16, 56),
                Size = new System.Drawing.Size(1050, 460),
                BackColor = System.Drawing.Color.Transparent,
            };

            int lblX = 16;
            int cmbX = 230;   // ラベル列右端
            int rowH = 44;

            // --- アイテム ---
            int row0 = 16;
            grp.Controls.Add(MkLbl("アイテム", fntNormal, new System.Drawing.Point(lblX, row0 + 4), false));
            _cmbItem = MkCmb(fntNormal, new System.Drawing.Point(cmbX, row0), new System.Drawing.Size(280, 27));
            _cmbItem.SelectedIndexChanged += CmbItem_SelectedIndexChanged;
            grp.Controls.Add(_cmbItem);

            // --- 中石形状 ---
            int row1 = row0 + rowH;
            grp.Controls.Add(MkLbl("中石形状", fntNormal, new System.Drawing.Point(lblX, row1 + 4), false));
            _cmbStoneShape = MkCmb(fntNormal, new System.Drawing.Point(cmbX, row1), new System.Drawing.Size(280, 27));
            grp.Controls.Add(_cmbStoneShape);

            // --- 中石セット ---
            int row2 = row1 + rowH;
            grp.Controls.Add(MkLbl("中石セット", fntNormal, new System.Drawing.Point(lblX, row2 + 4), false));
            _cmbSetStyle = MkCmb(fntNormal, new System.Drawing.Point(cmbX, row2), new System.Drawing.Size(280, 27));
            grp.Controls.Add(_cmbSetStyle);

            // --- 地金 ---
            int row3 = row2 + rowH;
            grp.Controls.Add(MkLbl("地　金", fntNormal, new System.Drawing.Point(lblX, row3 + 4), false));
            _cmbJigane = MkCmb(fntNormal, new System.Drawing.Point(cmbX, row3), new System.Drawing.Size(280, 27));
            grp.Controls.Add(_cmbJigane);

            // --- 品番（7桁） ---
            int row4 = row3 + rowH;
            grp.Controls.Add(MkLbl("品番", fntNormal, new System.Drawing.Point(lblX, row4 + 4), false));
            int hx = cmbX;
            _cmbHinban1 = MkCmb(fntNormal, new System.Drawing.Point(hx, row4), cmbSz); grp.Controls.Add(_cmbHinban1);
            _cmbHinban2 = MkCmb(fntNormal, new System.Drawing.Point(hx + 56, row4), cmbSz); grp.Controls.Add(_cmbHinban2);
            _cmbHinban3 = MkCmb(fntNormal, new System.Drawing.Point(hx + 112, row4), cmbSz); grp.Controls.Add(_cmbHinban3);
            _cmbHinban4 = MkCmb(fntNormal, new System.Drawing.Point(hx + 168, row4), cmbSz); grp.Controls.Add(_cmbHinban4);
            _cmbHinban5 = MkCmb(fntNormal, new System.Drawing.Point(hx + 224, row4), cmbSz); grp.Controls.Add(_cmbHinban5);
            _cmbHinban6 = MkCmb(fntNormal, new System.Drawing.Point(hx + 280, row4), cmbSz); grp.Controls.Add(_cmbHinban6);
            _cmbHinban7 = MkCmb(fntNormal, new System.Drawing.Point(hx + 336, row4), cmbSz); grp.Controls.Add(_cmbHinban7);

            // --- サイズ (xx.x mm × xx.x mm) ---
            // ※ 「．」ラベルが大きいため各コンボを十分に右へずらす
            int row5 = row4 + rowH;
            grp.Controls.Add(MkLbl("サイズ", fntNormal, new System.Drawing.Point(lblX, row5 + 4), false));
            int sx = cmbX;
            _cmbSizeUeJu = MkCmb(fntNormal, new System.Drawing.Point(sx, row5), cmbSz); grp.Controls.Add(_cmbSizeUeJu);
            _cmbSizeUeIchi = MkCmb(fntNormal, new System.Drawing.Point(sx + 56, row5), cmbSz); grp.Controls.Add(_cmbSizeUeIchi);
            grp.Controls.Add(MkLbl("．", fntNormal, new System.Drawing.Point(sx + 112, row5 + 3), true));
            _cmbSizeUeKo = MkCmb(fntNormal, new System.Drawing.Point(sx + 130, row5), cmbSz); grp.Controls.Add(_cmbSizeUeKo);
            grp.Controls.Add(MkLbl("mm×", fntNormal, new System.Drawing.Point(sx + 186, row5 + 3), true));
            _cmbSizeShitaJu = MkCmb(fntNormal, new System.Drawing.Point(sx + 250, row5), cmbSz); grp.Controls.Add(_cmbSizeShitaJu);
            _cmbSizeShitaIchi = MkCmb(fntNormal, new System.Drawing.Point(sx + 306, row5), cmbSz); grp.Controls.Add(_cmbSizeShitaIchi);
            grp.Controls.Add(MkLbl("．", fntNormal, new System.Drawing.Point(sx + 362, row5 + 3), true));
            _cmbSizeShitaKo = MkCmb(fntNormal, new System.Drawing.Point(sx + 380, row5), cmbSz); grp.Controls.Add(_cmbSizeShitaKo);
            grp.Controls.Add(MkLbl("mm", fntNormal, new System.Drawing.Point(sx + 436, row5 + 3), true));

            // --- リングサイズ / グレード ---
            // ※ ラベルが長いためコンボを十分に右にずらす
            int row6 = row5 + rowH;
            grp.Controls.Add(MkLbl("リングサイズ\nまたはグレード", fntNormal,
                new System.Drawing.Point(lblX, row6), false));
            grp.Controls.Add(MkLbl("＃", fntNormal, new System.Drawing.Point(cmbX, row6 + 4), true));
            _cmbRingsizeJu = MkCmb(fntNormal, new System.Drawing.Point(cmbX + 24, row6), cmbSz); grp.Controls.Add(_cmbRingsizeJu);
            _cmbRingsizeIchi = MkCmb(fntNormal, new System.Drawing.Point(cmbX + 80, row6), cmbSz); grp.Controls.Add(_cmbRingsizeIchi);
            grp.Controls.Add(MkLbl("グレード", fntNormal, new System.Drawing.Point(cmbX + 145, row6 + 4), true));
            _cmbGrade = MkCmb(fntNormal, new System.Drawing.Point(cmbX + 245, row6), new System.Drawing.Size(72, 27));
            grp.Controls.Add(_cmbGrade);

            grp.Height = row6 + rowH + 20;

            // ---- ボタン ----
            _btnSearch = new Button
            {
                Text = "品番検索",
                Font = fntBtn,
                Location = new System.Drawing.Point(380, grp.Bottom + 20),
                Size = new System.Drawing.Size(160, 57),
                BackColor = System.Drawing.Color.FromArgb(192, 224, 255),
                Cursor = Cursors.Hand,
            };
            _btnSearch.Click += BtnSearch_Click;

            _btnBack = new Button
            {
                Text = "メニュー",
                Font = fntBtn,
                Location = new System.Drawing.Point(680, grp.Bottom + 20),
                Size = new System.Drawing.Size(160, 57),
                BackColor = System.Drawing.Color.FromArgb(255, 192, 192),
                Cursor = Cursors.Hand,
            };
            _btnBack.Click += BtnBack_Click;

            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                lblTitle, grp, _btnSearch, _btnBack,
            });

            this.ResumeLayout(false);
        }

        // ================================================================
        // ヘルパー
        // ================================================================
        private static System.Windows.Forms.Label MkLbl(
            string text, System.Drawing.Font font,
            System.Drawing.Point loc, bool autoSize)
        {
            return new System.Windows.Forms.Label
            {
                Text = text,
                Font = font,
                BackColor = System.Drawing.Color.Transparent,
                AutoSize = autoSize,
                Location = loc,
            };
        }

        private static ComboBox MkCmb(
            System.Drawing.Font font,
            System.Drawing.Point loc,
            System.Drawing.Size size)
        {
            return new ComboBox
            {
                Font = font,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = loc,
                Size = size,
            };
        }

        // ================================================================
        // フォームロード
        // ================================================================
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.WindowState = FormWindowState.Maximized;
            InitControl();
        }

        // ================================================================
        // コントロール初期化 (VB6: Init_Control)
        // ================================================================
        private void InitControl()
        {
            // アイテム
            _cmbItem.Items.AddRange(new object[] {
                "リング", "ブローチ", "ペンダント", "タイタック",
                "ピアス", "イアリング", "その他（バチカン）"
            });
            _cmbItem.SelectedIndex = 0;

            // 中石形状
            _cmbStoneShape.Items.AddRange(new object[] {
                "ラウンド（円）", "オーバル（楕円）", "ボール（球）",
                "エメラルド（四角）", "マーキース", "ドロップ"
            });
            _cmbStoneShape.SelectedIndex = 0;

            // 中石セットスタイル
            _cmbSetStyle.Items.AddRange(new object[] { "爪留め", "爪無し（レール留め等）" });
            _cmbSetStyle.SelectedIndex = 0;

            // 地金
            _cmbJigane.Items.AddRange(new object[] {
                "プラチナ", "Ｋ１８ＹＧ", "Ｋ１８ＷＧ", "コンビ", "シルバー", "Ｋ１０"
            });
            _cmbJigane.SelectedIndex = 0;

            // 品番一,二: A/B/C/I/J/K/N/V
            string[] alpha = { "A", "B", "C", "I", "J", "K", "N", "V" };
            _cmbHinban1.Items.AddRange(alpha); _cmbHinban1.SelectedIndex = 0;
            _cmbHinban2.Items.AddRange(alpha); _cmbHinban2.SelectedIndex = 0;

            // 品番三〜七: 0-9
            string[] digits = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            _cmbHinban3.Items.AddRange(digits); _cmbHinban3.SelectedIndex = 0;
            _cmbHinban4.Items.AddRange(digits); _cmbHinban4.SelectedIndex = 0;
            _cmbHinban5.Items.AddRange(digits); _cmbHinban5.SelectedIndex = 0;
            _cmbHinban6.Items.AddRange(digits); _cmbHinban6.SelectedIndex = 0;
            _cmbHinban7.Items.AddRange(digits); _cmbHinban7.SelectedIndex = 0;

            // サイズ: 0-9
            foreach (var cmb in new[] {
                _cmbSizeUeJu, _cmbSizeUeIchi, _cmbSizeUeKo,
                _cmbSizeShitaJu, _cmbSizeShitaIchi, _cmbSizeShitaKo })
            {
                cmb.Items.AddRange(digits);
                cmb.SelectedIndex = 0;
            }

            // リングサイズ: 0-9
            _cmbRingsizeJu.Items.AddRange(digits); _cmbRingsizeJu.SelectedIndex = 0;
            _cmbRingsizeIchi.Items.AddRange(digits); _cmbRingsizeIchi.SelectedIndex = 0;

            // グレード
            _cmbGrade.Items.AddRange(new object[] { "A", "B", "C" });
            _cmbGrade.SelectedIndex = 0;
            _cmbGrade.Enabled = false;

            UpdateItemDependent();
        }

        // ================================================================
        // アイテム変更 (VB6: アイテム_Click)
        // ================================================================
        private void CmbItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateItemDependent();
        }

        private void UpdateItemDependent()
        {
            string item = _cmbItem.Text;
            bool isRing = (item == "リング");
            bool isVatican = (item == "その他（バチカン）");

            _cmbRingsizeJu.Enabled = isRing;
            _cmbRingsizeIchi.Enabled = isRing;
            _cmbGrade.Enabled = !isRing;

            bool sizeEnabled = !isVatican;
            _cmbStoneShape.Enabled = sizeEnabled;
            _cmbSetStyle.Enabled = sizeEnabled;
            _cmbSizeUeJu.Enabled = sizeEnabled;
            _cmbSizeUeIchi.Enabled = sizeEnabled;
            _cmbSizeUeKo.Enabled = sizeEnabled;
            _cmbSizeShitaJu.Enabled = sizeEnabled;
            _cmbSizeShitaIchi.Enabled = sizeEnabled;
            _cmbSizeShitaKo.Enabled = sizeEnabled;
        }

        // ================================================================
        // 入力チェック (VB6: check_control)
        // ================================================================
        private bool CheckControl()
        {
            string item = _cmbItem.Text;

            if (item != "その他（バチカン）")
            {
                bool ueZero = _cmbSizeUeJu.Text == "0"
                           && _cmbSizeUeIchi.Text == "0"
                           && _cmbSizeUeKo.Text == "0";
                bool shitaZero = _cmbSizeShitaJu.Text == "0"
                              && _cmbSizeShitaIchi.Text == "0"
                              && _cmbSizeShitaKo.Text == "0";
                if (ueZero || shitaZero)
                {
                    MessageBox.Show("「サイズ」を入力して下さい。", "入力チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            if (item == "リング")
            {
                if (_cmbRingsizeJu.Text == "0" && _cmbRingsizeIchi.Text == "0")
                {
                    MessageBox.Show(
                        "「アイテム」を「リング」で検索する場合、＃（サイズ）を入力して下さい。",
                        "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            return true;
        }

        // ================================================================
        // 品番検索ボタン (VB6: btn_search_Click)
        // ================================================================
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            if (!CheckControl()) return;

            MessageBox.Show(
                "数千点以上のデータからご希望のデザインをお探しします。",
                "検索開始", MessageBoxButtons.OK, MessageBoxIcon.Information);

            AppState.Db.ExecuteNonQuery("DELETE * FROM [ワーク検索テーブル]");
            AppState.Db.ExecuteNonQuery("DELETE * FROM [検索結果テーブル]");

            string strsql = BuildSearchSql();

            DataTable dtResult = AppState.Db.ExecuteQuery(strsql);
            if (dtResult == null || dtResult.Rows.Count == 0)
            {
                MessageBox.Show(
                    "誠に申し訳ございませんが\nご希望のデザインはございません。",
                    "「固定情報」該当データ無し", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            AppState.Db.ExecuteNonQuery("INSERT INTO [ワーク検索テーブル] " + strsql);

            RemoveDuplicateHinban();

            int searchCnt = SelectRandomResults();

            MessageBox.Show(
                $"ご希望のデザインは {searchCnt} 種類ございます。",
                "検索結果件数表示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            if (searchCnt == 0) return;

            this.Hide();
            var movieForm = new FormMovie2();
            movieForm.FormClosed += (s2, e2) => this.Close();
            movieForm.Show();
        }

        // ================================================================
        // 検索SQL構築
        // ※ LIKE条件は .NET OleDb に合わせて % ワイルドカードを使用
        // ※ D()関数は整数に丸めて出力
        // ================================================================
        private string BuildSearchSql()
        {
            var sb = new System.Text.StringBuilder();
            string item = _cmbItem.Text;

            sb.Append("SELECT * FROM [検索テーブル] WHERE ");

            // アイテム → b LIKE '%N%'
            string itemCode;
            switch (item)
            {
                case "リング": itemCode = "6"; break;
                case "ブローチ": itemCode = "8"; break;
                case "ペンダント": itemCode = "7"; break;
                case "タイタック": itemCode = "1"; break;
                case "ピアス": itemCode = "2"; break;
                case "イアリング": itemCode = "3"; break;
                case "その他（バチカン）": itemCode = "0"; break;
                default: itemCode = "6"; break;
            }
            sb.Append($"[b] LIKE '%{itemCode}%' ");

            if (item == "その他（バチカン）")
            {
                sb.Append("AND [l] = 'XXX' AND [m] = 'XXX' ");
            }
            else
            {
                // 中石形状 → c LIKE '%N%'
                string shapeCode;
                switch (_cmbStoneShape.Text)
                {
                    case "ラウンド（円）": shapeCode = "1"; break;
                    case "オーバル（楕円）": shapeCode = "2"; break;
                    case "ボール（球）": shapeCode = "3"; break;
                    case "エメラルド（四角）": shapeCode = "4"; break;
                    case "マーキース": shapeCode = "5"; break;
                    case "ドロップ": shapeCode = "6"; break;
                    default: shapeCode = ""; break;
                }
                if (shapeCode != "")
                    sb.Append($"AND [c] LIKE '%{shapeCode}%' ");

                // 中石セットスタイル → h LIKE '%N%'
                string styleCode;
                switch (_cmbSetStyle.Text)
                {
                    case "爪留め": styleCode = "1"; break;
                    case "爪無し（レール留め等）": styleCode = "2"; break;
                    default: styleCode = ""; break;
                }
                if (styleCode != "")
                    sb.Append($"AND [h] LIKE '%{styleCode}%' ");
            }

            // 地金 → d LIKE '%N%'
            string jiganeCode;
            switch (_cmbJigane.Text)
            {
                case "プラチナ": jiganeCode = "1"; break;
                case "Ｋ１８ＹＧ": jiganeCode = "2"; break;
                case "Ｋ１８ＷＧ": jiganeCode = "3"; break;
                case "コンビ": jiganeCode = "4"; break;
                case "シルバー": jiganeCode = "5"; break;
                case "Ｋ１０": jiganeCode = "6"; break;
                default: jiganeCode = ""; break;
            }
            if (jiganeCode != "")
                sb.Append($"AND [d] LIKE '%{jiganeCode}%' ");

            // 品番 → a = '品番7桁'
            string hinban = _cmbHinban1.Text + _cmbHinban2.Text
                          + _cmbHinban3.Text + _cmbHinban4.Text
                          + _cmbHinban5.Text + _cmbHinban6.Text
                          + _cmbHinban7.Text;
            sb.Append($"AND [a] = '{hinban}' ");

            // サイズ条件（バチカン以外）
            if (item != "その他（バチカン）")
            {
                double sizeA = ToDigit(_cmbSizeUeJu) * 100
                             + ToDigit(_cmbSizeUeIchi) * 10
                             + ToDigit(_cmbSizeUeKo);
                double sizeB = ToDigit(_cmbSizeShitaJu) * 100
                             + ToDigit(_cmbSizeShitaIchi) * 10
                             + ToDigit(_cmbSizeShitaKo);

                double bSize = Math.Max(sizeA, sizeB);
                double sSize = Math.Min(sizeA, sizeB);
                double tSize = sizeA + sizeB;

                if (_cmbStoneShape.Text == "ボール（球）")
                    sb.Append(BuildSizeConditionBall(bSize, sSize, tSize));
                else if (_cmbSetStyle.Text == "爪無し（レール留め等）")
                    sb.Append(BuildSizeCondition4(bSize, sSize, tSize));
                else
                    sb.Append(BuildSizeCondition3(bSize, sSize, tSize));
            }

            // リングサイズ → CLng([w]) / CLng([x]) で数値比較
            if (item == "リング")
            {
                int mSize = ToDigit(_cmbRingsizeJu) * 10 + ToDigit(_cmbRingsizeIchi);
                sb.Append($"AND ({mSize} BETWEEN CLng([w]) AND CLng([x])) ");
            }

            return sb.ToString();
        }

        // ================================================================
        // サイズ条件: 爪留め（式１・式３）AB2分岐あり
        // ================================================================
        private static string BuildSizeCondition3(double bSize, double sSize, double tSize)
        {
            return "AND (("
                + "NOT ([a] LIKE 'AB2%') "
                + $"AND [bigsize] >= {D(bSize / 1.18)} "
                + $"AND [bigsize] <= {D(bSize / 0.75)} "
                + $"AND [smallsize] >= {D(sSize / 1.18)} "
                + $"AND [smallsize] <= {D(sSize / 0.75)} "
                + $"AND [totalsize] >= {D(tSize - 5)} "
                + $"AND [totalsize] <= {D(tSize + 10)}) "
                + "OR ([a] LIKE 'AB2%' "
                + $"AND [bigsize] >= {D(bSize / 1.062)} "
                + $"AND [bigsize] <= {D(bSize / 0.675)} "
                + $"AND [smallsize] >= {D(sSize / 1.062)} "
                + $"AND [smallsize] <= {D(sSize / 0.675)} "
                + $"AND [totalsize] >= {D((tSize - 5) / 0.9)} "
                + $"AND [totalsize] <= {D((tSize + 10) / 0.9)})) ";
        }

        // ================================================================
        // サイズ条件: 爪無し（式２・式４）AB2分岐あり
        // ================================================================
        private static string BuildSizeCondition4(double bSize, double sSize, double tSize)
        {
            return "AND (("
                + "NOT ([a] LIKE 'AB2%') "
                + $"AND [bigsize] >= {D(bSize / 1.18)} "
                + $"AND [bigsize] <= {D(bSize / 0.75)} "
                + $"AND [smallsize] >= {D(sSize / 1.18)} "
                + $"AND [smallsize] <= {D(sSize / 0.75)} "
                + $"AND [totalsize] >= {D(tSize)} "
                + $"AND [totalsize] <= {D(tSize + 10)}) "
                + "OR ([a] LIKE 'AB2%' "
                + $"AND [bigsize] >= {D(bSize / 1.062)} "
                + $"AND [bigsize] <= {D(bSize / 0.675)} "
                + $"AND [smallsize] >= {D(sSize / 1.062)} "
                + $"AND [smallsize] <= {D(sSize / 0.675)} "
                + $"AND [totalsize] >= {D(tSize / 0.9)} "
                + $"AND [totalsize] <= {D((tSize + 10) / 0.9)})) ";
        }

        // ================================================================
        // サイズ条件: ボール（球）AB2分岐なし
        // ================================================================
        private static string BuildSizeConditionBall(double bSize, double sSize, double tSize)
        {
            return $"AND [bigsize] >= {D(bSize / 1.18)} "
                 + $"AND [bigsize] <= {D(bSize / 0.75)} "
                 + $"AND [smallsize] >= {D(sSize / 1.18)} "
                 + $"AND [smallsize] <= {D(sSize / 0.75)} "
                 + $"AND [totalsize] >= {D(tSize - 5)} "
                 + $"AND [totalsize] <= {D(tSize + 10)} ";
        }

        // ================================================================
        // 品番重複排除
        // ================================================================
        private void RemoveDuplicateHinban()
        {
            DataTable dt = AppState.Db.ExecuteQuery(
                "SELECT * FROM [ワーク検索テーブル] ORDER BY [index]");
            if (dt == null || dt.Rows.Count == 0) return;

            string prevHinban = "";
            foreach (DataRow row in dt.Rows)
            {
                string cur = row["a"].ToString();
                if (string.Compare(cur, prevHinban, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    AppState.Db.ExecuteNonQuery(
                        $"DELETE FROM [ワーク検索テーブル] WHERE [index] = {row["index"]}");
                }
                else
                {
                    prevHinban = cur;
                }
            }
        }

        // ================================================================
        // ランダムに最大7件を検索結果テーブルへ
        // ================================================================
        private int SelectRandomResults()
        {
            DataTable dtWork = AppState.Db.ExecuteQuery(
                "SELECT * FROM [ワーク検索テーブル] ORDER BY [index]");
            if (dtWork == null || dtWork.Rows.Count == 0) return 0;

            var indexList = new List<object>();
            foreach (DataRow row in dtWork.Rows)
                indexList.Add(row["index"]);

            var rnd = new Random();
            for (int i = indexList.Count - 1; i > 0; i--)
            {
                int j = rnd.Next(0, i + 1);
                var tmp = indexList[i];
                indexList[i] = indexList[j];
                indexList[j] = tmp;
            }

            int count = 0;
            int maxCount = Math.Min(7, indexList.Count);
            for (int i = 0; i < maxCount; i++)
            {
                AppState.Db.ExecuteNonQuery(
                    "INSERT INTO [検索結果テーブル] "
                    + "SELECT * FROM [ワーク検索テーブル] "
                    + $"WHERE [index] = {indexList[i]}");
                count++;
            }
            return count;
        }

        // ================================================================
        // メニューボタン (VB6: btn_back_Click)
        // ================================================================
        private void BtnBack_Click(object sender, EventArgs e)
        {
            AppState.FlagHinban = false;
            foreach (Form f in Application.OpenForms)
            {
                if (f is FormMenu) { f.Show(); break; }
            }
            this.Hide();
        }

        // ================================================================
        // ユーティリティ
        // ================================================================
        private static int ToDigit(ComboBox cmb)
        {
            return int.TryParse(cmb.Text, out int v) ? v : 0;
        }

        /// <summary>
        /// double を整数に丸めてSQL文字列化
        /// ※ Access SQLは小数桁数が多いと正しく動作しないため整数化する
        /// </summary>
        private static string D(double v) =>
            ((long)Math.Round(v)).ToString();
    }
}