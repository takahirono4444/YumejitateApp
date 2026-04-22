using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace YumejitateApp
{
    /// <summary>
    /// オーダーガイド見積り印刷フォーム (VB6: form_print_tedukuri.frm の移植)
    /// オーダーメイドテーブルから作業書番号を選択し、データ表示・印刷・削除を行う。
    ///
    /// VB6の重要な設計パターン:
    ///   Label コントロールの Name プロパティ = オーダーメイドテーブルのフィールド名
    ///   → Disp_Order_Data で全 Label を走査し、o.Caption = recTB.Fields(o.Name) でデータを設定
    ///   → C# では Dictionary&lt;string, Label&gt;(_dataLabels) で同様の動作を実現
    /// </summary>
    public class FormPrintTedukuri : Form
    {
        // ----------------------------------------------------------------
        // 操作コントロール
        // ----------------------------------------------------------------
        private ComboBox _cmbSave;       // VB6: 保存 (Style=2 DropDownList, 作業書番号選択)
        private Button _btnDispData;   // VB6: btn_disp_data  データ表示
        private Button _btnDeleteData; // VB6: btn_delete_data データ削除
        private Button _btnPrint;      // VB6: btn_print      印刷
        private Button _btnMenu;       // VB6: btn_menu       メニューへ戻る

        // メインスクロールパネル
        private Panel _pnlMain;

        // ----------------------------------------------------------------
        // データ表示ラベル辞書
        //   key   = オーダーメイドテーブルのフィールド名 (= VB6 ラベル名)
        //   value = 対応する Label コントロール
        // ----------------------------------------------------------------
        private readonly Dictionary<string, Label> _dataLabels =
            new Dictionary<string, Label>(StringComparer.Ordinal);

        // ----------------------------------------------------------------
        // 背景色定数
        //   ColorEmpty  = &H80000005 (システム Window 色) ← データ空のとき
        //   ColorFilled = &HFFC0C0  (薄ピンク)            ← データありのとき
        // ----------------------------------------------------------------
        private static readonly Color ColorEmpty = SystemColors.Window;
        private static readonly Color ColorFilled = Color.FromArgb(0xFF, 0xC0, 0xC0);

        // ----------------------------------------------------------------
        // 印刷用フィールド (PrintDocument 複数ページ対応)
        // ----------------------------------------------------------------
        private PrintDocument _printDoc;
        private List<string> _printLines;
        private int _printLineIdx;
        private Font _printFontTitle; // FontSize=20, Bold (VB6 タイトル行)
        private Font _printFontBody;  // FontSize=11        (VB6 本文)

        // ================================================================
        // コンストラクタ
        // ================================================================
        public FormPrintTedukuri()
        {
            InitializeComponent();
        }

        // ================================================================
        // InitializeComponent
        // ================================================================
        private void InitializeComponent()
        {
            this.SuspendLayout();

            // ---- フォーム基本設定 ----
            // VB6: BackColor=&H00D8FFFF&, Caption="夢仕立て-オーダーガイド見積り"
            this.Text = "夢仕立て - オーダーガイド見積り";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;

            // ---- ボトムバー (固定高さ) ----
            BuildBottomBar();

            // ---- メインスクロールパネル ----
            _pnlMain = new Panel
            {
                AutoScroll = true,
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF),
                Padding = new Padding(8, 8, 8, 8),
            };
            this.Controls.Add(_pnlMain);

            // ---- 各セクションをスクロールパネルに配置 ----
            int y = 8;
            y = BuildSectionRef(y);        // 参照データ
            y = BuildSectionArm(y, 1);     // 腕１
            y = BuildSectionArm(y, 2);     // 腕２
            y = BuildSectionHansen(y, 1);  // 板線材１
            y = BuildSectionHansen(y, 2);  // 板線材２
            y = BuildSectionIshiza(y, 1);  // 石座１
            y = BuildSectionIshiza(y, 2);  // 石座２
            y = BuildSectionIshiza(y, 3);  // 石座３
            y = BuildSectionIshidome(y, 1);// 石留め１
            y = BuildSectionIshidome(y, 2);// 石留め２
            y = BuildSectionIshidome(y, 3);// 石留め３
            y = BuildSectionIshidome(y, 4);// 石留め４
            y = BuildSectionDia(y, 1);     // ダイヤ１
            y = BuildSectionDia(y, 2);     // ダイヤ２
            y = BuildSectionDia(y, 3);     // ダイヤ３
            y = BuildSectionDia(y, 4);     // ダイヤ４
            y = BuildSectionRo(y, 1);      // ロー付け１
            y = BuildSectionRo(y, 2);      // ロー付け２
            y = BuildSectionKako(y);       // 加工難易度／加工グレード
            y = BuildSectionGokei(y);      // 合計

            this.ResumeLayout(false);
        }

        // ================================================================
        // ボトムバー構築 (VB6: 画面下部のコントロール群)
        // ================================================================
        private void BuildBottomBar()
        {
            var pnl = new Panel
            {
                Height = 58,
                Dock = DockStyle.Bottom,
                BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF),
            };

            // 作業書番号選択コンボ (VB6: 保存, Style=2 DropDownList)
            _cmbSave = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("ＭＳ Ｐゴシック", 12f, FontStyle.Bold),
                Width = 200,
                Location = new Point(10, 14),
            };

            // データ表示ボタン (VB6: btn_disp_data)
            _btnDispData = new Button
            {
                Text = "データ表示",
                Font = new Font("ＭＳ Ｐゴシック", 11f, FontStyle.Bold),
                Size = new Size(130, 36),
                Location = new Point(220, 11),
                BackColor = Color.FromArgb(0xC0, 0xFF, 0xFF),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnDispData.FlatAppearance.BorderColor = Color.SteelBlue;
            _btnDispData.FlatAppearance.BorderSize = 2;
            _btnDispData.Click += BtnDispData_Click;

            // データ削除ボタン (VB6: btn_delete_data)
            _btnDeleteData = new Button
            {
                Text = "データ削除",
                Font = new Font("ＭＳ Ｐゴシック", 11f, FontStyle.Bold),
                Size = new Size(130, 36),
                Location = new Point(360, 11),
                BackColor = Color.FromArgb(0xFF, 0xC0, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnDeleteData.FlatAppearance.BorderColor = Color.Crimson;
            _btnDeleteData.FlatAppearance.BorderSize = 2;
            _btnDeleteData.Click += BtnDeleteData_Click;

            // 印刷ボタン (VB6: btn_print)
            _btnPrint = new Button
            {
                Text = "印刷",
                Font = new Font("ＭＳ Ｐゴシック", 14f, FontStyle.Bold),
                Size = new Size(90, 45),
                Location = new Point(500, 6),
                BackColor = Color.FromArgb(0xFF, 0xFF, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnPrint.FlatAppearance.BorderColor = Color.Goldenrod;
            _btnPrint.FlatAppearance.BorderSize = 2;
            _btnPrint.Click += BtnPrint_Click;

            // メニューボタン (VB6: btn_menu, 右下固定)
            _btnMenu = new Button
            {
                Text = "メニュー",
                Font = new Font("ＭＳ Ｐゴシック", 12f, FontStyle.Bold),
                Size = new Size(110, 45),
                BackColor = Color.FromArgb(0xFF, 0xC0, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            };
            _btnMenu.FlatAppearance.BorderColor = Color.Crimson;
            _btnMenu.FlatAppearance.BorderSize = 2;
            _btnMenu.Click += BtnMenu_Click;

            pnl.Controls.Add(_cmbSave);
            pnl.Controls.Add(_btnDispData);
            pnl.Controls.Add(_btnDeleteData);
            pnl.Controls.Add(_btnPrint);
            pnl.Controls.Add(_btnMenu);

            // メニューボタンを右端に配置
            pnl.SizeChanged += (s, e) =>
                _btnMenu.Location = new Point(pnl.ClientSize.Width - 120, 6);

            this.Controls.Add(pnl);
        }

        // ================================================================
        // セクション共通ヘルパーメソッド
        // ================================================================

        /// <summary>
        /// セクションヘッダーラベルを追加する。
        /// </summary>
        private void AddSectionHeader(int y, string title)
        {
            var lbl = new Label
            {
                Text = title,
                Font = new Font("ＭＳ Ｐゴシック", 10f, FontStyle.Bold),
                BackColor = Color.FromArgb(0x70, 0xA0, 0xC0),
                ForeColor = Color.White,
                AutoSize = false,
                Width = 840,
                Height = 22,
                Location = new Point(8, y),
                TextAlign = ContentAlignment.MiddleLeft,
            };
            _pnlMain.Controls.Add(lbl);
        }

        /// <summary>
        /// フィールド名ラベル + 値ラベル（データ表示）の1行を追加する。
        /// 値ラベルは _dataLabels に登録される。
        /// VB6: Label.Name = DBフィールド名, BackColor=&H80000005, BackStyle=1
        /// </summary>
        private Label AddDataRow(int x, int y, string caption, string fieldName, int valWidth = 180)
        {
            // フィールド名ラベル (静的テキスト)
            var lblCap = new Label
            {
                Text = caption,
                Font = new Font("ＭＳ Ｐゴシック", 9f),
                BackColor = Color.Transparent,
                AutoSize = false,
                Width = 130,
                Height = 22,
                Location = new Point(x, y),
                TextAlign = ContentAlignment.MiddleRight,
            };
            _pnlMain.Controls.Add(lblCap);

            // 値ラベル (データ表示・VB6のラベルに相当)
            //   Name        = フィールド名 (Disp_Order_Data での Field アクセスに使用)
            //   BackColor   = ColorEmpty (&H80000005 相当)
            //   BorderStyle = FixedSingle (&H80000005 の BorderStyle=1 相当)
            var lblVal = new Label
            {
                Name = fieldName,
                Text = "",
                Font = new Font("ＭＳ Ｐゴシック", 9f, FontStyle.Bold),
                BackColor = ColorEmpty,
                BorderStyle = BorderStyle.FixedSingle,
                AutoSize = false,
                Width = valWidth,
                Height = 22,
                Location = new Point(x + 133, y),
                TextAlign = ContentAlignment.MiddleLeft,
            };
            _pnlMain.Controls.Add(lblVal);
            _dataLabels[fieldName] = lblVal;
            return lblVal;
        }

        // ================================================================
        // 各セクション構築
        // ================================================================

        // ---- 参照データ ----
        private int BuildSectionRef(int y)
        {
            AddSectionHeader(y, "■ 参照データ");
            y += 24;
            AddDataRow(8, y, "品番：", "参照品番", 150);
            AddDataRow(360, y, "PT900 (g)：", "参照PT900", 80);
            AddDataRow(570, y, "K18 (g)：", "参照K18", 80);
            y += 24;
            AddDataRow(8, y, "WG/PG (g)：", "参照WGPG", 80);
            AddDataRow(360, y, "K10 (g)：", "参照K10", 80);
            AddDataRow(570, y, "CODE：", "参照CODE", 80);
            y += 24;
            AddDataRow(8, y, "金額：", "参照金額", 150);
            y += 28;
            return y;
        }

        // ---- 腕 N (VB6: 腕地金種類N, 腕形状N, 腕天幅N, 腕天厚N, 腕底幅N, 腕底厚N,
        //                  腕特殊加工N, 腕リングサイズN, 腕プラマイN, 腕本数N) ----
        private int BuildSectionArm(int y, int n)
        {
            AddSectionHeader(y, $"■ 腕{n}");
            y += 24;
            AddDataRow(8, y, "地金種類：", $"腕地金種類{n}", 130);
            AddDataRow(360, y, "腕形状：", $"腕形状{n}", 130);
            y += 24;
            AddDataRow(8, y, "腕天幅 (mm)：", $"腕天幅{n}", 80);
            AddDataRow(360, y, "腕天厚 (mm)：", $"腕天厚{n}", 80);
            y += 24;
            AddDataRow(8, y, "腕底幅 (mm)：", $"腕底幅{n}", 80);
            AddDataRow(360, y, "腕底厚 (mm)：", $"腕底厚{n}", 80);
            y += 24;
            AddDataRow(8, y, "特殊加工：", $"腕特殊加工{n}", 130);
            AddDataRow(360, y, "リングサイズ (号)：", $"腕リングサイズ{n}", 80);
            y += 24;
            AddDataRow(8, y, "腕プラマイ：", $"腕プラマイ{n}", 60);
            AddDataRow(360, y, "腕本数：", $"腕本数{n}", 80);
            y += 28;
            return y;
        }

        // ---- 板線材 N (VB6: 板線材部材種類N, 板線材地金種類N, 板線材形状N,
        //                     板線材長径N, 板線材短径N, 板線材厚さN, 板線材個数N) ----
        private int BuildSectionHansen(int y, int n)
        {
            AddSectionHeader(y, $"■ 板線材{n}");
            y += 24;
            AddDataRow(8, y, "部材種類：", $"板線材部材種類{n}", 130);
            AddDataRow(360, y, "地金種類：", $"板線材地金種類{n}", 130);
            y += 24;
            AddDataRow(8, y, "形状：", $"板線材形状{n}", 130);
            AddDataRow(360, y, "長径 (mm)：", $"板線材長径{n}", 80);
            y += 24;
            AddDataRow(8, y, "短径 (mm)：", $"板線材短径{n}", 80);
            AddDataRow(360, y, "厚さ (mm)：", $"板線材厚さ{n}", 80);
            y += 24;
            AddDataRow(8, y, "個数：", $"板線材個数{n}", 80);
            y += 28;
            return y;
        }

        // ---- 石座 N (VB6: 石座地金種類N, 石座石の形状N, 石座石の長径N, 石座石の短径N,
        //                   石座主石座の種類N, 石座腰高N, 石座プラマイN, 石座個数N) ----
        private int BuildSectionIshiza(int y, int n)
        {
            AddSectionHeader(y, $"■ 石座{n}");
            y += 24;
            AddDataRow(8, y, "地金種類：", $"石座地金種類{n}", 130);
            AddDataRow(360, y, "石の形状：", $"石座石の形状{n}", 130);
            y += 24;
            AddDataRow(8, y, "石の長径 (mm)：", $"石座石の長径{n}", 80);
            AddDataRow(360, y, "石の短径 (mm)：", $"石座石の短径{n}", 80);
            y += 24;
            AddDataRow(8, y, "主石座の種類：", $"石座主石座の種類{n}", 160);
            AddDataRow(360, y, "腰高：", $"石座腰高{n}", 80);
            y += 24;
            AddDataRow(8, y, "プラマイ：", $"石座プラマイ{n}", 60);
            AddDataRow(360, y, "個数：", $"石座個数{n}", 80);
            y += 28;
            return y;
        }

        // ---- 石留め N (VB6: 石留め留め方法N, 石留め地金種類N, 石留め石の形状N,
        //                     石留め石のサイズN, 石留め個数N) ----
        private int BuildSectionIshidome(int y, int n)
        {
            AddSectionHeader(y, $"■ 石留め{n}");
            y += 24;
            AddDataRow(8, y, "留め方法：", $"石留め留め方法{n}", 160);
            AddDataRow(360, y, "地金種類：", $"石留め地金種類{n}", 130);
            y += 24;
            AddDataRow(8, y, "石の形状：", $"石留め石の形状{n}", 130);
            AddDataRow(360, y, "石のサイズ：", $"石留め石のサイズ{n}", 130);
            y += 24;
            AddDataRow(8, y, "個数/本数：", $"石留め個数{n}", 80);
            y += 28;
            return y;
        }

        // ---- ダイヤ N (VB6: ダイヤグレードN, ダイヤサイズN, ダイヤプラマイN, ダイヤ個数N) ----
        private int BuildSectionDia(int y, int n)
        {
            AddSectionHeader(y, $"■ ダイヤ{n}");
            y += 24;
            AddDataRow(8, y, "ダイヤグレード：", $"ダイヤグレード{n}", 130);
            AddDataRow(360, y, "サイズ (ct)：", $"ダイヤサイズ{n}", 80);
            y += 24;
            AddDataRow(8, y, "プラマイ：", $"ダイヤプラマイ{n}", 60);
            AddDataRow(360, y, "個数：", $"ダイヤ個数{n}", 80);
            y += 28;
            return y;
        }

        // ---- ロー付け N (VB6: ロー付け種類N, ロー付け個数N) ----
        private int BuildSectionRo(int y, int n)
        {
            AddSectionHeader(y, $"■ ロー付け{n}");
            y += 24;
            AddDataRow(8, y, "ロー付け種類：", $"ロー付け種類{n}", 160);
            AddDataRow(360, y, "ロー付け個数：", $"ロー付け個数{n}", 80);
            y += 28;
            return y;
        }

        // ---- 加工難易度 / 加工グレード ----
        private int BuildSectionKako(int y)
        {
            AddSectionHeader(y, "■ 加工難易度 / 加工グレード");
            y += 24;
            AddDataRow(8, y, "加工難易度：", "加工難易度", 320);
            AddDataRow(580, y, "加工グレード：", "加工グレード", 160);
            y += 28;
            return y;
        }

        // ---- 合計 (VB6: 合計PT900, 合計K18, 合計WGPG, 合計K10, 合計金額, 隠語) ----
        private int BuildSectionGokei(int y)
        {
            AddSectionHeader(y, "■ 合計");
            y += 24;
            AddDataRow(8, y, "PT900 (g)：", "合計PT900", 80);
            AddDataRow(360, y, "K18 (g)：", "合計K18", 80);
            y += 24;
            AddDataRow(8, y, "WG/PG (g)：", "合計WGPG", 80);
            AddDataRow(360, y, "K10 (g)：", "合計K10", 80);
            y += 24;
            AddDataRow(8, y, "合計金額 (税抜)：", "合計金額", 160);
            y += 24;
            AddDataRow(8, y, "隠語：", "隠語", 400);
            y += 28;
            return y;
        }

        // ================================================================
        // OnLoad / OnResize
        // ================================================================
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.WindowState = FormWindowState.Maximized;
            InitControl();
        }

        // ================================================================
        // VB6: Init_Control
        //   1. Init_All_Control で全ラベル初期化
        //   2. オーダーメイドテーブルから作業書番号を全件取得してコンボに追加
        //   3. 先頭を選択して Disp_Order_Data / Color_Control_Rtn を呼び出す
        // ================================================================
        private void InitControl()
        {
            InitAllControl();
            _cmbSave.Items.Clear();

            try
            {
                DataTable dt = AppState.Db.ExecuteQuery(
                    "SELECT [作業書番号] FROM [オーダーメイドテーブル]");
                foreach (DataRow row in dt.Rows)
                    _cmbSave.Items.Add(row["作業書番号"].ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("データ読み込みエラー：" + ex.Message,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (_cmbSave.Items.Count == 0) return;

            // VB6: 保存.Text = 保存.List(0)
            _cmbSave.SelectedIndex = 0;
            DispOrderData();
            ColorControlRtn();
        }

        // ================================================================
        // VB6: Init_All_Control
        //   全データ表示ラベルをクリアし、背景色を SystemWindow (ColorEmpty) に戻す。
        //   BackColor=&H80000005 And BackStyle=1 のラベルが対象 (C#では全 _dataLabels)
        // ================================================================
        private void InitAllControl()
        {
            foreach (var kvp in _dataLabels)
            {
                kvp.Value.Text = "";
                kvp.Value.BackColor = ColorEmpty;
            }
        }

        // ================================================================
        // VB6: Disp_Order_Data
        //   選択中の作業書番号に対応するレコードを取得し、
        //   ラベル名 = フィールド名 のパターンで Caption をセットする。
        // ================================================================
        private void DispOrderData()
        {
            if (string.IsNullOrEmpty(_cmbSave.Text)) return;

            string sql = "SELECT * FROM [オーダーメイドテーブル] WHERE [作業書番号] = '"
                       + _cmbSave.Text.Replace("'", "''") + "'";
            DataTable dt;
            try { dt = AppState.Db.ExecuteQuery(sql); }
            catch (Exception ex)
            {
                MessageBox.Show("データ取得エラー：" + ex.Message,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("該当データがありません。");
                return;
            }

            // VB6: For Each o In Me.Controls → o.Caption = recTB.Fields(o.Name)
            DataRow row = dt.Rows[0];
            foreach (var kvp in _dataLabels)
            {
                string field = kvp.Key;
                if (dt.Columns.Contains(field))
                    kvp.Value.Text = row[field]?.ToString() ?? "";
            }
        }

        // ================================================================
        // VB6: Color_Control_Rtn
        //   データが存在するセクションのラベルを薄ピンク (ColorFilled) でハイライト。
        //   各セクションの「キーフィールド」が非ゼロ / 非空のときにセクション全体を着色。
        // ================================================================
        private void ColorControlRtn()
        {
            // 参照データ
            HighlightIfNotEmpty("参照品番");
            HighlightIfNonZero("参照PT900");
            HighlightIfNonZero("参照K18");
            HighlightIfNonZero("参照WGPG");
            HighlightIfNonZero("参照K10");
            HighlightIfNotEmpty("参照CODE");
            HighlightIfNonZero("参照金額");

            // 腕1/2
            for (int n = 1; n <= 2; n++)
            {
                // VB6: 腕天幅N ≠ 0 And 腕天厚N ≠ 0 And 腕底幅N ≠ 0 And 腕底厚N ≠ 0
                //       And 腕リングサイズN ≠ 0 And 腕本数N ≠ 0
                if (IsNonZero($"腕天幅{n}") && IsNonZero($"腕天厚{n}") &&
                    IsNonZero($"腕底幅{n}") && IsNonZero($"腕底厚{n}") &&
                    IsNonZero($"腕リングサイズ{n}") && IsNonZero($"腕本数{n}"))
                {
                    Highlight($"腕地金種類{n}");
                    Highlight($"腕形状{n}");
                    Highlight($"腕天幅{n}");
                    Highlight($"腕天厚{n}");
                    Highlight($"腕底幅{n}");
                    Highlight($"腕底厚{n}");
                    Highlight($"腕特殊加工{n}");
                    Highlight($"腕リングサイズ{n}");
                    Highlight($"腕プラマイ{n}");
                    Highlight($"腕本数{n}");
                }
            }

            // 板線材1/2
            for (int n = 1; n <= 2; n++)
            {
                if (IsNonZero($"板線材長径{n}") && IsNonZero($"板線材短径{n}") &&
                    IsNonZero($"板線材厚さ{n}") && IsNonZero($"板線材個数{n}"))
                {
                    Highlight($"板線材部材種類{n}");
                    Highlight($"板線材地金種類{n}");
                    Highlight($"板線材形状{n}");
                    Highlight($"板線材長径{n}");
                    Highlight($"板線材短径{n}");
                    Highlight($"板線材厚さ{n}");
                    Highlight($"板線材個数{n}");
                }
            }

            // 石座1/2/3
            for (int n = 1; n <= 3; n++)
            {
                if (IsNonZero($"石座石の長径{n}") && IsNonZero($"石座石の短径{n}") &&
                    IsNonZero($"石座個数{n}"))
                {
                    Highlight($"石座地金種類{n}");
                    Highlight($"石座石の形状{n}");
                    Highlight($"石座石の長径{n}");
                    Highlight($"石座石の短径{n}");
                    Highlight($"石座主石座の種類{n}");
                    Highlight($"石座腰高{n}");
                    Highlight($"石座プラマイ{n}");
                    Highlight($"石座個数{n}");
                }
            }

            // 石留め1/2/3/4
            for (int n = 1; n <= 4; n++)
            {
                if (IsNonZero($"石留め個数{n}"))
                {
                    Highlight($"石留め留め方法{n}");
                    Highlight($"石留め地金種類{n}");
                    Highlight($"石留め石の形状{n}");
                    Highlight($"石留め石のサイズ{n}");
                    Highlight($"石留め個数{n}");
                }
            }

            // ダイヤ1/2/3/4
            for (int n = 1; n <= 4; n++)
            {
                if (IsNonZero($"ダイヤ個数{n}"))
                {
                    Highlight($"ダイヤグレード{n}");
                    Highlight($"ダイヤサイズ{n}");
                    Highlight($"ダイヤプラマイ{n}");
                    Highlight($"ダイヤ個数{n}");
                }
            }

            // ロー付け1/2
            for (int n = 1; n <= 2; n++)
            {
                if (IsNonZero($"ロー付け個数{n}"))
                {
                    Highlight($"ロー付け種類{n}");
                    Highlight($"ロー付け個数{n}");
                }
            }

            // 加工難易度 (VB6: コメントアウトされた条件なし → 常にハイライト)
            Highlight("加工難易度");

            // 合計
            HighlightIfNonZero("合計PT900");
            HighlightIfNonZero("合計K18");
            HighlightIfNonZero("合計WGPG");
            HighlightIfNonZero("合計K10");

            // 合計金額が非ゼロのとき加工グレード・隠語もハイライト
            if (IsNonZero("合計金額"))
            {
                Highlight("加工グレード");
                Highlight("合計金額");
                Highlight("隠語");
            }
        }

        // ================================================================
        // ハイライトヘルパー
        // ================================================================

        /// <summary>フィールド値が非ゼロなら ColorFilled に設定する。</summary>
        private bool IsNonZero(string fieldName)
        {
            if (!_dataLabels.TryGetValue(fieldName, out Label lbl)) return false;
            return double.TryParse(lbl.Text, out double v) && v != 0.0;
        }

        /// <summary>フィールド値が空文字でないか判定する。</summary>
        private bool IsNotEmpty(string fieldName)
        {
            if (!_dataLabels.TryGetValue(fieldName, out Label lbl)) return false;
            return !string.IsNullOrEmpty(lbl.Text);
        }

        /// <summary>指定フィールドのラベルを薄ピンクでハイライトする。</summary>
        private void Highlight(string fieldName)
        {
            if (_dataLabels.TryGetValue(fieldName, out Label lbl))
                lbl.BackColor = ColorFilled;
        }

        private void HighlightIfNonZero(string fieldName)
        {
            if (IsNonZero(fieldName)) Highlight(fieldName);
        }

        private void HighlightIfNotEmpty(string fieldName)
        {
            if (IsNotEmpty(fieldName)) Highlight(fieldName);
        }

        // ================================================================
        // ラベルテキスト取得ヘルパー
        // ================================================================
        private string GetVal(string fieldName)
        {
            if (_dataLabels.TryGetValue(fieldName, out Label lbl)) return lbl.Text;
            return "";
        }

        // ================================================================
        // イベントハンドラ
        // ================================================================

        /// <summary>
        /// データ表示ボタン (VB6: btn_disp_data_Click)
        /// 初期化 → データ取得 → カラーリング
        /// </summary>
        private void BtnDispData_Click(object sender, EventArgs e)
        {
            InitAllControl();
            DispOrderData();
            ColorControlRtn();
        }

        /// <summary>
        /// データ削除ボタン (VB6: btn_delete_data_Click)
        /// 確認後 DELETE、リスト再読み込み
        /// </summary>
        private void BtnDeleteData_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_cmbSave.Text))
            {
                MessageBox.Show("該当データがありません。");
                return;
            }

            // VB6: vbYesNo 確認
            if (MessageBox.Show(
                    $"作業書番号『{_cmbSave.Text}』の、オーダーメイドデータを削除します。よろしいですか？",
                    "データ削除確認", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            // VB6: dao_database.Execute "delete * from オーダーメイドテーブル where ..."
            string sql = "DELETE * FROM [オーダーメイドテーブル] WHERE [作業書番号] = '"
                       + _cmbSave.Text.Replace("'", "''") + "'";
            try
            {
                AppState.Db.ExecuteNonQuery(sql);
            }
            catch (Exception ex)
            {
                MessageBox.Show("削除エラー：" + ex.Message,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show("削除しました。", "オーダーメイドデータ削除完了");

            // リスト再読み込み (VB6: 保存.Clear → additem loop → 保存.Text = 保存.List(0))
            InitAllControl();
            _cmbSave.Items.Clear();

            try
            {
                DataTable dt = AppState.Db.ExecuteQuery(
                    "SELECT [作業書番号] FROM [オーダーメイドテーブル]");
                foreach (DataRow row in dt.Rows)
                    _cmbSave.Items.Add(row["作業書番号"].ToString());
            }
            catch { }

            if (_cmbSave.Items.Count == 0) return;

            _cmbSave.SelectedIndex = 0;
            DispOrderData();
            ColorControlRtn();
        }

        /// <summary>
        /// 印刷ボタン (VB6: btn_print_Click)
        /// データ存在確認 → A4縦確認 → 印刷確認 → Form_Print 相当を実行
        /// </summary>
        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_cmbSave.Text)) return;

            // データ存在確認
            string chkSql = "SELECT * FROM [オーダーメイドテーブル] WHERE [作業書番号] = '"
                           + _cmbSave.Text.Replace("'", "''") + "'";
            DataTable chkDt;
            try { chkDt = AppState.Db.ExecuteQuery(chkSql); }
            catch { chkDt = new DataTable(); }

            if (chkDt.Rows.Count == 0)
            {
                MessageBox.Show("該当データがありません。");
                return;
            }

            // VB6: A4縦確認ダイアログ (Yes=設定済み, No=設定ダイアログ表示, Cancel=中止)
            DialogResult ret = MessageBox.Show(
                "プリンターの設定を、A4縦方向に設定してありますか？",
                "プリンター設定確認",
                MessageBoxButtons.YesNoCancel);

            if (ret == DialogResult.Cancel) return;

            if (ret == DialogResult.No)
            {
                // VB6: sPrnSetUP → CommonDialog1.ShowPrinter
                using (var pd = new PrintDialog())
                {
                    pd.AllowSomePages = false;
                    pd.ShowDialog(this);
                }
            }

            // VB6: 印刷確認
            if (MessageBox.Show("印刷しますか？", "印刷確認",
                    MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            // VB6: Call Form_Print
            ExecutePrint();
        }

        /// <summary>
        /// メニューボタン (VB6: btn_menu_Click → Unload Me / form_menu.Show)
        /// </summary>
        private void BtnMenu_Click(object sender, EventArgs e)
        {
            foreach (Form f in Application.OpenForms)
            {
                if (f is FormMenu) { f.Show(); break; }
            }
            this.Close();
        }

        // ================================================================
        // 印刷処理 (VB6: Form_Print サブルーチン)
        //   VB6: Printer.Print "テキスト" → PrintDocument で DrawString 相当
        //   フォントサイズ: タイトル行=20pt Bold、本文=11pt
        // ================================================================

        /// <summary>印刷行リストを構築して PrintDocument で出力する。</summary>
        private void ExecutePrint()
        {
            _printLines = BuildPrintLines();
            _printLineIdx = 0;

            // VB6: Printer.FontSize=20 Bold → タイトルフォント
            _printFontTitle = new Font("ＭＳ Ｐゴシック", 20f, FontStyle.Bold);
            // VB6: Printer.FontSize=11 → 本文フォント
            _printFontBody = new Font("ＭＳ Ｐゴシック", 11f);

            _printDoc = new PrintDocument();
            _printDoc.DocumentName = "オーダーメイドお見積り";
            _printDoc.PrintPage += PrintDoc_PrintPage;
            _printDoc.EndPrint += PrintDoc_EndPrint;

            try
            {
                _printDoc.Print();
            }
            catch (Exception ex)
            {
                MessageBox.Show("印刷エラー：" + ex.Message,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// VB6: Form_Print 内の Printer.Print 呼び出し列を文字列リストとして構築する。
        ///   先頭文字が '\x01' の行 = タイトル行 (FontBold=True, FontSize=20)
        /// </summary>
        private List<string> BuildPrintLines()
        {
            var lines = new List<string>();

            // VB6: Printer.FontBold=True / Printer.Print "オーダーメイドお見積り"
            // '\x01' はタイトル行マーカー (PrintPage で太字大フォントを使用)
            lines.Add("\x01オーダーメイドお見積り");
            lines.Add("");

            // VB6: Printer.Print "■作業書番号：" & Me.保存.Text
            lines.Add("■作業書番号：" + _cmbSave.Text);
            lines.Add("");
            lines.Add("■参照データ----------------------");
            lines.Add("品番：" + GetVal("参照品番"));
            lines.Add("PT900：" + GetVal("参照PT900") + "g");
            lines.Add("K18：" + GetVal("参照K18") + "g");
            lines.Add("WG/PG：" + GetVal("参照WGPG") + "g");
            lines.Add("CODE：" + GetVal("参照CODE") + "g");
            lines.Add("金額：" + GetVal("参照金額"));

            // 腕1/2
            for (int n = 1; n <= 2; n++)
            {
                if (!IsNonZero($"腕天幅{n}") || !IsNonZero($"腕天厚{n}") ||
                    !IsNonZero($"腕底幅{n}") || !IsNonZero($"腕底厚{n}") ||
                    !IsNonZero($"腕リングサイズ{n}") || !IsNonZero($"腕本数{n}")) continue;

                lines.Add("");
                lines.Add($"■腕{n}----------------------");
                lines.Add("地金種類：" + GetVal($"腕地金種類{n}"));
                lines.Add("腕形状：" + GetVal($"腕形状{n}"));
                lines.Add("腕天幅：" + GetVal($"腕天幅{n}") + "mm");
                lines.Add("腕天厚：" + GetVal($"腕天厚{n}") + "mm");
                lines.Add("腕底幅：" + GetVal($"腕底幅{n}") + "mm");
                lines.Add("腕底厚：" + GetVal($"腕底厚{n}") + "mm");
                lines.Add("特殊加工：" + GetVal($"腕特殊加工{n}"));
                lines.Add("腕リングサイズ：" + GetVal($"腕リングサイズ{n}") + "号");
                // VB6: Me.腕プラマイN.Caption & Me.腕本数N.Caption
                lines.Add("腕本数：" + GetVal($"腕プラマイ{n}") + GetVal($"腕本数{n}"));
            }

            // 板線材1/2
            for (int n = 1; n <= 2; n++)
            {
                if (!IsNonZero($"板線材長径{n}") || !IsNonZero($"板線材短径{n}") ||
                    !IsNonZero($"板線材厚さ{n}") || !IsNonZero($"板線材個数{n}")) continue;

                lines.Add("");
                lines.Add($"■板線材{n}----------------------");
                lines.Add("部材種類：" + GetVal($"板線材部材種類{n}"));
                lines.Add("地金種類：" + GetVal($"板線材地金種類{n}"));
                lines.Add("形状：" + GetVal($"板線材形状{n}"));
                lines.Add("長径：" + GetVal($"板線材長径{n}") + "mm");
                lines.Add("短径：" + GetVal($"板線材短径{n}") + "mm");
                lines.Add("厚さ：" + GetVal($"板線材厚さ{n}") + "mm");
                lines.Add("個数：" + GetVal($"板線材個数{n}"));
            }

            // 石座1/2/3
            for (int n = 1; n <= 3; n++)
            {
                if (!IsNonZero($"石座石の長径{n}") || !IsNonZero($"石座石の短径{n}") ||
                    !IsNonZero($"石座個数{n}")) continue;

                lines.Add("");
                lines.Add($"■石座{n}----------------------");
                lines.Add("地金種類：" + GetVal($"石座地金種類{n}"));
                lines.Add("石の形状：" + GetVal($"石座石の形状{n}"));
                lines.Add("石の長径：" + GetVal($"石座石の長径{n}") + "mm");
                lines.Add("石の短径：" + GetVal($"石座石の短径{n}") + "mm");
                lines.Add("主石座の種類：" + GetVal($"石座主石座の種類{n}"));
                lines.Add("腰高：" + GetVal($"石座腰高{n}"));
                // VB6: Me.石座プラマイN.Caption & Me.石座個数N.Caption
                lines.Add("個数：" + GetVal($"石座プラマイ{n}") + GetVal($"石座個数{n}"));
            }

            // 石留め1/2/3/4
            for (int n = 1; n <= 4; n++)
            {
                if (!IsNonZero($"石留め個数{n}")) continue;

                lines.Add("");
                lines.Add($"■石留め{n}----------------------");
                string method = GetVal($"石留め留め方法{n}");
                lines.Add("留め方法：" + method);
                // VB6: 留め方法が "芯爪建留(本)" のときのみ地金種類を印刷
                if (method == "芯爪建留(本)")
                    lines.Add("地金種類：" + GetVal($"石留め地金種類{n}"));
                lines.Add("石の形状：" + GetVal($"石留め石の形状{n}"));
                lines.Add("石のサイズ：" + GetVal($"石留め石のサイズ{n}"));
                // VB6: "芯爪建留(本)" → "本数：", それ以外 → "石数："
                if (method == "芯爪建留(本)")
                    lines.Add("本数：" + GetVal($"石留め個数{n}"));
                else
                    lines.Add("石数：" + GetVal($"石留め個数{n}"));
            }

            // ダイヤ1/2/3/4
            for (int n = 1; n <= 4; n++)
            {
                if (!IsNonZero($"ダイヤ個数{n}")) continue;

                lines.Add("");
                lines.Add($"■ダイヤ{n}----------------------");
                lines.Add("ダイヤグレード：" + GetVal($"ダイヤグレード{n}"));
                lines.Add("ダイヤサイズ：" + GetVal($"ダイヤサイズ{n}") + "ct");
                // VB6: Me.ダイヤプラマイN.Caption & Me.ダイヤ個数N.Caption
                lines.Add("ダイヤ個数：" + GetVal($"ダイヤプラマイ{n}") + GetVal($"ダイヤ個数{n}"));
            }

            // ロー付け1/2
            for (int n = 1; n <= 2; n++)
            {
                if (!IsNonZero($"ロー付け個数{n}")) continue;

                lines.Add("");
                lines.Add($"■ロー付け{n}----------------------");
                lines.Add("ロー付け種類：" + GetVal($"ロー付け種類{n}"));
                lines.Add("ロー付け個数：" + GetVal($"ロー付け個数{n}"));
            }

            // 加工難易度 (VB6: コメントアウトされた条件なし → 常に印刷)
            lines.Add("");
            lines.Add("■加工難易度----------------------");
            lines.Add(GetVal("加工難易度"));

            // 加工グレード (VB6: 合計金額 ≠ 0 のときのみ)
            if (IsNonZero("合計金額"))
            {
                lines.Add("");
                lines.Add("■加工グレード----------------------");
                lines.Add(GetVal("加工グレード"));
            }

            // 合計
            lines.Add("");
            lines.Add("■合計----------------------");
            lines.Add("PT900：" + GetVal("合計PT900") + "g");
            lines.Add("K18：" + GetVal("合計K18") + "g");
            lines.Add("WG/PG：" + GetVal("合計WGPG") + "g");
            lines.Add("K10：" + GetVal("合計K10") + "g");
            // VB6: Printer.Print "合計金額：" & Me.合計金額.Caption & "(税抜き)"
            lines.Add("合計金額：" + GetVal("合計金額") + "(税抜き)");
            lines.Add("");
            // VB6: Printer.Print Me.隠語.Caption
            lines.Add(GetVal("隠語"));

            return lines;
        }

        /// <summary>
        /// PrintPage イベントハンドラ (複数ページ対応)
        /// _printLines を上から順に描画し、余白を超えたら HasMorePages=true。
        /// </summary>
        private void PrintDoc_PrintPage(object sender, PrintPageEventArgs e)
        {
            float y = e.MarginBounds.Top;
            float x = e.MarginBounds.Left;
            float bottom = e.MarginBounds.Bottom;
            var brush = Brushes.Black;

            while (_printLineIdx < _printLines.Count)
            {
                string raw = _printLines[_printLineIdx];
                bool isTitle = raw.Length > 0 && raw[0] == '\x01';
                string text = isTitle ? raw.Substring(1) : raw;
                Font font = isTitle ? _printFontTitle : _printFontBody;

                float lineH = font.GetHeight(e.Graphics);

                if (y + lineH > bottom)
                {
                    e.HasMorePages = true;
                    return;
                }

                e.Graphics.DrawString(text, font, brush, x, y);
                y += lineH;
                _printLineIdx++;
            }

            e.HasMorePages = false;
        }

        /// <summary>印刷終了後にフォントを解放する。</summary>
        private void PrintDoc_EndPrint(object sender, PrintEventArgs e)
        {
            _printFontTitle?.Dispose();
            _printFontTitle = null;
            _printFontBody?.Dispose();
            _printFontBody = null;
        }
    }
}
