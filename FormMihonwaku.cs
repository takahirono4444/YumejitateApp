using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

// ================================================================
// FormMihonwaku.cs
// VB6: form_mihonwaku.frm → C# + WinForms 移行
// 機能: 見本枠検索画面
//       アイテム・石形状・地金・サイズ等の条件から
//       見本枠検索テーブルを検索し、検索結果を form_movie2 に渡す
// 対象フレームワーク: .NET Framework 4.8
// C#バージョン: 7.3
// ================================================================

namespace YumejitateApp
{
    public class FormMihonwaku : Form
    {
        // ----------------------------------------------------------------
        // UIコントロール
        // ----------------------------------------------------------------

        // アイテム選択
        private ComboBox _cmbItem;         // アイテム
        // 石の形状
        private ComboBox _cmbStoneShape;   // 上から見た石の形
        // 中石セットスタイル
        private ComboBox _cmbSetStyle;     // 中石セットスタイル
        // 地金
        private ComboBox _cmbJigane;       // 地金
        // 商品番号
        private ComboBox _cmbHinban;       // Combo_商品番号

        // サイズ入力（上辺：十の位・一の位・小数位）
        private ComboBox _cmbSizeUeJuu;    // サイズ上十
        private ComboBox _cmbSizeUeIchi;   // サイズ上一
        private ComboBox _cmbSizeUeKo;     // サイズ上小
        // サイズ入力（下辺：十の位・一の位・小数位）
        private ComboBox _cmbSizeShitaJuu; // サイズ下十
        private ComboBox _cmbSizeShitaIchi;// サイズ下一
        private ComboBox _cmbSizeShitaKo; // サイズ下小

        // リングサイズ（十の位・一の位）
        private ComboBox _cmbRingsizeJuu;  // cmb_ringsize十
        private ComboBox _cmbRingsizeIchi; // cmb_ringsize一
        // グレード
        private ComboBox _cmbGrade;        // cmb_grade

        // ボタン
        private Button _btnSearch; // 見本枠検索
        private Button _btnBack;   // メニュー（戻る）

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------
        public FormMihonwaku()
        {
            InitializeComponent();
        }

        // ----------------------------------------------------------------
        // フォーム初期化（デザイン）
        // ----------------------------------------------------------------
        private void InitializeComponent()
        {
            this.Text = "夢仕立て-見本枠検索画面";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new System.Drawing.Size(900, 700);
            this.BackColor = System.Drawing.Color.FromArgb(0xD8, 0xFF, 0xFF); // VB6: &H00D8FFFF
            this.Font = new System.Drawing.Font("メイリオ", 11F);

            this.Load += new EventHandler(FormMihonwaku_Load);

            // --- タイトルラベル ---
            var lblTitle = new Label();
            lblTitle.Text = "見本枠検索";
            lblTitle.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Italic);
            lblTitle.AutoSize = true;
            lblTitle.Location = new System.Drawing.Point(380, 20);
            this.Controls.Add(lblTitle);

            int labelX = 40;
            int comboX = 220;
            int rowH = 50;
            int y = 80;

            // --- アイテム ---
            AddLabel("アイテム", labelX, y);
            _cmbItem = AddComboBox(comboX, y, 260);
            _cmbItem.SelectedIndexChanged += new EventHandler(CmbItem_SelectedIndexChanged);
            y += rowH;

            // --- 上から見た石の形 ---
            AddLabel("中石形状", labelX, y);
            _cmbStoneShape = AddComboBox(comboX, y, 260);
            _cmbStoneShape.SelectedIndexChanged += new EventHandler(CmbStoneShape_SelectedIndexChanged);
            y += rowH;

            // --- 中石セットスタイル ---
            AddLabel("中石セットスタイル", labelX, y);
            _cmbSetStyle = AddComboBox(comboX, y, 260);
            y += rowH;

            // --- 地金 ---
            AddLabel("地　金", labelX, y);
            _cmbJigane = AddComboBox(comboX, y, 260);
            y += rowH;

            // --- 商品番号 ---
            AddLabel("商品番号", labelX, y);
            _cmbHinban = AddComboBox(comboX, y, 180);
            y += rowH;

            // --- サイズ（上辺） ---
            AddLabel("サイズ", labelX, y);
            var lblSizeUe = new Label();
            lblSizeUe.Text = "上辺：";
            lblSizeUe.Location = new System.Drawing.Point(comboX, y + 5);
            lblSizeUe.AutoSize = true;
            this.Controls.Add(lblSizeUe);

            _cmbSizeUeJuu = AddDigitCombo(comboX + 55, y);
            var lblDot1b = new Label();
            lblDot1b.Location = new System.Drawing.Point(comboX + 118, y + 5);
            lblDot1b.AutoSize = true; this.Controls.Add(lblDot1b);
            _cmbSizeUeIchi = AddDigitCombo(comboX + 130, y);
            var lblDot1c = new Label(); lblDot1c.Text = "．";
            lblDot1c.Location = new System.Drawing.Point(comboX + 193, y + 5);
            lblDot1c.AutoSize = true; this.Controls.Add(lblDot1c);
            _cmbSizeUeKo = AddDigitCombo(comboX + 230, y);
            var lblMm1 = new Label(); lblMm1.Text = "mm";
            lblMm1.Location = new System.Drawing.Point(comboX + 295, y + 5);
            lblMm1.AutoSize = true; this.Controls.Add(lblMm1);
            y += rowH;

            // --- サイズ（下辺） ---
            var lblSizeShita = new Label();
            lblSizeShita.Text = "下辺：";
            lblSizeShita.Location = new System.Drawing.Point(comboX, y + 5);
            lblSizeShita.AutoSize = true;
            this.Controls.Add(lblSizeShita);

            _cmbSizeShitaJuu = AddDigitCombo(comboX + 55, y);
            var lblDot2b = new Label();
            lblDot2b.Location = new System.Drawing.Point(comboX + 118, y + 5);
            lblDot2b.AutoSize = true; this.Controls.Add(lblDot2b);
            _cmbSizeShitaIchi = AddDigitCombo(comboX + 130, y);
            var lblDot2c = new Label(); lblDot2c.Text = "．";
            lblDot2c.Location = new System.Drawing.Point(comboX + 193, y + 5);
            lblDot2c.AutoSize = true; this.Controls.Add(lblDot2c);
            _cmbSizeShitaKo = AddDigitCombo(comboX + 230, y);
            var lblMm2 = new Label(); lblMm2.Text = "mm";
            lblMm2.Location = new System.Drawing.Point(comboX + 295, y + 5);
            lblMm2.AutoSize = true; this.Controls.Add(lblMm2);
            y += rowH;

            // --- リングサイズまたはグレード ---
            AddLabel("リングサイズ\nまたはグレード", labelX, y);
            var lblHash = new Label(); lblHash.Text = "＃";
            lblHash.Location = new System.Drawing.Point(comboX, y + 5);
            lblHash.AutoSize = true; this.Controls.Add(lblHash);
            _cmbRingsizeJuu = AddDigitCombo(comboX + 20, y);
            _cmbRingsizeIchi = AddDigitCombo(comboX + 80, y);
            // グレード
            var lblGrade = new Label(); lblGrade.Text = "グレード";
            lblGrade.Location = new System.Drawing.Point(comboX + 150, y + 5);
            lblGrade.AutoSize = true; this.Controls.Add(lblGrade);
            _cmbGrade = AddComboBox(comboX + 230, y, 80);
            y += rowH + 20;

            // --- ボタン ---
            _btnSearch = new Button();
            _btnSearch.Text = "見本枠検索";
            _btnSearch.Location = new System.Drawing.Point(comboX, y);
            _btnSearch.Size = new System.Drawing.Size(160, 55);
            _btnSearch.Font = new System.Drawing.Font("メイリオ", 13F, System.Drawing.FontStyle.Bold);
            _btnSearch.BackColor = System.Drawing.Color.LightSkyBlue;
            _btnSearch.Click += new EventHandler(BtnSearch_Click);
            this.Controls.Add(_btnSearch);

            _btnBack = new Button();
            _btnBack.Text = "メニュー";
            _btnBack.Location = new System.Drawing.Point(comboX + 200, y);
            _btnBack.Size = new System.Drawing.Size(140, 55);
            _btnBack.Font = new System.Drawing.Font("メイリオ", 13F, System.Drawing.FontStyle.Bold);
            _btnBack.BackColor = System.Drawing.Color.LightGray;
            _btnBack.Click += new EventHandler(BtnBack_Click);
            this.Controls.Add(_btnBack);
        }

        // ================================================================
        // ヘルパー: ラベル追加
        // ================================================================
        private void AddLabel(string text, int x, int y)
        {
            var lbl = new Label();
            lbl.Text = text;
            lbl.Location = new System.Drawing.Point(x, y + 5);
            lbl.Size = new System.Drawing.Size(170, 38);
            lbl.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Bold);
            this.Controls.Add(lbl);
        }

        // ================================================================
        // ヘルパー: DropDownList形式のComboBox追加
        // ================================================================
        private ComboBox AddComboBox(int x, int y, int width)
        {
            var cmb = new ComboBox();
            cmb.Location = new System.Drawing.Point(x, y);
            cmb.Size = new System.Drawing.Size(width, 30);
            cmb.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Bold);
            this.Controls.Add(cmb);
            return cmb;
        }

        // ================================================================
        // ヘルパー: 0-9桁入力用ComboBox追加
        // ================================================================
        private ComboBox AddDigitCombo(int x, int y)
        {
            var cmb = new ComboBox();
            cmb.Location = new System.Drawing.Point(x, y);
            cmb.Size = new System.Drawing.Size(55, 30);
            cmb.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Bold);
            for (int d = 0; d <= 9; d++)
                cmb.Items.Add(d.ToString());
            cmb.SelectedIndex = 0;
            this.Controls.Add(cmb);
            return cmb;
        }

        // ================================================================
        // Form_Load
        // VB6: Call Init_Control → WindowState = 2（最大化）
        // ================================================================
        private void FormMihonwaku_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            InitControl();
        }

        // ================================================================
        // Init_Control: コントロール初期化
        // VB6: Init_Control() - コンボボックスにアイテムを追加し、DBから商品番号を読み込む
        // ================================================================
        private void InitControl()
        {
            try
            {
                // --- アイテム ---
                _cmbItem.Items.Add("リング");
                _cmbItem.Items.Add("ブローチ");
                _cmbItem.Items.Add("ペンダント");
                _cmbItem.Items.Add("タイタック");
                _cmbItem.Items.Add("ピアス");
                _cmbItem.Items.Add("イアリング");
                _cmbItem.Items.Add("その他（バチカン）");
                _cmbItem.SelectedIndex = 0;

                // --- 上から見た石の形 ---
                _cmbStoneShape.Items.Add("ラウンド（円）");
                _cmbStoneShape.Items.Add("オーバル（楕円）");
                _cmbStoneShape.Items.Add("ボール（球）");
                _cmbStoneShape.Items.Add("エメラルド（四角）");
                _cmbStoneShape.Items.Add("マーキース");
                _cmbStoneShape.Items.Add("ドロップ");
                _cmbStoneShape.Items.Add("石無し");
                _cmbStoneShape.SelectedIndex = 0;

                // --- 地金 ---
                _cmbJigane.Items.Add("プラチナ");
                _cmbJigane.Items.Add("Ｋ１８ＹＧ");
                _cmbJigane.Items.Add("Ｋ１８ＷＧ");
                _cmbJigane.Items.Add("コンビ");
                _cmbJigane.Items.Add("シルバー");
                _cmbJigane.Items.Add("Ｋ１０");
                _cmbJigane.SelectedIndex = 0;

                // --- 中石セットスタイル ---
                _cmbSetStyle.Items.Add("おまかせ");
                _cmbSetStyle.Items.Add("爪留め");
                _cmbSetStyle.Items.Add("爪無し（レール留め等）");
                _cmbSetStyle.SelectedIndex = 0;

                // --- 商品番号（DBから読み込み）---
                // VB6: "select a from 見本枠商品番号テーブル order by a"
                _cmbHinban.Items.Add("おまかせ");
                var dt = AppState.Db.ExecuteQuery("SELECT [a] FROM [見本枠商品番号テーブル] ORDER BY [a]");
                foreach (DataRow row in dt.Rows)
                    _cmbHinban.Items.Add(row["a"].ToString());
                _cmbHinban.SelectedIndex = 0;

                // --- グレード ---
                _cmbGrade.Items.Add("A");
                _cmbGrade.Items.Add("B");
                _cmbGrade.Items.Add("C");
                _cmbGrade.SelectedIndex = 0;
                // 初期状態ではグレードは無効（アイテムがリング以外の場合）
                // VB6: cmb_grade.Enabled = False
                _cmbGrade.Enabled = false;

                // リングサイズはリングのときのみ有効（初期はリングなので有効）
                UpdateRingsizeGradeEnabled();
            }
            catch (Exception ex)
            {
                MessageBox.Show("初期化処理でエラーが発生しました。\n" + ex.Message,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ================================================================
        // アイテム変更イベント
        // VB6: アイテム_Click()
        // リングのときはリングサイズを有効化、グレードを無効化。それ以外は逆。
        // ================================================================
        private void CmbItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateRingsizeGradeEnabled();
        }

        private void UpdateRingsizeGradeEnabled()
        {
            bool isRing = (_cmbItem.Text == "リング");
            _cmbRingsizeJuu.Enabled = isRing;
            _cmbRingsizeIchi.Enabled = isRing;
            _cmbGrade.Enabled = !isRing;
        }

        // ================================================================
        // 上から見た石の形変更イベント
        // VB6: 上から見た石の形_Click()
        // 石無しのときはサイズ入力と中石セットスタイルを無効化
        // ================================================================
        private void CmbStoneShape_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool isNoStone = (_cmbStoneShape.Text == "石無し");
            _cmbSetStyle.Enabled = !isNoStone;
            _cmbSizeUeJuu.Enabled = !isNoStone;
            _cmbSizeUeIchi.Enabled = !isNoStone;
            _cmbSizeUeKo.Enabled = !isNoStone;
            _cmbSizeShitaJuu.Enabled = !isNoStone;
            _cmbSizeShitaIchi.Enabled = !isNoStone;
            _cmbSizeShitaKo.Enabled = !isNoStone;
        }

        // ================================================================
        // 入力チェック
        // VB6: check_control() As Boolean
        // ================================================================
        private bool CheckControl()
        {
            // サイズ入力チェック（石無し以外）
            if (_cmbStoneShape.Text != "石無し")
            {
                bool ueAllZero = (_cmbSizeUeJuu.SelectedIndex == 0 &&
                                  _cmbSizeUeIchi.SelectedIndex == 0 &&
                                  _cmbSizeUeKo.SelectedIndex == 0);
                bool shitaAllZero = (_cmbSizeShitaJuu.SelectedIndex == 0 &&
                                     _cmbSizeShitaIchi.SelectedIndex == 0 &&
                                     _cmbSizeShitaKo.SelectedIndex == 0);
                if (ueAllZero || shitaAllZero)
                {
                    MessageBox.Show("「サイズ」を入力して下さい。", "入力チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            // リングサイズチェック
            if (_cmbItem.Text == "リング")
            {
                if (_cmbRingsizeJuu.SelectedIndex == 0 && _cmbRingsizeIchi.SelectedIndex == 0)
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
        // 見本枠検索ボタン
        // VB6: btn_search_Click()
        // ================================================================
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            // --- 入力チェック ---
            if (!CheckControl()) return;

            // --- グローバルフラグ初期化 ---
            AppState.FlagMihonRev1 = false;
            AppState.MihonRev1Tsize = 0;

            MessageBox.Show("数千点以上のデータからご希望のデザインをお探しします。",
                "検索開始", MessageBoxButtons.OK, MessageBoxIcon.Information);

            try
            {
                // --- ワークテーブル・検索結果テーブルをクリア ---
                AppState.Db.ExecuteNonQuery("DELETE * FROM [ワーク検索テーブル]");
                AppState.Db.ExecuteNonQuery("DELETE * FROM [検索結果テーブル]");

                // --- 検索SQL構築 ---
                string baseWhere = BuildBaseWhere();
                string sizeWhere = BuildSizeWhere();
                string ringsizeWhere = BuildRingSizeWhere();

                string whereClause = baseWhere + sizeWhere + ringsizeWhere;

                // 先頭のANDを除去してWHEREに変換
                whereClause = whereClause.Trim();
                if (whereClause.StartsWith("AND "))
                    whereClause = whereClause.Substring(4);

                string strsql;
                if (string.IsNullOrWhiteSpace(whereClause))
                    strsql = "SELECT * FROM [見本枠検索テーブル]";
                else
                    strsql = "SELECT * FROM [見本枠検索テーブル] WHERE " + whereClause;

                // --- 1回目の検索（精度高め） ---
                var dt = AppState.Db.ExecuteQuery(strsql);
                if (dt.Rows.Count == 0)
                {
                    // --- 改良ロジック（ヒットなしの場合、サイズ条件を緩める）---
                    // VB6: 見本枠検索-改良①処理
                    string sizeWhereRev1 = BuildSizeWhereRev1();
                    string whereClauseRev1 = (baseWhere + sizeWhereRev1 + ringsizeWhere).Trim();
                    if (whereClauseRev1.StartsWith("AND "))
                        whereClauseRev1 = whereClauseRev1.Substring(4);
                    strsql = "SELECT * FROM [見本枠検索テーブル] WHERE " + whereClauseRev1;

                    dt = AppState.Db.ExecuteQuery(strsql);
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show(
                            "誠に申し訳ございませんが" + Environment.NewLine +
                            "ご希望のデザインはございません。",
                            "「固定情報」該当データ無し",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                // --- ワーク検索テーブルへ挿入 ---
                string insertWork = "INSERT INTO [ワーク検索テーブル] " + strsql;
                AppState.Db.ExecuteNonQuery(insertWork);

                // --- 品番の重複排除（totalsize降順で先頭のみ残す）---
                RemoveDuplicateHinban();

                // --- ランダムに最大7件を検索結果テーブルへ移動 ---
                SelectRandomResults(7);

                // --- 件数表示 ---
                var dtResult = AppState.Db.ExecuteQuery("SELECT * FROM [検索結果テーブル]");
                int cnt = dtResult.Rows.Count;
                MessageBox.Show(
                    "ご希望のデザインは" + cnt + "種類ございます。",
                    "検索結果件数表示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                // --- 検索結果画面へ遷移 ---
                // VB6: form_mihonwaku.Visible = False; form_movie2.Show
                var movie2Form = new FormMovie2();
                this.Hide();
                movie2Form.FormClosed += (s, args) => this.Show();
                movie2Form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("検索処理でエラーが発生しました。\n" + ex.Message,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ================================================================
        // 基本検索条件（アイテム・石形状・セットスタイル・地金・商品番号）を構築
        // VB6: btn_search_Click 内の前半部分
        // ================================================================
        private string BuildBaseWhere()
        {
            string where = "";

            // アイテム条件
            string itemCond = GetItemCondition();
            if (!string.IsNullOrEmpty(itemCond))
                where += itemCond + " ";

            // 石形状条件
            string stoneCond = GetStoneShapeCondition();
            if (!string.IsNullOrEmpty(stoneCond))
                where += stoneCond + " ";

            // 中石セットスタイル条件（石無し以外）
            if (_cmbStoneShape.Text != "石無し")
            {
                string styleCond = GetSetStyleCondition();
                if (!string.IsNullOrEmpty(styleCond))
                    where += styleCond + " ";
            }

            // 地金条件
            string jiganeCond = GetJiganeCondition();
            if (!string.IsNullOrEmpty(jiganeCond))
                where += jiganeCond + " ";

            // 商品番号条件
            if (_cmbHinban.Text != "おまかせ")
                where += "AND [a] = '" + _cmbHinban.Text.Replace("'", "''") + "' ";

            return where;
        }

        private string GetItemCondition()
        {
            return ""; // 一時的に条件なし
        }

        private string GetStoneShapeCondition()
        {
            switch (_cmbStoneShape.Text)
            {
                case "ラウンド（円）": return "AND [c] LIKE '%1%'";
                case "オーバル（楕円）": return "AND [c] LIKE '%2%'";
                case "ボール（球）": return "AND [c] LIKE '%3%'";
                case "エメラルド（四角）": return "AND [c] LIKE '%4%'";
                case "マーキース": return "AND [c] LIKE '%5%'";
                case "ドロップ": return "AND [c] LIKE '%6%'";
                case "石無し": return "AND [c] = '9'";
                default: return "";
            }
        }

        private string GetSetStyleCondition()
        {
            switch (_cmbSetStyle.Text)
            {
                case "爪留め": return "AND [h] LIKE '%1%'";
                case "爪無し（レール留め等）": return "AND [h] LIKE '%2%'";
                default: return ""; // おまかせは条件なし
            }
        }

        private string GetJiganeCondition()
        {
            switch (_cmbJigane.Text)
            {
                case "プラチナ": return "AND [d] LIKE '%1%'";
                case "Ｋ１８ＹＧ": return "AND [d] LIKE '%2%'";
                case "Ｋ１８ＷＧ": return "AND [d] LIKE '%3%'";
                case "コンビ": return "AND [d] LIKE '%4%'";
                case "シルバー": return "AND [d] LIKE '%5%'";
                case "Ｋ１０": return "AND [d] LIKE '%6%'";
                default: return "";
            }
        }

        // ================================================================
        // サイズ検索条件を構築（通常版・精度高め）
        // VB6: btn_search_Click サイズ処理（式１・式３）
        //
        // size_a = サイズ上十*100 + サイズ上一*10 + サイズ上小
        // size_b = サイズ下十*100 + サイズ下一*10 + サイズ下小
        // b_size = max(size_a, size_b)  ← 大きい方
        // s_size = min(size_a, size_b)  ← 小さい方
        // t_size = size_a + size_b      ← 合計
        //
        // 非ボール（式１・AB2以外）:
        //   bigsize >= b_size/1.18, <= b_size/0.75
        //   smallsize >= s_size/1.18, <= s_size/0.75
        //   totalsize >= t_size-5, <= t_size+10
        // ================================================================
        private string BuildSizeWhere()
        {
            if (_cmbStoneShape.Text == "石無し") return "";

            double sizeA = GetSizeValue(_cmbSizeUeJuu, _cmbSizeUeIchi, _cmbSizeUeKo);
            double sizeB = GetSizeValue(_cmbSizeShitaJuu, _cmbSizeShitaIchi, _cmbSizeShitaKo);
            double bSize = Math.Max(sizeA, sizeB);
            double sSize = Math.Min(sizeA, sizeB);
            double tSize = sizeA + sizeB;

            AppState.MihonRev1Tsize = tSize;

            int bigLow = (int)Math.Round(bSize / 1.18);
            int bigHigh = (int)Math.Round(bSize / 0.75);
            int smlLow = (int)Math.Round(sSize / 1.18);
            int smlHigh = (int)Math.Round(sSize / 0.75);
            int totLow = (int)Math.Round(tSize - 5);
            int totHigh = (int)Math.Round(tSize + 10);

            return "AND [bigsize] >= " + bigLow +
                   " AND [bigsize] <= " + bigHigh +
                   " AND [smallsize] >= " + smlLow +
                   " AND [smallsize] <= " + smlHigh +
                   " AND [totalsize] >= " + totLow +
                   " AND [totalsize] <= " + totHigh + " ";
        }

        // ================================================================
        // サイズ検索条件（改良版・ヒットなし時の緩い条件）
        // VB6: 見本枠検索-改良①処理
        //
        // totalsize のみで絞り込む（bigsize/smallsizeは使わない）
        //   totalsize >= t_size/1.3, <= t_size
        // ================================================================
        private string BuildSizeWhereRev1()
        {
            if (_cmbStoneShape.Text == "石無し") return "";

            double sizeA = GetSizeValue(_cmbSizeUeJuu, _cmbSizeUeIchi, _cmbSizeUeKo);
            double sizeB = GetSizeValue(_cmbSizeShitaJuu, _cmbSizeShitaIchi, _cmbSizeShitaKo);
            double tSize = sizeA + sizeB;
            AppState.MihonRev1Tsize = tSize;

            int totLow = (int)Math.Round(tSize / 1.3);
            int totHigh = (int)Math.Round(tSize);

            return "AND [totalsize] >= " + totLow +
                   " AND [totalsize] <= " + totHigh + " ";
        }

        // ================================================================
        // リングサイズ条件を構築
        // VB6: m_size = ringsize十*10 + ringsize一
        //      wk_sql = "and (m_size between val(w) and val(x))"
        // ================================================================
        private string BuildRingSizeWhere()
        {
            if (_cmbItem.Text != "リング") return "";

            int mSize = _cmbRingsizeJuu.SelectedIndex * 10
                      + _cmbRingsizeIchi.SelectedIndex;

            // w,xはVARCHAR型のためCLng()で数値変換して比較
            return "AND (CLng([w]) <= " + mSize + " AND CLng([x]) >= " + mSize + ") ";
        }

        // ================================================================
        // 品番の重複排除
        // VB6: ワーク検索テーブルを品番でソートして同品番が連続した場合に削除
        //      （totalsize降順で先頭を残す）
        // ================================================================
        private void RemoveDuplicateHinban()
        {
            var dtWk = AppState.Db.ExecuteQuery(
                "SELECT * FROM [ワーク検索テーブル] ORDER BY [a], [totalsize] DESC");

            string prevHinban = "";
            foreach (DataRow row in dtWk.Rows)
            {
                string curHinban = row["a"].ToString();
                if (curHinban == prevHinban)
                {
                    // 同品番の2件目以降は削除
                    AppState.Db.ExecuteNonQuery(
                        "DELETE FROM [ワーク検索テーブル] WHERE [index] = " + row["index"]);
                }
                else
                {
                    prevHinban = curHinban;
                }
            }
        }

        // ================================================================
        // ランダムに最大 maxCount 件を検索結果テーブルへ移動
        // VB6: 乱数テーブルを使ったシャッフル処理
        //      maxCount=7 件になるまでランダム選択を繰り返す
        // ================================================================
        private void SelectRandomResults(int maxCount)
        {
            var rnd = new Random();
            int selected = 0;

            while (selected < maxCount)
            {
                var dtWk = AppState.Db.ExecuteQuery(
                    "SELECT * FROM [ワーク検索テーブル] ORDER BY [index]");
                int remaining = dtWk.Rows.Count;
                if (remaining == 0) break;

                // ランダムに1件選択
                int pick = rnd.Next(0, remaining);
                string idx = dtWk.Rows[pick]["index"].ToString();

                // 検索結果テーブルへコピー
                AppState.Db.ExecuteNonQuery(
                    "INSERT INTO [検索結果テーブル] SELECT * FROM [ワーク検索テーブル] WHERE [index] = " + idx);

                // ワークから削除
                AppState.Db.ExecuteNonQuery(
                    "DELETE FROM [ワーク検索テーブル] WHERE [index] = " + idx);

                selected++;
            }
        }

        // ================================================================
        // メニューボタン（戻る）
        // VB6: btn_back_Click → Flag_Mihon = False; form_mihonwaku.Visible = False; form_menu.Visible = True
        // ================================================================
        private void BtnBack_Click(object sender, EventArgs e)
        {
            AppState.FlagMihon = false;
            this.Close();
        }

        // ================================================================
        // ユーティリティ: サイズコンボから数値を計算
        // VB6: CDbl(サイズ十.Text)*100 + CDbl(サイズ一.Text)*10 + CDbl(サイズ小.Text)
        // ================================================================
        private double GetSizeValue(ComboBox juu, ComboBox ichi, ComboBox ko)
        {
            return juu.SelectedIndex * 100.0
                 + ichi.SelectedIndex * 10.0
                 + ko.SelectedIndex;
        }

        // ================================================================
        // ユーティリティ: double を Access SQL の数値文字列へ変換
        // VB6ではロケールによる小数点が異なるため、Invariant Culture で書式化
        // ================================================================
        private string FormatDouble(double value)
        {
            return ((int)Math.Round(value)).ToString();
        }
    }

    // ================================================================
    // FormMovie2: 検索結果表示フォーム（スタブ）
    // VB6: form_movie2.frm に相当
    // 実装時に form_movie2.frm を読み込んで置き換えること
    // ================================================================
    public class FormMovie2 : Form
    {
        public FormMovie2()
        {
            this.Text = "検索結果表示";
            this.Size = new System.Drawing.Size(1024, 768);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = System.Drawing.Color.White;

            var lbl = new Label();
            lbl.Text = "検索結果表示画面（form_movie2 の移植先）";
            lbl.Font = new System.Drawing.Font("メイリオ", 14F);
            lbl.AutoSize = true;
            lbl.Location = new System.Drawing.Point(200, 300);
            this.Controls.Add(lbl);

            var btnClose = new Button();
            btnClose.Text = "閉じる";
            btnClose.Location = new System.Drawing.Point(450, 400);
            btnClose.Size = new System.Drawing.Size(120, 45);
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);
        }
    }
}

// ================================================================
// AppState への追加フィールド（form_mihonwaku グローバル変数）
// VB6: Module レベルの Public 変数
//   Flag_Mihon_rev1  As Boolean
//   Mihon_rev1_tsize As Double
// C#: AppState 静的クラスに追加プロパティとして定義
// ================================================================

// ※ AppState クラスは FormMenu.cs に定義済み。
//   以下のプロパティを AppState クラスに追加すること：
//
//   public static bool   FlagMihonRev1  { get; set; }
//   public static double MihonRev1Tsize { get; set; }