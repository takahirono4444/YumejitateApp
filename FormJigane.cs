using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace YumejitateApp
{
    // ================================================================
    // FormJigane.cs
    // VB6: form_jigane.frm の移植
    // 地金相場の入力とその目安（Pt1000・K24の買取額試算）
    // ================================================================
    public class FormJigane : Form
    {
        // ----------------------------------------------------------------
        // 定数
        // ----------------------------------------------------------------

        // DBフィールド名（地金テーブル / 地金相場テーブル / 地金買い相場テーブル共通）
        private static readonly string[] PtFields = { "pt万", "pt千", "pt百", "pt十", "pt一" };
        private static readonly string[] K18Fields = { "k18万", "k18千", "k18百", "k18十", "k18一" };

        // 相場テーブルの読込時オフセット（VB6: 2015/01/15 調整値）
        // 地金相場テーブルは「実勢価格 + オフセット」で格納されている
        private const double PtOffset = 118.0;
        private const double K24Offset = 78.0;

        // カテゴリ名（計算ロジックに対応させるための順序）
        private static readonly string[] Categories = { "リング", "チェーン", "インゴット", "その他" };

        // 買取係数（VB6 btn_keisan_Click 内の計算式から抽出）
        // 順序: [リング, チェーン, インゴット, その他]
        private static readonly double[] MaxFactors = { 0.80, 0.65, 0.85, 0.55 };
        private static readonly double[] MinFactors = { 0.60, 0.45, 0.75, 0.35 };

        // ----------------------------------------------------------------
        // フィールド
        // ----------------------------------------------------------------

        // Pt1000相場入力（5桁コンボ: 万千百十一）
        // VB6: PT万, PT千, PT百, PT十, PT一 各ComboBox
        private ComboBox[] _cmbPt = new ComboBox[5];

        // K24相場入力（5桁コンボ）
        // VB6: K24万, K24千, K24百, K24十, K24一 各ComboBox
        private ComboBox[] _cmbK24 = new ComboBox[5];

        // 重量入力 セット1（カテゴリ × 金属種別）
        // [category, 0]=PT数量, [category, 1]=K18数量
        private NumericUpDown[,] _nudQty1 = new NumericUpDown[4, 2];
        // [category, 0]=PT重量(g), [category, 1]=K18重量(g)
        private NumericUpDown[,] _nudGrams1 = new NumericUpDown[4, 2];

        // 重量入力 セット2（VB6では同画面に2セット分配置されていた）
        private NumericUpDown[,] _nudQty2 = new NumericUpDown[4, 2];
        private NumericUpDown[,] _nudGrams2 = new NumericUpDown[4, 2];

        // 表示ラベル
        private Label _lblCurrentPt;   // 現在のPt1000価格（地金買い相場テーブルより）
        private Label _lblCurrentK24;  // 現在のK24価格
        private Label _lblTotMid;      // 集計中（最大・最小の平均）
        private Label _lblTotLow;      // 集計低（最小値）
        private Label _lblTotMidTax;   // 集計中+消費税
        private Label _lblTotLowTax;   // 集計低+消費税

        // ボタン
        private Button _btnKeisan;  // 計算・保存（VB6: btn_keisan）
        private Button _btnBack;    // 戻る（VB6: btn_back）

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------

        public FormJigane()
        {
            InitializeComponent();
        }

        // ----------------------------------------------------------------
        // コントロール初期化
        // ----------------------------------------------------------------

        private void InitializeComponent()
        {
            // フォーム基本設定
            // VB6: BackColor = &H00D8FFFF（水色系）
            this.Text = "夢仕立て - 地金相場の入力とその目安";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            this.StartPosition = FormStartPosition.WindowsDefaultBounds;
            this.Font = new Font("メイリオ", 11f, FontStyle.Regular);

            this.Load += new EventHandler(FormJigane_Load);

            // スクロール可能パネル（コントロールが多いため）
            var scrollPanel = new Panel();
            scrollPanel.Dock = DockStyle.Fill;
            scrollPanel.AutoScroll = true;
            scrollPanel.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);

            // 各セクションを縦に積み上げて配置
            int y = 10;
            y = BuildPriceSection(scrollPanel, y);
            y = BuildWeightSection(scrollPanel, y, 1, _nudQty1, _nudGrams1);
            y = BuildWeightSection(scrollPanel, y, 2, _nudQty2, _nudGrams2);
            BuildResultSection(scrollPanel, y);

            this.Controls.Add(scrollPanel);
        }

        // ================================================================
        // セクション構築ヘルパー
        // ================================================================

        // ---- 相場入力セクション -----------------------------------------

        /// <summary>
        /// Pt1000・K24相場の5桁コンボボックス入力エリアを構築する。
        /// VB6: PT万/PT千/PT百/PT十/PT一 および K24万/K24千/K24百/K24十/K24一 コンボ相当。
        /// </summary>
        private int BuildPriceSection(Panel parent, int startY)
        {
            var grp = new GroupBox();
            grp.Text = "相場入力（各桁を選択）";
            grp.Location = new Point(10, startY);
            grp.Size = new Size(960, 130);
            grp.Font = new Font("メイリオ", 11f, FontStyle.Bold);
            grp.BackColor = Color.FromArgb(200, 240, 255);

            string[] digitLabels = { "万", "千", "百", "十", "一" };

            // ---- Pt1000 行 ----
            var lblPt = new Label
            {
                Text = "Pt1000:",
                Location = new Point(10, 30),
                AutoSize = true,
                Font = new Font("メイリオ", 12f, FontStyle.Bold),
                ForeColor = Color.DarkBlue
            };
            grp.Controls.Add(lblPt);

            for (int i = 0; i < 5; i++)
            {
                var lbl = new Label
                {
                    Text = digitLabels[i],
                    Location = new Point(125 + i * 80, 15),
                    AutoSize = true,
                    Font = new Font("メイリオ", 10f, FontStyle.Regular)
                };
                grp.Controls.Add(lbl);
                _cmbPt[i] = CreateDigitCombo(new Point(110 + i * 80, 30));
                grp.Controls.Add(_cmbPt[i]);
            }

            _lblCurrentPt = new Label
            {
                Text = "現在: ---",
                Location = new Point(530, 30),
                AutoSize = true,
                Font = new Font("メイリオ", 14f, FontStyle.Bold),
                ForeColor = Color.DarkBlue
            };
            grp.Controls.Add(_lblCurrentPt);

            // ---- K24 行 ----
            var lblK = new Label
            {
                Text = "K24:",
                Location = new Point(10, 80),
                AutoSize = true,
                Font = new Font("メイリオ", 12f, FontStyle.Bold),
                ForeColor = Color.DarkGoldenrod
            };
            grp.Controls.Add(lblK);

            for (int i = 0; i < 5; i++)
            {
                var lbl = new Label
                {
                    Text = digitLabels[i],
                    Location = new Point(125 + i * 80, 65),
                    AutoSize = true,
                    Font = new Font("メイリオ", 10f, FontStyle.Regular)
                };
                grp.Controls.Add(lbl);
                _cmbK24[i] = CreateDigitCombo(new Point(110 + i * 80, 80));
                grp.Controls.Add(_cmbK24[i]);
            }

            _lblCurrentK24 = new Label
            {
                Text = "現在: ---",
                Location = new Point(530, 80),
                AutoSize = true,
                Font = new Font("メイリオ", 14f, FontStyle.Bold),
                ForeColor = Color.DarkGoldenrod
            };
            grp.Controls.Add(_lblCurrentK24);

            parent.Controls.Add(grp);
            return startY + grp.Height + 10;
        }

        /// <summary>0〜9のリストを持つ桁入力用ComboBoxを生成する。</summary>
        private ComboBox CreateDigitCombo(Point location)
        {
            var cmb = new ComboBox();
            cmb.Location = location;
            cmb.Size = new Size(68, 30);
            cmb.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb.Font = new Font("メイリオ", 14f, FontStyle.Bold);
            for (int i = 0; i <= 9; i++) cmb.Items.Add(i);
            cmb.SelectedIndex = 0;
            return cmb;
        }

        // ---- 重量入力セクション -----------------------------------------

        /// <summary>
        /// リング/チェーン/インゴット/その他 × Pt/K18 の重量入力エリアを構築する。
        /// VB6では各桁ごとのComboBoxで入力していたが、C#ではNumericUpDownに簡略化。
        /// 計算式はVB6と完全に同一。
        /// </summary>
        private int BuildWeightSection(Panel parent, int startY, int setNo,
                                       NumericUpDown[,] qty, NumericUpDown[,] grams)
        {
            var grp = new GroupBox();
            grp.Text = string.Format("重量入力　セット{0}　（Pt: 個数×重量g　/ K18: 個数×重量g）", setNo);
            grp.Location = new Point(10, startY);
            grp.Size = new Size(1100, 220);
            grp.Font = new Font("メイリオ", 10f, FontStyle.Bold);
            grp.BackColor = Color.FromArgb(220, 255, 220);

            // ヘッダー行
            AddLabel(grp, "区分", new Point(10, 28), new Font("メイリオ", 10f, FontStyle.Bold));
            AddLabel(grp, "Pt 個数", new Point(120, 28), new Font("メイリオ", 10f, FontStyle.Regular));
            AddLabel(grp, "Pt 重量(g)", new Point(260, 28), new Font("メイリオ", 10f, FontStyle.Regular));
            AddLabel(grp, "K18 個数", new Point(420, 28), new Font("メイリオ", 10f, FontStyle.Regular));
            AddLabel(grp, "K18 重量(g)", new Point(560, 28), new Font("メイリオ", 10f, FontStyle.Regular));

            // カテゴリ行（リング/チェーン/インゴット/その他）
            for (int c = 0; c < 4; c++)
            {
                int rowY = 55 + c * 40;

                AddLabel(grp, Categories[c], new Point(10, rowY + 5),
                         new Font("メイリオ", 12f, FontStyle.Bold));

                // Pt 個数（VB6: 万digit×1000 + 百digit×100 + 十digit×10 → 0〜9990）
                qty[c, 0] = CreateNud(new Point(120, rowY), 0, 9990, 10, 0);
                grp.Controls.Add(qty[c, 0]);

                // Pt 重量g（VB6: G百digit×100 + G十digit×10 + G一digit + G0.1digit×0.1 → 0.0〜999.9）
                grams[c, 0] = CreateNud(new Point(260, rowY), 0m, 999.9m, 0.1m, 1);
                grp.Controls.Add(grams[c, 0]);

                // K18 個数（VB6: 十digit×10 + 一digit → 0〜99）
                qty[c, 1] = CreateNud(new Point(420, rowY), 0, 99, 1, 0);
                grp.Controls.Add(qty[c, 1]);

                // K18 重量g（Ptと同じ構造）
                grams[c, 1] = CreateNud(new Point(560, rowY), 0m, 999.9m, 0.1m, 1);
                grp.Controls.Add(grams[c, 1]);
            }

            parent.Controls.Add(grp);
            return startY + grp.Height + 10;
        }

        /// <summary>NumericUpDownを生成して返す。</summary>
        private NumericUpDown CreateNud(Point location, decimal min, decimal max, decimal increment, int decimalPlaces)
        {
            var nud = new NumericUpDown();
            nud.Location = location;
            nud.Size = new Size(125, 28);
            nud.Minimum = min;
            nud.Maximum = max;
            nud.Increment = increment;
            nud.DecimalPlaces = decimalPlaces;
            nud.Value = 0;
            nud.Font = new Font("メイリオ", 11f, FontStyle.Regular);
            return nud;
        }

        // ---- 結果表示・ボタンセクション ---------------------------------

        /// <summary>計算結果ラベルと計算/戻るボタンを配置するパネルを構築する。</summary>
        private void BuildResultSection(Panel parent, int startY)
        {
            var pnl = new Panel();
            pnl.Location = new Point(10, startY);
            pnl.Size = new Size(960, 130);
            pnl.BackColor = Color.FromArgb(255, 255, 200);
            pnl.BorderStyle = BorderStyle.FixedSingle;

            var bigBold = new Font("メイリオ", 14f, FontStyle.Bold);

            // 集計結果ラベル（VB6: 集計中, 集計低, 集計中税込, 集計低税込）
            _lblTotMid = new Label { Text = "集計中: ---", Location = new Point(10, 12), AutoSize = true, Font = bigBold, ForeColor = Color.Navy };
            _lblTotLow = new Label { Text = "集計低: ---", Location = new Point(280, 12), AutoSize = true, Font = bigBold, ForeColor = Color.DarkRed };
            _lblTotMidTax = new Label { Text = "集計中+税: ---", Location = new Point(10, 65), AutoSize = true, Font = bigBold, ForeColor = Color.Navy };
            _lblTotLowTax = new Label { Text = "集計低+税: ---", Location = new Point(280, 65), AutoSize = true, Font = bigBold, ForeColor = Color.DarkRed };

            // 計算・保存ボタン（VB6: btn_keisan）
            _btnKeisan = new Button
            {
                Text = "計算・保存",
                Location = new Point(650, 20),
                Size = new Size(140, 50),
                Font = new Font("メイリオ", 13f, FontStyle.Bold),
                BackColor = Color.FromArgb(100, 200, 255),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnKeisan.Click += new EventHandler(BtnKeisan_Click);

            // 戻るボタン（VB6: btn_back）
            _btnBack = new Button
            {
                Text = "戻る",
                Location = new Point(810, 20),
                Size = new Size(120, 50),
                Font = new Font("メイリオ", 13f, FontStyle.Bold),
                BackColor = Color.FromArgb(255, 180, 180),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnBack.Click += new EventHandler(BtnBack_Click);

            pnl.Controls.AddRange(new Control[]
            {
                _lblTotMid, _lblTotLow, _lblTotMidTax, _lblTotLowTax, _btnKeisan, _btnBack
            });
            parent.Controls.Add(pnl);
        }

        // ---- ラベル生成ヘルパー -----------------------------------------

        private static void AddLabel(Control parent, string text, Point location, Font font)
        {
            var lbl = new Label { Text = text, Location = location, AutoSize = true, Font = font };
            parent.Controls.Add(lbl);
        }

        // ================================================================
        // 起動時処理（VB6の Form_Load → Init_Control に相当）
        // ================================================================

        private void FormJigane_Load(object sender, EventArgs e)
        {
            try
            {
                // AppState.Dbが接続されているか確認
                if (AppState.Db == null || !AppState.Db.IsConnected)
                {
                    MessageBox.Show("DB未接続です。", "エラー",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 地金テーブルから現在値を読み込んでコンボボックスに設定
                try
                {
                    DataTable dt = AppState.Db.ExecuteQuery(
                        "SELECT * FROM [地金テーブル]");

                    if (dt.Rows.Count > 0 && dt.Columns.Count >= 10)
                    {
                        DataRow row = dt.Rows[0];
                        for (int i = 0; i < 5; i++)
                        {
                            int idx = 0;
                            int.TryParse(row[i].ToString(), out idx);
                            if (idx >= 0 && idx <= 9)
                                _cmbPt[i].SelectedIndex = idx;
                            else
                                _cmbPt[i].SelectedIndex = 0;
                        }
                        for (int i = 0; i < 5; i++)
                        {
                            int idx = 0;
                            int.TryParse(row[i + 5].ToString(), out idx);
                            if (idx >= 0 && idx <= 9)
                                _cmbK24[i].SelectedIndex = idx;
                            else
                                _cmbK24[i].SelectedIndex = 0;
                        }
                    }
                    else
                    {
                        // データなし→全て0にセット
                        for (int i = 0; i < 5; i++)
                        {
                            _cmbPt[i].SelectedIndex = 0;
                            _cmbK24[i].SelectedIndex = 0;
                        }
                    }
                }
                catch
                {
                    for (int i = 0; i < 5; i++)
                    {
                        _cmbPt[i].SelectedIndex = 0;
                        _cmbK24[i].SelectedIndex = 0;
                    }
                }

                // 現在の相場表示を更新
                UpdateCurrentPriceLabels();
            }
            catch (Exception ex)
            {
                MessageBox.Show("初期化エラー：" + ex.Message, "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ----------------------------------------------------------------
        // DBから相場読み込み
        // VB6: Init_Control 内の地金相場テーブル読込処理に相当
        // ----------------------------------------------------------------

        private void LoadPriceFromDb()
        {
            try
            {
                var db = AppState.Db;
                if (db == null || !db.IsConnected)
                {
                    MessageBox.Show("DBに接続されていません。", "エラー",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 地金相場テーブルから生値を取得
                // VB6: strsql = "select * from 地金相場テーブル"
                DataTable dt = db.ExecuteQuery("SELECT * FROM [地金相場テーブル]");
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("システムエラーです。", "地金相場テーブル内データ無し",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                DataRow row = dt.Rows[0];

                // Pt1000 生値を再構成 → オフセット引いて実勢価格に変換
                // VB6: dbl_Pt1000 = CDbl(pt万)*10000 + CDbl(pt千)*1000 + ... - 118
                double rawPt = ToDouble(row["pt万"]) * 10000 + ToDouble(row["pt千"]) * 1000
                             + ToDouble(row["pt百"]) * 100 + ToDouble(row["pt十"]) * 10
                             + ToDouble(row["pt一"]);
                double actualPt = rawPt - PtOffset;

                if (actualPt < 0 || actualPt > 99999)
                {
                    MessageBox.Show("システムエラーです。", "地金相場計算",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // K24 同様
                double rawK24 = ToDouble(row["k18万"]) * 10000 + ToDouble(row["k18千"]) * 1000
                              + ToDouble(row["k18百"]) * 100 + ToDouble(row["k18十"]) * 10
                              + ToDouble(row["k18一"]);
                double actualK24 = rawK24 - K24Offset;

                if (actualK24 < 0 || actualK24 > 99999)
                {
                    MessageBox.Show("システムエラーです。", "地金相場計算",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 各桁を対応するComboBoxに設定
                // VB6: str_Pt1000 = Format(CStr(dbl_Pt1000), "00000")
                //      PT万.Text = CStr(PT万.List(CInt(Mid(str_Pt1000, 1, 1))))
                string strPt = ((long)actualPt).ToString("00000");
                string strK24 = ((long)actualK24).ToString("00000");

                for (int i = 0; i < 5; i++)
                {
                    _cmbPt[i].SelectedIndex = int.Parse(strPt[i].ToString());
                    _cmbK24[i].SelectedIndex = int.Parse(strK24[i].ToString());
                }

                // 現在の買い相場ラベルを更新（地金買い相場テーブルから）
                UpdateCurrentPriceLabels();
            }
            catch (Exception ex)
            {
                MessageBox.Show("DB読み込みエラー: " + ex.Message, "エラー",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ----------------------------------------------------------------
        // ラベル更新（地金買い相場テーブルの現在値を表示）
        // VB6: Init_Control 末尾 および btn_keisan_Click 末尾の処理に相当
        // ----------------------------------------------------------------

        private void UpdateCurrentPriceLabels()
        {
            try
            {
                DataTable dt = AppState.Db.ExecuteQuery(
                    "SELECT * FROM [地金買い相場テーブル]");

                if (dt.Rows.Count == 0 || dt.Columns.Count == 0)
                {
                    _lblCurrentPt.Text = "現在: 未設定";
                    _lblCurrentK24.Text = "現在: 未設定";
                    return;
                }

                DataRow row = dt.Rows[0];

                // フィールドが存在するか確認してから取得
                string strPt = "";
                string strK18 = "";

                if (dt.Columns.Contains("pt万"))
                {
                    strPt = row["pt万"].ToString() + row["pt千"].ToString()
                          + row["pt百"].ToString() + row["pt十"].ToString()
                          + row["pt一"].ToString();
                }

                if (dt.Columns.Contains("k18万"))
                {
                    strK18 = row["k18万"].ToString() + row["k18千"].ToString()
                           + row["k18百"].ToString() + row["k18十"].ToString()
                           + row["k18一"].ToString();
                }

                long ptVal = 0;
                long k18Val = 0;
                long.TryParse(strPt, out ptVal);
                long.TryParse(strK18, out k18Val);

                _lblCurrentPt.Text = ptVal == 0 ? "現在: 未設定"
                                                   : "現在: " + ptVal.ToString("#,##0") + " 円/g";
                _lblCurrentK24.Text = k18Val == 0 ? "現在: 未設定"
                                                   : "現在: " + k18Val.ToString("#,##0") + " 円/g";
            }
            catch (Exception ex)
            {
                _lblCurrentPt.Text = "現在: 未設定";
                _lblCurrentK24.Text = "現在: 未設定";
                System.Diagnostics.Debug.WriteLine("ラベル更新失敗: " + ex.Message);
            }
        }

        // ================================================================
        // 計算・保存ボタン（VB6: btn_keisan_Click に相当）
        // ================================================================

        private void BtnKeisan_Click(object sender, EventArgs e)
        {
            // ---- 1. ComboBoxから相場値を再構成 --------------------------
            // VB6: pt1000 = Left(PT万.Text,1)*10000 + ... + Left(PT一.Text,1)
            long pt1000 = (long)_cmbPt[0].SelectedIndex * 10000
                        + (long)_cmbPt[1].SelectedIndex * 1000
                        + (long)_cmbPt[2].SelectedIndex * 100
                        + (long)_cmbPt[3].SelectedIndex * 10
                        + (long)_cmbPt[4].SelectedIndex;

            long k24 = (long)_cmbK24[0].SelectedIndex * 10000
                     + (long)_cmbK24[1].SelectedIndex * 1000
                     + (long)_cmbK24[2].SelectedIndex * 100
                     + (long)_cmbK24[3].SelectedIndex * 10
                     + (long)_cmbK24[4].SelectedIndex;

            // ---- 2. 重量入力から買取額試算（2セット分を加算）-----------
            // VB6: totMax/totMin の計算式を忠実に再現
            double totMax = 0.0;
            double totMin = 0.0;
            CalcSet(_nudQty1, _nudGrams1, pt1000, k24, ref totMax, ref totMin);
            CalcSet(_nudQty2, _nudGrams2, pt1000, k24, ref totMax, ref totMin);

            // ---- 3. 結果をラベルに表示 ----------------------------------
            // VB6: 集計中 = (totMax + totMin) / 2
            //      集計低 = totMin
            //      集計中税込 = 1.1 * 集計中
            //      集計低税込 = 1.1 * 集計低
            double mid = (totMax + totMin) / 2.0;
            _lblTotMid.Text = "集計中: " + ((long)mid).ToString("0,000") + " 円";
            _lblTotLow.Text = "集計低: " + ((long)totMin).ToString("0,000") + " 円";
            _lblTotMidTax.Text = "集計中+税: " + ((long)(1.1 * mid)).ToString("0,000") + " 円";
            _lblTotLowTax.Text = "集計低+税: " + ((long)(1.1 * totMin)).ToString("0,000") + " 円";

            // ---- 4. 地金テーブルに保存 ----------------------------------
            // VB6: delete * from 地金テーブル
            //      insert into 地金テーブル values(PT各桁のListIndex, K24各桁のListIndex)
            try
            {
                SaveToJinkineTable();

                // 地金買い相場テーブル = 地金テーブルのコピー
                // VB6: delete * from 地金買い相場テーブル
                //      insert into 地金買い相場テーブル select * from 地金テーブル
                SaveToBuyMarketTable();

                // ラベル表示を更新
                UpdateCurrentPriceLabels();
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存中にエラーが発生しました。\n\n" + ex.Message, "保存エラー",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ----------------------------------------------------------------
        // 1セット分の買取額計算
        // VB6 btn_keisan_Click の計算式から抽出
        //
        // 計算構造:
        //   Pt: 個数(qty) × 重量g(grams) = 総重量(mg相当)
        //       買取額 = pt1000 × 総重量 × 係数 / 1000
        //   K18: 個数(qty) × 重量g(grams) = 総重量(g) → /24 でK24単位に換算
        //       買取額 = k24 × 総重量/24 × 係数
        // ----------------------------------------------------------------

        private static void CalcSet(NumericUpDown[,] qty, NumericUpDown[,] grams,
                                    long pt1000, long k24,
                                    ref double totMax, ref double totMin)
        {
            for (int c = 0; c < 4; c++)
            {
                // Pt側
                // VB6: ringPt = (万digit×1000 + 百digit×100 + 十digit×10) × (G百digit×100 + ...)
                double ptWeight = (double)qty[c, 0].Value * (double)grams[c, 0].Value;
                totMax += pt1000 * ptWeight * MaxFactors[c] / 1000.0;
                totMin += pt1000 * ptWeight * MinFactors[c] / 1000.0;

                // K18側
                // VB6: ringK = (十digit×10 + 一digit) × (G百digit×100 + ...) / 24
                double kWeight = (double)qty[c, 1].Value * (double)grams[c, 1].Value / 24.0;
                totMax += k24 * kWeight * MaxFactors[c];
                totMin += k24 * kWeight * MinFactors[c];
            }
        }

        // ----------------------------------------------------------------
        // 地金テーブルへの保存
        // VB6: delete * from 地金テーブル → insert into 地金テーブル values(...)
        // ----------------------------------------------------------------

        private void SaveToJinkineTable()
        {
            var db = AppState.Db;

            // 地金相場テーブルに保存（pt万〜pt一、K18万〜k18一）
            db.ExecuteNonQuery("DELETE * FROM [地金相場テーブル]");

            var prms = new OleDbParameter[10];
            prms[0] = new OleDbParameter("pt万", _cmbPt[0].SelectedIndex.ToString());
            prms[1] = new OleDbParameter("pt千", _cmbPt[1].SelectedIndex.ToString());
            prms[2] = new OleDbParameter("pt百", _cmbPt[2].SelectedIndex.ToString());
            prms[3] = new OleDbParameter("pt十", _cmbPt[3].SelectedIndex.ToString());
            prms[4] = new OleDbParameter("pt一", _cmbPt[4].SelectedIndex.ToString());
            prms[5] = new OleDbParameter("K18万", _cmbK24[0].SelectedIndex.ToString());
            prms[6] = new OleDbParameter("k18千", _cmbK24[1].SelectedIndex.ToString());
            prms[7] = new OleDbParameter("k18百", _cmbK24[2].SelectedIndex.ToString());
            prms[8] = new OleDbParameter("k18十", _cmbK24[3].SelectedIndex.ToString());
            prms[9] = new OleDbParameter("k18一", _cmbK24[4].SelectedIndex.ToString());

            db.ExecuteNonQuery(
                "INSERT INTO [地金相場テーブル] VALUES(?,?,?,?,?,?,?,?,?,?)",
                prms);
        }

        // ----------------------------------------------------------------
        // 地金買い相場テーブルへのコピー
        // VB6: delete * from 地金買い相場テーブル
        //      insert into 地金買い相場テーブル select * from 地金テーブル
        // ----------------------------------------------------------------

        private void SaveToBuyMarketTable()
        {
            var db = AppState.Db;
            db.ExecuteNonQuery("DELETE * FROM [地金買い相場テーブル]");
            db.ExecuteNonQuery(
                "INSERT INTO [地金買い相場テーブル] SELECT * FROM [地金相場テーブル]");
        }

        // ================================================================
        // 戻るボタン（VB6: btn_back_Click → Unload form_jigane; form_menu.Show）
        // C#: FormMenu.OpenChildForm が FormClosed で自動的に Show() を呼ぶ
        // ================================================================

        private void BtnBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // ----------------------------------------------------------------
        // 型変換ヘルパー
        // ----------------------------------------------------------------

        /// <summary>DBのフィールド値を double に変換する。変換失敗時は 0 を返す。</summary>
        private static double ToDouble(object dbValue)
        {
            double result = 0.0;
            double.TryParse(dbValue.ToString(), out result);
            return result;
        }
    }
}
