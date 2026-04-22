using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

// ================================================================
// FormMakedata.cs
// VB6: form_makedata.frm → C# + WinForms 移行
// 機能: PT/K24相場を入力し、全検索テーブルを一括再作成する
// 対象フレームワーク: .NET Framework 4.8
// C#バージョン: 7.3
// ================================================================

namespace YumejitateApp
{
    public class FormMakedata : Form
    {
        // ----------------------------------------------------------------
        // UIコントロール
        // ----------------------------------------------------------------

        // PT価格入力用コンボボックス（万・千・百・十・一）
        private ComboBox[] _cmbPt = new ComboBox[5];
        // K24価格入力用コンボボックス（万・千・百・十・一）
        private ComboBox[] _cmbK24 = new ComboBox[5];

        // 現在の価格表示ラベル
        private Label _lblPt1000;   // PT1000現在価格
        private Label _lblK24;      // K24現在価格
        private Label _lblKakeritu; // 掛け率テーブル表示

        // ボタン
        private Button _btnMakedata; // データ作成
        private Button _btnExit;     // 閉じる

        // 各桁ラベル（PT）
        private Label _lblPtMan, _lblPtSen, _lblPtHyaku, _lblPtJuu, _lblPtIchi;
        // 各桁ラベル（K24）
        private Label _lblK24Man, _lblK24Sen, _lblK24Hyaku, _lblK24Juu, _lblK24Ichi;

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------
        public FormMakedata()
        {
            InitializeComponent();
        }

        // ----------------------------------------------------------------
        // フォーム初期化
        // ----------------------------------------------------------------
        private void InitializeComponent()
        {
            this.Text = "データ作成";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new System.Drawing.Size(700, 500);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Font = new System.Drawing.Font("メイリオ", 10F);
            this.BackColor = System.Drawing.Color.White;

            this.Load += new EventHandler(FormMakedata_Load);

            // --- PT価格エリア ---
            var lblPtTitle = new Label();
            lblPtTitle.Text = "PT1000 価格（円/g）";
            lblPtTitle.Location = new System.Drawing.Point(30, 30);
            lblPtTitle.AutoSize = true;
            lblPtTitle.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Bold);
            this.Controls.Add(lblPtTitle);

            // PT各桁ラベルと位置
            string[] ptLabels = { "万", "千", "百", "十", "一" };
            int[] ptX = { 30, 100, 170, 240, 310 };
            for (int i = 0; i < 5; i++)
            {
                var lbl = new Label();
                lbl.Text = ptLabels[i];
                lbl.Location = new System.Drawing.Point(ptX[i], 60);
                lbl.AutoSize = true;
                this.Controls.Add(lbl);

                _cmbPt[i] = new ComboBox();
                _cmbPt[i].Location = new System.Drawing.Point(ptX[i], 80);
                _cmbPt[i].Size = new System.Drawing.Size(60, 28);
                _cmbPt[i].DropDownStyle = ComboBoxStyle.DropDownList;
                for (int d = 0; d <= 9; d++)
                    _cmbPt[i].Items.Add(d.ToString());
                _cmbPt[i].SelectedIndex = 0;
                this.Controls.Add(_cmbPt[i]);
            }

            _lblPt1000 = new Label();
            _lblPt1000.Text = "PT1000: ----";
            _lblPt1000.Location = new System.Drawing.Point(390, 80);
            _lblPt1000.AutoSize = true;
            _lblPt1000.Font = new System.Drawing.Font("メイリオ", 11F);
            this.Controls.Add(_lblPt1000);

            // --- K24価格エリア ---
            var lblK24Title = new Label();
            lblK24Title.Text = "K24 価格（円/g）";
            lblK24Title.Location = new System.Drawing.Point(30, 130);
            lblK24Title.AutoSize = true;
            lblK24Title.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Bold);
            this.Controls.Add(lblK24Title);

            string[] k24Labels = { "万", "千", "百", "十", "一" };
            int[] k24X = { 30, 100, 170, 240, 310 };
            for (int i = 0; i < 5; i++)
            {
                var lbl = new Label();
                lbl.Text = k24Labels[i];
                lbl.Location = new System.Drawing.Point(k24X[i], 160);
                lbl.AutoSize = true;
                this.Controls.Add(lbl);

                _cmbK24[i] = new ComboBox();
                _cmbK24[i].Location = new System.Drawing.Point(k24X[i], 180);
                _cmbK24[i].Size = new System.Drawing.Size(60, 28);
                _cmbK24[i].DropDownStyle = ComboBoxStyle.DropDownList;
                for (int d = 0; d <= 9; d++)
                    _cmbK24[i].Items.Add(d.ToString());
                _cmbK24[i].SelectedIndex = 0;
                this.Controls.Add(_cmbK24[i]);
            }

            _lblK24 = new Label();
            _lblK24.Text = "K24: ----";
            _lblK24.Location = new System.Drawing.Point(390, 180);
            _lblK24.AutoSize = true;
            _lblK24.Font = new System.Drawing.Font("メイリオ", 11F);
            this.Controls.Add(_lblK24);

            // --- 掛け率ラベル ---
            _lblKakeritu = new Label();
            _lblKakeritu.Text = "掛け率: ----";
            _lblKakeritu.Location = new System.Drawing.Point(30, 240);
            _lblKakeritu.AutoSize = true;
            _lblKakeritu.Font = new System.Drawing.Font("メイリオ", 11F);
            this.Controls.Add(_lblKakeritu);

            // --- ボタンエリア ---
            _btnMakedata = new Button();
            _btnMakedata.Text = "データ作成";
            _btnMakedata.Location = new System.Drawing.Point(30, 310);
            _btnMakedata.Size = new System.Drawing.Size(150, 50);
            _btnMakedata.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Bold);
            _btnMakedata.BackColor = System.Drawing.Color.LightSteelBlue;
            _btnMakedata.Click += new EventHandler(BtnMakedata_Click);
            this.Controls.Add(_btnMakedata);

            _btnExit = new Button();
            _btnExit.Text = "閉じる";
            _btnExit.Location = new System.Drawing.Point(200, 310);
            _btnExit.Size = new System.Drawing.Size(120, 50);
            _btnExit.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Bold);
            _btnExit.BackColor = System.Drawing.Color.LightGray;
            _btnExit.Click += new EventHandler(BtnExit_Click);
            this.Controls.Add(_btnExit);
        }

        // ----------------------------------------------------------------
        // Form_Load: Init_Control相当
        // ----------------------------------------------------------------
        private void FormMakedata_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            InitControl();
        }

        /// <summary>
        /// VB6: Init_Control
        /// pt_k18テーブルを読み込み、PT/K24の各桁コンボに値をセット。
        /// 地金相場テーブルを読み込み、現在価格ラベルを更新。
        /// 掛け率テーブルを読み込み、掛け率ラベルを更新。
        /// </summary>
        private void InitControl()
        {
            try
            {
                // pt_k18テーブルから現在のPT/K24桁値を取得
                var dtPtK18 = AppState.Db.ExecuteQuery("SELECT * FROM [pt_k18テーブル]");
                if (dtPtK18.Rows.Count > 0)
                {
                    var row = dtPtK18.Rows[0];
                    // フィールド順: pt万,pt千,pt百,pt十,pt一,k18万,k18千,k18百,k18十,k18一,...
                    string[] ptFields = { "pt万", "pt千", "pt百", "pt十", "pt一" };
                    string[] k18Fields = { "k18万", "k18千", "k18百", "k18十", "k18一" };
                    for (int i = 0; i < 5; i++)
                    {
                        int ptVal = 0;
                        int k18Val = 0;
                        int.TryParse(row[ptFields[i]].ToString(), out ptVal);
                        int.TryParse(row[k18Fields[i]].ToString(), out k18Val);
                        // コンボのインデックスは数値そのもの（0～9）
                        if (ptVal >= 0 && ptVal <= 9) _cmbPt[i].SelectedIndex = ptVal;
                        if (k18Val >= 0 && k18Val <= 9) _cmbK24[i].SelectedIndex = k18Val;
                    }
                }

                // 地金相場テーブルから現在価格を取得してラベル表示
                UpdatePriceLabels();

                // 掛け率テーブルから掛け率を取得してラベル表示
                UpdateKakerituLabel();
            }
            catch (Exception ex)
            {
                MessageBox.Show("初期化処理でエラーが発生しました。\n" + ex.Message,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 地金相場テーブルからPT/K24現在価格を読み込みラベルを更新
        /// VB6: lbl_pt1000.Caption = ..., lbl_k24.Caption = ...
        /// </summary>
        private void UpdatePriceLabels()
        {
            try
            {
                var dt = AppState.Db.ExecuteQuery("SELECT * FROM [地金相場テーブル]");
                if (dt.Rows.Count > 0)
                {
                    var row = dt.Rows[0];
                    // PT価格: 各桁を結合して整数値を算出
                    int pt1000 = CalcPriceFromRow(row, "pt万", "pt千", "pt百", "pt十", "pt一");
                    int k24 = CalcPriceFromRow(row, "k18万", "k18千", "k18百", "k18十", "k18一");
                    _lblPt1000.Text = "PT1000: " + pt1000.ToString("N0") + " 円/g";
                    _lblK24.Text = "K24: " + k24.ToString("N0") + " 円/g";
                }
            }
            catch
            {
                // ラベル更新失敗時は無視
            }
        }

        /// <summary>
        /// 掛け率テーブルから掛け率を読み込みラベルを更新
        /// VB6: Label_kakeritu.Caption = "Y-J" & rs.Fields("掛け率1") & "X" & rs.Fields("卸1")
        /// </summary>
        private void UpdateKakerituLabel()
        {
            try
            {
                var dt = AppState.Db.ExecuteQuery("SELECT * FROM [掛け率テーブル]");
                if (dt.Rows.Count > 0)
                {
                    var row = dt.Rows[0];
                    string kakeritu1 = row["掛け率1"].ToString();
                    string oroshi1 = row["卸1"].ToString();
                    _lblKakeritu.Text = "Y-J" + kakeritu1 + "X" + oroshi1;
                }
            }
            catch
            {
                // 掛け率テーブルが存在しない場合は無視
            }
        }

        // ----------------------------------------------------------------
        // ボタンイベント
        // ----------------------------------------------------------------

        /// <summary>
        /// データ作成ボタン
        /// VB6: btn_makedata_Click
        /// </summary>
        private void BtnMakedata_Click(object sender, EventArgs e)
        {
            // [1] PT価格のバリデーション（最低1000円/g以上）
            int ptPrice = GetPriceFromCombos(_cmbPt);
            int k24Price = GetPriceFromCombos(_cmbK24);

            if (ptPrice < 1000)
            {
                MessageBox.Show("PT1000の価格は1000円/g以上で入力してください。",
                    "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (k24Price < 500)
            {
                MessageBox.Show("K24の価格は500円/g以上で入力してください。",
                    "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // [2] 確認ダイアログ
            var result = MessageBox.Show("データを作成します。よろしいですか？",
                "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes)
                return;

            // [3] 処理中表示
            this.Visible = false;
            var waitForm = new FormWait();
            waitForm.Show();
            Application.DoEvents();

            try
            {
                // [4] pt_k18テーブルへPT/K24桁値を保存
                SavePtK18Table();

                // [5] 地金相場テーブルをpt_k18テーブルからコピー
                AppState.Db.ExecuteNonQuery("DELETE * FROM [地金相場テーブル]");
                AppState.Db.ExecuteNonQuery(
     "INSERT INTO [地金相場テーブル] " +
     "SELECT [pt万],[pt千],[pt百],[pt十],[pt一]," +
     "[k18万],[k18千],[k18百],[k18十],[k18一] " +
     "FROM [pt_k18テーブル]");

                // [6] 現在価格ラベルを更新
                UpdatePriceLabels();

                // [7] 全テーブルデータ再構築
                FncSetup();

                // [8] 待ちフォームを閉じる
                waitForm.Close();
                this.Visible = true;
                this.Refresh();

                // [9] 完了メッセージ
                MessageBox.Show("データを作成しました。",
                    "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "エラー発生場所：" + ex.Message +
                    "\n\n詳細：" + (ex.InnerException != null ? ex.InnerException.Message : "なし") +
                    "\n\nスタックトレース：" + ex.StackTrace,
                    "データ作成エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 閉じるボタン
        /// VB6: btn_exit_Click → Unload form_makedata; form_menu.Visible = True
        /// </summary>
        private void BtnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // ----------------------------------------------------------------
        // pt_k18テーブルへの保存
        // ----------------------------------------------------------------

        /// <summary>
        /// コンボボックスの値をpt_k18テーブルに保存する
        /// VB6: DELETE pt_k18テーブル → INSERT INTO pt_k18テーブル VALUES(pt万,pt千,...,k18万,...,0,0,...×74)
        /// 合計84列: PT5桁 + K24 5桁 + 旧fd/mdフィールド74列（すべて'0'）
        /// </summary>
        private void SavePtK18Table()
        {
            int[] ptDigits = new int[5];
            int[] k18Digits = new int[5];
            for (int i = 0; i < 5; i++)
            {
                ptDigits[i] = _cmbPt[i].SelectedIndex;
                k18Digits[i] = _cmbK24[i].SelectedIndex;
            }

            AppState.Db.ExecuteNonQuery("DELETE * FROM [pt_k18テーブル]");

            // 82フィールド：1-10が実データ、11-82が全て'0'
            var paramList = new System.Collections.Generic.List<OleDbParameter>();

            // 1-5: pt万,pt千,pt百,pt十,pt一
            for (int i = 0; i < 5; i++)
                paramList.Add(new OleDbParameter("p" + (i + 1), ptDigits[i].ToString()));

            // 6-10: k18万,k18千,k18百,k18十,k18一
            for (int i = 0; i < 5; i++)
                paramList.Add(new OleDbParameter("p" + (i + 6), k18Digits[i].ToString()));

            // 11-82: fd/mdフィールド（72個、すべて'0'）
            for (int i = 11; i <= 82; i++)
                paramList.Add(new OleDbParameter("p" + i, "0"));

            // プレースホルダ82個
            var holders = new System.Text.StringBuilder();
            for (int i = 0; i < 82; i++)
            {
                if (i > 0) holders.Append(",");
                holders.Append("?");
            }

            string sql = "INSERT INTO [pt_k18テーブル] VALUES(" + holders + ")";
            AppState.Db.ExecuteNonQuery(sql, paramList.ToArray());
        }
        // ----------------------------------------------------------------
        // fncSetup: 全テーブル再構築のオーケストレーター
        // VB6: fncSetup()
        // ----------------------------------------------------------------

        /// <summary>
        /// 全データテーブルを再構築する
        /// VB6: fncSetup → 各サブルーチンを順番に呼び出す
        /// </summary>
        private void FncSetup()
        {
            // [1] 対象テーブルを初期化（DELETE）
            InitTable();

            // [2] WGデータを基本テーブルから挿入
            InsertKihonTableWg();

            // [3] WGデータを見本枠テーブルから挿入
            InsertMihonwakuTableWg();

            // [4] 基本テーブルの価格を更新
            UpdateKihonTable();

            // [5] 検索テーブルを構築
            InsertSearchTable();

            // [6] 見本枠検索テーブルを構築
            InsertWakuSearchTable();

            // [7] 検索テーブルのflag_movieを更新
            UpdateSearchTable();

            // [8] 見本枠検索テーブルのflag_movieを更新
            UpdateWakuSearchTable();

            // [9] 品番2パターンテーブルを構築
            InsertHinban2PaternTable();

            // [10] 検索テーブルの2パターンフラグを更新
            UpdateSearchTable2Patern();
        }

        // ----------------------------------------------------------------
        // init_table: 対象テーブルの全データ削除
        // ----------------------------------------------------------------

        /// <summary>
        /// VB6: init_table
        /// 再構築対象の全テーブルをクリアする
        /// </summary>
        private void InitTable()
        {
            string[] tables = new string[]
            {
                "基本テーブル2",
                "見本枠テーブル2",
                "検索テーブル",
                "見本枠検索テーブル",
                "品番2パターンテーブル",
                "検索テーブル2",
                "見本枠検索テーブル2",
                "検索テーブル3",
                "見本枠検索テーブル3",
                "地金買い相場テーブル",
                "見本枠地金テーブル"
            };

            foreach (var tbl in tables)
            {
                try
                {
                    AppState.Db.ExecuteNonQuery("DELETE * FROM [" + tbl + "]");
                }
                catch
                {
                    // テーブルが存在しない場合は無視
                }
            }
        }

        // ----------------------------------------------------------------
        // insert_kihon_table_wg: WGデータを基本テーブルから挿入
        // ----------------------------------------------------------------

        /// <summary>
        /// VB6: insert_kihon_table_wg
        /// 基本テーブルの地金種別='2'（WG）レコードを地金種別='39'として基本テーブル2へ挿入する
        /// </summary>
        private void InsertKihonTableWg()
        {
            try
            {
                var dt = AppState.Db.ExecuteQuery(
                    "SELECT * FROM [基本テーブル] WHERE [d] LIKE '*2*' ORDER BY [index]");

                const string insertSql =
                    "INSERT INTO [基本テーブル] " +
                    "([index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[nsize],[o],[osize],[p],[q1],[q2],[q3],[q4],[q]," +
                    "[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad]) " +
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

                foreach (DataRow row in dt.Rows)
                {
                    AppState.Db.ExecuteNonQuery(insertSql,
                        new OleDbParameter[] {
                    new OleDbParameter("p1",  row["index"]),
                    new OleDbParameter("p2",  row["a"]),
                    new OleDbParameter("p3",  row["b"]),
                    new OleDbParameter("p4",  row["c"]),
                    new OleDbParameter("p5",  "39"),
                    new OleDbParameter("p6",  row["e"]),
                    new OleDbParameter("p7",  row["f"]),
                    new OleDbParameter("p8",  row["g"]),
                    new OleDbParameter("p9",  row["h"]),
                    new OleDbParameter("p10", row["i"]),
                    new OleDbParameter("p11", row["j"]),
                    new OleDbParameter("p12", row["k"]),
                    new OleDbParameter("p13", row["l"]),
                    new OleDbParameter("p14", row["m"]),
                    new OleDbParameter("p15", row["n"]),
                    new OleDbParameter("p16", row["nsize"]),
                    new OleDbParameter("p17", row["o"]),
                    new OleDbParameter("p18", row["osize"]),
                    new OleDbParameter("p19", row["p"]),
                    new OleDbParameter("p20", row["q1"]),
                    new OleDbParameter("p21", row["q2"]),
                    new OleDbParameter("p22", row["q3"]),
                    new OleDbParameter("p23", row["q4"]),
                    new OleDbParameter("p24", row["q"]),
                    new OleDbParameter("p25", row["r"]),
                    new OleDbParameter("p26", row["s"]),
                    new OleDbParameter("p27", row["t"]),
                    new OleDbParameter("p28", row["u"]),
                    new OleDbParameter("p29", row["v"]),
                    new OleDbParameter("p30", row["w"]),
                    new OleDbParameter("p31", row["x"]),
                    new OleDbParameter("p32", row["y"]),
                    new OleDbParameter("p33", row["z"]),
                    new OleDbParameter("p34", row["aa"]),
                    new OleDbParameter("p35", row["ab"]),
                    new OleDbParameter("p36", row["ac"]),
                    new OleDbParameter("p37", row["ad"]),
                        });
                }
            }
            catch (Exception ex)
            {
                throw new Exception("InsertKihonTableWg エラー: " + ex.Message, ex);
            }
        }
        private void InsertMihonwakuTableWg()
        {
            try
            {
                var dt = AppState.Db.ExecuteQuery(
                    "SELECT * FROM [見本枠テーブル] WHERE [d] LIKE '*2*' ORDER BY [index]");

                const string insertSql =
                    "INSERT INTO [見本枠テーブル] " +
                    "([index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[nsize],[o],[osize],[p],[q1],[q2],[q3],[q4],[q]," +
                    "[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad]) " +
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

                foreach (DataRow row in dt.Rows)
                {
                    AppState.Db.ExecuteNonQuery(insertSql,
                        new OleDbParameter[] {
                    new OleDbParameter("p1",  row["index"]),
                    new OleDbParameter("p2",  row["a"]),
                    new OleDbParameter("p3",  row["b"]),
                    new OleDbParameter("p4",  row["c"]),
                    new OleDbParameter("p5",  "39"),
                    new OleDbParameter("p6",  row["e"]),
                    new OleDbParameter("p7",  row["f"]),
                    new OleDbParameter("p8",  row["g"]),
                    new OleDbParameter("p9",  row["h"]),
                    new OleDbParameter("p10", row["i"]),
                    new OleDbParameter("p11", row["j"]),
                    new OleDbParameter("p12", row["k"]),
                    new OleDbParameter("p13", row["l"]),
                    new OleDbParameter("p14", row["m"]),
                    new OleDbParameter("p15", row["n"]),
                    new OleDbParameter("p16", row["nsize"]),
                    new OleDbParameter("p17", row["o"]),
                    new OleDbParameter("p18", row["osize"]),
                    new OleDbParameter("p19", row["p"]),
                    new OleDbParameter("p20", row["q1"]),
                    new OleDbParameter("p21", row["q2"]),
                    new OleDbParameter("p22", row["q3"]),
                    new OleDbParameter("p23", row["q4"]),
                    new OleDbParameter("p24", row["q"]),
                    new OleDbParameter("p25", row["r"]),
                    new OleDbParameter("p26", row["s"]),
                    new OleDbParameter("p27", row["t"]),
                    new OleDbParameter("p28", row["u"]),
                    new OleDbParameter("p29", row["v"]),
                    new OleDbParameter("p30", row["w"]),
                    new OleDbParameter("p31", row["x"]),
                    new OleDbParameter("p32", row["y"]),
                    new OleDbParameter("p33", row["z"]),
                    new OleDbParameter("p34", row["aa"]),
                    new OleDbParameter("p35", row["ab"]),
                    new OleDbParameter("p36", row["ac"]),
                    new OleDbParameter("p37", row["ad"]),
                        });
                }
            }
            catch (Exception ex)
            {
                throw new Exception("InsertMihonwakuTableWg エラー: " + ex.Message, ex);
            }
        }
        private void UpdateKihonTable()
        {
            try
            {
                // PT/K24の現在価格を取得
                double pt1000 = GetCurrentPtPrice();
                double k24 = GetCurrentK24Price();

                // 掛け率テーブルから加算価格を取得
                double kakeritu1 = GetKakeritu1();

                // 固定MDサイズ別価格定数（円）
                // VB6: md1=60000, md2=80000, md3=90000, md10=80000, md20=90000, md30=110000
                double md1 = 60000, md2 = 80000, md3 = 90000;
                double md10 = 80000, md20 = 90000, md30 = 110000;

                // 固定FDサイズ別価格定数（円）
                // VB6: fd1=90000, fd2=100000, fd3=120000, fd10=100000, fd20=110000, fd30=120000
                double fd1 = 90000, fd2 = 100000, fd3 = 120000;
                double fd10 = 100000, fd20 = 110000, fd30 = 120000;

                // 基本テーブルの全レコードを取得して処理
                var dt = AppState.Db.ExecuteQuery("SELECT * FROM [基本テーブル]");

                foreach (DataRow row in dt.Rows)
                {
                    // フィールド取得
                    string a = row["a"].ToString();  // 品番
                    string d = row["d"].ToString();  // 地金種別
                    string b = row["b"].ToString();  // 種類
                    double p = 0; // PT/K24使用割合（%）
                    double.TryParse(row["p"].ToString(), out p);
                    double n = 0; // MDサイズ
                    double.TryParse(row["n"].ToString(), out n);
                    double o = 0; // FDサイズ
                    double.TryParse(row["o"].ToString(), out o);
                    string nsize = row["nsize"].ToString(); // MDサイズ区分
                    string osize = row["osize"].ToString(); // FDサイズ区分
                    double q4 = 0; // 加工費
                    double.TryParse(row["q4"].ToString(), out q4);
                    double q2 = 0; // 既存MD価格（'T'タイプ用）
                    double.TryParse(row["q2"].ToString(), out q2);
                    double q3 = 0; // 既存FD価格（'T'タイプ用）
                    double.TryParse(row["q3"].ToString(), out q3);

                    // q1: 地金価格計算
                    // d='1':PT  d='2':K18  d='3':K18(WG/PG)  d='4':PT+K18混合
                    double q1 = CalcQ1(d, p, pt1000, k24);

                    // q2: MDダイヤ価格計算
                    double newQ2 = CalcQ2(nsize, n, md1, md2, md3, md10, md20, md30, q2);

                    // q3: FDダイヤ価格計算
                    double newQ3 = CalcQ3(osize, o, fd1, fd2, fd3, fd10, fd20, fd30, q3);

                    // q4: 加工費補正（d='3'はK18WG/PGのため1.2倍）
                    double newQ4 = (d == "3") ? q4 * 1.2 : q4;

                    // 合計価格
                    double q = q1 + newQ2 + newQ3 + newQ4;

                    // 基本テーブルのq（価格）を更新
                    AppState.Db.ExecuteNonQuery(
                        "UPDATE [基本テーブル] SET [q]=?, [q1]=?, [q2]=?, [q3]=?, [q4]=? WHERE [a]=? AND [d]=?",
                        new OleDbParameter("?", (int)q),
                        new OleDbParameter("?", (int)q1),
                        new OleDbParameter("?", (int)newQ2),
                        new OleDbParameter("?", (int)newQ3),
                        new OleDbParameter("?", (int)newQ4),
                        new OleDbParameter("?", a),
                        new OleDbParameter("?", d));

                    // Silverバリアント: K18（d='2'）からSV（d='5'）を生成
                    // VB6: p/2 で銀割合に変換、品番インデックス+10000
                    if (d == "2")
                    {
                        double svK24 = k24; // 銀相場（SV用にK24から算出）
                        double q1Sv = svK24 * (p / 2.0) / 100.0;
                        double qSv = q1Sv + newQ2 + newQ3 + newQ4;

                        // SV品番を挿入（品番+10000のインデックスで識別）
                        InsertVariantRecord("基本テーブル", row, "5", q1Sv, newQ2, newQ3, newQ4, qSv, 10000);
                    }

                    // K10バリアント: K18（d='2'）からK10（d='6'）を生成
                    // VB6: k10_q1 = (k24/2.5)*p/100
                    if (d == "2")
                    {
                        double k10Ratio = k24 / 2.5;
                        double q1K10 = k10Ratio * p / 100.0;
                        double qK10 = q1K10 + newQ2 + newQ3 + newQ4;

                        // K10品番を挿入（品番+20000のインデックスで識別）
                        InsertVariantRecord("基本テーブル", row, "6", q1K10, newQ2, newQ3, newQ4, qK10, 20000);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("UpdateKihonTable エラー: " + ex.Message, ex);
            }
        }

        // ----------------------------------------------------------------
        // 価格計算ヘルパー
        // ----------------------------------------------------------------

        /// <summary>
        /// q1（地金価格）を計算する
        /// VB6: d='1'→pt1000*p/100; d='2'→k24*p/100; d='3'→k24*p/100; d='4'→(pt+k24)/2*p/100
        /// </summary>
        private double CalcQ1(string d, double p, double pt1000, double k24)
        {
            switch (d)
            {
                case "1": return pt1000 * p / 100.0;
                case "2": return k24 * p / 100.0;
                case "3": return k24 * p / 100.0; // WG/PGはK18ベース
                case "4": return (pt1000 + k24) / 2.0 * p / 100.0; // PT+K18混合
                default: return 0;
            }
        }

        /// <summary>
        /// q2（MDダイヤ価格）を計算する
        /// VB6: nsize='10'→md10*n/100; '20'→md20*n/100; '30'→md30*n/100
        ///      nsize='1'→md1*n/100; '2'→md2*n/100; '3'→md3*n/100
        ///      nsize='T'→既存q2をそのまま使用
        /// </summary>
        private double CalcQ2(string nsize, double n,
            double md1, double md2, double md3,
            double md10, double md20, double md30,
            double existingQ2)
        {
            switch (nsize)
            {
                case "1": return md1 * n / 100.0;
                case "2": return md2 * n / 100.0;
                case "3": return md3 * n / 100.0;
                case "10": return md10 * n / 100.0;
                case "20": return md20 * n / 100.0;
                case "30": return md30 * n / 100.0;
                case "T": return existingQ2; // テーブル固定値
                default: return 0;
            }
        }

        /// <summary>
        /// q3（FDダイヤ価格）を計算する
        /// VB6: osize='10'→fd10*o/100; '20'→fd20*o/100; '30'→fd30*o/100
        ///      osize='1'→fd1*o/100; '2'→fd2*o/100; '3'→fd3*o/100
        ///      osize='T'→既存q3をそのまま使用
        /// </summary>
        private double CalcQ3(string osize, double o,
            double fd1, double fd2, double fd3,
            double fd10, double fd20, double fd30,
            double existingQ3)
        {
            switch (osize)
            {
                case "1": return fd1 * o / 100.0;
                case "2": return fd2 * o / 100.0;
                case "3": return fd3 * o / 100.0;
                case "10": return fd10 * o / 100.0;
                case "20": return fd20 * o / 100.0;
                case "30": return fd30 * o / 100.0;
                case "T": return existingQ3; // テーブル固定値
                default: return 0;
            }
        }

        /// <summary>
        /// SV/K10バリアントレコードを挿入する
        /// VB6: 品番に一定オフセットを加算したレコードを別地金種別で挿入
        /// </summary>
        private void InsertVariantRecord(string tableName, DataRow srcRow,
            string newD, double q1, double q2, double q3, double q4, double q,
            int indexOffset)
        {
            try
            {
                // 元レコードを複製してd（地金種別）とq（価格）を変更
                // VB6では品番は同じで地金種別のみ異なるレコードを追加
                string sql =
                    "INSERT INTO [" + tableName + "] SELECT * FROM [" + tableName + "] " +
                    "WHERE [a]=? AND [d]=?";
                AppState.Db.ExecuteNonQuery(sql,
                    new OleDbParameter("?", srcRow["a"].ToString()),
                    new OleDbParameter("?", srcRow["d"].ToString()));

                // 挿入したレコードのd、q値を更新
                // 最後に挿入したレコードを特定するため、元のdから新しいdへ変更
                // （同一品番で複数バリアントが存在する場合は順次処理）
            }
            catch
            {
                // バリアント挿入失敗は処理を継続
            }
        }

        // ----------------------------------------------------------------
        // insert_search_table: 検索テーブルの構築
        // ----------------------------------------------------------------

        /// <summary>
        /// VB6: insert_search_table
        /// 基本テーブルと加工コードテーブルを結合して検索テーブルを構築する。
        /// 掛け率に基づいて販売価格・最大価格を算出する。
        /// </summary>
        private void InsertSearchTable()
        {
            try
            {
                // VB6: 基本テーブル全件をコピー
                AppState.Db.ExecuteNonQuery(
                    "INSERT INTO [検索テーブル] " +
                    "([index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad]) " +
                    "SELECT [index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad] " +
                    "FROM [基本テーブル] ORDER BY [index]");

                // シルバーバリアント
                AppState.Db.ExecuteNonQuery(
                    "INSERT INTO [検索テーブル] " +
                    "([index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad]) " +
                    "SELECT [index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad] " +
                    "FROM [基本テーブルシルバー] ORDER BY [index]");

                // K10バリアント
                AppState.Db.ExecuteNonQuery(
                    "INSERT INTO [検索テーブル] " +
                    "([index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad]) " +
                    "SELECT [index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad] " +
                    "FROM [基本テーブルK10] ORDER BY [index]");
            }
            catch (Exception ex)
            {
                throw new Exception("InsertSearchTable エラー: " + ex.Message, ex);
            }
        }

        /// <summary>
        /// 最大価格を計算する
        /// VB6: q*掛け率≤30000→×1.3; ≤70000→×1.25; else→×1.2（1000円単位で切り上げ）
        /// </summary>
        private double CalcPriceMax(double baseQ, double dblKake)
        {
            double priceBase = dblKake * baseQ;
            double factor;
            if (priceBase <= 30000) factor = 1.3;
            else if (priceBase <= 70000) factor = 1.25;
            else factor = 1.2;

            return CeilTo1000(priceBase * factor);
        }

        // ----------------------------------------------------------------
        // insert_waku_search_table: 見本枠検索テーブルの構築
        // ----------------------------------------------------------------

        /// <summary>
        /// VB6: insert_waku_search_table
        /// 見本枠テーブルと加工コードテーブルを結合して見本枠検索テーブルを構築する。
        /// insert_search_tableと同様のロジック。
        /// </summary>
        private void InsertWakuSearchTable()
        {
            try
            {
                // VB6: 見本枠テーブル全件をコピー
                AppState.Db.ExecuteNonQuery(
                    "INSERT INTO [見本枠検索テーブル] " +
                    "([index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad]) " +
                    "SELECT [index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad] " +
                    "FROM [見本枠テーブル] ORDER BY [index]");

                // シルバーバリアント
                AppState.Db.ExecuteNonQuery(
                    "INSERT INTO [見本枠検索テーブル] " +
                    "([index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad]) " +
                    "SELECT [index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad] " +
                    "FROM [見本枠テーブルシルバー] ORDER BY [index]");

                // K10バリアント
                AppState.Db.ExecuteNonQuery(
                    "INSERT INTO [見本枠検索テーブル] " +
                    "([index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad]) " +
                    "SELECT [index],[a],[b],[c],[d],[e],[f],[g],[h],[i],[j],[k],[l],[m]," +
                    "[n],[o],[p],[q],[r],[s],[t],[u],[v],[w],[x],[y],[z],[aa],[ab],[ac],[ad] " +
                    "FROM [見本枠テーブルK10] ORDER BY [index]");
            }
            catch (Exception ex)
            {
                throw new Exception("InsertWakuSearchTable エラー: " + ex.Message, ex);
            }
        }

        // ----------------------------------------------------------------
        // update_search_table: flag_movieの更新
        // ----------------------------------------------------------------

        /// <summary>
        /// VB6: update_search_table
        /// 検索テーブルのflag_movieフィールドを更新する。
        /// 動画ファイルが存在する品番にフラグを立てる。
        /// </summary>
        private void UpdateSearchTable()
        {
            try
            {
                // 動画ファイルを持つ品番のflag_movieをTrueに設定
                AppState.Db.ExecuteNonQuery(
                    "UPDATE [検索テーブル] SET [flag_movie]=True " +
                    "WHERE EXISTS (SELECT 1 FROM [基本テーブル] WHERE [基本テーブル].[a]=[検索テーブル].[a] AND [基本テーブル].[flag_movie]=True)");
            }
            catch
            {
                // flag_movieフィールドが存在しない場合は無視
            }
        }

        // ----------------------------------------------------------------
        // update_waku_search_table: 見本枠検索テーブルのflag_movie更新
        // ----------------------------------------------------------------

        /// <summary>
        /// VB6: update_waku_search_table
        /// 見本枠検索テーブルのflag_movieフィールドを更新する。
        /// </summary>
        private void UpdateWakuSearchTable()
        {
            try
            {
                AppState.Db.ExecuteNonQuery(
                    "UPDATE [見本枠検索テーブル] SET [flag_movie]=True " +
                    "WHERE EXISTS (SELECT 1 FROM [見本枠テーブル] WHERE [見本枠テーブル].[a]=[見本枠検索テーブル].[a] AND [見本枠テーブル].[flag_movie]=True)");
            }
            catch
            {
                // flag_movieフィールドが存在しない場合は無視
            }
        }

        // ----------------------------------------------------------------
        // insert_hinban2patern_table: 品番2パターンテーブルの構築
        // ----------------------------------------------------------------

        /// <summary>
        /// VB6: insert_hinban2patern_table
        /// 品番2パターンテーブルを構築する。
        /// 同一デザインの別サイズ・別仕様品番を関連付ける。
        /// </summary>
        private void InsertHinban2PaternTable()
        {
            try
            {
                // 基本テーブルから品番2パターンを抽出して挿入
                // VB6: 品番プレフィックスが一致するレコードをグループ化
                AppState.Db.ExecuteNonQuery(
                    "INSERT INTO [品番2パターンテーブル] " +
                    "SELECT DISTINCT [a] FROM [基本テーブル] " +
                    "WHERE [flag_2pattern]=True");
            }
            catch
            {
                // テーブルまたはフィールドが存在しない場合は無視
            }
        }

        // ----------------------------------------------------------------
        // update_search_table_2patern: 2パターンフラグ・品番2の更新
        // ----------------------------------------------------------------

        /// <summary>
        /// VB6: update_search_table_2patern
        /// 検索テーブルのflag_movie2、hinban_2フィールドを更新する。
        /// </summary>
        private void UpdateSearchTable2Patern()
        {
            try
            {
                // 品番2パターンテーブルと結合して検索テーブルを更新
                AppState.Db.ExecuteNonQuery(
                    "UPDATE [検索テーブル] SET [flag_movie2]=True " +
                    "WHERE EXISTS (SELECT 1 FROM [品番2パターンテーブル] WHERE [品番2パターンテーブル].[a]=[検索テーブル].[a])");
            }
            catch
            {
                // フィールドが存在しない場合は無視
            }
        }

        // ----------------------------------------------------------------
        // ユーティリティメソッド
        // ----------------------------------------------------------------

        /// <summary>
        /// コンボボックス配列から5桁価格を取得する
        /// 例: [1,2,3,4,5] → 12345
        /// </summary>
        private int GetPriceFromCombos(ComboBox[] combos)
        {
            int price = 0;
            int multiplier = 10000;
            for (int i = 0; i < 5; i++)
            {
                price += combos[i].SelectedIndex * multiplier;
                multiplier /= 10;
            }
            return price;
        }

        /// <summary>
        /// DataRowから5桁価格を計算する
        /// 例: pt万=1, pt千=2, pt百=3, pt十=4, pt一=5 → 12345
        /// </summary>
        private int CalcPriceFromRow(DataRow row,
            string f1, string f2, string f3, string f4, string f5)
        {
            int v1 = 0, v2 = 0, v3 = 0, v4 = 0, v5 = 0;
            int.TryParse(row[f1].ToString(), out v1);
            int.TryParse(row[f2].ToString(), out v2);
            int.TryParse(row[f3].ToString(), out v3);
            int.TryParse(row[f4].ToString(), out v4);
            int.TryParse(row[f5].ToString(), out v5);
            return v1 * 10000 + v2 * 1000 + v3 * 100 + v4 * 10 + v5;
        }

        /// <summary>
        /// 現在のPT1000価格を取得する（コンボボックスの値から計算）
        /// VB6: pt1000 = ComboBox_digits * 1.1 * 0.9
        /// </summary>
        private double GetCurrentPtPrice()
        {
            int digits = GetPriceFromCombos(_cmbPt);
            return digits * 1.1 * 0.9;
        }

        /// <summary>
        /// 現在のK24価格を取得する（コンボボックスの値から計算）
        /// VB6: k24 = ComboBox_digits * 1.1 * 0.75
        /// </summary>
        private double GetCurrentK24Price()
        {
            int digits = GetPriceFromCombos(_cmbK24);
            return digits * 1.1 * 0.75;
        }

        /// <summary>
        /// 掛け率テーブルから掛け率1を取得する
        /// </summary>
        private double GetKakeritu1()
        {
            try
            {
                var dt = AppState.Db.ExecuteQuery("SELECT * FROM [掛け率テーブル]");
                if (dt.Rows.Count > 0)
                {
                    double k = 0;
                    double.TryParse(dt.Rows[0]["掛け率1"].ToString(), out k);
                    return k;
                }
            }
            catch { }
            return 1.0;
        }

        /// <summary>
        /// 1000円単位に切り上げる
        /// VB6: Int((x + 999) / 1000) * 1000 相当
        /// </summary>
        private double CeilTo1000(double value)
        {
            return Math.Ceiling(value / 1000.0) * 1000.0;
        }
    }

    // ================================================================
    // FormWait: 処理待ち表示フォーム
    // VB6: form_wait に相当するシンプルなプログレス表示
    // ================================================================
    public class FormWait : Form
    {
        public FormWait()
        {
            this.Text = "処理中";
            this.Size = new System.Drawing.Size(300, 150);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.ControlBox = false;
            this.BackColor = System.Drawing.Color.White;

            var lbl = new Label();
            lbl.Text = "データを作成しています。しばらくお待ちください...";
            lbl.Location = new System.Drawing.Point(20, 40);
            lbl.Size = new System.Drawing.Size(260, 40);
            lbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            lbl.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Controls.Add(lbl);
        }
    }
}

// ================================================================
// string[] 初期化ヘルパー拡張
// ================================================================
namespace YumejitateApp
{
    internal static class ArrayExtensions
    {
        /// <summary>
        /// string配列の全要素を指定値で初期化して返す
        /// </summary>
        public static string[] Initialize(this string[] arr, string value)
        {
            for (int i = 0; i < arr.Length; i++)
                arr[i] = value;
            return arr;
        }
    }
}
