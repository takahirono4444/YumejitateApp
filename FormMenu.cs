using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace YumejitateApp
{
    // ================================================================
    // FormMenu.cs
    // VB6: form_menu.frm の移植
    // メインメニュー画面
    // ================================================================
    public class FormMenu : Form
    {
        // ----------------------------------------------------------------
        // 定数
        // ----------------------------------------------------------------

        /// <summary>DBファイルの絶対パス</summary>
        private const string DB_PATH =
            @"C:\Users\yumejitate\Desktop\YumejitateApp\夢仕立て.mdb";

        /// <summary>ソフトウェアバージョン（VB6の Label1 相当）</summary>
        private const string SOFTWARE_VERSION = "Soft 2023/11/10";

        // ----------------------------------------------------------------
        // フィールド
        // ----------------------------------------------------------------

        /// <summary>DB接続管理（VB6の dao_database 相当）</summary>
        private DatabaseManager _db;

        // コントロール
        private TableLayoutPanel _tableLayout;
        private Panel _bottomPanel;
        private Label _lblVersion;
        private Label _lblDataDate;

        // ボタン（VB6のボタン名と対応させる）
        private Button _btnData;           // btn_data          データ作成
        private Button _btnMihon;          // btn_mihon          H1商品検索
        private Button _btnDejicame;       // btn_dejicame       デジカメ画像
        private Button _btnGousei;         // btn_gousei         画像編集
        private Button _btnJigane;         // btn_jigane         地金相場
        private Button _btnSearch;         // btn_search         AB商品検索
        private Button _btnCamera;         // btn_camera         カメラモニタ
        private Button _btnGazouDisp;      // btn_gazou_disp     画像表示
        private Button _btnTedukuri;       // btn_tedukuri       オーダーガイド
        private Button _btnHinban;         // btn_hinban         AB品番検索
        private Button _btnPrint;          // btn_print          画像印刷
        private Button _btnNr;             // btn_nr             NRコレクション検索
        private Button _btnPrintTedukuri;  // btn_print_tedukuri オーダーガイド印刷
        private Button _btnRepair;         // btn_repair         修理
        private Button _btnKaisaku;        // btn_kaisaku        チェーン
        private Button _btnExit;           // btn_exit           終了

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------

        public FormMenu()
        {
            InitializeComponent();
        }

        // ----------------------------------------------------------------
        // コントロール初期化
        // ----------------------------------------------------------------

        private void InitializeComponent()
        {
            // ---- フォーム基本設定 ----------------------------------------
            this.Text = "夢仕立て - メニュー画面";
            this.BackColor = Color.White;
            this.WindowState = FormWindowState.Maximized;
            this.StartPosition = FormStartPosition.WindowsDefaultBounds;
            this.Font = new Font("メイリオ", 14f, FontStyle.Bold);

            this.Load += new EventHandler(FormMenu_Load);
            this.FormClosing += new FormClosingEventHandler(FormMenu_FormClosing);

            // ---- ボタンの生成 --------------------------------------------
            _btnData = CreateButton("データ作成", Color.FromArgb(220, 235, 252));
            _btnMihon = CreateButton("H1商品検索", Color.FromArgb(220, 235, 252));
            _btnDejicame = CreateButton("デジカメ画像", Color.FromArgb(220, 235, 252));
            _btnGousei = CreateButton("画像編集", Color.FromArgb(220, 235, 252));

            _btnJigane = CreateButton("地金相場", Color.FromArgb(252, 235, 220));
            _btnSearch = CreateButton("AB商品検索", Color.FromArgb(252, 235, 220));
            _btnCamera = CreateButton("カメラモニタ", Color.FromArgb(252, 235, 220));
            _btnGazouDisp = CreateButton("画像表示", Color.FromArgb(252, 235, 220));

            _btnTedukuri = CreateButton("オーダー\nガイド", Color.FromArgb(235, 252, 220));
            _btnHinban = CreateButton("AB品番検索", Color.FromArgb(235, 252, 220));
            _btnPrint = CreateButton("画像印刷", Color.FromArgb(235, 252, 220));
            _btnNr = CreateButton("NRコレクション\n検索", Color.FromArgb(235, 252, 220));

            _btnPrintTedukuri = CreateButton("オーダー\nガイド印刷", Color.FromArgb(252, 252, 200));
            _btnRepair = CreateButton("修理", Color.FromArgb(252, 252, 200));
            _btnKaisaku = CreateButton("チェーン", Color.FromArgb(252, 252, 200));
            _btnExit = CreateButton("終了", Color.FromArgb(255, 180, 180));

            // ---- クリックイベント登録 ------------------------------------
            _btnData.Click += new EventHandler(BtnData_Click);
            _btnMihon.Click += new EventHandler(BtnMihon_Click);
            _btnDejicame.Click += new EventHandler(BtnDejicame_Click);
            _btnGousei.Click += new EventHandler(BtnGousei_Click);
            _btnJigane.Click += new EventHandler(BtnJigane_Click);
            _btnSearch.Click += new EventHandler(BtnSearch_Click);
            _btnCamera.Click += new EventHandler(BtnCamera_Click);
            _btnGazouDisp.Click += new EventHandler(BtnGazouDisp_Click);
            _btnTedukuri.Click += new EventHandler(BtnTedukuri_Click);
            _btnHinban.Click += new EventHandler(BtnHinban_Click);
            _btnPrint.Click += new EventHandler(BtnPrint_Click);
            _btnNr.Click += new EventHandler(BtnNr_Click);
            _btnPrintTedukuri.Click += new EventHandler(BtnPrintTedukuri_Click);
            _btnRepair.Click += new EventHandler(BtnRepair_Click);
            _btnKaisaku.Click += new EventHandler(BtnKaisaku_Click);
            _btnExit.Click += new EventHandler(BtnExit_Click);

            // ---- TableLayoutPanel（4列 × 4行）---------------------------
            // VB6の座標から復元したグループ配置:
            //   行0: データ作成 / H1商品検索 / デジカメ画像 / 画像編集
            //   行1: 地金相場   / AB商品検索 / カメラモニタ / 画像表示
            //   行2: オーダーガイド / AB品番検索 / 画像印刷 / NRコレクション検索
            //   行3: オーダーガイド印刷 / 修理 / チェーン / 終了
            _tableLayout = new TableLayoutPanel();
            _tableLayout.Dock = DockStyle.Fill;
            _tableLayout.ColumnCount = 4;
            _tableLayout.RowCount = 4;
            _tableLayout.Padding = new Padding(20, 20, 20, 10);
            _tableLayout.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;

            // 各列を均等幅に設定
            for (int i = 0; i < 4; i++)
            {
                _tableLayout.ColumnStyles.Add(
                    new ColumnStyle(SizeType.Percent, 25f));
            }
            // 各行を均等高さに設定
            for (int i = 0; i < 4; i++)
            {
                _tableLayout.RowStyles.Add(
                    new RowStyle(SizeType.Percent, 25f));
            }

            // ボタンを行列順に追加
            // 行0
            _tableLayout.Controls.Add(_btnData, 0, 0);
            _tableLayout.Controls.Add(_btnMihon, 1, 0);
            _tableLayout.Controls.Add(_btnDejicame, 2, 0);
            _tableLayout.Controls.Add(_btnGousei, 3, 0);
            // 行1
            _tableLayout.Controls.Add(_btnJigane, 0, 1);
            _tableLayout.Controls.Add(_btnSearch, 1, 1);
            _tableLayout.Controls.Add(_btnCamera, 2, 1);
            _tableLayout.Controls.Add(_btnGazouDisp, 3, 1);
            // 行2
            _tableLayout.Controls.Add(_btnTedukuri, 0, 2);
            _tableLayout.Controls.Add(_btnHinban, 1, 2);
            _tableLayout.Controls.Add(_btnPrint, 2, 2);
            _tableLayout.Controls.Add(_btnNr, 3, 2);
            // 行3
            _tableLayout.Controls.Add(_btnPrintTedukuri, 0, 3);
            _tableLayout.Controls.Add(_btnRepair, 1, 3);
            _tableLayout.Controls.Add(_btnKaisaku, 2, 3);
            _tableLayout.Controls.Add(_btnExit, 3, 3);

            // ---- 下部パネル（バージョン・日付ラベル）--------------------
            _bottomPanel = new Panel();
            _bottomPanel.Dock = DockStyle.Bottom;
            _bottomPanel.Height = 50;
            _bottomPanel.BackColor = Color.White;

            _lblVersion = new Label();
            _lblVersion.Text = SOFTWARE_VERSION;
            _lblVersion.Font = new Font("メイリオ", 13f, FontStyle.Regular);
            _lblVersion.ForeColor = Color.Gray;
            _lblVersion.AutoSize = true;
            _lblVersion.Location = new Point(20, 12);

            _lblDataDate = new Label();
            _lblDataDate.Text = "Data ----/--/--";
            _lblDataDate.Font = new Font("メイリオ", 13f, FontStyle.Regular);
            _lblDataDate.ForeColor = Color.Gray;
            _lblDataDate.AutoSize = true;
            _lblDataDate.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _lblDataDate.Location = new Point(400, 12);

            _bottomPanel.Controls.Add(_lblVersion);
            _bottomPanel.Controls.Add(_lblDataDate);

            // ---- フォームにコントロールを追加 ----------------------------
            // Bottom を先に追加しないと TableLayoutPanel が Bottom を覆う
            this.Controls.Add(_bottomPanel);
            this.Controls.Add(_tableLayout);
        }

        // ----------------------------------------------------------------
        // ヘルパー：ボタン生成
        // ----------------------------------------------------------------

        /// <summary>共通スタイルのボタンを生成して返す。</summary>
        private Button CreateButton(string caption, Color backColor)
        {
            var btn = new Button();
            btn.Text = caption;
            btn.Dock = DockStyle.Fill;
            btn.BackColor = backColor;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderColor = Color.Silver;
            btn.FlatAppearance.BorderSize = 1;
            btn.Font = new Font("メイリオ", 14f, FontStyle.Bold);
            btn.Margin = new Padding(6);
            btn.Cursor = Cursors.Hand;
            return btn;
        }

        // ================================================================
        // 起動時処理（VB6の Form_Load に相当）
        // ================================================================

        private void FormMenu_Load(object sender, EventArgs e)
        {
            // [1] 多重起動チェック（VB6: If App.PrevInstance Then Unload Me）
            string exeName = Path.GetFileNameWithoutExtension(
                System.Windows.Forms.Application.ExecutablePath);

            if (Process.GetProcessesByName(exeName).Length > 1)
            {
                MessageBox.Show(
                    "システムはすでに立ち上がっています。",
                    "システム二重起動チェック",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                System.Windows.Forms.Application.Exit();
                return;
            }

            // [2] グローバルフラグの初期化（VB6: Flag_Hinban = False など）
            AppState.FlagHinban = false;
            AppState.FlagMihon = false;
            AppState.FlagDispPicture = false;

            // [3] データベース接続（VB6: DBEngine.OpenDatabase(... ";pwd=shigeshi")）
            try
            {
                _db = new DatabaseManager();
                _db.Connect(DB_PATH);
                AppState.Db = _db;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "データベースへの接続に失敗しました。\n\n" + ex.Message,
                    "DB接続エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
                return;
            }

            // [4] ウィンドウ最大化（Form初期化済みのため念のため再設定）
            this.WindowState = FormWindowState.Maximized;

            // [5] データ更新日付の取得と表示
            //     VB6: SELECT * FROM データ日付テーブル → Label_data.Caption = "Data " & 日付
            try
            {
                DataTable dt = _db.ExecuteQuery("SELECT * FROM [データ日付テーブル]");
                if (dt.Rows.Count > 0)
                {
                    string date = dt.Rows[0]["日付"].ToString();
                    _lblDataDate.Text = "Data " + date;
                }
                else
                {
                    MessageBox.Show(
                        "夢仕立てデータの更新日付が取得できませんでした。",
                        "情報",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch
            {
                // 日付取得失敗は致命的エラーではないので続行
                _lblDataDate.Text = "Data (取得失敗)";
            }
        }

        // ================================================================
        // 終了時処理（VB6の Form_Unload に相当）
        // ================================================================

        private void FormMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_db == null || !_db.IsConnected)
                return;

            try
            {
                // [1] 地金テーブルをリセット（10列すべて '0'）
                //     VB6: delete * from 地金テーブル → insert into 地金テーブル values('0'×10)
                ResetJinkineTable();

                // [2] pt_k18テーブルをリセット（84列すべて '0'）
                //     VB6: delete * from pt_k18テーブル → insert into pt_k18テーブル values('0'×84)
                ResetPtK18Table();
            }
            catch (Exception ex)
            {
                // リセット失敗はログ出力のみ（終了は続行）
                System.Diagnostics.Debug.WriteLine("テーブルリセット失敗: " + ex.Message);
            }

            // [3] データベースクローズ（VB6: dao_database.Close）
            _db.Dispose();
            _db = null;

            // [4] 開いている子フォームをすべて閉じる
            //     VB6: Unload form_mihonwaku / form_hinban / form_search / form_gousei / form_tedukuri
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f != this)
                    f.Close();
            }
        }

        // ----------------------------------------------------------------
        // 地金テーブルのリセット（VB6の Form_Unload 内処理）
        // ----------------------------------------------------------------

        private void ResetJinkineTable()
        {
            // 地金テーブルは10カラム構成
            _db.ExecuteNonQuery("DELETE * FROM [地金テーブル]");

            var prms = new OleDbParameter[10];
            for (int i = 0; i < 10; i++)
                prms[i] = new OleDbParameter("p" + i, "0");

            _db.ExecuteNonQuery(
                "INSERT INTO [地金テーブル] VALUES(?,?,?,?,?,?,?,?,?,?)",
                prms);
        }

        // ----------------------------------------------------------------
        // pt_k18テーブルのリセット（VB6の Form_Unload 内処理）
        // ----------------------------------------------------------------

        private void ResetPtK18Table()
        {
            // pt_k18テーブルは84カラム構成（VB6コードより）
            _db.ExecuteNonQuery("DELETE * FROM [pt_k18テーブル]");

            const int columnCount = 84;
            var prms = new OleDbParameter[columnCount];
            var holders = new string[columnCount];

            for (int i = 0; i < columnCount; i++)
            {
                prms[i] = new OleDbParameter("p" + i, "0");
                holders[i] = "?";
            }

            _db.ExecuteNonQuery(
                "INSERT INTO [pt_k18テーブル] VALUES(" + string.Join(",", holders) + ")",
                prms);
        }

        // ================================================================
        // ボタンクリックハンドラ
        // VB6パターン: form_menu.Visible = False → 子フォーム.Show
        // C#移植: this.Hide() → 子フォームを非モーダル表示 → 閉じたら this.Show()
        // ================================================================

        // ---- 検索系 -----------------------------------------------------

        /// <summary>AB品番検索（VB6: Flag_Hinban = True → form_hinban.Show）</summary>
        private void BtnHinban_Click(object sender, EventArgs e)
        {
            AppState.FlagHinban = true;
            OpenChildForm(new FormHinban());
        }

        /// <summary>AB商品検索（VB6: form_search.Show）</summary>
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormSearch());
        }

        /// <summary>H1商品検索（VB6: Flag_Mihon = True → form_mihonwaku.Show）</summary>
        private void BtnMihon_Click(object sender, EventArgs e)
        {
            AppState.FlagMihon = true;
            OpenChildForm(new FormMihonwaku());
        }

        /// <summary>NRコレクション検索（VB6: form_nr.Show）</summary>
        private void BtnNr_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormNr());
        }

        // ---- 地金・データ -----------------------------------------------

        /// <summary>地金相場（VB6: form_jigane.Show）</summary>
        private void BtnJigane_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormJigane());
        }

        /// <summary>データ作成（VB6: form_makedata.Show）</summary>
        private void BtnData_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormMakedata());
        }

        // ---- オーダー・加工 ---------------------------------------------

        /// <summary>
        /// オーダーガイド（VB6: form_tedukuri.Show）
        /// VB6では開く前に参照データラベルを初期化していた
        /// </summary>
        private void BtnTedukuri_Click(object sender, EventArgs e)
        {
            var form = new FormTedukuri();
            // VB6互換: 参照データの初期化（form_tedukuriのラベルをリセット）
            form.ResetSampleLabels();
            OpenChildForm(form);
        }

        /// <summary>オーダーガイド印刷（VB6: form_print_tedukuri.Show）</summary>
        private void BtnPrintTedukuri_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormPrintTedukuri());
        }

        /// <summary>修理（VB6: form_repair.Show）</summary>
        private void BtnRepair_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormRepair());
        }

        /// <summary>チェーン（VB6: form_kaisaku.Show）</summary>
        private void BtnKaisaku_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormKaisaku());
        }

        // ---- 画像系 -----------------------------------------------------

        /// <summary>画像編集（VB6: form_gousei.Show）</summary>
        private void BtnGousei_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormGousei());
        }

        /// <summary>画像表示（VB6: Flag_Disp_Picture = True → form_disp_picture.Show）</summary>
        private void BtnGazouDisp_Click(object sender, EventArgs e)
        {
            AppState.FlagDispPicture = true;
            OpenChildForm(new FormDispPicture());
        }

        /// <summary>画像印刷（VB6: Form_Print.Show）</summary>
        private void BtnPrint_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormPrint());
        }

        /// <summary>
        /// カメラモニタ（VB6: frmVideo.Show）
        /// ※ VB6は DirectShow / StCamD.Bas 依存。C#では AForge.NET 等への置き換えが必要。
        /// </summary>
        private void BtnCamera_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormVideo());
        }

        /// <summary>
        /// デジカメ画像（VB6: form_capture3.Show）
        /// ※ カメラキャプチャ機能。移行難易度: 高。
        /// </summary>
        private void BtnDejicame_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormCapture());
        }

        // ---- 終了 -------------------------------------------------------

        /// <summary>終了（VB6: Unload form_menu → Form_Unload が発火）</summary>
        private void BtnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // ================================================================
        // フォーム遷移ヘルパー
        // VB6パターン: form_menu.Visible = False → 子フォーム.Show（非モーダル）
        // C#実装: this.Hide() → 子フォーム表示 → 子フォームが閉じたら this.Show()
        // ================================================================

        /// <summary>
        /// 子フォームを非モーダルで開く。
        /// 子フォームが閉じられるとメニューに自動的に戻る。
        /// </summary>
        private void OpenChildForm(Form childForm)
        {
            this.Hide();

            // 子フォームが閉じられたときにメニューに戻る
            childForm.FormClosed += (s, args) =>
            {
                // 子フォームのフラグをリセット
                AppState.FlagHinban = false;
                AppState.FlagMihon = false;
                AppState.FlagDispPicture = false;

                this.Show();
            };

            childForm.Show();
        }
    }

    // ================================================================
    // AppState.cs（グローバル状態管理）
    // VB6の Module（.bas）レベルのグローバル変数に相当
    // ================================================================
    public static class AppState
    {
        /// <summary>品番検索フォームから戻る際の挙動制御フラグ（VB6: Flag_Hinban）</summary>
        public static bool FlagHinban { get; set; }

        /// <summary>見本枠フォームから戻る際の挙動制御フラグ（VB6: Flag_Mihon）</summary>
        public static bool FlagMihon { get; set; }

        /// <summary>画像表示フォームから戻る際の挙動制御フラグ（VB6: Flag_Disp_Picture）</summary>
        public static bool FlagDispPicture { get; set; }

        /// <summary>アプリ全体で共有するDB接続インスタンス（VB6: dao_database）</summary>
        public static DatabaseManager Db { get; set; }

        /// <summary>見本枠検索改良フラグ（VB6: Flag_Mihon_rev1）</summary>
        public static bool FlagMihonRev1 { get; set; }

        /// <summary>見本枠検索改良用トータルサイズ（VB6: Mihon_rev1_tsize）</summary>
        public static double MihonRev1Tsize { get; set; }

        /// <summary>オーダーサンプル関連（VB6グローバル変数）</summary>
        public static double OrderSamplePt900 { get; set; }
        public static double OrderSampleK18 { get; set; }
        public static double OrderSampleWgPg { get; set; }
        public static double OrderSampleK10 { get; set; }
        public static int OrderSampleCode { get; set; }
        public static double OrderSamplePrice { get; set; }
        public static string OrderSampleHinban { get; set; }
    }
}