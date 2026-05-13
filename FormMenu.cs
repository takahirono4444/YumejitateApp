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
    // メインメニュー画面 - リデザイン版
    // ================================================================
    public class FormMenu : Form
    {
        // ----------------------------------------------------------------
        // 定数
        // ----------------------------------------------------------------
        private const string DB_PATH =
            @"C:\Users\yumejitate\Desktop\YumejitateApp\夢仕立て.mdb";
        private const string SOFTWARE_VERSION = "Soft 2023/11/10";

        // ----------------------------------------------------------------
        // フィールド
        // ----------------------------------------------------------------
        private DatabaseManager _db;

        // ---- レイアウト用パネル ----
        private Panel _pnlTitle;   // 上部タイトルエリア
        private Panel _pnlCenter;  // ボタン配置エリア（中央寄せ用）
        private TableLayoutPanel _tableLayout;
        private Panel _pnlBottom;  // 下部情報バー

        // ---- タイトルラベル ----
        private Label _lblSystemTitle; // 夢仕立SYSTEM
        private Label _lblKoboTitle;   // 夢仕立工房
        private PictureBox _picLogo;   // Yマーク

        // ---- 情報ラベル ----
        private Label _lblVersion;     // Soft xxxx/xx/xx
        private Label _lblDataDate;    // Data xxxx/xx/xx

        // ---- ボタン ----
        private Button _btnData;
        private Button _btnMihon;
        private Button _btnDejicame;
        private Button _btnGousei;
        private Button _btnJigane;
        private Button _btnSearch;
        private Button _btnCamera;
        private Button _btnGazouDisp;
        private Button _btnTedukuri;
        private Button _btnHinban;
        private Button _btnPrint;
        private Button _btnNr;
        private Button _btnPrintTedukuri;
        private Button _btnRepair;
        private Button _btnKaisaku;
        private Button _btnExit;

        // ボタン1個のサイズ
        private const int BTN_W = 170;
        private const int BTN_H = 62;
        private const int BTN_MARGIN = 5;

        // ================================================================
        // コンストラクタ
        // ================================================================
        public FormMenu()
        {
            InitializeComponent();
        }



        // ================================================================
        // InitializeComponent
        // ================================================================
        private void InitializeComponent()
        {
            this.Text = "夢仕立て - メニュー画面";
            this.BackColor = Color.FromArgb(0, 172, 172);
            this.WindowState = FormWindowState.Maximized;
            this.StartPosition = FormStartPosition.WindowsDefaultBounds;
            this.Font = new Font("メイリオ", 12f, FontStyle.Bold);
            this.Load += new EventHandler(FormMenu_Load);
            this.FormClosing += new FormClosingEventHandler(FormMenu_FormClosing);
            this.Resize += new EventHandler(FormMenu_Resize);

            // ============================================================
            // 上部タイトルエリア
            // ============================================================
            _pnlTitle = new Panel();
            _pnlTitle.BackColor = Color.FromArgb(18, 32, 80);
            _pnlTitle.Dock = DockStyle.Top;
            _pnlTitle.Height = 0; // Resize で計算

            // 「夢仕立SYSTEM」大タイトル
            _lblSystemTitle = new Label();
            _lblSystemTitle.Text = "夢仕立SYSTEM";
            _lblSystemTitle.Font = new Font("メイリオ", 54f, FontStyle.Bold | FontStyle.Italic);
            _lblSystemTitle.ForeColor = Color.FromArgb(255, 215, 0); // Gold
            _lblSystemTitle.BackColor = Color.Transparent;
            _lblSystemTitle.AutoSize = true;
            _lblSystemTitle.Location = new Point(0, 0); // Resize で配置

            // 「夢仕立工房」サブタイトル
            _lblKoboTitle = new Label();
            _lblKoboTitle.Text = "夢仕立工房";
            _lblKoboTitle.Font = new Font("メイリオ", 24f, FontStyle.Bold);
            _lblKoboTitle.ForeColor = Color.FromArgb(220, 230, 255);
            _lblKoboTitle.BackColor = Color.Transparent;
            _lblKoboTitle.AutoSize = true;
            _lblKoboTitle.Location = new Point(0, 0);

            // Yマーク
            // ロゴ画像
            _picLogo = new PictureBox();
            _picLogo.SizeMode = PictureBoxSizeMode.Zoom;
            _picLogo.BackColor = Color.Transparent;
            _picLogo.Size = new Size(140, 140);
            _picLogo.Location = new Point(0, 0); // LayoutControlsで配置

            string logoPath = Path.Combine(
                Application.StartupPath, "_Y_刻印ロゴ.jpg");
            if (File.Exists(logoPath))
            {
                try { _picLogo.Image = Image.FromFile(logoPath); }
                catch { }
            }

            _pnlTitle.Controls.Add(_picLogo);
            _pnlTitle.Controls.Add(_lblSystemTitle);
            _pnlTitle.Controls.Add(_lblKoboTitle);

            _pnlTitle.Controls.Add(_lblSystemTitle);
            _pnlTitle.Controls.Add(_lblKoboTitle);

            // ============================================================
            // ボタン配置エリア（中央寄せ用コンテナ）
            // ============================================================
            _pnlCenter = new Panel();
            _pnlCenter.BackColor = Color.FromArgb(18, 32, 80);
            _pnlCenter.Location = new Point(0, 0); // Resize で配置

            // ---- TableLayoutPanel（4列×4行） ----
            _tableLayout = new TableLayoutPanel();
            _tableLayout.ColumnCount = 4;
            _tableLayout.RowCount = 4;
            _tableLayout.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;
            _tableLayout.BackColor = Color.FromArgb(18, 32, 80);
            _tableLayout.Location = new Point(0, 0);

            for (int i = 0; i < 4; i++)
            {
                _tableLayout.ColumnStyles.Add(
                    new ColumnStyle(SizeType.Absolute, BTN_W + BTN_MARGIN * 2));
                _tableLayout.RowStyles.Add(
                    new RowStyle(SizeType.Absolute, BTN_H + BTN_MARGIN * 2));
            }

            // テーブルサイズを固定
            _tableLayout.Size = new Size(
                4 * (BTN_W + BTN_MARGIN * 2),
                4 * (BTN_H + BTN_MARGIN * 2));

            // ---- ボタン生成 ----
            _btnData = MakeBtn("データ作成", Color.FromArgb(70, 100, 180));
            _btnMihon = MakeBtn("H1商品検索", Color.FromArgb(70, 100, 180));
            _btnDejicame = MakeBtn("デジカメ画像", Color.FromArgb(70, 100, 180));
            _btnGousei = MakeBtn("画像編集", Color.FromArgb(70, 100, 180));

            _btnJigane = MakeBtn("地金相場", Color.FromArgb(60, 130, 100));
            _btnSearch = MakeBtn("AB商品検索", Color.FromArgb(60, 130, 100));
            _btnCamera = MakeBtn("カメラモニタ", Color.FromArgb(60, 130, 100));
            _btnGazouDisp = MakeBtn("画像表示", Color.FromArgb(60, 130, 100));

            _btnTedukuri = MakeBtn("オーダー\nガイド", Color.FromArgb(140, 90, 50));
            _btnHinban = MakeBtn("AB品番検索", Color.FromArgb(140, 90, 50));
            _btnPrint = MakeBtn("画像印刷", Color.FromArgb(140, 90, 50));
            _btnNr = MakeBtn("NRコレクション\n検索", Color.FromArgb(140, 90, 50));

            _btnPrintTedukuri = MakeBtn("オーダー\nガイド印刷", Color.FromArgb(100, 70, 140));
            _btnRepair = MakeBtn("修理", Color.FromArgb(100, 70, 140));
            _btnKaisaku = MakeBtn("チェーン", Color.FromArgb(100, 70, 140));
            _btnExit = MakeBtn("終了", Color.FromArgb(160, 40, 40));

            // クリックイベント
            _btnData.Click += (s, e) => OpenChildForm(new FormMakedata());
            _btnMihon.Click += (s, e) => { AppState.FlagMihon = true; OpenChildForm(new FormMihonwaku()); };
            _btnDejicame.Click += (s, e) => OpenChildForm(new FormCapture());
            _btnGousei.Click += (s, e) => OpenChildForm(new FormGousei());
            _btnJigane.Click += (s, e) => OpenChildForm(new FormJigane());
            _btnSearch.Click += (s, e) => OpenChildForm(new FormSearch());
            _btnCamera.Click += (s, e) => OpenChildForm(new FormVideo());
            _btnGazouDisp.Click += (s, e) => { AppState.FlagDispPicture = true; OpenChildForm(new FormDispPicture()); };
            _btnTedukuri.Click += BtnTedukuri_Click;
            _btnHinban.Click += (s, e) => { AppState.FlagHinban = true; OpenChildForm(new FormHinban()); };
            _btnPrint.Click += (s, e) => OpenChildForm(new FormPrint());
            _btnNr.Click += (s, e) => OpenChildForm(new FormNr());
            _btnPrintTedukuri.Click += (s, e) => OpenChildForm(new FormPrintTedukuri());
            _btnRepair.Click += (s, e) => OpenChildForm(new FormRepair());
            _btnKaisaku.Click += (s, e) => OpenChildForm(new FormKaisaku());
            _btnExit.Click += (s, e) => this.Close();

            // ボタンをテーブルに追加
            _tableLayout.Controls.Add(_btnData, 0, 0);
            _tableLayout.Controls.Add(_btnMihon, 1, 0);
            _tableLayout.Controls.Add(_btnDejicame, 2, 0);
            _tableLayout.Controls.Add(_btnGousei, 3, 0);
            _tableLayout.Controls.Add(_btnJigane, 0, 1);
            _tableLayout.Controls.Add(_btnSearch, 1, 1);
            _tableLayout.Controls.Add(_btnCamera, 2, 1);
            _tableLayout.Controls.Add(_btnGazouDisp, 3, 1);
            _tableLayout.Controls.Add(_btnTedukuri, 0, 2);
            _tableLayout.Controls.Add(_btnHinban, 1, 2);
            _tableLayout.Controls.Add(_btnPrint, 2, 2);
            _tableLayout.Controls.Add(_btnNr, 3, 2);
            _tableLayout.Controls.Add(_btnPrintTedukuri, 0, 3);
            _tableLayout.Controls.Add(_btnRepair, 1, 3);
            _tableLayout.Controls.Add(_btnKaisaku, 2, 3);
            _tableLayout.Controls.Add(_btnExit, 3, 3);

            _pnlCenter.Controls.Add(_tableLayout);

            // ============================================================
            // 下部情報バー
            // ============================================================
            _pnlBottom = new Panel();
            _pnlBottom.Dock = DockStyle.Bottom;
            _pnlBottom.Height = 56;
            _pnlBottom.BackColor = Color.FromArgb(10, 20, 55);

            // 区切り線（上辺）
            var sep = new Panel();
            sep.BackColor = Color.FromArgb(80, 100, 180);
            sep.Dock = DockStyle.Top;
            sep.Height = 2;
            _pnlBottom.Controls.Add(sep);

            _lblVersion = new Label();
            _lblVersion.Text = SOFTWARE_VERSION;
            _lblVersion.Font = new Font("メイリオ", 12f, FontStyle.Regular);
            _lblVersion.ForeColor = Color.FromArgb(180, 190, 220);
            _lblVersion.BackColor = Color.Transparent;
            _lblVersion.AutoSize = true;
            _lblVersion.Location = new Point(24, 16);
            _pnlBottom.Controls.Add(_lblVersion);

            _lblDataDate = new Label();
            _lblDataDate.Text = "Data ----/--/--";
            _lblDataDate.Font = new Font("メイリオ", 12f, FontStyle.Regular);
            _lblDataDate.ForeColor = Color.FromArgb(180, 190, 220);
            _lblDataDate.BackColor = Color.Transparent;
            _lblDataDate.AutoSize = true;
            _lblDataDate.Location = new Point(280, 16);
            _pnlBottom.Controls.Add(_lblDataDate);

            // ============================================================
            // フォームにコントロールを追加
            // ============================================================
            this.Controls.Add(_pnlBottom);   // Bottom を先に
            this.Controls.Add(_pnlCenter);
            this.Controls.Add(_pnlTitle);    // Top を最後（最前面）
        }

        // ================================================================
        // ボタン生成ヘルパー
        // ================================================================
        private Button MakeBtn(string text, Color back)
        {
            var btn = new Button();
            btn.Text = text;
            btn.Size = new Size(BTN_W, BTN_H);
            btn.Margin = new Padding(BTN_MARGIN);
            btn.BackColor = back;
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderColor = Color.FromArgb(255, 255, 255, 80);
            btn.FlatAppearance.BorderSize = 1;
            btn.Font = new Font("メイリオ", 11f, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
            btn.Dock = DockStyle.Fill;
            return btn;
        }

        // ================================================================
        // リサイズ：タイトル・ボタンを動的配置
        // ================================================================
        private void FormMenu_Resize(object sender, EventArgs e)
        {
            LayoutControls();
        }

        private void LayoutControls()
        {
            int w = this.ClientSize.Width;
            int h = this.ClientSize.Height;
            if (w == 0 || h == 0) return;

            int bottomH = _pnlBottom.Height;
            int availH = h - bottomH;

            // ---- タイトルエリアの高さ（上半分）----
            int titleH = availH / 2;
            _pnlTitle.Height = titleH;

            // タイトルラベルをタイトルエリア内に中央配置
            // ラベルサイズを強制計算
            using (var g = this.CreateGraphics())
            {
                var sysSize = g.MeasureString(
                    _lblSystemTitle.Text, _lblSystemTitle.Font);
                var kobSize = g.MeasureString(
                    _lblKoboTitle.Text, _lblKoboTitle.Font);

                int logoSize = 140;
                int textW = (int)Math.Max(sysSize.Width, kobSize.Width);
                int groupW = logoSize + 24 + textW;
                int startX = (w - groupW) / 2;
                int midY = titleH / 2;

                _picLogo.Size = new Size(logoSize, logoSize);
                _picLogo.Location = new Point(startX, midY - logoSize / 2);

                _lblSystemTitle.Location = new Point(
                    startX + logoSize + 24,
                    midY - (int)sysSize.Height - 4);

                _lblKoboTitle.Location = new Point(
                    startX + logoSize + 24,
                    midY + 4);
            }

            // ---- ボタンエリアを下半分中央に配置 ----
            int tblW = _tableLayout.Width;
            int tblH = _tableLayout.Height;

            int centerX = (w - tblW) / 2;
            int centerY = titleH + (availH - titleH - tblH) / 2;

            _tableLayout.Location = new Point(0, 0);
            _pnlCenter.Location = new Point(centerX, centerY);
            _pnlCenter.Size = new Size(tblW, tblH);
        }

        // ================================================================
        // フォームロード
        // ================================================================
        private void FormMenu_Load(object sender, EventArgs e)
        {
            // 多重起動チェック
            string exeName = Path.GetFileNameWithoutExtension(
                System.Windows.Forms.Application.ExecutablePath);
            if (Process.GetProcessesByName(exeName).Length > 1)
            {
                MessageBox.Show("システムはすでに立ち上がっています。",
                    "システム二重起動チェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Windows.Forms.Application.Exit();
                return;
            }

           
            // フラグ初期化
            AppState.FlagHinban = false;
            AppState.FlagMihon = false;
            AppState.FlagDispPicture = false;

            // DB接続
            try
            {
                _db = new DatabaseManager();
                _db.Connect(DB_PATH);
                AppState.Db = _db;
            }
            catch (Exception ex)
            {
                MessageBox.Show("データベースへの接続に失敗しました。\n\n" + ex.Message,
                    "DB接続エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
                return;
            }

            this.WindowState = FormWindowState.Maximized;

            // データ更新日付の取得
            try
            {
                DataTable dt = _db.ExecuteQuery("SELECT * FROM [データ日付テーブル]");
                if (dt.Rows.Count > 0)
                    _lblDataDate.Text = "Data " + dt.Rows[0]["日付"].ToString();
                else
                    _lblDataDate.Text = "Data ----/--/--";
            }
            catch
            {
                _lblDataDate.Text = "Data (取得失敗)";
            }

            // 初回レイアウト
            LayoutControls();
        }

        // ================================================================
        // 終了時処理
        // ================================================================
        private void FormMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_db == null || !_db.IsConnected) return;
            try
            {
                ResetJinkineTable();
                ResetPtK18Table();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("テーブルリセット失敗: " + ex.Message);
            }
            _db.Dispose();
            _db = null;
        }

        // ================================================================
        // テーブルリセット
        // ================================================================
        private void ResetJinkineTable()
        {
            _db.ExecuteNonQuery("DELETE * FROM [地金相場テーブル]");
            var p = new OleDbParameter[10];
            for (int i = 0; i < 10; i++) p[i] = new OleDbParameter("p" + i, "0");
            _db.ExecuteNonQuery("INSERT INTO [地金相場テーブル] VALUES(?,?,?,?,?,?,?,?,?,?)", p);
        }

        private void ResetPtK18Table()
        {
            _db.ExecuteNonQuery("DELETE * FROM [pt_k18テーブル]");
            const int n = 84;
            var p = new OleDbParameter[n];
            var h = new string[n];
            for (int i = 0; i < n; i++) { p[i] = new OleDbParameter("p" + i, "0"); h[i] = "?"; }
            _db.ExecuteNonQuery("INSERT INTO [pt_k18テーブル] VALUES(" + string.Join(",", h) + ")", p);
        }

        // ================================================================
        // オーダーガイドボタン（ラベル初期化あり）
        // ================================================================
        private void BtnTedukuri_Click(object sender, EventArgs e)
        {
            var form = new FormTedukuri();
            form.ResetSampleLabels();
            OpenChildForm(form);
        }

        // ================================================================
        // フォーム遷移ヘルパー
        // ================================================================
        private void OpenChildForm(Form childForm)
        {
            this.Hide();
            childForm.FormClosed += (s, args) =>
            {
                AppState.FlagHinban = false;
                AppState.FlagMihon = false;
                AppState.FlagDispPicture = false;
                this.Show();
            };
            childForm.Show();
        }
    }

    // ================================================================
    // AppState（グローバル状態管理）
    // ================================================================
    public static class AppState
    {
        public static bool FlagHinban { get; set; }
        public static bool FlagMihon { get; set; }
        public static bool FlagDispPicture { get; set; }
        public static DatabaseManager Db { get; set; }
        public static bool FlagMihonRev1 { get; set; }
        public static double MihonRev1Tsize { get; set; }
        public static double OrderSamplePt900 { get; set; }
        public static double OrderSampleK18 { get; set; }
        public static double OrderSampleWgPg { get; set; }
        public static double OrderSampleK10 { get; set; }
        public static int OrderSampleCode { get; set; }
        public static double OrderSamplePrice { get; set; }
        public static string OrderSampleHinban { get; set; }
        public static string SearchItem { get; set; }
        public static string SearchStoneShape { get; set; }
        public static string SearchJigane { get; set; }
        public static string SearchGrade { get; set; }
        public static int SearchRingsizeJuu { get; set; }
        public static int SearchRingsizeIchi { get; set; }
    }
}