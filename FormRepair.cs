using System;
using System.Drawing;
using System.Windows.Forms;

namespace YumejitateApp
{
    /// <summary>
    /// 修理メニューフォーム (VB6: form_repair.frm の移植)
    /// 各修理種別の子フォームへ遷移するナビゲーション画面。
    /// </summary>
    public class FormRepair : Form
    {
        // ----------------------------------------------------------------
        // コントロール
        // ----------------------------------------------------------------
        private Button _btnSizenaoshi;  // サイズ直し
        private Button _btnIshidome;    // 石留め/石外し
        private Button _btnShintsume;   // 芯爪立て替えと石留め
        private Button _btnHenkei;      // 変形修理
        private Button _btnBrandnew;    // 新品仕上げ
        private Button _btnBack;        // メニューへ戻る

        private Label _lblTitle;        // "指輪修理"
        private Label _lblSubtitle;     // "修理の種類を選んでください"

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------
        public FormRepair()
        {
            InitializeComponent();
        }

        // ----------------------------------------------------------------
        // InitializeComponent
        // ----------------------------------------------------------------
        private void InitializeComponent()
        {
            this.SuspendLayout();

            // ---- フォーム基本設定 ----
            // VB6: BackColor=&H00D8FFFF& (水色), Caption="夢仕立て-メニュー画面"
            this.Text = "夢仕立て - 指輪修理";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF); // &H00D8FFFF
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;

            // ---- タイトルラベル (VB6: Label2 "指輪修理") ----
            _lblTitle = new Label
            {
                Text = "指輪修理",
                Font = new Font("ＭＳ Ｐゴシック", 18f, FontStyle.Italic),
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(20, 15),
            };

            // ---- サブタイトルラベル ----
            _lblSubtitle = new Label
            {
                Text = "修理の種類を選んでください",
                Font = new Font("メイリオ", 12f),
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(20, 55),
                ForeColor = Color.FromArgb(64, 64, 128),
            };

            // ---- ボタン共通スタイル ----
            // VB6: Height=1335, Width=2055 twips → ÷15 ≒ 89×137 px
            // 視認性を高めるため実際は 160×90 に拡大して配置
            var btnFont = new Font("ＭＳ Ｐゴシック", 14.25f, FontStyle.Bold);
            var btnSize = new Size(200, 100);
            var btnBack = Color.FromArgb(0xD8, 0xFF, 0xFF); // フォーム背景と同系

            // VB6 ボタン座標(twips)をピクセルに変換(÷15)し、
            // 最大化表示のため相対配置（Anchor）を使用する。
            // 元座標グループ:
            //   行1 (Top≈192px): btn_sizenaoshi(Left=272), btn_ishidome(Left=440)
            //   行2 (Top≈312px): btn_shintsume(Left=272), btn_henkei(Left=440), btn_brandnew(Left=608)

            // ---- サイズ直し (VB6: btn_sizenaoshi, Top=2880, Left=4080 twips) ----
            _btnSizenaoshi = new Button
            {
                Text = "サイズ直し",
                Font = btnFont,
                Size = btnSize,
                BackColor = Color.FromArgb(0xC0, 0xFF, 0xFF),
                FlatStyle = FlatStyle.Flat,
                Location = new Point(80, 120),
                Cursor = Cursors.Hand,
            };
            _btnSizenaoshi.FlatAppearance.BorderColor = Color.SteelBlue;
            _btnSizenaoshi.FlatAppearance.BorderSize = 2;
            _btnSizenaoshi.Click += BtnSizenaoshi_Click;

            // ---- 石留め/石外し (VB6: btn_ishidome, Top=2880, Left=6600 twips) ----
            _btnIshidome = new Button
            {
                Text = "石留め/石外し",
                Font = btnFont,
                Size = btnSize,
                BackColor = Color.FromArgb(0xC0, 0xFF, 0xFF),
                FlatStyle = FlatStyle.Flat,
                Location = new Point(300, 120),
                Cursor = Cursors.Hand,
            };
            _btnIshidome.FlatAppearance.BorderColor = Color.SteelBlue;
            _btnIshidome.FlatAppearance.BorderSize = 2;
            _btnIshidome.Click += BtnIshidome_Click;

            // ---- 芯爪立て替えと石留め (VB6: btn_shintsume, Top=4680, Left=4080 twips) ----
            _btnShintsume = new Button
            {
                Text = "芯爪立て替えと\n石留め",
                Font = btnFont,
                Size = btnSize,
                BackColor = Color.FromArgb(0xC0, 0xFF, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Location = new Point(80, 240),
                Cursor = Cursors.Hand,
            };
            _btnShintsume.FlatAppearance.BorderColor = Color.SeaGreen;
            _btnShintsume.FlatAppearance.BorderSize = 2;
            _btnShintsume.Click += BtnShintsume_Click;

            // ---- 変形修理 (VB6: btn_henkei, Top=4680, Left=6600 twips) ----
            _btnHenkei = new Button
            {
                Text = "変形修理",
                Font = btnFont,
                Size = btnSize,
                BackColor = Color.FromArgb(0xC0, 0xFF, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Location = new Point(300, 240),
                Cursor = Cursors.Hand,
            };
            _btnHenkei.FlatAppearance.BorderColor = Color.SeaGreen;
            _btnHenkei.FlatAppearance.BorderSize = 2;
            _btnHenkei.Click += BtnHenkei_Click;

            // ---- 新品仕上げ (VB6: btn_brandnew, Top=4680, Left=9120 twips) ----
            _btnBrandnew = new Button
            {
                Text = "新品仕上げ",
                Font = btnFont,
                Size = btnSize,
                BackColor = Color.FromArgb(0xFF, 0xFF, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Location = new Point(520, 240),
                Cursor = Cursors.Hand,
            };
            _btnBrandnew.FlatAppearance.BorderColor = Color.Goldenrod;
            _btnBrandnew.FlatAppearance.BorderSize = 2;
            _btnBrandnew.Click += BtnBrandnew_Click;

            // ---- メニューへ戻る (VB6: btn_back, Top=8880, Left=12240 twips) ----
            _btnBack = new Button
            {
                Text = "メニュー",
                Font = new Font("ＭＳ Ｐゴシック", 15.75f, FontStyle.Bold),
                Size = new Size(160, 60),
                BackColor = Color.FromArgb(0xFF, 0xC0, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Location = new Point(600, 400),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            };
            _btnBack.FlatAppearance.BorderColor = Color.Crimson;
            _btnBack.FlatAppearance.BorderSize = 2;
            _btnBack.Click += BtnBack_Click;

            // ---- コントロールをフォームに追加 ----
            this.Controls.Add(_lblTitle);
            this.Controls.Add(_lblSubtitle);
            this.Controls.Add(_btnSizenaoshi);
            this.Controls.Add(_btnIshidome);
            this.Controls.Add(_btnShintsume);
            this.Controls.Add(_btnHenkei);
            this.Controls.Add(_btnBrandnew);
            this.Controls.Add(_btnBack);

            this.ResumeLayout(false);
        }

        // ----------------------------------------------------------------
        // フォームロード: 最大化（VB6: Form_Load → WindowState=2）
        // ----------------------------------------------------------------
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.WindowState = FormWindowState.Maximized;

            // 最大化後にボタン配置を中央寄せに再調整
            CenterControls();
        }

        // ----------------------------------------------------------------
        // フォームリサイズ時にボタン配置を再計算
        // ----------------------------------------------------------------
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            CenterControls();
        }

        // ----------------------------------------------------------------
        // ボタンをフォーム中央に配置（最大化対応）
        // VB6では固定座標だったが、C#では画面解像度に合わせて中央配置する
        // ----------------------------------------------------------------
        private void CenterControls()
        {
            if (this.ClientSize.Width == 0 || this.ClientSize.Height == 0) return;

            int centerX = this.ClientSize.Width / 2;
            int centerY = this.ClientSize.Height / 2;

            int btnW = 200;
            int btnH = 100;
            int gapX = 20;
            int gapY = 20;

            // 行1: サイズ直し / 石留め/石外し（2ボタン横並び）
            int row1Y = centerY - 140;
            int row1StartX = centerX - btnW - gapX / 2;
            _btnSizenaoshi.SetBounds(row1StartX, row1Y, btnW, btnH);
            _btnIshidome.SetBounds(row1StartX + btnW + gapX, row1Y, btnW, btnH);

            // 行2: 芯爪立て替え / 変形修理 / 新品仕上げ（3ボタン横並び）
            int row2Y = row1Y + btnH + gapY;
            int row2StartX = centerX - (btnW * 3 + gapX * 2) / 2;
            _btnShintsume.SetBounds(row2StartX, row2Y, btnW, btnH);
            _btnHenkei.SetBounds(row2StartX + btnW + gapX, row2Y, btnW, btnH);
            _btnBrandnew.SetBounds(row2StartX + (btnW + gapX) * 2, row2Y, btnW, btnH);

            // タイトル・サブタイトル
            _lblTitle.SetBounds(centerX - 100, 15, 200, 35);
            _lblSubtitle.SetBounds(centerX - 150, 55, 300, 25);

            // メニューボタン（右下）
            _btnBack.SetBounds(
                this.ClientSize.Width - 180,
                this.ClientSize.Height - 90,
                160, 60);
        }

        // ----------------------------------------------------------------
        // イベントハンドラ
        // ----------------------------------------------------------------

        /// <summary>
        /// サイズ直し (VB6: btn_sizenaoshi_Click → form_sizenaoshi.Show)
        /// </summary>
        private void BtnSizenaoshi_Click(object sender, EventArgs e)
        {
            // VB6: Unload Me → form_sizenaoshi.Show
            OpenChildForm(new FormSizanaoshi());
        }

        /// <summary>
        /// 石留め/石外し (VB6: btn_ishidome_Click → form_ishidome.Show)
        /// </summary>
        private void BtnIshidome_Click(object sender, EventArgs e)
        {
            // VB6: Unload Me → form_ishidome.Show
            OpenChildForm(new FormIshidome());
        }

        /// <summary>
        /// 芯爪立て替えと石留め (VB6: btn_shintsume_Click → form_shintsume.Show)
        /// </summary>
        private void BtnShintsume_Click(object sender, EventArgs e)
        {
            // VB6: Unload Me → form_shintsume.Show
            OpenChildForm(new FormShintsume());
        }

        /// <summary>
        /// 変形修理 (VB6: btn_henkei_Click → form_henkei.Show)
        /// </summary>
        private void BtnHenkei_Click(object sender, EventArgs e)
        {
            // VB6: Unload Me → form_henkei.Show
            OpenChildForm(new FormHenkei());
        }

        /// <summary>
        /// 新品仕上げ (VB6: btn_brandnew_Click → form_brand_new.Show)
        /// </summary>
        private void BtnBrandnew_Click(object sender, EventArgs e)
        {
            // VB6: Unload Me → form_brand_new.Show
            OpenChildForm(new FormBrandNew());
        }

        /// <summary>
        /// メニューへ戻る (VB6: btn_back_Click → Unload Me → form_menu.Show)
        /// </summary>
        private void BtnBack_Click(object sender, EventArgs e)
        {
            // FormMenu を検索して表示
            foreach (Form f in Application.OpenForms)
            {
                if (f is FormMenu)
                {
                    f.Show();
                    break;
                }
            }
            this.Close();
        }

        // ----------------------------------------------------------------
        // 子フォーム遷移ヘルパー
        // VB6: Unload Me → 子フォーム.Show（非モーダル）
        // C# : このフォームを非表示にして子フォームを開き、
        //       子フォームが閉じたときに再表示する。
        // ----------------------------------------------------------------
        private void OpenChildForm(Form childForm)
        {
            this.Hide();
            childForm.FormClosed += (s, e) =>
            {
                // 子フォームが閉じたらこのフォームも閉じる（VB6のUnload Meに相当）
                this.Close();
            };
            childForm.Show();
        }
    }
}
