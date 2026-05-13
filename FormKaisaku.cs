using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace YumejitateApp
{
    /// <summary>
    /// チェーン(きへい他) 改作見積フォーム (VB6: form_kaisaku.frm の移植)
    /// 改作JIGANEテーブルと地金売り相場テーブルを参照して価格を計算する。
    /// LEADLib.LEAD は System.Drawing.PictureBox に置き換え済み。
    /// </summary>
    public class FormKaisaku : Form
    {
        // ----------------------------------------------------------------
        // コントロール - 日付入力コンボ (千年/百年/十年/一年/十月/一月/十日/一日)
        // ----------------------------------------------------------------
        private ComboBox _cmbSennen;   // 千年
        private ComboBox _cmbHyakuNen; // 百年
        private ComboBox _cmbJuNen;    // 十年
        private ComboBox _cmbIchiNen;  // 一年
        private ComboBox _cmbJuTsuki;  // 十月
        private ComboBox _cmbIchiTsuki;// 一月
        private ComboBox _cmbJuNichi;  // 十日
        private ComboBox _cmbIchiNichi;// 一日

        // ----------------------------------------------------------------
        // コントロール - 入力コンボ
        // ----------------------------------------------------------------
        private ComboBox _cmbShohinShu;  // 商品の種類
        private ComboBox _cmbJigane;     // 地金
        private ComboBox _cmbMaxHaba;    // 最大幅
        private ComboBox _cmbNagasaJu;   // 長さ十
        private ComboBox _cmbNagasaIchi; // 長さ一

        // ----------------------------------------------------------------
        // コントロール - 出力ラベル (計算・DB参照結果)
        // ----------------------------------------------------------------
        private Label _lblPt1000;          // PT1000 相場
        private Label _lblK24;             // K24 相場
        private Label _lblDesignNo;        // デザインNO
        private Label _lblMaxKei;          // 最大径
        private Label _lblJiganeJuryo;     // 地金重量 (非表示・計算用)
        private Label _lblKotin;           // 工賃 (非表示・計算用)
        private Label _lblHyojiJiganeJuryo;// 表示地金重量
        private Label _lblGokeiSeikyu;     // 合計ご請求額

        // ----------------------------------------------------------------
        // コントロール - 画像プレビュー (VB6: LEAD1)
        // ----------------------------------------------------------------
        private PictureBox _pictureBox1;   // LEAD1 → PictureBox

        // ----------------------------------------------------------------
        // コントロール - エラー表示
        // ----------------------------------------------------------------
        private TextBox _txtErr;           // txt_err

        // ----------------------------------------------------------------
        // コントロール - ボタン
        // ----------------------------------------------------------------
        private Button _btnKeisan;  // 計算
        private Button _btnExit;    // メニュー

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------
        public FormKaisaku()
        {
            InitializeComponent();
        }

        // ----------------------------------------------------------------
        // InitializeComponent: コントロール構築
        // ----------------------------------------------------------------
        private void InitializeComponent()
        {
            this.SuspendLayout();

            // ---- フォーム基本設定 ----
            // VB6: Caption="夢仕立て-NRコレクション", BackColor=&H00D8FFFF&
            this.Text = "夢仕立て - チェーン(きへい他)";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.AutoScroll = true;

            var fntNormal = new Font("ＭＳ Ｐゴシック", 14.25f, FontStyle.Bold);
            var fntTitle = new Font("ＭＳ Ｐゴシック", 18f, FontStyle.Italic);
            var fntLarge = new Font("ＭＳ Ｐゴシック", 15.75f, FontStyle.Bold);

            // ---- タイトル (VB6: Label2 "チェーン(きへい他)", Top=240twip≈16px) ----
            var lblTitle = new Label
            {
                Text = "チェーン(きへい他)",
                Font = fntTitle,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(384, 16),
            };

            // ================================================================
            // 日付入力エリア (VB6: 千年〜一日 コンボ, Top=1920twip≈128px)
            // ================================================================
            var lblDateTitle = new Label
            {
                Text = "本日の地金　売り　相場",
                Font = fntNormal,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(16, 104),
            };

            // コンボボックス共通サイズ (VB6: Width=700twip≈47px → 見やすく50px)
            var cmbSize = new Size(50, 27);

            _cmbSennen = MakeDigitCombo(new Point(16, 128), cmbSize);
            _cmbHyakuNen = MakeDigitCombo(new Point(76, 128), cmbSize);
            _cmbJuNen = MakeDigitCombo(new Point(136, 128), cmbSize);
            _cmbIchiNen = MakeDigitCombo(new Point(196, 128), cmbSize);

            var lblNen = new Label { Text = "年", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(224, 132) };

            _cmbJuTsuki = MakeDigitCombo(new Point(296, 128), cmbSize);
            _cmbIchiTsuki = MakeDigitCombo(new Point(356, 128), cmbSize);
            var lblTsuki = new Label { Text = "月", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(354, 132) };

            _cmbJuNichi = MakeDigitCombo(new Point(456, 128), cmbSize);
            _cmbIchiNichi = MakeDigitCombo(new Point(516, 128), cmbSize);
            var lblNichi = new Label { Text = "日", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(484, 132) };

            // ================================================================
            // 地金売り相場表示エリア (VB6: lbl_pt1000, lbl_k24)
            // ================================================================
            var pnlSoba = new GroupBox
            {
                Text = "現在のデータ入力値",
                Font = fntNormal,
                Location = new Point(672, 72),
                Size = new Size(460, 90),
                BackColor = Color.Transparent,
            };

            var lblPt1000Ttl = new Label { Text = "PT１０００", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(10, 20) };
            var lblYen1 = new Label { Text = "¥", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(5, 48) };
            _lblPt1000 = new Label
            {
                Text = "0",
                Font = fntLarge,
                TextAlign = ContentAlignment.MiddleRight,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Location = new Point(30, 44),
                Size = new Size(145, 27),
            };
            var lblPtUnit = new Label { Text = "／ｇ", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(178, 48) };

            var lblK24Ttl = new Label { Text = "Ｋ２４", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(230, 20) };
            var lblYen2 = new Label { Text = "¥", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(225, 48) };
            _lblK24 = new Label
            {
                Text = "0",
                Font = fntLarge,
                TextAlign = ContentAlignment.MiddleRight,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Location = new Point(250, 44),
                Size = new Size(145, 27),
            };
            var lblK24Unit = new Label { Text = "／ｇ", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(398, 48) };

            pnlSoba.Controls.AddRange(new Control[] {
                lblPt1000Ttl, lblYen1, _lblPt1000, lblPtUnit,
                lblK24Ttl, lblYen2, _lblK24, lblK24Unit,
            });

            // ================================================================
            // 条件入力エリア (VB6: 商品の種類/地金/最大幅/デザインNO/最大径, Top=3240twip≈216px)
            // ================================================================
            var lblShohinShu = new Label { Text = "商品の種類", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(16, 196) };
            _cmbShohinShu = new ComboBox
            {
                Font = fntNormal,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(120, 192),
                Size = new Size(160, 27),
            };
            _cmbShohinShu.SelectedIndexChanged += CmbShohinShu_SelectedIndexChanged;

            var lblJigane = new Label { Text = "地金", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(290, 196) };
            _cmbJigane = new ComboBox
            {
                Font = fntNormal,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(340, 192),
                Size = new Size(100, 27),
            };
            _cmbJigane.SelectedIndexChanged += CmbJigane_SelectedIndexChanged;

            // VB6: 線径(mm) ラベル (Label7, Top=2880twip≈192px)
            var lblSenKei = new Label { Text = "線径(mm)", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(452, 196) };
            _cmbMaxHaba = new ComboBox
            {
                Font = fntNormal,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(560, 192),
                Size = new Size(100, 27),
            };
            _cmbMaxHaba.SelectedIndexChanged += CmbMaxHaba_SelectedIndexChanged;

            var lblDesignNoTtl = new Label { Text = "デザインNo．", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(672, 196) };
            _lblDesignNo = new Label
            {
                Text = "",
                Font = fntLarge,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Location = new Point(792, 192),
                Size = new Size(130, 27),
            };

            var lblMaxKeiTtl = new Label { Text = "最大径約(mm)", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(932, 196) };
            _lblMaxKei = new Label
            {
                Text = "",
                Font = fntLarge,
                TextAlign = ContentAlignment.MiddleRight,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Location = new Point(1115, 192),
                Size = new Size(115, 27),
            };

            // ================================================================
            // 画像プレビュー (VB6: LEAD1, Left=1560twip≈104px, Top=3960twip≈264px)
            // ================================================================
            _pictureBox1 = new PictureBox
            {
                Location = new Point(104, 264),
                Size = new Size(305, 137),
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.Black,
                BorderStyle = BorderStyle.FixedSingle,
            };

            // ================================================================
            // 長さ入力 (VB6: 長さ十/長さ一, Top=5040twip≈336px)
            // ================================================================
            var lblNagasa = new Label { Text = "長さ(cm)", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(464, 312) };
            _cmbNagasaJu = MakeDigitCombo(new Point(464, 336), new Size(50, 27));
            _cmbNagasaIchi = MakeDigitCombo(new Point(516, 336), new Size(50, 27));
            var lblCm = new Label { Text = "cm", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(568, 340) };

            // 非表示ラベル (計算用中間値): 地金重量, 工賃
            _lblJiganeJuryo = new Label { Text = "0", Visible = false };
            _lblKotin = new Label { Text = "0", Visible = false };

            // ================================================================
            // 結果表示エリア (VB6: 表示地金重量, 合計ご請求額, Top≈455-490px)
            // ================================================================
            var lblJiganeJuryoTtl = new Label
            {
                Text = "地金重量約(g)",
                Font = fntNormal,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(560, 456),
            };
            _lblHyojiJiganeJuryo = new Label
            {
                Text = "0.00",
                Font = fntLarge,
                TextAlign = ContentAlignment.MiddleRight,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Location = new Point(710, 452),
                Size = new Size(97, 27),
            };

            var lblGokeiTtl = new Label
            {
                Text = "合計ご請求額(税込)",
                Font = fntNormal,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(352, 456),
            };
            var lblYen3 = new Label { Text = "¥", Font = fntNormal, BackColor = Color.Transparent, AutoSize = true, Location = new Point(331, 490) };
            _lblGokeiSeikyu = new Label
            {
                Text = "0",
                Font = fntLarge,
                TextAlign = ContentAlignment.MiddleRight,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Location = new Point(352, 486),
                Size = new Size(178, 27),
            };

            // ================================================================
            // エラー表示テキストボックス (VB6: txt_err, Top=6240twip≈416px)
            // ================================================================
            _txtErr = new TextBox
            {
                Font = fntNormal,
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF),
                ForeColor = Color.Blue,
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                Location = new Point(280, 416),
                Size = new Size(465, 27),
                Text = "",
            };

            // ================================================================
            // 注記ラベル (VB6: Label26[0]/[1]/[2])
            // ================================================================
            // Index=0 "合計ご請求額(税込)" はすでに lblGokeiTtl で対応済み
            var lblNote1 = new Label
            {
                Text = "※．ご注文の製品は量産品と異なり、１０％程の増減が生じます。",
                Font = fntNormal,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(224, 560),
            };
            var lblNote2 = new Label
            {
                Text = "※．ご注文の製品は全て、スライダー製品となります。",
                Font = fntNormal,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(224, 528),
            };

            // ================================================================
            // ボタン (VB6: btn_keisan Top=8760twip≈584px, btn_exit Top=8760twip)
            // ================================================================
            _btnKeisan = new Button
            {
                Text = "計算",
                Font = new Font("ＭＳ Ｐゴシック", 15.75f, FontStyle.Bold),
                Location = new Point(416, 608),
                Size = new Size(137, 57),
                BackColor = Color.FromArgb(192, 255, 192),
                Cursor = Cursors.Hand,
            };
            _btnKeisan.Click += BtnKeisan_Click;

            _btnExit = new Button
            {
                Text = "メニュー",
                Font = new Font("ＭＳ Ｐゴシック", 15.75f, FontStyle.Bold),
                Location = new Point(688, 608),
                Size = new Size(137, 57),
                BackColor = Color.FromArgb(255, 192, 192),
                Cursor = Cursors.Hand,
            };
            _btnExit.Click += BtnExit_Click;

            // ================================================================
            // フォームにコントロールを追加
            // ================================================================
            this.Controls.AddRange(new Control[]
            {
                lblTitle,
                lblDateTitle,
                _cmbSennen, _cmbHyakuNen, _cmbJuNen, _cmbIchiNen,
                lblNen,
                _cmbJuTsuki, _cmbIchiTsuki,
                lblTsuki,
                _cmbJuNichi, _cmbIchiNichi,
                lblNichi,
                pnlSoba,
                lblShohinShu, _cmbShohinShu,
                lblJigane, _cmbJigane,
                lblSenKei, _cmbMaxHaba,
                lblDesignNoTtl, _lblDesignNo,
                lblMaxKeiTtl, _lblMaxKei,
                _pictureBox1,
                lblNagasa, _cmbNagasaJu, _cmbNagasaIchi, lblCm,
                _lblJiganeJuryo, _lblKotin,
                lblJiganeJuryoTtl, _lblHyojiJiganeJuryo,
                lblGokeiTtl, lblYen3, _lblGokeiSeikyu,
                _txtErr,
                lblNote1, lblNote2,
                _btnKeisan, _btnExit,
            });

            this.ResumeLayout(false);
        }

        // ----------------------------------------------------------------
        // 0-9 入力用コンボボックス生成ヘルパー
        // ----------------------------------------------------------------
        private static ComboBox MakeDigitCombo(Point loc, Size sz)
        {
            var cmb = new ComboBox
            {
                Font = new Font("ＭＳ Ｐゴシック", 14.25f, FontStyle.Bold),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = loc,
                Size = sz,
            };
            for (int i = 0; i <= 9; i++) cmb.Items.Add(i.ToString());
            return cmb;
        }

        // ----------------------------------------------------------------
        // フォームロード (VB6: Form_Load → Init_Control → WindowState=2)
        // ----------------------------------------------------------------
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.WindowState = FormWindowState.Maximized;
            InitControl();
        }

        // ----------------------------------------------------------------
        // 画面コントロール初期化処理 (VB6: Init_Control)
        // ----------------------------------------------------------------
        private void InitControl()
        {
            // ---- 今日の日付をコンボに自動設定 ----
            // VB6: str_date = Date → 例 "2026/04/17"
            //      千年.List(Mid(str_date,1,1)) → "2"
            string today = DateTime.Today.ToString("yyyy/MM/dd");
            SetDateCombo(today);

            // ---- 商品の種類リスト ----
            _cmbShohinShu.Items.Clear();
            _cmbShohinShu.Items.AddRange(new object[] { "ベネチアン", "ハーフラウンド", "喜平２面", "丸小豆" });
            _cmbShohinShu.SelectedIndex = 0;

            // ---- 地金リスト ----
            _cmbJigane.Items.Clear();
            _cmbJigane.Items.AddRange(new object[] { "Pt850", "K18", "WG", "K10" });
            _cmbJigane.SelectedIndex = 0;

            // ---- 長さ初期値 (VB6: 長さ十=4, 長さ一=5 → "45" cm) ----
            _cmbNagasaJu.SelectedIndex = 4;
            _cmbNagasaIchi.SelectedIndex = 5;

            // ---- 合計初期値 ----
            _lblGokeiSeikyu.Text = "0";
            _lblHyojiJiganeJuryo.Text = "0.00";

            // ---- DBから改作JIGANEテーブルを初期ロード ----
            LoadJiganeTable();

            // ---- 商品画像をロード ----
            LoadProductImage(_cmbShohinShu.Text);

            // ---- 地金売り相場テーブルをロード ----
            LoadSobaTable();

            _txtErr.Text = "";
        }

        // ----------------------------------------------------------------
        // 今日の日付をコンボボックスに設定
        // VB6: 千年.Text = 千年.List(Mid(str_date, 1, 1)) 等
        //      str_date 例: "2026/04/17" (VB6 の Date 関数の出力形式)
        // ----------------------------------------------------------------
        private void SetDateCombo(string yyyyMMdd)
        {
            // yyyyMMdd = "2026/04/17" (length=10)
            // VB6: Mid(str_date,1,1)="2", Mid(str_date,2,1)="0", Mid(str_date,3,1)="2", Mid(str_date,4,1)="6"
            //      Mid(str_date,6,1)="0", Mid(str_date,7,1)="4"
            //      Mid(str_date,9,1)="1", Mid(str_date,10,1)="7"
            if (yyyyMMdd.Length < 10) return;
            SelectDigit(_cmbSennen, yyyyMMdd[0]);
            SelectDigit(_cmbHyakuNen, yyyyMMdd[1]);
            SelectDigit(_cmbJuNen, yyyyMMdd[2]);
            SelectDigit(_cmbIchiNen, yyyyMMdd[3]);
            SelectDigit(_cmbJuTsuki, yyyyMMdd[5]);
            SelectDigit(_cmbIchiTsuki, yyyyMMdd[6]);
            SelectDigit(_cmbJuNichi, yyyyMMdd[8]);
            SelectDigit(_cmbIchiNichi, yyyyMMdd[9]);
        }

        private static void SelectDigit(ComboBox cmb, char digit)
        {
            int idx = digit - '0';
            if (idx >= 0 && idx <= 9) cmb.SelectedIndex = idx;
        }

        // ----------------------------------------------------------------
        // 改作JIGANEテーブルをDBからロードして最大幅リストを構築
        // VB6: SELECT * FROM 改作JIGANEテーブル WHERE 商品の種類=? AND 地金=? ORDER BY index
        // ----------------------------------------------------------------
        private void LoadJiganeTable()
        {
            string shohin = _cmbShohinShu.Text;
            string jigane = _cmbJigane.Text;
            if (string.IsNullOrEmpty(shohin) || string.IsNullOrEmpty(jigane)) return;

            _cmbMaxHaba.Items.Clear();

            string sql = "SELECT * FROM [改作JIGANEテーブル]"
                       + $" WHERE 商品の種類 = '{shohin}'"
                       + $" AND 地金 = '{jigane}'"
                       + " ORDER BY index";

            DataTable dt = AppState.Db.ExecuteQuery(sql);
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show(
                    "システムエラーです。K10は丸小豆とベネチアンのみです。該当品の計算は無効です。検索をやり直してください。",
                    "改作JIGANEテーブル取得", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 最大幅リストを構築
            foreach (DataRow row in dt.Rows)
                _cmbMaxHaba.Items.Add(row["最大幅"].ToString());

            _cmbMaxHaba.SelectedIndex = 0;

            // 最初のレコードで各表示ラベルを更新
            DataRow first = dt.Rows[0];
            _lblDesignNo.Text = first["デザインNO"].ToString();
            _lblJiganeJuryo.Text = first["地金重量"].ToString();
            _lblKotin.Text = first["工賃"].ToString();
            _lblMaxKei.Text = first["最大径"].ToString();
        }

        // ----------------------------------------------------------------
        // 選択中の最大幅に対応するレコードをDBから取得して表示更新
        // VB6: 最大幅_Click → SELECT WHERE 商品の種類 AND 最大幅 AND 地金
        // ----------------------------------------------------------------
        private void UpdateByMaxHaba()
        {
            if (_cmbMaxHaba.Text == "" || _cmbJigane.Text == "") return;

            string sql = "SELECT * FROM [改作JIGANEテーブル]"
                       + $" WHERE 商品の種類 = '{_cmbShohinShu.Text}'"
                       + $" AND 最大幅 = {ToDoubleStr(_cmbMaxHaba.Text)}"
                       + $" AND 地金 = '{_cmbJigane.Text}'"
                       + " ORDER BY index";

            DataTable dt = AppState.Db.ExecuteQuery(sql);
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show(
                    "システムエラーです。K10は丸小豆とベネチアンのみです。該当品の計算は無効です。検索をやり直してください。",
                    "改作JIGANEテーブル取得", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataRow row = dt.Rows[0];
            _lblDesignNo.Text = row["デザインNO"].ToString();
            _lblJiganeJuryo.Text = row["地金重量"].ToString();
            _lblKotin.Text = row["工賃"].ToString();
            _lblMaxKei.Text = row["最大径"].ToString();
        }

        // ----------------------------------------------------------------
        // 地金変更時の更新 (VB6: 地金_Click)
        // SELECT WHERE 商品の種類 AND 最大幅 AND デザインNO AND 地金
        // ----------------------------------------------------------------
        private void UpdateByJigane()
        {
            if (_cmbMaxHaba.Text == "") return;

            string sql = "SELECT * FROM [改作JIGANEテーブル]"
                       + $" WHERE 商品の種類 = '{_cmbShohinShu.Text}'"
                       + $" AND 最大幅 = {ToDoubleStr(_cmbMaxHaba.Text)}"
                       + $" AND デザインNO = '{_lblDesignNo.Text}'"
                       + $" AND 地金 = '{_cmbJigane.Text}'"
                       + " ORDER BY index";

            DataTable dt = AppState.Db.ExecuteQuery(sql);
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show(
                    "システムエラーです。K10は丸小豆とベネチアンのみです。該当品の計算は無効です。検索をやり直してください。",
                    "改作JIGANEテーブル取得", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataRow row = dt.Rows[0];
            _lblJiganeJuryo.Text = row["地金重量"].ToString();
            _lblKotin.Text = row["工賃"].ToString();
            _lblMaxKei.Text = row["最大径"].ToString();
        }

        // ----------------------------------------------------------------
        // 地金売り相場テーブルをロード (VB6: Init_Control 内)
        // SELECT * FROM 地金売り相場テーブル
        // pt万,pt千,pt百,pt十,pt一 → lbl_pt1000
        // k18万,k18千,k18百,k18十,k18一 → lbl_k24
        // ----------------------------------------------------------------
        private void LoadSobaTable()
        {
            DataTable dt = AppState.Db.ExecuteQuery("SELECT * FROM [地金売り相場テーブル]");
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("システムエラーです。", "地金売り相場テーブル該当データ無し",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataRow row = dt.Rows[0];
            // VB6: wk_str = pt万 & pt千 & pt百 & pt十 & pt一
            string ptStr = row["pt万"].ToString() + row["pt千"].ToString()
                          + row["pt百"].ToString() + row["pt十"].ToString()
                          + row["pt一"].ToString();
            string k18Str = row["k18万"].ToString() + row["k18千"].ToString()
                          + row["k18百"].ToString() + row["k18十"].ToString()
                          + row["k18一"].ToString();

            if (long.TryParse(ptStr, out long ptVal))
                _lblPt1000.Text = ptVal.ToString("#,##0");
            if (long.TryParse(k18Str, out long k18Val))
                _lblK24.Text = k18Val.ToString("#,##0");
        }

        // ----------------------------------------------------------------
        // 商品画像ロード (VB6: LEAD1.Load App.Path + "\kaisaku\" + 商品の種類 + ".JPG")
        // ----------------------------------------------------------------
        private void LoadProductImage(string shohinShu)
        {
            string path = Path.Combine(Application.StartupPath, "kaisaku", shohinShu + ".JPG");
            if (!File.Exists(path))
                path = Path.Combine(Application.StartupPath, "kaisaku", shohinShu + ".jpg");

            if (_pictureBox1.Image != null)
            {
                _pictureBox1.Image.Dispose();
                _pictureBox1.Image = null;
            }

            if (File.Exists(path))
            {
                try
                {
                    using (var tmp = Image.FromFile(path))
                        _pictureBox1.Image = new Bitmap(tmp);
                }
                catch { /* 画像読み込み失敗は無視 */ }
            }
        }

        // ----------------------------------------------------------------
        // コンボボックス イベントハンドラ
        // ----------------------------------------------------------------

        /// <summary>
        /// 商品の種類変更 (VB6: 商品の種類_Click)
        /// </summary>
        private void CmbShohinShu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_cmbJigane.Text == "") return;
            _cmbMaxHaba.Items.Clear();
            LoadJiganeTable();
            LoadProductImage(_cmbShohinShu.Text);
        }

        /// <summary>
        /// 地金変更 (VB6: 地金_Click)
        /// </summary>
        private void CmbJigane_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateByJigane();
        }

        /// <summary>
        /// 最大幅変更 (VB6: 最大幅_Click)
        /// </summary>
        private void CmbMaxHaba_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateByMaxHaba();
        }

        // ----------------------------------------------------------------
        // 計算ボタン (VB6: btn_keisan_Click)
        // ----------------------------------------------------------------
        private void BtnKeisan_Click(object sender, EventArgs e)
        {
            // ---- 地金相場取得 ----
            // VB6: lbl_pt1000.Caption / lbl_k24.Caption からカンマ除去してDouble変換
            if (!TryParsePrice(_lblPt1000.Text, out double dblPt1000) ||
                !TryParsePrice(_lblK24.Text, out double dblK24))
            {
                MessageBox.Show("地金相場が取得できません。", "計算エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ---- 長さ取得 (VB6: 長さ十*10 + 長さ一) ----
            double dblLong = SelectedDigit(_cmbNagasaJu) * 10.0
                           + SelectedDigit(_cmbNagasaIchi);

            // ---- 全体掛け率取得 (VB6: 掛け率テーブル.掛け率1 / 10) ----
            DataTable dtKakeru = AppState.Db.ExecuteQuery("SELECT * FROM [掛け率テーブル]");
            if (dtKakeru == null || dtKakeru.Rows.Count == 0)
            {
                MessageBox.Show("システムエラー", "計算処理：掛け率テーブル取得",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            double zentaiKakeritu = Convert.ToDouble(dtKakeru.Rows[0]["掛け率1"]) / 10.0;

            // ---- 加工コード95 の掛け率取得 ----
            DataTable dtKakou = AppState.Db.ExecuteQuery(
                "SELECT * FROM [加工コード] WHERE コード = 95");
            if (dtKakou == null || dtKakou.Rows.Count == 0)
            {
                MessageBox.Show("システムエラー", "計算処理：加工コード取得",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            double kakouCD = Convert.ToDouble(dtKakou.Rows[0]["掛け率"]);

            // VB6: Zentai_Kakeritu = Zentai_Kakeritu * Kakou_CD
            zentaiKakeritu *= kakouCD;

            // ---- 地金重量・工賃取得 ----
            if (!double.TryParse(_lblJiganeJuryo.Text, out double jiganeJuryo))
            {
                MessageBox.Show("地金重量が取得できません。", "計算エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!double.TryParse(_lblKotin.Text.Replace(",", ""), out double kotin))
            {
                MessageBox.Show("工賃が取得できません。", "計算エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ---- 価格計算 ----
            // VB6 の計算式:
            //   単位重量(g/cm) = 地金重量 / 1000
            //   全体重量(g)    = 単位重量 * 長さ
            //   Pt850: (pt1000 * 0.85 * 全体重量 + 全体重量 * 工賃) * 1.15 * 3.2 * 1.1
            //   K18:   (k24   * 0.75 * 全体重量 + 全体重量 * 工賃) * 1.15 * 3.2 * 1.1
            //   WG:    (k24   * 0.75 * 全体重量 + 全体重量 * 工賃) * 1.15 * 3.3 * 1.1
            //   K10:   (k24   * (10/24) * 全体重量 + 全体重量 * 工賃) * 1.15 * 3.3 * 1.1
            double unitWeight = jiganeJuryo / 1000.0; // g/cm
            double totalWeight = unitWeight * dblLong;  // g

            double wkDbl;
            string jigane = _cmbJigane.Text;
            if (string.Compare(jigane, "Pt850", StringComparison.OrdinalIgnoreCase) == 0)
            {
                wkDbl = (dblPt1000 * (850.0 / 1000.0) * totalWeight
                       + totalWeight * kotin) * 1.15 * 3.2 * 1.1;
            }
            else if (string.Compare(jigane, "K18", StringComparison.OrdinalIgnoreCase) == 0)
            {
                wkDbl = (dblK24 * (18.0 / 24.0) * totalWeight
                       + totalWeight * kotin) * 1.15 * 3.2 * 1.1;
            }
            else if (string.Compare(jigane, "WG", StringComparison.OrdinalIgnoreCase) == 0)
            {
                wkDbl = (dblK24 * (18.0 / 24.0) * totalWeight
                       + totalWeight * kotin) * 1.15 * 3.3 * 1.1;
            }
            else if (string.Compare(jigane, "K10", StringComparison.OrdinalIgnoreCase) == 0)
            {
                wkDbl = (dblK24 * (10.0 / 24.0) * totalWeight
                       + totalWeight * kotin) * 1.15 * 3.3 * 1.1;
            }
            else
            {
                wkDbl = 0;
            }

            // ---- 表示地金重量 (VB6: Jigane/1000 * dbl_Long) ----
            _lblHyojiJiganeJuryo.Text = (jiganeJuryo / 1000.0 * dblLong).ToString("#0.00");

            // ---- 合計ご請求額 ----
            _lblGokeiSeikyu.Text = ((long)Math.Round(wkDbl, 0)).ToString("#,##0");

            // ---- エラーメッセージ (VB6: wk_dbl=0 → "この製品は特注品となります。") ----
            _txtErr.Text = (wkDbl == 0)
                ? "この製品は特注品となります。"
                : "";
        }

        // ----------------------------------------------------------------
        // メニューボタン (VB6: btn_exit_Click → Unload form_kaisaku → form_menu.Visible=True)
        // ----------------------------------------------------------------
        private void BtnExit_Click(object sender, EventArgs e)
        {
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
        // フォームクローズ時のリソース解放
        // ----------------------------------------------------------------
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            if (_pictureBox1?.Image != null)
            {
                _pictureBox1.Image.Dispose();
                _pictureBox1.Image = null;
            }
        }

        // ----------------------------------------------------------------
        // ユーティリティ
        // ----------------------------------------------------------------

        /// <summary>
        /// コンボボックスの選択値（0-9 数字）を int で返す
        /// </summary>
        private static int SelectedDigit(ComboBox cmb)
        {
            if (cmb.SelectedIndex >= 0 && int.TryParse(cmb.Text, out int v)) return v;
            return 0;
        }

        /// <summary>
        /// カンマ付き数値文字列を double に変換
        /// </summary>
        private static bool TryParsePrice(string text, out double val)
        {
            return double.TryParse(text.Replace(",", ""), out val);
        }

        /// <summary>
        /// 数値文字列を SQL 用の double 文字列に変換（ロケール依存なし）
        /// </summary>
        private static string ToDoubleStr(string s)
        {
            if (double.TryParse(s, out double d))
                return d.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return "0";
        }
    }
}
