using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace YumejitateApp
{
    /// <summary>
    /// NRコレクション検索フォーム (VB6: form_nr.frm の移植)
    /// NRコレクション品の品番・地金コードを入力し、
    /// 地金代金・工賃・合計見積予算を計算して表示する。
    /// </summary>
    public class FormNr : Form
    {
        // ----------------------------------------------------------------
        // 日付入力コンボボックス (年/月/日 各桁)
        // VB6: 千年 四年 十年 一年 / 十月 一月 / 十日 一日
        // ----------------------------------------------------------------
        private ComboBox _cmbSennen;    // 年 千の位
        private ComboBox _cmbHyakunen; // 年 百の位 (VB6: 四年)
        private ComboBox _cmbJuNen;    // 年 十の位
        private ComboBox _cmbIchiNen;  // 年 一の位
        private ComboBox _cmbJuTsuki;  // 月 十の位
        private ComboBox _cmbIchiTsuki;// 月 一の位
        private ComboBox _cmbJuNichi;  // 日 十の位
        private ComboBox _cmbIchiNichi;// 日 一の位

        // ----------------------------------------------------------------
        // PT1000 相場入力コンボ (5桁: 万千四十一)
        // VB6: PT万 PT千 PT四 PT十 PT一
        // 価格 = 万×10000 + 千×1000 + 四×100 + 十×10 + 一
        // ----------------------------------------------------------------
        private ComboBox _cmbPtMan;  // PT 万の位
        private ComboBox _cmbPtSen;  // PT 千の位
        private ComboBox _cmbPtShi;  // PT 四の位(百の位相当, VB6: PT四)
        private ComboBox _cmbPtJu;   // PT 十の位
        private ComboBox _cmbPtIchi; // PT 一の位

        // ----------------------------------------------------------------
        // K24 相場入力コンボ (5桁: 万千四十一)
        // VB6: K24万 K24千 K24四 K24十 K24一
        // ----------------------------------------------------------------
        private ComboBox _cmbK24Man;  // K24 万の位
        private ComboBox _cmbK24Sen;  // K24 千の位
        private ComboBox _cmbK24Shi;  // K24 四の位(百の位相当)
        private ComboBox _cmbK24Ju;   // K24 十の位
        private ComboBox _cmbK24Ichi; // K24 一の位

        // ----------------------------------------------------------------
        // 品番入力コンボ (5桁)
        // VB6: 品番一(0-9/T/H) 品番二〜五(0-9)
        // ----------------------------------------------------------------
        private ComboBox _cmbHinban1; // 品番 第1桁 (0-9, T, H)
        private ComboBox _cmbHinban2; // 品番 第2桁
        private ComboBox _cmbHinban3; // 品番 第3桁
        private ComboBox _cmbHinban4; // 品番 第4桁
        private ComboBox _cmbHinban5; // 品番 第5桁

        // ----------------------------------------------------------------
        // その他入力コンボ
        // ----------------------------------------------------------------
        private ComboBox _cmbJiganeCode; // 地金コード (1/2/4/5)
        private ComboBox _cmbSizeJu;     // サイズ 十の位
        private ComboBox _cmbSizeIchi;   // サイズ 一の位

        // ----------------------------------------------------------------
        // 計算結果表示ラベル (枠付き右寄せ)
        // ----------------------------------------------------------------
        private Label _lblDesignNo;          // デザインNO (VB6: デザインNO)
        private Label _lblJiganeJuryo;       // 地金付き重量 (VB6: 地金付き重量)
        private Label _lblJiganeDaikin;      // 地金代金     (VB6: 地金代金)
        private Label _lblJunkinHitsuyo;     // 純金必要量   (VB6: 純金必要量)
        private Label _lblKotin;             // 工賃         (VB6: 工賃)
        private Label _lblMd;               // メレCT       (VB6: MD)
        private Label _lblFd;               // FDカラット    (VB6: FD)
        private Label _lblGokeimitsumoriYosan; // 合計見積予算 (VB6: 合計見積予算)

        // ----------------------------------------------------------------
        // 動的に内容が変わるラベル
        // ----------------------------------------------------------------
        private Label _lblItemType;       // "リング" ↔ "チェーン" (VB6: label_アイテム)
        private Label _lblSizeType;       // "サイズ" ↔ "単価重量" (VB6: Label_サイズ)
        private Label _lblJiganeCodeDisp; // "純白"/"純"/"合白"/"合純" (VB6: label_地金コード)

        // ----------------------------------------------------------------
        // 固定ラベル (見出し・単位)
        // ----------------------------------------------------------------
        private Label _lblTitle;
        private Label _lblSobaLabel;      // "本日の地金相場 価格"
        private Label _lblPt1000Label;    // "PT1000"
        private Label _lblK24Label;       // "K24"
        private Label _lblNen;            // "年"
        private Label _lblTsuki;          // "月"
        private Label _lblNichi;          // "日"
        private Label _lblPerGram;        // "/g"
        private Label _lblPerGram2;       // "/g" (K24側)
        private Label _lblPtDot;          // "・" (PT 小数点区切)
        private Label _lblK24Dot;         // "・" (K24 小数点区切)
        private Label _lblHinbanLabel;    // "改造対象品番"
        private Label _lblJiganeCodeLabel;// "地金コード"
        private Label _lblSizeLabelHdr;   // "サイズ(ヘッダ)" → 品番一で切り替わる
        private Label _lblDesignNoHdr;    // "デザインNo."
        private Label _lblJiganeJuryoHdr; // "地金付き重量"
        private Label _lblJiganeDaikinHdr;// "地金代金"
        private Label _lblJunkinHdr;      // "純金必要量"
        private Label _lblKotinHdr;       // "工賃"
        private Label _lblMdFdHdr;        // "メレ・FD"
        private Label _lblMdUnit;         // "ct" (MD単位)
        private Label _lblFdUnit;         // "ct" (FD単位)
        private Label _lblGokeimitsumoriHdr; // "合計見積予算（最低値）"
        private Label _lblNote;           // 誤差注記

        // ボタン
        private Button _btnKeisan; // 計算 (VB6: btn_keisan)
        private Button _btnExit;   // メニュー (VB6: btn_exit)

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------
        public FormNr()
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
            // VB6: BackColor=&H00D8FFFF& (水色), Caption="夢仕立て-NRコレクション"
            this.Text = "夢仕立て - NRコレクション";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;

            var labelFont = new Font("ＭＳ Ｐゴシック", 14.25f, FontStyle.Bold);
            var smallFont = new Font("ＭＳ Ｐゴシック", 12f, FontStyle.Bold);
            var titleFont = new Font("ＭＳ Ｐゴシック", 18f, FontStyle.Italic);
            var resultFont = new Font("ＭＳ Ｐゴシック", 15.75f, FontStyle.Bold);
            var btnFont = new Font("ＭＳ Ｐゴシック", 15.75f, FontStyle.Bold);

            // ---- ヘルパー: 1桁コンボ(0-9) ----
            ComboBox MkDigit()
            {
                var c = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Font = labelFont,
                    Size = new Size(50, 27),
                    Location = new Point(0, 0),
                };
                for (int i = 0; i <= 9; i++) c.Items.Add(i.ToString());
                return c;
            }

            // ---- ヘルパー: 固定ラベル ----
            Label MkLbl(string txt, int w = 200, bool bold = true)
            {
                return new Label
                {
                    Text = txt,
                    Font = bold ? labelFont : smallFont,
                    BackColor = Color.Transparent,
                    AutoSize = false,
                    Size = new Size(w, 26),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Location = new Point(0, 0),
                };
            }

            // ---- ヘルパー: 結果ラベル (枠付き右寄せ) ----
            Label MkResult(int w = 180)
            {
                return new Label
                {
                    Text = "0",
                    Font = resultFont,
                    BackColor = SystemColors.Window,
                    ForeColor = SystemColors.WindowText,
                    BorderStyle = BorderStyle.Fixed3D,
                    Size = new Size(w, 28),
                    TextAlign = ContentAlignment.MiddleRight,
                    Location = new Point(0, 0),
                };
            }

            // ================================================================
            // コントロール生成
            // ================================================================

            _lblTitle = new Label { Text = "NRコレクション", Font = titleFont, BackColor = Color.Transparent, AutoSize = true, Location = new Point(0, 0) };
            _lblSobaLabel = MkLbl("本日の地金相場  価格  価格", 330);
            _lblPt1000Label = MkLbl("PT1000", 100);
            _lblK24Label = MkLbl("K24", 60);
            _lblNen = MkLbl("年", 30);
            _lblTsuki = MkLbl("月", 30);
            _lblNichi = MkLbl("日", 30);
            _lblPerGram = MkLbl("/g", 40);
            _lblPerGram2 = MkLbl("/g", 40);
            _lblPtDot = MkLbl("・", 20);
            _lblK24Dot = MkLbl("・", 20);

            // 日付コンボ (年: 千/百/十/一, 月: 十/一, 日: 十/一)
            _cmbSennen = MkDigit();
            _cmbHyakunen = MkDigit();
            _cmbJuNen = MkDigit();
            _cmbIchiNen = MkDigit();
            _cmbJuTsuki = MkDigit();
            _cmbIchiTsuki = MkDigit();
            _cmbJuNichi = MkDigit();
            _cmbIchiNichi = MkDigit();

            // PT相場コンボ (万/千/四/十/一)
            _cmbPtMan = MkDigit();
            _cmbPtSen = MkDigit();
            _cmbPtShi = MkDigit();
            _cmbPtJu = MkDigit();
            _cmbPtIchi = MkDigit();

            // K24相場コンボ (万/千/四/十/一)
            _cmbK24Man = MkDigit();
            _cmbK24Sen = MkDigit();
            _cmbK24Shi = MkDigit();
            _cmbK24Ju = MkDigit();
            _cmbK24Ichi = MkDigit();

            // 品番コンボ
            _cmbHinban1 = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = labelFont, Size = new Size(50, 27), Location = new Point(0, 0) };
            _cmbHinban2 = MkDigit();
            _cmbHinban3 = MkDigit();
            _cmbHinban4 = MkDigit();
            _cmbHinban5 = MkDigit();
            // 品番一: 0-9, T, H (VB6: ループ後に AddItem "T", "H")
            for (int i = 0; i <= 9; i++) _cmbHinban1.Items.Add(i.ToString());
            _cmbHinban1.Items.Add("T");
            _cmbHinban1.Items.Add("H");
            _cmbHinban1.SelectedIndexChanged += CmbHinban1_SelectedIndexChanged;

            // 地金コード: 1, 2, 4, 5
            _cmbJiganeCode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = labelFont, Size = new Size(60, 27), Location = new Point(0, 0) };
            _cmbJiganeCode.Items.AddRange(new object[] { 1, 2, 4, 5 });
            _cmbJiganeCode.SelectedIndexChanged += CmbJiganeCode_SelectedIndexChanged;

            // サイズコンボ
            _cmbSizeJu = MkDigit();
            _cmbSizeIchi = MkDigit();

            // ---- 見出しラベル ----
            _lblHinbanLabel = MkLbl("改造対象品番", 160);
            _lblJiganeCodeLabel = MkLbl("地金コード", 100);
            _lblSizeLabelHdr = MkLbl("サイズ", 100); // 品番一で変更: "単価重量"
            _lblDesignNoHdr = MkLbl("デザインNo.", 160);
            _lblJiganeJuryoHdr = MkLbl("地金付き重量", 160);
            _lblJiganeDaikinHdr = MkLbl("地金代金", 120);
            _lblJunkinHdr = MkLbl("純金必要量", 140);
            _lblKotinHdr = MkLbl("工賃", 60);
            _lblMdFdHdr = MkLbl("メレ・FD", 100);
            _lblMdUnit = MkLbl("ct", 30);
            _lblFdUnit = MkLbl("ct", 30);
            _lblGokeimitsumoriHdr = MkLbl("合計見積予算（最低値）", 230);
            _lblNote = new Label
            {
                Text = "オーダーの商品は量産品と異なり、10%程度の誤差が生じます。",
                Font = smallFont,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(0, 0),
            };

            // ---- 動的ラベル ----
            _lblItemType = new Label { Text = "チェーン", Font = labelFont, BackColor = Color.Transparent, AutoSize = true, Location = new Point(0, 0) };
            _lblSizeType = new Label { Text = "単価重量", Font = labelFont, BackColor = Color.Transparent, AutoSize = true, Location = new Point(0, 0) };
            _lblJiganeCodeDisp = new Label { Text = "", Font = labelFont, BackColor = Color.Transparent, AutoSize = true, Location = new Point(0, 0) };

            // ---- 計算結果ラベル ----
            _lblDesignNo = new Label { Text = "", Font = resultFont, BackColor = SystemColors.Window, BorderStyle = BorderStyle.Fixed3D, Size = new Size(160, 28), TextAlign = ContentAlignment.MiddleCenter, Location = new Point(0, 0) };
            _lblJiganeJuryo = MkResult(160);
            _lblJiganeDaikin = MkResult(200);
            _lblJunkinHitsuyo = MkResult(160);
            _lblKotin = MkResult(200);
            _lblMd = MkResult(100);
            _lblFd = MkResult(100);
            _lblGokeimitsumoriYosan = new Label
            {
                Text = "0",
                Font = resultFont,
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.WindowText,
                BorderStyle = BorderStyle.Fixed3D,
                Size = new Size(220, 28),
                TextAlign = ContentAlignment.MiddleRight,
                Location = new Point(0, 0),
            };

            // ---- ボタン ----
            _btnKeisan = new Button
            {
                Text = "計算",
                Font = btnFont,
                Size = new Size(140, 60),
                BackColor = Color.FromArgb(0xC0, 0xFF, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnKeisan.FlatAppearance.BorderColor = Color.SeaGreen;
            _btnKeisan.FlatAppearance.BorderSize = 2;
            _btnKeisan.Click += BtnKeisan_Click;

            _btnExit = new Button
            {
                Text = "メニュー",
                Font = btnFont,
                Size = new Size(140, 60),
                BackColor = Color.FromArgb(0xFF, 0xC0, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnExit.FlatAppearance.BorderColor = Color.Crimson;
            _btnExit.FlatAppearance.BorderSize = 2;
            _btnExit.Click += BtnExit_Click;

            // ---- コントロール追加 ----
            this.Controls.AddRange(new Control[]
            {
                _lblTitle, _lblSobaLabel, _lblPt1000Label, _lblK24Label,
                _lblNen, _lblTsuki, _lblNichi, _lblPerGram, _lblPerGram2,
                _lblPtDot, _lblK24Dot,
                // 日付
                _cmbSennen, _cmbHyakunen, _cmbJuNen, _cmbIchiNen,
                _cmbJuTsuki, _cmbIchiTsuki,
                _cmbJuNichi, _cmbIchiNichi,
                // PT/K24
                _cmbPtMan, _cmbPtSen, _cmbPtShi, _cmbPtJu, _cmbPtIchi,
                _cmbK24Man, _cmbK24Sen, _cmbK24Shi, _cmbK24Ju, _cmbK24Ichi,
                // 品番
                _cmbHinban1, _cmbHinban2, _cmbHinban3, _cmbHinban4, _cmbHinban5,
                _cmbJiganeCode, _cmbSizeJu, _cmbSizeIchi,
                // 見出しラベル
                _lblHinbanLabel, _lblJiganeCodeLabel, _lblSizeLabelHdr,
                _lblDesignNoHdr, _lblJiganeJuryoHdr, _lblJiganeDaikinHdr,
                _lblJunkinHdr, _lblKotinHdr, _lblMdFdHdr,
                _lblMdUnit, _lblFdUnit,
                _lblGokeimitsumoriHdr, _lblNote,
                // 動的ラベル
                _lblItemType, _lblSizeType, _lblJiganeCodeDisp,
                // 結果ラベル
                _lblDesignNo, _lblJiganeJuryo, _lblJiganeDaikin,
                _lblJunkinHitsuyo, _lblKotin, _lblMd, _lblFd,
                _lblGokeimitsumoriYosan,
                // ボタン
                _btnKeisan, _btnExit,
            });

            this.ResumeLayout(false);
        }

        // ----------------------------------------------------------------
        // フォームロード (VB6: Form_Load → Init_Control)
        // ----------------------------------------------------------------
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.WindowState = FormWindowState.Maximized;
            InitControl();
            CenterControls();
        }

        // ----------------------------------------------------------------
        // フォームリサイズ
        // ----------------------------------------------------------------
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            CenterControls();
        }

        // ----------------------------------------------------------------
        // コントロール初期化 (VB6: Init_Control)
        // ----------------------------------------------------------------
        private void InitControl()
        {
            // ---- 日付コンボに今日の日付をセット ----
            // VB6: str_date = Date → yyyy/mm/dd 形式
            //      千年.Text = 千年.List(Mid(str_date, 1, 1)) ← 1桁目
            //      四年.Text = 四年.List(Mid(str_date, 2, 1)) ← 2桁目
            //      十年.Text = 十年.List(Mid(str_date, 3, 1)) ← 3桁目
            //      一年.Text = 一年.List(Mid(str_date, 4, 1)) ← 4桁目
            //      十月.Text = 十月.List(Mid(str_date, 6, 1)) ← 6桁目 (スラッシュをスキップ)
            //      一月.Text = 一月.List(Mid(str_date, 7, 1)) ← 7桁目
            //      十日.Text = 十日.List(Mid(str_date, 9, 1)) ← 9桁目
            //      一日.Text = 一日.List(Mid(str_date, 10, 1)) ← 10桁目
            string d = DateTime.Today.ToString("yyyy/MM/dd");
            _cmbSennen.SelectedIndex = int.Parse(d[0].ToString());
            _cmbHyakunen.SelectedIndex = int.Parse(d[1].ToString());
            _cmbJuNen.SelectedIndex = int.Parse(d[2].ToString());
            _cmbIchiNen.SelectedIndex = int.Parse(d[3].ToString());
            _cmbJuTsuki.SelectedIndex = int.Parse(d[5].ToString()); // d[4] = '/'
            _cmbIchiTsuki.SelectedIndex = int.Parse(d[6].ToString());
            _cmbJuNichi.SelectedIndex = int.Parse(d[8].ToString()); // d[7] = '/'
            _cmbIchiNichi.SelectedIndex = int.Parse(d[9].ToString());

            // ---- PT/K24 相場を pt_k18テーブルから読み込む ----
            // VB6: SELECT * FROM pt_k18テーブル
            //      PT万.Text = PT万.List(CInt(recTB.Fields("pt万"))) など
            try
            {
                DataTable ptDt = AppState.Db.ExecuteQuery("SELECT * FROM [pt_k18テーブル]");
                if (ptDt.Rows.Count == 0)
                {
                    MessageBox.Show("システムエラーです。", "pt_k18テーブル該当データなし",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    DataRow r = ptDt.Rows[0];
                    _cmbPtMan.SelectedIndex = ToSafeIndex(r, "pt万", _cmbPtMan.Items.Count);
                    _cmbPtSen.SelectedIndex = ToSafeIndex(r, "pt千", _cmbPtSen.Items.Count);
                    _cmbPtShi.SelectedIndex = ToSafeIndex(r, "pt四", _cmbPtShi.Items.Count);
                    _cmbPtJu.SelectedIndex = ToSafeIndex(r, "pt十", _cmbPtJu.Items.Count);
                    _cmbPtIchi.SelectedIndex = ToSafeIndex(r, "pt一", _cmbPtIchi.Items.Count);
                    _cmbK24Man.SelectedIndex = ToSafeIndex(r, "k18万", _cmbK24Man.Items.Count);
                    _cmbK24Sen.SelectedIndex = ToSafeIndex(r, "k18千", _cmbK24Sen.Items.Count);
                    _cmbK24Shi.SelectedIndex = ToSafeIndex(r, "k18四", _cmbK24Shi.Items.Count);
                    _cmbK24Ju.SelectedIndex = ToSafeIndex(r, "k18十", _cmbK24Ju.Items.Count);
                    _cmbK24Ichi.SelectedIndex = ToSafeIndex(r, "k18一", _cmbK24Ichi.Items.Count);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("pt_k18テーブル読込エラー: " + ex.Message, "システムエラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // フォールバック: 全コンボを 0 に設定
                foreach (var cmb in new[] {
                    _cmbPtMan, _cmbPtSen, _cmbPtShi, _cmbPtJu, _cmbPtIchi,
                    _cmbK24Man, _cmbK24Sen, _cmbK24Shi, _cmbK24Ju, _cmbK24Ichi })
                {
                    if (cmb.Items.Count > 0) cmb.SelectedIndex = 0;
                }
            }

            // ---- 品番・地金コード・サイズの初期値 ----
            _cmbHinban1.SelectedIndex = 0;
            _cmbHinban2.SelectedIndex = 0;
            _cmbHinban3.SelectedIndex = 0;
            _cmbHinban4.SelectedIndex = 0;
            _cmbHinban5.SelectedIndex = 0;
            _cmbJiganeCode.SelectedIndex = 0; // "1"
            _cmbSizeJu.SelectedIndex = 0;
            _cmbSizeIchi.SelectedIndex = 0;

            // ---- 計算結果ラベルをクリア ----
            // VB6: デザインNO.Caption = "" など
            ResetResultLabels();
        }

        // ----------------------------------------------------------------
        // 計算結果ラベルをリセット (VB6: Init_Control および btn_keisan_Click 冒頭)
        // ----------------------------------------------------------------
        private void ResetResultLabels()
        {
            _lblDesignNo.Text = "";
            _lblJiganeJuryo.Text = "0.0";
            _lblJiganeDaikin.Text = "0";
            _lblJunkinHitsuyo.Text = "0.0";
            _lblKotin.Text = "0";
            _lblMd.Text = "0";
            _lblFd.Text = "0";
            _lblGokeimitsumoriYosan.Text = "0";
        }

        // ----------------------------------------------------------------
        // コントロール中央配置 (最大化・リサイズ対応)
        // ----------------------------------------------------------------
        private void CenterControls()
        {
            if (this.ClientSize.Width == 0 || this.ClientSize.Height == 0) return;

            int cx = this.ClientSize.Width / 2;
            int lx = cx - 500;  // レイアウト左端
            int y = 15;
            const int dy = 44;

            // タイトル
            _lblTitle.Location = new Point(cx - 90, y);
            y += 35;

            // ----------------------------------------------------------------
            // [行A] 地金相場ヘッダ
            // ----------------------------------------------------------------
            _lblSobaLabel.SetBounds(lx, y, 330, 26);
            y += 28;

            // ----------------------------------------------------------------
            // [行B] 日付 & PT/K24相場 入力行
            //   [千年][百年][十年][一年]年 [十月][一月]月 [十日][一日]日
            //   PT1000: [万][千][四][十][一]/g  K24: [万][千][四][十][一]/g
            // ----------------------------------------------------------------
            int bx = lx; // 日付入力開始X
            int by = y;

            // 日付コンボ
            _cmbSennen.SetBounds(bx, by, 50, 27);
            _cmbHyakunen.SetBounds(bx + 52, by, 50, 27);
            _cmbJuNen.SetBounds(bx + 104, by, 50, 27);
            _cmbIchiNen.SetBounds(bx + 156, by, 50, 27);
            _lblNen.SetBounds(bx + 208, by + 3, 30, 22);
            _cmbJuTsuki.SetBounds(bx + 242, by, 50, 27);
            _cmbIchiTsuki.SetBounds(bx + 294, by, 50, 27);
            _lblTsuki.SetBounds(bx + 346, by + 3, 30, 22);
            _cmbJuNichi.SetBounds(bx + 380, by, 50, 27);
            _cmbIchiNichi.SetBounds(bx + 432, by, 50, 27);
            _lblNichi.SetBounds(bx + 484, by + 3, 30, 22);

            // PT1000 入力
            int px = lx + 530;
            _lblPt1000Label.SetBounds(px - 80, by + 3, 80, 22);
            _cmbPtMan.SetBounds(px, by, 50, 27);
            _cmbPtSen.SetBounds(px + 52, by, 50, 27);
            _cmbPtShi.SetBounds(px + 104, by, 50, 27);
            _lblPtDot.SetBounds(px + 155, by + 3, 20, 22);
            _cmbPtJu.SetBounds(px + 156, by, 50, 27);
            _cmbPtIchi.SetBounds(px + 208, by, 50, 27);
            _lblPerGram.SetBounds(px + 261, by + 3, 35, 22);

            // K24 入力
            int kx = px + 300;
            _lblK24Label.SetBounds(kx - 50, by + 3, 50, 22);
            _cmbK24Man.SetBounds(kx, by, 50, 27);
            _cmbK24Sen.SetBounds(kx + 52, by, 50, 27);
            _cmbK24Shi.SetBounds(kx + 104, by, 50, 27);
            _lblK24Dot.SetBounds(kx + 155, by + 3, 20, 22);
            _cmbK24Ju.SetBounds(kx + 156, by, 50, 27);
            _cmbK24Ichi.SetBounds(kx + 208, by, 50, 27);
            _lblPerGram2.SetBounds(kx + 261, by + 3, 35, 22);

            y += dy;

            // ----------------------------------------------------------------
            // [行C] 品番・地金コード・サイズ 入力行
            //   [ラベル] [品番1][品番2][品番3][品番4][品番5]  [地金コード][表示名]  [サイズ十][サイズ一]
            // ----------------------------------------------------------------
            int inputY = y;
            _lblHinbanLabel.SetBounds(lx, inputY + 3, 160, 22);
            int hx = lx + 165;
            _cmbHinban1.SetBounds(hx, inputY, 50, 27);
            _cmbHinban2.SetBounds(hx + 52, inputY, 50, 27);
            _cmbHinban3.SetBounds(hx + 104, inputY, 50, 27);
            _cmbHinban4.SetBounds(hx + 156, inputY, 50, 27);
            _cmbHinban5.SetBounds(hx + 208, inputY, 50, 27);

            _lblItemType.SetBounds(hx, inputY + 30, 120, 22);

            int jx = hx + 270;
            _lblJiganeCodeLabel.SetBounds(jx, inputY + 3, 100, 22);
            _cmbJiganeCode.SetBounds(jx + 105, inputY, 60, 27);
            _lblJiganeCodeDisp.SetBounds(jx + 170, inputY + 3, 80, 22);

            int sx = jx + 270;
            _lblSizeLabelHdr.SetBounds(sx, inputY + 3, 100, 22);
            _cmbSizeJu.SetBounds(sx + 105, inputY, 50, 27);
            _cmbSizeIchi.SetBounds(sx + 158, inputY, 50, 27);
            _lblSizeType.SetBounds(sx + 105, inputY + 30, 120, 22);

            y += dy + 10;

            // ----------------------------------------------------------------
            // [行D〜H] 計算結果表示エリア
            // ----------------------------------------------------------------

            // デザインNO
            int ry = y;
            _lblDesignNoHdr.SetBounds(lx, ry + 3, 160, 22);
            _lblDesignNo.SetBounds(lx + 165, ry, 160, 28);
            y += dy;

            // 地金付き重量 / 地金代金
            ry = y;
            _lblJiganeJuryoHdr.SetBounds(lx, ry + 3, 160, 22);
            _lblJiganeJuryo.SetBounds(lx + 165, ry, 160, 28);
            _lblJiganeDaikinHdr.SetBounds(lx + 360, ry + 3, 120, 22);
            _lblJiganeDaikin.SetBounds(lx + 485, ry, 200, 28);
            y += dy;

            // 純金必要量 / 工賃
            ry = y;
            _lblJunkinHdr.SetBounds(lx, ry + 3, 140, 22);
            _lblJunkinHitsuyo.SetBounds(lx + 165, ry, 160, 28);
            _lblKotinHdr.SetBounds(lx + 360, ry + 3, 60, 22);
            _lblKotin.SetBounds(lx + 425, ry, 200, 28);
            y += dy;

            // メレ・FD (MD / FD)
            ry = y;
            _lblMdFdHdr.SetBounds(lx, ry + 3, 100, 22);
            _lblMd.SetBounds(lx + 165, ry, 100, 28);
            _lblMdUnit.SetBounds(lx + 268, ry + 3, 30, 22);
            _lblFd.SetBounds(lx + 310, ry, 100, 28);
            _lblFdUnit.SetBounds(lx + 413, ry + 3, 30, 22);
            y += dy;

            // 合計見積予算
            ry = y;
            _lblGokeimitsumoriHdr.SetBounds(lx, ry + 3, 230, 22);
            _lblGokeimitsumoriYosan.SetBounds(lx + 235, ry, 220, 28);
            y += dy;

            // 注記
            _lblNote.SetBounds(lx + 235, y, 500, 22);
            y += dy;

            // ----------------------------------------------------------------
            // ボタン
            // ----------------------------------------------------------------
            int btnY = Math.Max(y + 10, this.ClientSize.Height - 100);
            _btnKeisan.SetBounds(cx - 160, btnY, 140, 60);
            _btnExit.SetBounds(cx + 20, btnY, 140, 60);
        }

        // ----------------------------------------------------------------
        // 品番一 変更時: アイテム/サイズラベル切り替え (VB6: 品番一_Click)
        // ----------------------------------------------------------------
        private void CmbHinban1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_cmbHinban1.SelectedIndex < 0) return;
            string h1 = _cmbHinban1.Text;
            if (h1 == "T")
            {
                // VB6: label_アイテム.Caption = "リング"
                //      Label_サイズ.Caption  = "サイズ"
                _lblItemType.Text = "リング";
                _lblSizeLabelHdr.Text = "サイズ";
            }
            else
            {
                // VB6: label_アイテム.Caption = "チェーン"
                //      Label_サイズ.Caption  = "単価重量"
                _lblItemType.Text = "チェーン";
                _lblSizeLabelHdr.Text = "単価重量";
            }
        }

        // ----------------------------------------------------------------
        // 地金コード 変更時: 種類ラベル切り替え (VB6: 地金コード_Click)
        // ----------------------------------------------------------------
        private void CmbJiganeCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_cmbJiganeCode.SelectedIndex < 0) return;
            int code = (int)_cmbJiganeCode.SelectedItem;
            switch (code)
            {
                case 1: _lblJiganeCodeDisp.Text = "純白"; break; // Pt1000 純白
                case 2: _lblJiganeCodeDisp.Text = "純"; break; // K24 純
                case 4: _lblJiganeCodeDisp.Text = "合白"; break; // Pt850 合白
                case 5: _lblJiganeCodeDisp.Text = "合純"; break; // K18 合純
                default: _lblJiganeCodeDisp.Text = ""; break;
            }
        }

        // ----------------------------------------------------------------
        // 計算ボタン (VB6: btn_keisan_Click)
        // ----------------------------------------------------------------
        private void BtnKeisan_Click(object sender, EventArgs e)
        {
            // 結果をリセット
            ResetResultLabels();

            double dblKake2, dblKake3;

            // ---- 掛け率テーブルを読み込む ----
            // VB6: SELECT * FROM 掛け率テーブル
            try
            {
                DataTable kakeDt = AppState.Db.ExecuteQuery("SELECT * FROM [掛け率テーブル]");
                if (kakeDt.Rows.Count == 0)
                {
                    MessageBox.Show("掛け率テーブルのデータがありません。", "システムエラー",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                dblKake2 = Convert.ToDouble(kakeDt.Rows[0]["掛け率2"]);
                dblKake3 = Convert.ToDouble(kakeDt.Rows[0]["掛け率3"]);
            }
            catch (Exception ex)
            {
                MessageBox.Show("掛け率テーブル読込エラー: " + ex.Message, "システムエラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // ---- 地金テーブルを品番・地金コードで検索 ----
            // VB6: SELECT * FROM 地金テーブル WHERE 品番 = '品番五桁' AND 地金コード = '地金コード'
            // ※ 地金テーブルはNRコレクション用の製品仕様テーブル
            string hinban5 = _cmbHinban1.Text + _cmbHinban2.Text + _cmbHinban3.Text
                           + _cmbHinban4.Text + _cmbHinban5.Text;
            string jiganeCodeStr = _cmbJiganeCode.SelectedItem?.ToString() ?? "1";

            DataTable jdDt;
            try
            {
                string sql = "SELECT * FROM [地金テーブル] WHERE "
                           + "[品番] = '" + hinban5 + "' "
                           + "AND [地金コード] = '" + jiganeCodeStr + "'";
                jdDt = AppState.Db.ExecuteQuery(sql);
                if (jdDt.Rows.Count == 0)
                {
                    MessageBox.Show("お探しの品番、地金コードはありません。", "地金テーブル内容チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("地金テーブル検索エラー: " + ex.Message, "システムエラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataRow jr = jdDt.Rows[0];

            // ---- サイズチェック (品番一 == "H" のとき) ----
            // VB6: If 品番一.Text = "H" Then If サイズ十=0 And サイズ一=0 → エラー
            string h1 = _cmbHinban1.Text;
            if (h1 == "H")
            {
                if (_cmbSizeJu.SelectedIndex == 0 && _cmbSizeIchi.SelectedIndex == 0)
                {
                    MessageBox.Show("長さを入力してください。", "入力チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // ---- 地金相場入力チェック ----
            // VB6: 地金コード=1 → PT価格が全部0ならエラー / 地金コード=2 → K24価格全部0ならエラー / else → PT価格チェック
            int jiganeCode = (int)_cmbJiganeCode.SelectedItem;
            int ptPrice = GetPtPrice();
            int k24Price = GetK24Price();

            if (jiganeCode == 1)
            {
                if (ptPrice == 0)
                {
                    MessageBox.Show("地金相場を入力してください。", "入力チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            else if (jiganeCode == 2)
            {
                if (k24Price == 0)
                {
                    MessageBox.Show("地金相場を入力してください。", "入力チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            else
            {
                if (ptPrice == 0)
                {
                    MessageBox.Show("地金相場を入力してください。", "入力チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // ================================================================
            // 計算処理
            // ================================================================

            // ---- 地金付き重量 (克付き重量) ----
            // VB6: wk_dbl = recTB.Fields("地金重量")
            //      If 品番一 = "T" Then wk_dbl /= 10
            //      Else wk_dbl = wk_dbl * (サイズ十*10 + サイズ一) / 1000
            //      Format "#0.0" して1桁丸め
            double jiganeJuryo = Convert.ToDouble(jr["地金重量"]);
            if (h1 == "T")
            {
                jiganeJuryo = jiganeJuryo / 10.0;
            }
            else
            {
                int sizeVal = _cmbSizeJu.SelectedIndex * 10 + _cmbSizeIchi.SelectedIndex;
                jiganeJuryo = jiganeJuryo * sizeVal / 1000.0;
            }
            jiganeJuryo = RoundToDecimal1(jiganeJuryo);
            _lblJiganeJuryo.Text = jiganeJuryo.ToString("#0.0");

            // ---- 純金必要量 ----
            // VB6: wk_dbl = wk_dbl / (CInt(recTB.Fields("純度")) / 100)
            //      Format "#0.0"
            double jundoRatio = Convert.ToInt32(jr["純度"]) / 100.0;
            double junkinHitsuyo = RoundToDecimal1(jiganeJuryo / jundoRatio);
            _lblJunkinHitsuyo.Text = junkinHitsuyo.ToString("#0.0");
            double wkDbl4 = junkinHitsuyo; // 工賃計算に使用

            // ---- 地金代金 ----
            // VB6:
            //   地金コード=1: 純金必要量 × 地金割合 × PT価格 / 1000 × kake2 / 10
            //   地金コード=2: 純金必要量 × 地金割合 × K24価格 / 24 × kake2 / 10
            //   else        : 純金必要量 × 850 × PT価格 / 1000 × kake2 / 10
            int jiganeWariai = Convert.ToInt32(jr["地金割合"]);
            double jiganeDaikin;
            if (jiganeCode == 1)
            {
                jiganeDaikin = junkinHitsuyo * jiganeWariai * ptPrice / 1000.0 * dblKake2 / 10.0;
            }
            else if (jiganeCode == 2)
            {
                jiganeDaikin = junkinHitsuyo * jiganeWariai * k24Price / 24.0 * dblKake2 / 10.0;
            }
            else
            {
                // 地金コード 4/5: Pt850 相当
                jiganeDaikin = junkinHitsuyo * 850.0 * ptPrice / 1000.0 * dblKake2 / 10.0;
            }
            long jiganeDaikinRounded = (long)Math.Round(jiganeDaikin, MidpointRounding.AwayFromZero);
            _lblJiganeDaikin.Text = jiganeDaikinRounded.ToString("#0");

            // ---- 工賃 ----
            // VB6:
            //   品番一 = "T": 工賃 = 単価工賃 × kake3 / 10
            //   else        : 工賃 = 純金必要量 × 単価工賃 × kake3 / 10
            double tankaKotin = Convert.ToDouble(jr["単価工賃"]);
            double kotin;
            if (h1 == "T")
            {
                kotin = tankaKotin * dblKake3 / 10.0;
            }
            else
            {
                kotin = wkDbl4 * tankaKotin * dblKake3 / 10.0;
            }
            _lblKotin.Text = ((long)Math.Round(kotin, MidpointRounding.AwayFromZero)).ToString();

            // ---- 合計見積予算 ----
            // VB6: 合計見積予算 = wk_dbl2 + wk_dbl3
            //      (wk_dbl2 = 地金代金(整数), wk_dbl3 = 工賃(double))
            long gokei = jiganeDaikinRounded + (long)Math.Round(kotin, MidpointRounding.AwayFromZero);
            _lblGokeimitsumoriYosan.Text = gokei.ToString("#,##0");

            // ---- MD (メレCT) ----
            // VB6: wk_dbl = Fields("メレCT"); if wk_dbl <> 0 then wk_dbl /= 100
            //      MD.Caption = Format(wk_dbl, "#0.00")
            double meCt = Convert.ToDouble(jr["メレCT"]);
            if (meCt != 0) meCt /= 100.0;
            _lblMd.Text = meCt.ToString("#0.00");

            // ---- FD (FDカラット) ----
            double fdCt = Convert.ToDouble(jr["FDカラット"]);
            if (fdCt != 0) fdCt /= 100.0;
            _lblFd.Text = fdCt.ToString("#0.00");

            // ---- デザインNO 表示 ----
            // VB6: "R" (品番一=="T") or "N" (その他) + 地金コード + 品番5桁
            string designPrefix = (h1 == "T") ? "R" : "N";
            _lblDesignNo.Text = designPrefix + jiganeCodeStr + hinban5;

            // ================================================================
            // PT/K24 相場インデックスを pt_k18テーブルに保存
            // VB6: DELETE * FROM pt_k18テーブル
            //      INSERT INTO pt_k18テーブル VALUES (PT万idx, ..., K24一idx, '0'×74)
            // ================================================================
            SavePtK18Table();
        }

        // ----------------------------------------------------------------
        // メニューへ戻るボタン (VB6: btn_exit_Click → Unload form_nr → form_menu.Visible = True)
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

        // ================================================================
        // ヘルパーメソッド
        // ================================================================

        /// <summary>
        /// PT 相場価格を 5桁の整数として返す。
        /// VB6: PT万 * 10000 + PT千 * 1000 + PT四 * 100 + PT十 * 10 + PT一
        /// </summary>
        private int GetPtPrice()
        {
            return _cmbPtMan.SelectedIndex * 10000
                 + _cmbPtSen.SelectedIndex * 1000
                 + _cmbPtShi.SelectedIndex * 100
                 + _cmbPtJu.SelectedIndex * 10
                 + _cmbPtIchi.SelectedIndex;
        }

        /// <summary>
        /// K24 相場価格を 5桁の整数として返す。
        /// VB6: K24万 * 10000 + K24千 * 1000 + K24四 * 100 + K24十 * 10 + K24一
        /// </summary>
        private int GetK24Price()
        {
            return _cmbK24Man.SelectedIndex * 10000
                 + _cmbK24Sen.SelectedIndex * 1000
                 + _cmbK24Shi.SelectedIndex * 100
                 + _cmbK24Ju.SelectedIndex * 10
                 + _cmbK24Ichi.SelectedIndex;
        }

        /// <summary>
        /// pt_k18テーブルに PT/K24 各桁のインデックスを保存する。
        /// VB6: DELETE * → INSERT INTO pt_k18テーブル VALUES(PT万idx, ..., '0' × 74)
        /// pt_k18テーブルは計 84 列。先頭10列が PT/K24 インデックス、残り74列は '0'。
        /// </summary>
        private void SavePtK18Table()
        {
            try
            {
                AppState.Db.ExecuteNonQuery("DELETE * FROM [pt_k18テーブル]");

                // 先頭10列: PT万/千/四/十/一, K24万/千/四/十/一 のインデックス値
                string ins = "INSERT INTO [pt_k18テーブル] VALUES ("
                    + "'" + _cmbPtMan.SelectedIndex + "',"
                    + "'" + _cmbPtSen.SelectedIndex + "',"
                    + "'" + _cmbPtShi.SelectedIndex + "',"
                    + "'" + _cmbPtJu.SelectedIndex + "',"
                    + "'" + _cmbPtIchi.SelectedIndex + "',"
                    + "'" + _cmbK24Man.SelectedIndex + "',"
                    + "'" + _cmbK24Sen.SelectedIndex + "',"
                    + "'" + _cmbK24Shi.SelectedIndex + "',"
                    + "'" + _cmbK24Ju.SelectedIndex + "',"
                    + "'" + _cmbK24Ichi.SelectedIndex + "',";

                // 残り74列はすべて '0' (VB6 の '0' × 74 を再現)
                var zeros = new System.Text.StringBuilder();
                for (int i = 0; i < 73; i++) zeros.Append("'0',");
                zeros.Append("'0')");
                ins += zeros.ToString();

                AppState.Db.ExecuteNonQuery(ins);
            }
            catch (Exception ex)
            {
                // 保存失敗は致命的ではないため警告のみ
                MessageBox.Show("pt_k18テーブル更新エラー: " + ex.Message, "警告",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// DataRow のフィールド値を安全に SelectedIndex として返す。
        /// 範囲外の場合は 0 を返す。
        /// </summary>
        private static int ToSafeIndex(DataRow row, string fieldName, int maxCount)
        {
            try
            {
                int idx = Convert.ToInt32(row[fieldName]);
                return (idx >= 0 && idx < maxCount) ? idx : 0;
            }
            catch { return 0; }
        }

        /// <summary>
        /// 小数点1桁に四捨五入する (VB6: Format(wk_dbl, "#0.0") → CDbl(wk_str))。
        /// </summary>
        private static double RoundToDecimal1(double v)
        {
            return Math.Round(v, 1, MidpointRounding.AwayFromZero);
        }
    }
}
