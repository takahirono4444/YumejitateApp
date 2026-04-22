using System;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace YumejitateApp
{
    // ================================================================
    // FormTedukuri_Part1.cs
    // VB6: form_tedukuri.frm → C# + WinForms 移植（Part 1 / 2）
    // オーダーメイドお見積り画面 — フィールド定義・画面構築・初期化
    // ================================================================
    public partial class FormTedukuri : Form
    {
        // ============================================================
        // 定数（VB6: 夢仕立て.bas の Public Const Hijuu_xxx）
        // ============================================================
        private const double HijuuPt900 = 20.5;   // Pt900 比重
        private const double HijuuK18 = 17.1;   // K18   比重
        private const double HijuuWgpg = 17.1;   // WG/PG 比重（K18 と同値）

        // 計算済み／未計算のカラー（VB6: &HFFFFC0 / &H80000005）
        private static readonly Color CCalc = Color.FromArgb(255, 255, 192);
        private static readonly Color CDefault = SystemColors.Window;

        // ============================================================
        // フォームレベル変数
        // VB6: Public Order_xxx As Double / Private Chk_Flg_Keisan As Boolean
        // ============================================================
        private double _orderPrice;     // 各部品の税抜き価格の累計
        private double _orderPt900;     // Pt900 重量合計 (g)
        private double _orderK18;       // K18   重量合計 (g)
        private double _orderWgpg;      // WG/PG 重量合計 (g)
        private bool _chkFlgKeisan;   // 計算エラーフラグ（True=中断）
        private bool _forceClose;     // btn_menu / btn_back による強制クローズ

        // ============================================================
        // コントロール宣言
        // ============================================================

        // ── タブコントロール ──────────────────────────────────────────
        private TabControl _tabMain;

        // ── 下部固定パネル ────────────────────────────────────────────
        private Panel _pnlBottom;

        // 参照データラベル（VB6: lbl_sample_hinban / pt900 / k18 / wgpg / code / price）
        private Label _lblSmpHinban, _lblSmpPt900, _lblSmpK18;
        private Label _lblSmpWgpg, _lblSmpCode, _lblSmpPrice;

        // 計算結果ラベル（VB6: lbl_price / lbl_pt900 / lbl_k18 / lbl_wgpg / lbl_ingo）
        private Label _lblPrice, _lblPt900, _lblK18, _lblWgpg, _lblIngo;

        // ボタン（VB6: btn_keisan / btn_clear / btn_print / btn_back / btn_menu / btn_save）
        private Button _btnKeisan, _btnClear, _btnPrint;
        private Button _btnBack, _btnMenu, _btnSave;

        // 保存番号コンボ（VB6: 保存1〜7）
        private ComboBox _cmb保1, _cmb保2, _cmb保3, _cmb保4;
        private ComboBox _cmb保5, _cmb保6, _cmb保7;

        // ── 腕１（VB6: 腕A1〜F1プラマイ）────────────────────────────
        private ComboBox _cU1A, _cU1B, _cU1D, _cU1Pm;         // 地金/形状/特殊加工/プラマイ
        private ComboBox _cU1TJu, _cU1TJi, _cU1TJs;            // 天幅 十一小
        private ComboBox _cU1TAJu, _cU1TAJi, _cU1TAJs;         // 天厚 十一小
        private ComboBox _cU1BJu, _cU1BJi, _cU1BJs;          // 底幅 十一小
        private ComboBox _cU1BAJu, _cU1BAJi, _cU1BAJs;         // 底厚 十一小
        private ComboBox _cU1SzJu, _cU1SzJi;                   // リングサイズ 十一
        private ComboBox _cU1F;                                  // 個数

        // ── 腕２ ──────────────────────────────────────────────────────
        private ComboBox _cU2A, _cU2B, _cU2D, _cU2Pm;
        private ComboBox _cU2TJu, _cU2TJi, _cU2TJs;
        private ComboBox _cU2TAJu, _cU2TAJi, _cU2TAJs;
        private ComboBox _cU2BJu, _cU2BJi, _cU2BJs;
        private ComboBox _cU2BAJu, _cU2BAJi, _cU2BAJs;
        private ComboBox _cU2SzJu, _cU2SzJi;
        private ComboBox _cU2F;

        // ── 板線材１（VB6: 板線材A1〜F1）────────────────────────────
        private ComboBox _cIt1A, _cIt1B, _cIt1C;               // 部材種類/地金/形状
        private ComboBox _cIt1LJu, _cIt1LJi, _cIt1LJs;         // 長径 十一小
        private ComboBox _cIt1SJu, _cIt1SJi, _cIt1SJs;         // 短径 十一小
        private ComboBox _cIt1AJu, _cIt1AJi, _cIt1AJs;         // 厚み 十一小
        private ComboBox _cIt1F;

        // ── 板線材２ ──────────────────────────────────────────────────
        private ComboBox _cIt2A, _cIt2B, _cIt2C;
        private ComboBox _cIt2LJu, _cIt2LJi, _cIt2LJs;
        private ComboBox _cIt2SJu, _cIt2SJi, _cIt2SJs;
        private ComboBox _cIt2AJu, _cIt2AJi, _cIt2AJs;
        private ComboBox _cIt2F;

        // ── 石座１（VB6: 石座A1〜F1プラマイ）───────────────────────
        private ComboBox _cIsz1A, _cIsz1B, _cIsz1D, _cIsz1E, _cIsz1Pm;
        private ComboBox _cIsz1LJu, _cIsz1LJi, _cIsz1LJs;     // 長径
        private ComboBox _cIsz1SJu, _cIsz1SJi, _cIsz1SJs;     // 短径
        private ComboBox _cIsz1F;

        // ── 石座２ ────────────────────────────────────────────────────
        private ComboBox _cIsz2A, _cIsz2B, _cIsz2D, _cIsz2E, _cIsz2Pm;
        private ComboBox _cIsz2LJu, _cIsz2LJi, _cIsz2LJs;
        private ComboBox _cIsz2SJu, _cIsz2SJi, _cIsz2SJs;
        private ComboBox _cIsz2F;

        // ── 石座３ ────────────────────────────────────────────────────
        private ComboBox _cIsz3A, _cIsz3B, _cIsz3D, _cIsz3E, _cIsz3Pm;
        private ComboBox _cIsz3LJu, _cIsz3LJi, _cIsz3LJs;
        private ComboBox _cIsz3SJu, _cIsz3SJi, _cIsz3SJs;
        private ComboBox _cIsz3F;

        // ── 石留め 1〜4（VB6: 石留めA1〜E1一 等）───────────────────
        private ComboBox _cIsm1D, _cIsm1A, _cIsm1B, _cIsm1C, _cIsm1EJu, _cIsm1EJi;
        private ComboBox _cIsm2D, _cIsm2A, _cIsm2B, _cIsm2C, _cIsm2EJu, _cIsm2EJi;
        private ComboBox _cIsm3D, _cIsm3A, _cIsm3B, _cIsm3C, _cIsm3EJu, _cIsm3EJi;
        private ComboBox _cIsm4D, _cIsm4A, _cIsm4B, _cIsm4C, _cIsm4EJu, _cIsm4EJi;

        // ── ダイヤ 1〜4（VB6: ダイヤA1〜C1一 等）───────────────────
        private ComboBox _cDia1A, _cDia1B, _cDia1Kigo, _cDia1Ju, _cDia1Ji;
        private ComboBox _cDia2A, _cDia2B, _cDia2Kigo, _cDia2Ju, _cDia2Ji;
        private ComboBox _cDia3A, _cDia3B, _cDia3Kigo, _cDia3Ju, _cDia3Ji;
        private ComboBox _cDia4A, _cDia4B, _cDia4Kigo, _cDia4Ju, _cDia4Ji;

        // ── ロー付け（VB6: ロー付けA1〜D1一）───────────────────────
        private ComboBox _cRoA1, _cRoB1Ju, _cRoB1Ji;   // 面ロー種類・個数
        private ComboBox _cRoC1, _cRoD1Ju, _cRoD1Ji;   // 点ロー種類・個数

        // ── 加工グレード（VB6: 加工グレードA1 / B1）────────────────
        private ComboBox _cKakouA1; // 加工難易度
        private ComboBox _cKakouB1; // 加工グレード

        // ============================================================
        // コンストラクタ
        // ============================================================
        public FormTedukuri()
        {
            InitializeComponent();
        }

        // ============================================================
        // コントロール初期化（VB6: Form_Load → Init_Control 相当）
        // ============================================================
        private void InitializeComponent()
        {
            // フォーム基本設定（VB6: Caption / BackColor / WindowState）
            this.Text = "夢仕立て-オーダーメイドお見積り";
            this.BackColor = Color.FromArgb(216, 255, 255); // &H00D8FFFF
            this.WindowState = FormWindowState.Maximized;
            this.StartPosition = FormStartPosition.WindowsDefaultBounds;
            this.Font = new Font("MS Pゴシック", 11f, FontStyle.Bold);
            this.Load += new EventHandler(FormTedukuri_Load);
            this.FormClosing += new FormClosingEventHandler(FormTedukuri_FormClosing);

            // 下部固定パネルを先に構築
            BuildBottomPanel();

            // タブコントロール
            BuildTabControl();

            // Bottom を先に追加しないと TabControl が Bottom エリアを覆う
            this.Controls.Add(_pnlBottom);
            this.Controls.Add(_tabMain);

            // Part2 のボタンイベントを登録
            WireEvents();
        }

        // ─────────────────────────────────────────────────────────────
        // 下部固定パネル構築
        // 参照データ・計算結果・保存番号・ボタンを常時表示
        // ─────────────────────────────────────────────────────────────
        private void BuildBottomPanel()
        {
            _pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 220,
                BackColor = Color.FromArgb(235, 235, 250),
                Padding = new Padding(6, 4, 6, 4)
            };

            // ── 参照データ GroupBox ───────────────────────────────────
            var grpSmp = new GroupBox
            {
                Text = "参照データ（画像表示から転送）",
                Location = new Point(6, 4),
                Size = new Size(500, 90),
                Font = new Font("MS Pゴシック", 10f)
            };
            _lblSmpHinban = MkSmpLbl("品番: ─────");
            _lblSmpPt900 = MkSmpLbl("Pt900: ─ g");
            _lblSmpK18 = MkSmpLbl("K18: ─ g");
            _lblSmpWgpg = MkSmpLbl("WG/PG: ─ g");
            _lblSmpCode = MkSmpLbl("Code: ─");
            _lblSmpPrice = MkSmpLbl("金額: ─ 円");
            var flpSmp = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                WrapContents = true,
                Padding = new Padding(3)
            };
            flpSmp.Controls.AddRange(new Control[] {
                _lblSmpHinban, _lblSmpPt900, _lblSmpK18,
                _lblSmpWgpg,   _lblSmpCode,  _lblSmpPrice
            });
            grpSmp.Controls.Add(flpSmp);

            // ── 計算結果 GroupBox ─────────────────────────────────────
            var grpRes = new GroupBox
            {
                Text = "計算結果",
                Location = new Point(512, 4),
                Size = new Size(440, 90),
                Font = new Font("MS Pゴシック", 10f)
            };
            _lblPrice = MkResLbl("合計金額: ─ 円（税込）", 410);
            _lblPt900 = MkResLbl("Pt900: ─ g", 195);
            _lblK18 = MkResLbl("K18: ─ g", 195);
            _lblWgpg = MkResLbl("WG/PG: ─ g", 195);
            _lblIngo = MkResLbl("─", 195);
            var flpRes = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                WrapContents = true,
                Padding = new Padding(3)
            };
            flpRes.Controls.AddRange(new Control[] {
                _lblPrice, _lblPt900, _lblK18, _lblWgpg, _lblIngo
            });
            grpRes.Controls.Add(flpRes);

            // ── 保存番号行 ────────────────────────────────────────────
            _cmb保1 = MkDigit(); _cmb保2 = MkDigit(); _cmb保3 = MkDigit();
            _cmb保4 = MkDigit(); _cmb保5 = MkDigit(); _cmb保6 = MkDigit();
            _cmb保7 = MkCmb(48, "A", "B", "C", "D", "E");

            var pnlSaveNo = new Panel { Location = new Point(6, 100), Size = new Size(600, 36) };
            var flpSaveNo = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight
            };
            flpSaveNo.Controls.Add(MkLbl("作業書番号:"));
            flpSaveNo.Controls.AddRange(new Control[] {
                _cmb保1, _cmb保2, _cmb保3, _cmb保4, _cmb保5, _cmb保6, _cmb保7
            });
            pnlSaveNo.Controls.Add(flpSaveNo);

            // ── ボタン行 ──────────────────────────────────────────────
            _btnKeisan = MkBtn("計算", 85, Color.LightGreen);
            _btnClear = MkBtn("クリアー", 85, Color.LightYellow);
            _btnPrint = MkBtn("印刷", 85, SystemColors.Control);
            _btnBack = MkBtn("戻る", 65, SystemColors.Control);
            _btnMenu = MkBtn("メニュー", 85, Color.LightSalmon);
            _btnSave = MkBtn("画面情報保存", 130, Color.FromArgb(255, 230, 180));

            var pnlBtn = new Panel { Location = new Point(6, 140), Size = new Size(700, 70) };
            var flpBtn = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight
            };
            flpBtn.Controls.AddRange(new Control[] {
                _btnKeisan, _btnClear, _btnPrint, _btnBack, _btnMenu, _btnSave
            });
            pnlBtn.Controls.Add(flpBtn);

            _pnlBottom.Controls.AddRange(new Control[] {
                grpSmp, grpRes, pnlSaveNo, pnlBtn
            });
        }

        // ─────────────────────────────────────────────────────────────
        // タブコントロール構築（4タブ）
        // ─────────────────────────────────────────────────────────────
        private void BuildTabControl()
        {
            _tabMain = new TabControl { Dock = DockStyle.Fill };
            _tabMain.Font = new Font("MS Pゴシック", 11f, FontStyle.Bold);

            _tabMain.TabPages.Add(BuildUdeTab());         // 腕1・腕2
            _tabMain.TabPages.Add(BuildItaIszTab());      // 板線材・石座
            _tabMain.TabPages.Add(BuildIsmdmeTab());      // 石留め
            _tabMain.TabPages.Add(BuildDiaRoTab());       // ダイヤ・ロー付け・加工グレード
        }

        // ─────────────────────────────────────────────────────────────
        // Tab 1: 腕１・腕２
        // ─────────────────────────────────────────────────────────────
        private TabPage BuildUdeTab()
        {
            var tp = new TabPage("腕（リング腕）");
            var pnl = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(8) };

            var grp1 = BuildUde1Group();
            grp1.Location = new Point(8, 8);

            var grp2 = BuildUde2Group();
            grp2.Location = new Point(8, grp1.Bottom + 10);

            pnl.Controls.AddRange(new Control[] { grp1, grp2 });
            tp.Controls.Add(pnl);
            return tp;
        }

        private GroupBox BuildUde1Group()
        {
            // 腕１: 地金 / 形状 / 天幅 / 天厚 / 底幅 / 底厚 / 特殊加工 / リングサイズ / 個数
            _cU1A = MkCmb(90, "Pt900", "K18", "WG/PG");
            _cU1B = MkCmb(80, "平打ち", "甲丸");
            _cU1D = MkCmb(110, "通常", "中浅抜き", "中深抜き");
            _cU1Pm = MkCmb(50, "＋", "−");
            _cU1TJu = MkDigit(); _cU1TJi = MkDigit(); _cU1TJs = MkDigit();
            _cU1TAJu = MkDigit(); _cU1TAJi = MkDigit(); _cU1TAJs = MkDigit();
            _cU1BJu = MkDigit(); _cU1BJi = MkDigit(); _cU1BJs = MkDigit();
            _cU1BAJu = MkDigit(); _cU1BAJi = MkDigit(); _cU1BAJs = MkDigit();
            _cU1SzJu = MkRsSzJ(); _cU1SzJi = MkDigit();
            _cU1F = MkDigit();

            return BuildUdeGroup("腕１（リング腕）",
                _cU1A, _cU1B, _cU1D, _cU1Pm,
                _cU1TJu, _cU1TJi, _cU1TJs,
                _cU1TAJu, _cU1TAJi, _cU1TAJs,
                _cU1BJu, _cU1BJi, _cU1BJs,
                _cU1BAJu, _cU1BAJi, _cU1BAJs,
                _cU1SzJu, _cU1SzJi, _cU1F);
        }

        private GroupBox BuildUde2Group()
        {
            _cU2A = MkCmb(90, "Pt900", "K18", "WG/PG");
            _cU2B = MkCmb(80, "平打ち", "甲丸");
            _cU2D = MkCmb(110, "通常", "中浅抜き", "中深抜き");
            _cU2Pm = MkCmb(50, "＋", "−");
            _cU2TJu = MkDigit(); _cU2TJi = MkDigit(); _cU2TJs = MkDigit();
            _cU2TAJu = MkDigit(); _cU2TAJi = MkDigit(); _cU2TAJs = MkDigit();
            _cU2BJu = MkDigit(); _cU2BJi = MkDigit(); _cU2BJs = MkDigit();
            _cU2BAJu = MkDigit(); _cU2BAJi = MkDigit(); _cU2BAJs = MkDigit();
            _cU2SzJu = MkRsSzJ(); _cU2SzJi = MkDigit();
            _cU2F = MkDigit();

            return BuildUdeGroup("腕２（リング腕）",
                _cU2A, _cU2B, _cU2D, _cU2Pm,
                _cU2TJu, _cU2TJi, _cU2TJs,
                _cU2TAJu, _cU2TAJi, _cU2TAJs,
                _cU2BJu, _cU2BJi, _cU2BJs,
                _cU2BAJu, _cU2BAJi, _cU2BAJs,
                _cU2SzJu, _cU2SzJi, _cU2F);
        }

        // 腕セクション GroupBox の共通ビルダー
        private GroupBox BuildUdeGroup(string title,
            ComboBox cmbA, ComboBox cmbB, ComboBox cmbD, ComboBox cmbPm,
            ComboBox tJu, ComboBox tJi, ComboBox tJs,
            ComboBox taJu, ComboBox taJi, ComboBox taJs,
            ComboBox bJu, ComboBox bJi, ComboBox bJs,
            ComboBox baJu, ComboBox baJi, ComboBox baJs,
            ComboBox szJu, ComboBox szJi, ComboBox cmbF)
        {
            var grp = new GroupBox { Text = title, AutoSize = true, Width = 900, Padding = new Padding(6) };
            var flp = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                Dock = DockStyle.Fill
            };

            flp.Controls.Add(MkFlpRow(
                MkLbl("地金:", 45), cmbA,
                MkLbl("  形状:", 50), cmbB,
                MkLbl("  特殊加工:", 75), cmbD));

            flp.Controls.Add(MkFlpRow(
                MkLbl("天幅:", 45), tJu, tJi, MkLbl(".", 8), tJs, MkLbl(" mm  ", 35),
                MkLbl("天厚:", 45), taJu, taJi, MkLbl(".", 8), taJs, MkLbl(" mm", 30)));

            flp.Controls.Add(MkFlpRow(
                MkLbl("底幅:", 45), bJu, bJi, MkLbl(".", 8), bJs, MkLbl(" mm  ", 35),
                MkLbl("底厚:", 45), baJu, baJi, MkLbl(".", 8), baJs, MkLbl(" mm", 30)));

            flp.Controls.Add(MkFlpRow(
                MkLbl("リングサイズ:", 90), szJu, szJi, MkLbl(" 号  ", 30),
                MkLbl("個数:", 45), cmbPm, cmbF));

            grp.Controls.Add(flp);
            return grp;
        }

        // ─────────────────────────────────────────────────────────────
        // Tab 2: 板線材・石座
        // ─────────────────────────────────────────────────────────────
        private TabPage BuildItaIszTab()
        {
            var tp = new TabPage("板線材・石座");
            var pnl = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(8) };

            var grpIt1 = BuildIta1Group(); grpIt1.Location = new Point(8, 8);
            var grpIt2 = BuildIta2Group(); grpIt2.Location = new Point(8, grpIt1.Bottom + 8);
            var grpIsz1 = BuildIsz1Group(); grpIsz1.Location = new Point(8, grpIt2.Bottom + 8);
            var grpIsz2 = BuildIsz2Group(); grpIsz2.Location = new Point(8, grpIsz1.Bottom + 8);
            var grpIsz3 = BuildIsz3Group(); grpIsz3.Location = new Point(8, grpIsz2.Bottom + 8);

            pnl.Controls.AddRange(new Control[] { grpIt1, grpIt2, grpIsz1, grpIsz2, grpIsz3 });
            tp.Controls.Add(pnl);
            return tp;
        }

        private GroupBox BuildIta1Group()
        {
            _cIt1A = MkCmb(60, "棒", "板");
            _cIt1B = MkCmb(90, "Pt900", "K18", "WG/PG");
            _cIt1C = MkCmb(100, "円、楕円", "角", "ドロップ");
            _cIt1LJu = MkDigit(); _cIt1LJi = MkDigit(); _cIt1LJs = MkDigit();
            _cIt1SJu = MkDigit(); _cIt1SJi = MkDigit(); _cIt1SJs = MkDigit();
            _cIt1AJu = MkDigit(); _cIt1AJi = MkDigit(); _cIt1AJs = MkDigit();
            _cIt1F = MkDigit();
            return BuildItaGroup("板線材１",
                _cIt1A, _cIt1B, _cIt1C,
                _cIt1LJu, _cIt1LJi, _cIt1LJs,
                _cIt1SJu, _cIt1SJi, _cIt1SJs,
                _cIt1AJu, _cIt1AJi, _cIt1AJs, _cIt1F);
        }

        private GroupBox BuildIta2Group()
        {
            _cIt2A = MkCmb(60, "棒", "板");
            _cIt2B = MkCmb(90, "Pt900", "K18", "WG/PG");
            _cIt2C = MkCmb(100, "円、楕円", "角", "ドロップ");
            _cIt2LJu = MkDigit(); _cIt2LJi = MkDigit(); _cIt2LJs = MkDigit();
            _cIt2SJu = MkDigit(); _cIt2SJi = MkDigit(); _cIt2SJs = MkDigit();
            _cIt2AJu = MkDigit(); _cIt2AJi = MkDigit(); _cIt2AJs = MkDigit();
            _cIt2F = MkDigit();
            return BuildItaGroup("板線材２",
                _cIt2A, _cIt2B, _cIt2C,
                _cIt2LJu, _cIt2LJi, _cIt2LJs,
                _cIt2SJu, _cIt2SJi, _cIt2SJs,
                _cIt2AJu, _cIt2AJi, _cIt2AJs, _cIt2F);
        }

        private GroupBox BuildItaGroup(string title,
            ComboBox cA, ComboBox cB, ComboBox cC,
            ComboBox lJu, ComboBox lJi, ComboBox lJs,
            ComboBox sJu, ComboBox sJi, ComboBox sJs,
            ComboBox aJu, ComboBox aJi, ComboBox aJs,
            ComboBox cF)
        {
            var grp = new GroupBox { Text = title, AutoSize = true, Width = 900, Padding = new Padding(6) };
            var flp = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                Dock = DockStyle.Fill
            };
            flp.Controls.Add(MkFlpRow(
                MkLbl("部材:", 45), cA,
                MkLbl("  地金:", 50), cB,
                MkLbl("  形状:", 50), cC));
            flp.Controls.Add(MkFlpRow(
                MkLbl("長径:", 45), lJu, lJi, MkLbl(".", 8), lJs, MkLbl(" mm  ", 35),
                MkLbl("短径:", 45), sJu, sJi, MkLbl(".", 8), sJs, MkLbl(" mm", 30)));
            flp.Controls.Add(MkFlpRow(
                MkLbl("厚み/長さ:", 75), aJu, aJi, MkLbl(".", 8), aJs, MkLbl(" mm  ", 35),
                MkLbl("個数:", 45), cF));
            grp.Controls.Add(flp);
            return grp;
        }

        private GroupBox BuildIsz1Group()
        {
            _cIsz1A = MkCmb(90, "Pt900", "K18", "WG/PG");
            _cIsz1B = MkCmb(120, "円", "楕円", "角", "ドロップ", "0.1ct未満円", "0.1ct未満ﾌｧﾝｼｰ");
            _cIsz1D = MkCmb(100, "通常爪", "覆輪", "チョコ", "取り巻き");
            _cIsz1E = MkCmb(80, "通常", "深腰", "浅腰");
            _cIsz1Pm = MkCmb(50, "＋", "−");
            _cIsz1LJu = MkDigit(); _cIsz1LJi = MkDigit(); _cIsz1LJs = MkDigit();
            _cIsz1SJu = MkDigit(); _cIsz1SJi = MkDigit(); _cIsz1SJs = MkDigit();
            _cIsz1F = MkDigit();
            return BuildIszGroup("石座１",
                _cIsz1A, _cIsz1B, _cIsz1D, _cIsz1E, _cIsz1Pm,
                _cIsz1LJu, _cIsz1LJi, _cIsz1LJs,
                _cIsz1SJu, _cIsz1SJi, _cIsz1SJs, _cIsz1F);
        }

        private GroupBox BuildIsz2Group()
        {
            _cIsz2A = MkCmb(90, "Pt900", "K18", "WG/PG");
            _cIsz2B = MkCmb(120, "円", "楕円", "角", "ドロップ", "0.1ct未満円", "0.1ct未満ﾌｧﾝｼｰ");
            _cIsz2D = MkCmb(100, "通常爪", "覆輪", "チョコ", "取り巻き");
            _cIsz2E = MkCmb(80, "通常", "深腰", "浅腰");
            _cIsz2Pm = MkCmb(50, "＋", "−");
            _cIsz2LJu = MkDigit(); _cIsz2LJi = MkDigit(); _cIsz2LJs = MkDigit();
            _cIsz2SJu = MkDigit(); _cIsz2SJi = MkDigit(); _cIsz2SJs = MkDigit();
            _cIsz2F = MkDigit();
            return BuildIszGroup("石座２",
                _cIsz2A, _cIsz2B, _cIsz2D, _cIsz2E, _cIsz2Pm,
                _cIsz2LJu, _cIsz2LJi, _cIsz2LJs,
                _cIsz2SJu, _cIsz2SJi, _cIsz2SJs, _cIsz2F);
        }

        private GroupBox BuildIsz3Group()
        {
            _cIsz3A = MkCmb(90, "Pt900", "K18", "WG/PG");
            _cIsz3B = MkCmb(120, "円", "楕円", "角", "ドロップ", "0.1ct未満円", "0.1ct未満ﾌｧﾝｼｰ");
            _cIsz3D = MkCmb(100, "通常爪", "覆輪", "チョコ", "取り巻き");
            _cIsz3E = MkCmb(80, "通常", "深腰", "浅腰");
            _cIsz3Pm = MkCmb(50, "＋", "−");
            _cIsz3LJu = MkDigit(); _cIsz3LJi = MkDigit(); _cIsz3LJs = MkDigit();
            _cIsz3SJu = MkDigit(); _cIsz3SJi = MkDigit(); _cIsz3SJs = MkDigit();
            _cIsz3F = MkDigit();
            return BuildIszGroup("石座３",
                _cIsz3A, _cIsz3B, _cIsz3D, _cIsz3E, _cIsz3Pm,
                _cIsz3LJu, _cIsz3LJi, _cIsz3LJs,
                _cIsz3SJu, _cIsz3SJi, _cIsz3SJs, _cIsz3F);
        }

        private GroupBox BuildIszGroup(string title,
            ComboBox cA, ComboBox cB, ComboBox cD, ComboBox cE, ComboBox cPm,
            ComboBox lJu, ComboBox lJi, ComboBox lJs,
            ComboBox sJu, ComboBox sJi, ComboBox sJs,
            ComboBox cF)
        {
            var grp = new GroupBox { Text = title, AutoSize = true, Width = 900, Padding = new Padding(6) };
            var flp = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                Dock = DockStyle.Fill
            };
            flp.Controls.Add(MkFlpRow(
                MkLbl("地金:", 45), cA,
                MkLbl("  石の形状:", 80), cB));
            flp.Controls.Add(MkFlpRow(
                MkLbl("石長径:", 60), lJu, lJi, MkLbl(".", 8), lJs, MkLbl(" mm  ", 35),
                MkLbl("石短径:", 60), sJu, sJi, MkLbl(".", 8), sJs, MkLbl(" mm", 30)));
            flp.Controls.Add(MkFlpRow(
                MkLbl("石座種類:", 70), cD,
                MkLbl("  腰高:", 50), cE,
                MkLbl("  ＋/－:", 50), cPm,
                MkLbl("  個数:", 50), cF));
            grp.Controls.Add(flp);
            return grp;
        }

        // ─────────────────────────────────────────────────────────────
        // Tab 3: 石留め 1〜4
        // ─────────────────────────────────────────────────────────────
        private TabPage BuildIsmdmeTab()
        {
            var tp = new TabPage("石留め");
            var pnl = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(8) };

            var grp1 = BuildIsm1Group(); grp1.Location = new Point(8, 8);
            var grp2 = BuildIsm2Group(); grp2.Location = new Point(8, grp1.Bottom + 8);
            var grp3 = BuildIsm3Group(); grp3.Location = new Point(8, grp2.Bottom + 8);
            var grp4 = BuildIsm4Group(); grp4.Location = new Point(8, grp3.Bottom + 8);

            pnl.Controls.AddRange(new Control[] { grp1, grp2, grp3, grp4 });
            tp.Controls.Add(pnl);
            return tp;
        }

        private static readonly string[] IsmSizes = {
            "0.05Ct以下：縦横合計5ミリ以下",  "0.1Ct以下：縦横合計6ミリ以下",
            "0.2Ct以下：縦横合計7ミリ以下",   "0.3Ct以下：縦横合計9ミリ以下",
            "0.5Ct以下：縦横合計10ミリ以下",  "1Ct以下：縦横合計13ミリ以下",
            "2Ct以下：縦横合計16ミリ以下",    "3Ct以下：縦横合計19ミリ以下",
            "4Ct以下：縦横合計21ミリ以下",    "5Ct以上：縦横合計22ミリ以上"
        };

        private GroupBox BuildIsm1Group()
        {
            _cIsm1D = MkCmb(130, "芯爪建留(本)", "留め/外し(石)", "接着留/外(石)");
            _cIsm1A = MkCmb(90, "Pt900", "K18", "WG/PG");
            _cIsm1B = MkCmb(130, "円楕円珠", "ﾌｧﾝｼｰ", "円楕円珠寄留", "ﾌｧﾝｼｰ寄留");
            _cIsm1C = MkCmb(280, IsmSizes);
            _cIsm1EJu = MkDigit(); _cIsm1EJi = MkDigit();
            _cIsm1D.SelectedIndexChanged += (s, e) =>
                _cIsm1A.Enabled = (_cIsm1D.Text == "芯爪建留(本)");
            return BuildIsmGroup("石留め１", _cIsm1D, _cIsm1A, _cIsm1B, _cIsm1C, _cIsm1EJu, _cIsm1EJi);
        }

        private GroupBox BuildIsm2Group()
        {
            _cIsm2D = MkCmb(130, "芯爪建留(本)", "留め/外し(石)", "接着留/外(石)");
            _cIsm2A = MkCmb(90, "Pt900", "K18", "WG/PG");
            _cIsm2B = MkCmb(130, "円楕円珠", "ﾌｧﾝｼｰ", "円楕円珠寄留", "ﾌｧﾝｼｰ寄留");
            _cIsm2C = MkCmb(280, IsmSizes);
            _cIsm2EJu = MkDigit(); _cIsm2EJi = MkDigit();
            _cIsm2D.SelectedIndexChanged += (s, e) =>
                _cIsm2A.Enabled = (_cIsm2D.Text == "芯爪建留(本)");
            return BuildIsmGroup("石留め２", _cIsm2D, _cIsm2A, _cIsm2B, _cIsm2C, _cIsm2EJu, _cIsm2EJi);
        }

        private GroupBox BuildIsm3Group()
        {
            _cIsm3D = MkCmb(130, "芯爪建留(本)", "留め/外し(石)", "接着留/外(石)");
            _cIsm3A = MkCmb(90, "Pt900", "K18", "WG/PG");
            _cIsm3B = MkCmb(130, "円楕円珠", "ﾌｧﾝｼｰ", "円楕円珠寄留", "ﾌｧﾝｼｰ寄留");
            _cIsm3C = MkCmb(280, IsmSizes);
            _cIsm3EJu = MkDigit(); _cIsm3EJi = MkDigit();
            _cIsm3D.SelectedIndexChanged += (s, e) =>
                _cIsm3A.Enabled = (_cIsm3D.Text == "芯爪建留(本)");
            return BuildIsmGroup("石留め３", _cIsm3D, _cIsm3A, _cIsm3B, _cIsm3C, _cIsm3EJu, _cIsm3EJi);
        }

        private GroupBox BuildIsm4Group()
        {
            _cIsm4D = MkCmb(130, "芯爪建留(本)", "留め/外し(石)", "接着留/外(石)");
            _cIsm4A = MkCmb(90, "Pt900", "K18", "WG/PG");
            _cIsm4B = MkCmb(130, "円楕円珠", "ﾌｧﾝｼｰ", "円楕円珠寄留", "ﾌｧﾝｼｰ寄留");
            _cIsm4C = MkCmb(280, IsmSizes);
            _cIsm4EJu = MkDigit(); _cIsm4EJi = MkDigit();
            _cIsm4D.SelectedIndexChanged += (s, e) =>
                _cIsm4A.Enabled = (_cIsm4D.Text == "芯爪建留(本)");
            return BuildIsmGroup("石留め４", _cIsm4D, _cIsm4A, _cIsm4B, _cIsm4C, _cIsm4EJu, _cIsm4EJi);
        }

        private GroupBox BuildIsmGroup(string title,
            ComboBox cD, ComboBox cA, ComboBox cB, ComboBox cC,
            ComboBox eJu, ComboBox eJi)
        {
            var grp = new GroupBox { Text = title, AutoSize = true, Width = 900, Padding = new Padding(6) };
            var flp = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                Dock = DockStyle.Fill
            };
            flp.Controls.Add(MkFlpRow(
                MkLbl("留め方法:", 75), cD,
                MkLbl("  地金:", 50), cA,
                MkLbl("  ※地金は「芯爪建留」のみ有効", 220)));
            flp.Controls.Add(MkFlpRow(
                MkLbl("石の形状:", 75), cB,
                MkLbl("  石サイズ:", 75), cC));
            flp.Controls.Add(MkFlpRow(
                MkLbl("個数:", 45), eJu, eJi));
            grp.Controls.Add(flp);
            return grp;
        }

        // ─────────────────────────────────────────────────────────────
        // Tab 4: ダイヤ・ロー付け・加工グレード
        // ─────────────────────────────────────────────────────────────
        private TabPage BuildDiaRoTab()
        {
            var tp = new TabPage("ダイヤ・ロー付け・加工グレード");
            var pnl = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(8) };

            var diaSize = new string[] {
                "約0.01", "約0.02", "約0.03", "約0.05",
                "約0.1",  "約0.2",  "約0.3",  "約0.5"
            };

            // ダイヤ1〜4
            _cDia1A = MkCmb(60, "B", "A", "S"); _cDia1B = MkCmb(80, diaSize);
            _cDia1Kigo = MkCmb(50, "＋", "−"); _cDia1Ju = MkDigit(); _cDia1Ji = MkDigit();

            _cDia2A = MkCmb(60, "B", "A", "S"); _cDia2B = MkCmb(80, diaSize);
            _cDia2Kigo = MkCmb(50, "＋", "−"); _cDia2Ju = MkDigit(); _cDia2Ji = MkDigit();

            _cDia3A = MkCmb(60, "B", "A", "S"); _cDia3B = MkCmb(80, diaSize);
            _cDia3Kigo = MkCmb(50, "＋", "−"); _cDia3Ju = MkDigit(); _cDia3Ji = MkDigit();

            _cDia4A = MkCmb(60, "B", "A", "S"); _cDia4B = MkCmb(80, diaSize);
            _cDia4Kigo = MkCmb(50, "＋", "−"); _cDia4Ju = MkDigit(); _cDia4Ji = MkDigit();

            var grpDia1 = BuildDiaGroup("ダイヤ１（メレー）", _cDia1A, _cDia1B, _cDia1Kigo, _cDia1Ju, _cDia1Ji);
            var grpDia2 = BuildDiaGroup("ダイヤ２（メレー）", _cDia2A, _cDia2B, _cDia2Kigo, _cDia2Ju, _cDia2Ji);
            var grpDia3 = BuildDiaGroup("ダイヤ３（メレー）", _cDia3A, _cDia3B, _cDia3Kigo, _cDia3Ju, _cDia3Ji);
            var grpDia4 = BuildDiaGroup("ダイヤ４（メレー）", _cDia4A, _cDia4B, _cDia4Kigo, _cDia4Ju, _cDia4Ji);

            grpDia1.Location = new Point(8, 8);
            grpDia2.Location = new Point(8, grpDia1.Bottom + 8);
            grpDia3.Location = new Point(8, grpDia2.Bottom + 8);
            grpDia4.Location = new Point(8, grpDia3.Bottom + 8);

            // ロー付け（面ロー）
            _cRoA1 = MkCmb(160, "面ロー", "レーザー面ロー");
            _cRoB1Ju = MkDigit(); _cRoB1Ji = MkDigit();
            var grpRoM = new GroupBox { Text = "ロー付け（面ロー）", AutoSize = true, Width = 900, Padding = new Padding(6) };
            var flpRoM = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, Dock = DockStyle.Fill };
            flpRoM.Controls.Add(MkFlpRow(MkLbl("種類:", 45), _cRoA1, MkLbl("  個数:", 55), _cRoB1Ju, _cRoB1Ji));
            grpRoM.Controls.Add(flpRoM);
            grpRoM.Location = new Point(8, grpDia4.Bottom + 8);

            // ロー付け（点ロー）
            _cRoC1 = MkCmb(160, "点ロー", "レーザー点ロー");
            _cRoD1Ju = MkDigit(); _cRoD1Ji = MkDigit();
            var grpRoT = new GroupBox { Text = "ロー付け（点ロー）", AutoSize = true, Width = 900, Padding = new Padding(6) };
            var flpRoT = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, Dock = DockStyle.Fill };
            flpRoT.Controls.Add(MkFlpRow(MkLbl("種類:", 45), _cRoC1, MkLbl("  個数:", 55), _cRoD1Ju, _cRoD1Ji));
            grpRoT.Controls.Add(flpRoT);
            grpRoT.Location = new Point(8, grpRoM.Bottom + 8);

            // 加工グレード
            _cKakouA1 = MkCmb(220,
                "部品合計",
                "96・95：部品集合＆整形",
                "94・93：部分手作り",
                "92：手作り(難易度1)",
                "91：手作り(難易度2)",
                "90：手作り(難易度3)",
                "89：手作り(特注)");
            _cKakouB1 = MkCmb(80, "SS", "S", "AS", "A", "AB", "B");

            var grpKakou = new GroupBox { Text = "加工難易度・加工グレード", AutoSize = true, Width = 900, Padding = new Padding(6) };
            var flpKakou = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, Dock = DockStyle.Fill };
            flpKakou.Controls.Add(MkFlpRow(
                MkLbl("加工難易度:", 85), _cKakouA1,
                MkLbl("  加工グレード:", 110), _cKakouB1));
            grpKakou.Controls.Add(flpKakou);
            grpKakou.Location = new Point(8, grpRoT.Bottom + 8);

            pnl.Controls.AddRange(new Control[] {
                grpDia1, grpDia2, grpDia3, grpDia4, grpRoM, grpRoT, grpKakou
            });
            tp.Controls.Add(pnl);
            return tp;
        }

        private GroupBox BuildDiaGroup(string title,
            ComboBox cA, ComboBox cB, ComboBox cKigo, ComboBox cJu, ComboBox cJi)
        {
            var grp = new GroupBox { Text = title, AutoSize = true, Width = 900, Padding = new Padding(6) };
            var flp = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                Dock = DockStyle.Fill
            };
            flp.Controls.Add(MkFlpRow(
                MkLbl("グレード:", 70), cA,
                MkLbl("  サイズ:", 60), cB, MkLbl(" ct", 25),
                MkLbl("  個数:", 55), cKigo, cJu, cJi));
            grp.Controls.Add(flp);
            return grp;
        }

        // ============================================================
        // コントロールファクトリーヘルパー
        // ============================================================

        // 0〜9 の DropDownList ComboBox（VB6: wk_int = 0〜9 の AddItem ループ）
        private static ComboBox MkDigit()
        {
            var c = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 42,
                Margin = new Padding(1, 2, 1, 2)
            };
            for (int i = 0; i <= 9; i++) c.Items.Add(i.ToString());
            c.SelectedIndex = 0;
            return c;
        }

        // リングサイズ十の位（0/1/2 のみ）
        private static ComboBox MkRsSzJ()
        {
            var c = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 42,
                Margin = new Padding(1, 2, 1, 2)
            };
            c.Items.AddRange(new object[] { "0", "1", "2" });
            c.SelectedIndex = 0;
            return c;
        }

        // 指定リストの DropDownList ComboBox
        private static ComboBox MkCmb(int width, params string[] items)
        {
            var c = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = width,
                Margin = new Padding(1, 2, 1, 2)
            };
            c.Items.AddRange(items);
            if (c.Items.Count > 0) c.SelectedIndex = 0;
            return c;
        }

        // ラベル（AutoSize、指定幅があればその幅に固定）
        private static Label MkLbl(string text, int width = 0)
        {
            var lbl = new Label
            {
                Text = text,
                AutoSize = (width == 0),
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(1, 4, 1, 2)
            };
            if (width > 0) { lbl.AutoSize = false; lbl.Width = width; }
            return lbl;
        }

        // 参照データ用ラベル（水色背景）
        private static Label MkSmpLbl(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = false,
                Width = 210,
                Height = 26,
                BackColor = Color.LightCyan,
                TextAlign = ContentAlignment.MiddleLeft,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(2)
            };
        }

        // 計算結果用ラベル（白背景）
        private static Label MkResLbl(string text, int width)
        {
            return new Label
            {
                Text = text,
                AutoSize = false,
                Width = width,
                Height = 26,
                BackColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("MS Pゴシック", 11f, FontStyle.Bold),
                Margin = new Padding(2)
            };
        }

        // ボタン
        private static Button MkBtn(string text, int width, Color backColor)
        {
            return new Button
            {
                Text = text,
                Width = width,
                Height = 60,
                BackColor = backColor,
                Font = new Font("MS Pゴシック", 12f, FontStyle.Bold),
                Margin = new Padding(3)
            };
        }

        // 横並びコントロール行（FlowLayoutPanel）
        private static FlowLayoutPanel MkFlpRow(params Control[] ctrls)
        {
            var flp = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = true,
                Margin = new Padding(0, 2, 0, 2)
            };
            flp.Controls.AddRange(ctrls);
            return flp;
        }

        // ============================================================
        // フォームロード
        // VB6: Form_Load → Init_Control → form_tedukuri.WindowState = 2
        // ============================================================
        private void FormTedukuri_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        // ============================================================
        // フォームクローズ
        // VB6: システムメニューの閉じるボタンを無効化していた動作を再現
        // ============================================================
        private void FormTedukuri_FormClosing(object sender, FormClosingEventArgs e)
        {
            // btn_menu / btn_back 経由は _forceClose = true が立っているので許可
            if (e.CloseReason == CloseReason.UserClosing && !_forceClose)
                e.Cancel = true;
        }

        // ============================================================
        // SetSampleData
        // VB6: form_disp_picture の btn_order_Click から呼び出される
        // 参照データラベルに値をセットし、黄色でハイライト
        // ============================================================
        /// <summary>
        /// 画像表示画面から商品データを受け取り、参照データラベルに表示する。
        /// </summary>
        /// <param name="hinban">品番</param>
        /// <param name="pt900">Pt900重量(g)</param>
        /// <param name="k18">K18重量(g)</param>
        /// <param name="wgpg">WG/PG重量(g)</param>
        /// <param name="k10">K10重量(g)（本フォームでは参考表示のみ）</param>
        /// <param name="code">加工コード</param>
        /// <param name="price">税込価格</param>
        public void SetSampleData(
            string hinban, double pt900, double k18,
            double wgpg, double k10, double code, double price)
        {
            // 参照データを各ラベルに反映
            _lblSmpHinban.Text = "品番: " + hinban;
            _lblSmpPt900.Text = "Pt900: " + pt900.ToString("0.00") + " g";
            _lblSmpK18.Text = "K18: " + k18.ToString("0.00") + " g";
            _lblSmpWgpg.Text = "WG/PG: " + wgpg.ToString("0.00") + " g";
            _lblSmpCode.Text = "Code: " + ((int)code).ToString();
            _lblSmpPrice.Text = "金額: " + price.ToString("#,##0") + " 円";

            // 値が入っているラベルを黄色でハイライト
            // （VB6: lbl_sample_xxx.BackColor = &HFFFFC0 → 印刷判定に使用）
            _lblSmpHinban.BackColor = (!string.IsNullOrEmpty(hinban)) ? CCalc : Color.LightCyan;
            _lblSmpPt900.BackColor = (pt900 != 0) ? CCalc : Color.LightCyan;
            _lblSmpK18.BackColor = (k18 != 0) ? CCalc : Color.LightCyan;
            _lblSmpWgpg.BackColor = (wgpg != 0) ? CCalc : Color.LightCyan;
            _lblSmpCode.BackColor = (code != 0) ? CCalc : Color.LightCyan;
            _lblSmpPrice.BackColor = (price != 0) ? CCalc : Color.LightCyan;

            // AppState にも保存（btn_save で参照）
            AppState.OrderSampleHinban = hinban;
            AppState.OrderSamplePt900 = pt900;
            AppState.OrderSampleK18 = k18;
            AppState.OrderSampleWgPg = wgpg;
            AppState.OrderSampleK10 = k10;
            AppState.OrderSampleCode = (int)code;
            AppState.OrderSamplePrice = price;
        }

        // ============================================================
        // ResetSampleLabels
        // VB6: FormMenu の btn_tedukuri_Click 直前に参照ラベルをリセット
        //      "form_tedukuri の lbl_sample_xxx.Caption = "" 相当"
        // ============================================================
        /// <summary>
        /// 参照データラベルを初期状態に戻す。
        /// </summary>
        public void ResetSampleLabels()
        {
            _lblSmpHinban.Text = "品番: ─────";
            _lblSmpPt900.Text = "Pt900: ─ g";
            _lblSmpK18.Text = "K18: ─ g";
            _lblSmpWgpg.Text = "WG/PG: ─ g";
            _lblSmpCode.Text = "Code: ─";
            _lblSmpPrice.Text = "金額: ─ 円";

            _lblSmpHinban.BackColor = Color.LightCyan;
            _lblSmpPt900.BackColor = Color.LightCyan;
            _lblSmpK18.BackColor = Color.LightCyan;
            _lblSmpWgpg.BackColor = Color.LightCyan;
            _lblSmpCode.BackColor = Color.LightCyan;
            _lblSmpPrice.BackColor = Color.LightCyan;

            AppState.OrderSampleHinban = "";
            AppState.OrderSamplePt900 = 0;
            AppState.OrderSampleK18 = 0;
            AppState.OrderSampleWgPg = 0;
            AppState.OrderSampleCode = 0;
            AppState.OrderSamplePrice = 0;
        }
    }
}
