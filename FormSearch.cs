using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace YumejitateApp
{
    /// <summary>
    /// AB商品検索フォーム (VB6: form_search.frm の移植)
    /// 商品特性（アイテム・地金・サイズ等）を条件に検索テーブルを絞り込み、
    /// ランダムに最大7件を検索結果テーブルへ格納したのち form_movie へ遷移する。
    /// </summary>
    public class FormSearch : Form
    {
        // ----------------------------------------------------------------
        // コントロール宣言
        // ----------------------------------------------------------------

        // 検索条件コンボボックス
        private ComboBox _cmbItem;         // アイテム         (VB6: アイテム)
        private ComboBox _cmbShapeView;    // 横から見た場合の形 (VB6: 横から見た場合の形)
        private ComboBox _cmbJigane;       // 地金             (VB6: 地金)
        private ComboBox _cmbDiaNum;       // ダイヤインの数    (VB6: ダイヤインの数)
        private ComboBox _cmbDiaStyle;     // ダイヤインのスタイル (VB6: ダイヤインのスタイル)
        private ComboBox _cmbDesignQty;    // デザイン数量      (VB6: デザイン数量)
        private ComboBox _cmbSetStyle;     // 石留セットスタイル  (VB6: 石留セットスタイル)

        // 重視した見込予算・方法
        private ComboBox _cmbYosanHyaku;   // 予算 百万の位     (VB6: 予算四桁)
        private ComboBox _cmbYosanJu;      // 予算 十万の位     (VB6: 予算十桁)
        private ComboBox _cmbYosanIchi;    // 予算 一万の位     (VB6: 予算一万)
        private ComboBox _cmbYosanType;    // 重視した見込予算  (VB6: 重視した見込予算)
        private ComboBox _cmbOrderMethod;  // 重視した方法      (VB6: 重視した方法)

        // サイズ入力コンボ（上サイズ × 下サイズ）
        private ComboBox _cmbSizeAJu;      // 上サイズ 十の位   (VB6: サイズ十\)
        private ComboBox _cmbSizeAIchi;    // 上サイズ 一の位   (VB6: サイズ一)
        private ComboBox _cmbSizeAKo;      // 上サイズ 小数点以下 (VB6: サイズ上小)
        private ComboBox _cmbSizeBJu;      // 下サイズ 十の位   (VB6: サイズ下\)
        private ComboBox _cmbSizeBIchi;    // 下サイズ 一の位   (VB6: サイズ下一)
        // VB6 の size_b 計算では サイズ下一 を小数点以下にも使用している（コピーミスと思われる）。
        // 本実装では 下サイズ小数点以下 専用コンボを用意するが、SQL では VB6 に忠実に _cmbSizeBIchi を使用する。
        private ComboBox _cmbSizeBKo;      // 下サイズ 小数点以下 (VB6: サイズ下一 ※同名で参照)
        private ComboBox _cmbSizeType;     // サイズ種別         (VB6: サイズ)

        // リングサイズ・グレード
        private ComboBox _cmbRingsizeJu;   // リングサイズ 十の位 (VB6: cmb_ringsize十)
        private ComboBox _cmbRingsizeIchi; // リングサイズ 一の位 (VB6: cmb_ringsize一)
        private ComboBox _cmbGrade;        // グレード           (VB6: cmb_grade)

        // ボタン
        private Button _btnSearch;         // 検索 (VB6: btn_search)
        private Button _btnBack;           // メニューへ戻る (VB6: btn_back)

        // ラベル（ナビゲーション用ラベルは CenterControls で配置）
        private Label _lblTitle;
        private Label _lblItem;
        private Label _lblShapeView;
        private Label _lblJigane;
        private Label _lblDiaNum;
        private Label _lblDiaStyle;
        private Label _lblDesignQty;
        private Label _lblSetStyle;
        private Label _lblYosanBudget;     // "予算" 見出し
        private Label _lblYosanMan;        // "万円" 単位ラベル
        private Label _lblYosanType;       // "重視した見込予算" ラベル
        private Label _lblOrderMethod;
        private Label _lblSize;
        private Label _lblSizeADot;        // "." 上サイズ区切り
        private Label _lblSizeAMmX;        // "mm×" 区切り
        private Label _lblSizeBDot;        // "." 下サイズ区切り
        private Label _lblSizeBMm;         // "mm" 末尾
        private Label _lblRingSizeGrade;
        private Label _lblRingsizeNote;    // "十の位" 補助
        private Label _lblGrade;
        private Label _lblRingsizeHint;    // "1,2,3,4,5,9 はリングのみ有効"

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------
        public FormSearch()
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
            // VB6: BackColor=&H00D8FFFF& (水色), Caption="夢仕立て-商品検索入力画面"
            this.Text = "夢仕立て - 商品検索入力画面";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;

            var labelFont = new Font("ＭＳ Ｐゴシック", 14.25f, FontStyle.Bold);
            var comboFont = new Font("ＭＳ Ｐゴシック", 14.25f, FontStyle.Bold);
            var smallFont = new Font("ＭＳ Ｐゴシック", 12f, FontStyle.Bold);
            var titleFont = new Font("ＭＳ Ｐゴシック", 18f, FontStyle.Italic);

            // ---- タイトル (VB6: Label2 "商品検索") ----
            _lblTitle = new Label
            {
                Text = "商品検索",
                Font = titleFont,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(20, 10),
            };

            // ---- ラベル作成ヘルパ ----
            Label MkLabel(string txt, int w = 230)
            {
                return new Label
                {
                    Text = txt,
                    Font = labelFont,
                    BackColor = Color.Transparent,
                    AutoSize = false,
                    Size = new Size(w, 26),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Location = new Point(0, 0), // CenterControls で再配置
                };
            }

            // ---- 広コンボ作成ヘルパ ----
            ComboBox MkWide(int w = 280)
            {
                return new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Font = comboFont,
                    Size = new Size(w, 27),
                    Location = new Point(0, 0),
                };
            }

            // ---- 1桁コンボ作成ヘルパ (0-9) ----
            ComboBox MkDigit()
            {
                var c = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Font = comboFont,
                    Size = new Size(50, 27),
                    Location = new Point(0, 0),
                };
                for (int i = 0; i <= 9; i++) c.Items.Add(i.ToString());
                return c;
            }

            // ---- 区切りラベル ----
            Label MkSep(string txt)
            {
                return new Label
                {
                    Text = txt,
                    Font = labelFont,
                    BackColor = Color.Transparent,
                    AutoSize = true,
                    Location = new Point(0, 0),
                };
            }

            // ---- 各コントロールの生成 ----
            _lblItem = MkLabel("アイテム");
            _cmbItem = MkWide(300);
            _cmbItem.SelectedIndexChanged += CmbItem_SelectedIndexChanged;

            _lblShapeView = MkLabel("横から見た場合の形", 250);
            _cmbShapeView = MkWide(300);

            _lblJigane = MkLabel("地金");
            _cmbJigane = MkWide(300);

            _lblDiaNum = MkLabel("ダイヤインの数");
            _cmbDiaNum = MkWide(300);

            _lblDiaStyle = MkLabel("ダイヤインのスタイル", 250);
            _cmbDiaStyle = MkWide(300);

            _lblDesignQty = MkLabel("デザイン数量");
            _cmbDesignQty = MkWide(300);

            _lblSetStyle = MkLabel("石留セットスタイル", 250);
            _cmbSetStyle = MkWide(300);

            // 予算
            _lblYosanBudget = MkLabel("重視した見込予算", 250);
            _cmbYosanHyaku = MkDigit();
            _cmbYosanJu = MkDigit();
            _cmbYosanIchi = MkDigit();
            _lblYosanMan = MkSep("万円");
            _cmbYosanType = MkWide(280);
            _lblYosanType = MkLabel("");  // スペーサ―（不使用）

            _lblOrderMethod = MkLabel("重視した方法");
            _cmbOrderMethod = MkWide(300);

            // サイズ
            _lblSize = MkLabel("サイズ");
            _cmbSizeAJu = MkDigit();
            _cmbSizeAIchi = MkDigit();
            _lblSizeADot = MkSep(".");
            _cmbSizeAKo = MkDigit();
            _lblSizeAMmX = MkSep("mm×");
            _cmbSizeBJu = MkDigit();
            _cmbSizeBIchi = MkDigit();
            _lblSizeBDot = MkSep(".");
            _cmbSizeBKo = MkDigit();
            _lblSizeBMm = MkSep("mm");
            _cmbSizeType = MkWide(280);

            // リングサイズ / グレード
            _lblRingSizeGrade = MkLabel("リングサイズまたは石のグレード", 310);
            _cmbRingsizeJu = MkDigit();
            _cmbRingsizeIchi = MkDigit();
            _lblGrade = MkLabel("グレード", 80);
            _cmbGrade = MkWide(100);
            _lblRingsizeNote = new Label
            {
                Text = "十の位",
                Font = smallFont,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(0, 0),
            };
            _lblRingsizeHint = new Label
            {
                // VB6: Label13 "1,2,3,4,5,9 はリングのみ有効"
                Text = "1,2,3,4,5,9 はリングのみ有効",
                Font = smallFont,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(0, 0),
            };

            // ---- ボタン ----
            _btnSearch = new Button
            {
                Text = "検索",
                Font = new Font("ＭＳ Ｐゴシック", 15.75f, FontStyle.Bold),
                Size = new Size(140, 60),
                BackColor = Color.FromArgb(0xC0, 0xC0, 0xFF),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnSearch.FlatAppearance.BorderColor = Color.SteelBlue;
            _btnSearch.FlatAppearance.BorderSize = 2;
            _btnSearch.Click += BtnSearch_Click;

            _btnBack = new Button
            {
                Text = "メニュー",
                Font = new Font("ＭＳ Ｐゴシック", 15.75f, FontStyle.Bold),
                Size = new Size(140, 60),
                BackColor = Color.FromArgb(0xFF, 0xC0, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnBack.FlatAppearance.BorderColor = Color.Crimson;
            _btnBack.FlatAppearance.BorderSize = 2;
            _btnBack.Click += BtnBack_Click;

            // ---- コントロールをフォームへ追加 ----
            this.Controls.AddRange(new Control[]
            {
                _lblTitle,
                _lblItem,      _cmbItem,
                _lblShapeView, _cmbShapeView,
                _lblJigane,    _cmbJigane,
                _lblDiaNum,    _cmbDiaNum,
                _lblDiaStyle,  _cmbDiaStyle,
                _lblDesignQty, _cmbDesignQty,
                _lblSetStyle,  _cmbSetStyle,
                _lblYosanBudget,
                _cmbYosanHyaku, _cmbYosanJu, _cmbYosanIchi,
                _lblYosanMan,  _cmbYosanType,
                _lblOrderMethod, _cmbOrderMethod,
                _lblSize,
                _cmbSizeAJu, _cmbSizeAIchi, _lblSizeADot, _cmbSizeAKo,
                _lblSizeAMmX,
                _cmbSizeBJu, _cmbSizeBIchi, _lblSizeBDot, _cmbSizeBKo,
                _lblSizeBMm, _cmbSizeType,
                _lblRingSizeGrade,
                _cmbRingsizeJu, _cmbRingsizeIchi,
                _lblRingsizeNote, _lblRingsizeHint,
                _lblGrade, _cmbGrade,
                _btnSearch, _btnBack,
            });

            this.ResumeLayout(false);
        }

        // ----------------------------------------------------------------
        // フォームロード
        // ----------------------------------------------------------------
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.WindowState = FormWindowState.Maximized;

            // VB6: Form_Load → Init_Control → コンボボックスにリストを追加
            InitControlItems();

            CenterControls();
        }

        // ----------------------------------------------------------------
        // フォームリサイズ時に再配置
        // ----------------------------------------------------------------
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            CenterControls();
        }

        // ----------------------------------------------------------------
        // コントロール初期化 (VB6: Init_Control)
        // 各コンボボックスにリストを追加し、デフォルト値を設定する。
        // ----------------------------------------------------------------
        private void InitControlItems()
        {
            // ---- アイテム (VB6: アイテム) ----
            // b フィールドで検索; T→1, P→2, E→3 にマッピングされる
            _cmbItem.Items.AddRange(new object[]
            {
                "6:リング",
                "7:ペンダント",
                "8:ブローチ",
                "9:その他（リングのみ）",
                "T:タイタック",
                "P:ピアス",
                "E:イアリング",
                "0:その他（バチカン）",
            });
            _cmbItem.SelectedIndex = 3; // "9:その他（リングのみ）"

            // ---- 横から見た場合の形 (VB6: 横から見た場合の形) ----
            // c フィールドで検索
            _cmbShapeView.Items.AddRange(new object[]
            {
                "1:ラウンド（丸）",
                "2:オーバル（楕円）",
                "3:ボール（球）",
                "4:エメラルド（四角）",
                "5:マーキース",
                "6:ドロップ",
                "9:その他（バチカン、球）",
            });
            _cmbShapeView.SelectedIndex = 6; // "9:その他..."

            // ---- 地金 (VB6: 地金) ----
            // d フィールドで検索; "3","6" → '*2*', "9" → 複合条件
            _cmbJigane.Items.AddRange(new object[]
            {
                "1:プラチナ",
                "2:18金デザイン",
                "3:18金プレーン",
                "4:コンビ",
                "5:シルバー",
                "6:18金0",
                "9:その他（銀）",
            });
            _cmbJigane.SelectedIndex = 5; // "6:18金0"

            // ---- ダイヤインの数 (VB6: ダイヤインの数) ----
            // e フィールドで検索
            _cmbDiaNum.Items.AddRange(new object[]
            {
                "1:シルバー",
                "2:テーパー（楕形）",
                "3:ファンシーカット（カクア、ペア他）",
                "4:ミックス",
                "5:その他",
                "9:その他（銀）",
            });
            _cmbDiaNum.SelectedIndex = 5; // "9:その他..."

            // ---- ダイヤインのスタイル (VB6: ダイヤインのスタイル) ----
            // f フィールドで検索
            _cmbDiaStyle.Items.AddRange(new object[]
            {
                "1:石留あり",
                "2:サイド（脇石のみ）のみ",
                "3:爪なしのみ",
                "4:その他",
                "9:その他（銀）",
            });
            _cmbDiaStyle.SelectedIndex = 4; // "9:その他..."

            // ---- デザイン数量 (VB6: デザイン数量) ----
            // g フィールドで検索
            _cmbDesignQty.Items.AddRange(new object[]
            {
                "1:超定番",
                "2:一文字定番",
                "3:デンタル（色違い）",
                "4:ファッション大量店向",
                "5:デザイナーズ（独占扱い）",
                "9:その他（銀）",
                "6:シンプル",
                "7:その他",
            });
            _cmbDesignQty.SelectedIndex = 5; // "9:その他..."

            // ---- 石留セットスタイル (VB6: 石留セットスタイル) ----
            // h フィールドで検索; code 2 → 爪なし（AB2 補正の切り替えに使用）
            _cmbSetStyle.Items.AddRange(new object[]
            {
                "1:爪留め",
                "2:爪無し（ビール留め等）",
                "9:その他（銀）",
            });
            _cmbSetStyle.SelectedIndex = 2; // "9:その他..."

            // ---- 重視した見込予算 桁コンボ (VB6: 予算四桁/予算十桁/予算一万) ----
            // yosan = (百万*100 + 十万*10 + 一万) * 10000 yen
            _cmbYosanHyaku.SelectedIndex = 0; // "0"
            _cmbYosanJu.SelectedIndex = 0;
            _cmbYosanIchi.SelectedIndex = 0;

            // ---- 重視した見込予算タイプ (VB6: 重視した見込予算) ----
            // price フィールドで範囲検索; "9" は予算指定なし
            _cmbYosanType.Items.AddRange(new object[]
            {
                "1:第1重視　予算70%UP",
                "2:第2重視　予算50%UP",
                "3:第3重視　予算30%UP",
                "9:その他（指定なし）",
            });
            _cmbYosanType.SelectedIndex = 3; // "9:..."

            // ---- 重視した方法 (VB6: 重視した方法) ----
            // j フィールドで検索
            _cmbOrderMethod.Items.AddRange(new object[]
            {
                "1:イージーオーダー",
                "2:カスタムオーダー",
                "9:その他（銀）",
            });
            _cmbOrderMethod.SelectedIndex = 2; // "9:..."

            // ---- サイズ桁コンボ (0-9 は MkDigit() で追加済み) ----
            _cmbSizeAJu.SelectedIndex = 0;
            _cmbSizeAIchi.SelectedIndex = 0;
            _cmbSizeAKo.SelectedIndex = 0;
            _cmbSizeBJu.SelectedIndex = 0;
            _cmbSizeBIchi.SelectedIndex = 0;
            _cmbSizeBKo.SelectedIndex = 0;

            // ---- サイズ種別 (VB6: サイズ) ----
            // k フィールドで検索
            _cmbSizeType.Items.AddRange(new object[]
            {
                "1:ファセットカット石の場合はボール",
                "2:カボッションセットは球",
            });
            _cmbSizeType.SelectedIndex = 0; // "1:..."

            // ---- リングサイズ ----
            _cmbRingsizeJu.SelectedIndex = 0; // "0"
            _cmbRingsizeIchi.SelectedIndex = 0; // "0"

            // ---- グレード (VB6: cmb_grade) ----
            _cmbGrade.Items.AddRange(new object[] { "A", "B", "C" });
            _cmbGrade.SelectedIndex = 0; // "A"
            _cmbGrade.Enabled = false; // アイテムが "9" or "6" のときは無効

            // アイテムの初期値("9")に合わせてリングサイズ/グレードの有効状態を設定
            UpdateItemDependentControls();
        }

        // ----------------------------------------------------------------
        // コントロール中央配置 (最大化・リサイズ対応)
        // VB6 では固定座標だったが C# では画面解像度に合わせて中央配置する。
        // ----------------------------------------------------------------
        private void CenterControls()
        {
            if (this.ClientSize.Width == 0 || this.ClientSize.Height == 0) return;

            // 左列(ラベル)と右列(コンボ)の起点
            int cx = this.ClientSize.Width / 2;
            int lx = cx - 580;   // ラベル左端
            int rx = cx - 330;   // コンボ右列左端
            int ry = 55;         // 先頭行 Y
            int dy = 46;         // 行間

            // タイトル
            _lblTitle.Location = new Point(cx - 80, 10);

            // ---- 各行の Y 座標 ----
            int yItem = ry;
            int yShape = ry + dy * 1;
            int yJigane = ry + dy * 2;
            int yDiaNum = ry + dy * 3;
            int yDiaStyle = ry + dy * 4;
            int yDesignQty = ry + dy * 5;
            int ySetStyle = ry + dy * 6;
            int yYosan = ry + dy * 7;
            int yOrder = ry + dy * 8;
            int ySize = ry + dy * 9;
            int yRing = ry + dy * 10;
            int yBtn = ry + dy * 11 + 15;

            // ---- アイテム ----
            _lblItem.SetBounds(lx, yItem + 2, 230, 26);
            _cmbItem.SetBounds(rx, yItem, 310, 27);

            // ---- 横から見た場合の形 ----
            _lblShapeView.SetBounds(lx, yShape + 2, 250, 26);
            _cmbShapeView.SetBounds(rx, yShape, 310, 27);

            // ---- 地金 ----
            _lblJigane.SetBounds(lx, yJigane + 2, 230, 26);
            _cmbJigane.SetBounds(rx, yJigane, 310, 27);

            // ---- ダイヤインの数 ----
            _lblDiaNum.SetBounds(lx, yDiaNum + 2, 230, 26);
            _cmbDiaNum.SetBounds(rx, yDiaNum, 310, 27);

            // ---- ダイヤインのスタイル ----
            _lblDiaStyle.SetBounds(lx, yDiaStyle + 2, 250, 26);
            _cmbDiaStyle.SetBounds(rx, yDiaStyle, 310, 27);

            // ---- デザイン数量 ----
            _lblDesignQty.SetBounds(lx, yDesignQty + 2, 230, 26);
            _cmbDesignQty.SetBounds(rx, yDesignQty, 310, 27);

            // ---- 石留セットスタイル ----
            _lblSetStyle.SetBounds(lx, ySetStyle + 2, 250, 26);
            _cmbSetStyle.SetBounds(rx, ySetStyle, 310, 27);

            // ---- 重視した見込予算 ----
            // [ラベル][百万桁][十万桁][一万桁][万円][重視タイプコンボ]
            _lblYosanBudget.SetBounds(lx, yYosan + 2, 250, 26);
            _cmbYosanHyaku.SetBounds(rx, yYosan, 50, 27);
            _cmbYosanJu.SetBounds(rx + 55, yYosan, 50, 27);
            _cmbYosanIchi.SetBounds(rx + 110, yYosan, 50, 27);
            _lblYosanMan.SetBounds(rx + 165, yYosan + 2, 50, 26);
            _cmbYosanType.SetBounds(rx + 215, yYosan, 290, 27);

            // ---- 重視した方法 ----
            _lblOrderMethod.SetBounds(lx, yOrder + 2, 230, 26);
            _cmbOrderMethod.SetBounds(rx, yOrder, 310, 27);

            // ---- サイズ ----
            // [ラベル][上十][上一][.][上小][mm×][下十][下一][.][下小][mm][種別]
            _lblSize.SetBounds(lx, ySize + 2, 230, 26);
            _cmbSizeAJu.SetBounds(rx, ySize, 50, 27);
            _cmbSizeAIchi.SetBounds(rx + 55, ySize, 50, 27);
            _lblSizeADot.SetBounds(rx + 107, ySize + 5, 10, 20);
            _cmbSizeAKo.SetBounds(rx + 115, ySize, 50, 27);
            _lblSizeAMmX.SetBounds(rx + 168, ySize + 2, 55, 26);
            _cmbSizeBJu.SetBounds(rx + 225, ySize, 50, 27);
            _cmbSizeBIchi.SetBounds(rx + 280, ySize, 50, 27);
            _lblSizeBDot.SetBounds(rx + 332, ySize + 5, 10, 20);
            _cmbSizeBKo.SetBounds(rx + 340, ySize, 50, 27);
            _lblSizeBMm.SetBounds(rx + 393, ySize + 2, 40, 26);
            _cmbSizeType.SetBounds(rx + 435, ySize, 290, 27);

            // ---- リングサイズ / グレード ----
            _lblRingSizeGrade.SetBounds(lx, yRing + 2, 310, 26);
            _cmbRingsizeJu.SetBounds(rx, yRing, 50, 27);
            _cmbRingsizeIchi.SetBounds(rx + 55, yRing, 50, 27);
            _lblRingsizeNote.SetBounds(rx + 110, yRing + 5, 50, 20);
            _lblGrade.SetBounds(rx + 170, yRing + 2, 80, 26);
            _cmbGrade.SetBounds(rx + 255, yRing, 100, 27);
            _lblRingsizeHint.SetBounds(rx + 370, yRing + 5, 300, 20);

            // ---- ボタン ----
            _btnSearch.SetBounds(cx - 160, yBtn, 140, 60);
            _btnBack.SetBounds(cx + 20, yBtn, 140, 60);
        }

        // ----------------------------------------------------------------
        // アイテム変更時: 有効/無効を更新 (VB6: アイテム_Click)
        // ----------------------------------------------------------------
        private void CmbItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateItemDependentControls();
        }

        /// <summary>
        /// アイテムの先頭文字に応じてコントロールの有効/無効を切り替える。
        /// VB6: アイテム_Click
        /// </summary>
        private void UpdateItemDependentControls()
        {
            if (_cmbItem.SelectedIndex < 0) return;
            string itemCode = _cmbItem.Text.Substring(0, 1);

            // ---- リングサイズ / グレード ----
            // アイテムが "9" または "6" のときリングサイズを有効にする (VB6 と同様)
            bool isRingLike = (itemCode == "9" || itemCode == "6");
            _cmbRingsizeJu.Enabled = isRingLike;
            _cmbRingsizeIchi.Enabled = isRingLike;
            _cmbGrade.Enabled = !isRingLike;

            // ---- バチカン（0）の場合: 関連コントロールを無効化 ----
            bool isVatican = (itemCode == "0");
            _cmbShapeView.Enabled = !isVatican;
            _cmbDiaStyle.Enabled = !isVatican;
            _cmbDesignQty.Enabled = !isVatican;
            _cmbSetStyle.Enabled = !isVatican;
            _cmbSizeAJu.Enabled = !isVatican;
            _cmbSizeAIchi.Enabled = !isVatican;
            _cmbSizeAKo.Enabled = !isVatican;
            _cmbSizeBJu.Enabled = !isVatican;
            _cmbSizeBIchi.Enabled = !isVatican;
            _cmbSizeBKo.Enabled = !isVatican;
        }

        // ----------------------------------------------------------------
        // 検索ボタン (VB6: btn_search_Click)
        // ----------------------------------------------------------------
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            // ---- 入力チェック ----
            if (!CheckControl()) return;

            // 検索開始メッセージ (VB6: MsgBox "検索中...")
            MessageBox.Show(
                "検索中以上のデータから御要望のデザインを探します。",
                "検索開始",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            // ---- ワークテーブルクリア ----
            // VB6: dao_database.Execute "delete * from ワーク検索テーブル"
            //      dao_database.Execute "delete * from 検索結果テーブル"
            AppState.Db.ExecuteNonQuery("DELETE * FROM [ワーク検索テーブル]");
            AppState.Db.ExecuteNonQuery("DELETE * FROM [検索結果テーブル]");

            // ---- SQL 構築 ----
            string strSql = BuildSearchSql(out string strSqlKotei,
                                                out string chkSqlD,
                                                out string chkSqlE,
                                                out string chkSqlF,
                                                out string chkSqlG,
                                                out string chkSqlI,
                                                out string chkSqlJ,
                                                out string chkSqlM);

            // ---- 検索実行 ----
            DataTable dt = AppState.Db.ExecuteQuery(strSql);

            if (dt.Rows.Count == 0)
            {
                // 結果なし: 原因を段階的に診断する (VB6 の診断ロジックを再現)
                DiagnoseNoResult(strSqlKotei, chkSqlD, chkSqlE, chkSqlF,
                                 chkSqlG, chkSqlI, chkSqlJ, chkSqlM);
                return;
            }

            // ---- ワークテーブルへ INSERT ----
            // VB6: "insert into ワーク検索テーブル " & strsql
            AppState.Db.ExecuteNonQuery("INSERT INTO [ワーク検索テーブル] " + strSql);

            // ---- 品番(a フィールド)の重複除去 ----
            RemoveDuplicateHinban();

            // ---- ランダムに最大7件を検索結果テーブルへ格納 ----
            int resultCnt = SelectRandomResults();

            // 検索結果件数を表示
            MessageBox.Show(
                $"御要望のデザインが {resultCnt} 種類見つかりました。",
                "検索結果件数表示",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            // ---- form_movie へ遷移 ----
            // VB6: form_search.Visible = False → form_movie.Show
            this.Hide();
            var movie = new FormMovie();
            movie.FormClosed += (s2, e2) => this.Close();
            movie.Show();
        }

        // ----------------------------------------------------------------
        // 検索 SQL の構築 (VB6: btn_search_Click 内の SQL 構築部分)
        // ----------------------------------------------------------------
        private string BuildSearchSql(
            out string strSqlKotei,
            out string chkSqlD, out string chkSqlE, out string chkSqlF,
            out string chkSqlG, out string chkSqlI, out string chkSqlJ,
            out string chkSqlM)
        {
            string itemCode = _cmbItem.Text.Substring(0, 1);

            // ---- アイテムコード変換 (VB6: T→1, P→2, E→3) ----
            string wkStr;
            if (itemCode == "T") wkStr = "1";
            else if (itemCode == "P") wkStr = "2";
            else if (itemCode == "E") wkStr = "3";
            else wkStr = itemCode;

            // ---- 基本条件 ----
            string sql = "SELECT * FROM [検索テーブル] WHERE b LIKE '*" + wkStr + "*' ";

            // バチカン除外条件 (VB6: "and l = 'XXX' and m = 'XXX'")
            if (itemCode == "0")
                sql += "AND l = 'XXX' AND m = 'XXX' ";

            // 横から見た場合の形 (c フィールド)
            if (itemCode != "0")
                sql += "AND c LIKE '*" + _cmbShapeView.Text.Substring(0, 1) + "*' ";

            // 地金 (d フィールド) ※ "3", "6" → '*2*'; "9" → 複合条件
            string jiganeCode = _cmbJigane.Text.Substring(0, 1);
            if (jiganeCode == "3" || jiganeCode == "6")
                sql += "AND d LIKE '*2*' ";
            else if (jiganeCode == "9")
                sql += "AND (d LIKE '*9*' AND d NOT LIKE '*3*' AND d NOT LIKE '*5*' AND d NOT LIKE '*6*') ";
            else
                sql += "AND d LIKE '*" + jiganeCode + "*' ";

            // ここまでが固定条件 (VB6: strSqlKotei の最初の値)
            strSqlKotei = sql;

            // ---- ダイヤインの数 (e フィールド) ----
            string wkSql = "AND e LIKE '*" + _cmbDiaNum.Text.Substring(0, 1) + "*' ";
            sql += wkSql;
            chkSqlD = wkSql;

            // ---- ダイヤインのスタイル (f フィールド) ----
            wkSql = "";
            if (itemCode != "0")
            {
                wkSql = "AND f LIKE '*" + _cmbDiaStyle.Text.Substring(0, 1) + "*' ";
                sql += wkSql;
            }
            chkSqlE = wkSql;

            // ---- デザイン数量 (g フィールド) ----
            wkSql = "";
            if (itemCode != "0")
            {
                wkSql = "AND g LIKE '*" + _cmbDesignQty.Text.Substring(0, 1) + "*' ";
                sql += wkSql;
            }
            chkSqlF = wkSql;

            // ---- 石留セットスタイル (h フィールド) ----
            wkSql = "";
            if (itemCode != "0")
            {
                wkSql = "AND h LIKE '*" + _cmbSetStyle.Text.Substring(0, 1) + "*' ";
                sql += wkSql;
            }
            chkSqlG = wkSql;

            // ---- 重視した見込予算 (price フィールド) ----
            long yosan = (long.Parse(_cmbYosanHyaku.Text) * 100
                        + long.Parse(_cmbYosanJu.Text) * 10
                        + long.Parse(_cmbYosanIchi.Text)) * 10000L;
            wkSql = "";
            string priceCode = _cmbYosanType.Text.Substring(0, 1);
            if (priceCode == "1") wkSql = $"AND price <= {D(yosan * 1.7)} AND price >= {yosan} ";
            else if (priceCode == "2") wkSql = $"AND price <= {D(yosan * 1.5)} AND price >= {yosan} ";
            else if (priceCode == "3") wkSql = $"AND price <= {D(yosan * 1.3)} AND price >= {yosan} ";
            // "9" (その他): 予算制限なし → wkSql 空
            sql += wkSql;
            chkSqlI = wkSql;

            // ---- 重視した方法 (j フィールド) ----
            wkSql = "AND j LIKE '*" + _cmbOrderMethod.Text.Substring(0, 1) + "*' ";
            sql += wkSql;
            chkSqlJ = wkSql;

            // ---- サイズ種別 (k フィールド) ----
            wkSql = "AND k LIKE '*" + _cmbSizeType.Text.Substring(0, 1) + "*' ";
            sql += wkSql;
            strSqlKotei += wkSql;

            // ---- サイズ入力値 (bigsize/smallsize/totalsize フィールド) ----
            if (itemCode != "0")
            {
                // VB6: size_a = CDbl(サイズ十\) * 100 + CDbl(サイズ一) * 10 + CDbl(サイズ上小)
                //      size_b = CDbl(サイズ下\) * 100 + CDbl(サイズ下一) * 10 + CDbl(サイズ下一)
                //      ※ size_b の小数点以下は VB6 では サイズ下一 を重複参照（コピーミス）
                double sizeA = double.Parse(_cmbSizeAJu.Text) * 100
                             + double.Parse(_cmbSizeAIchi.Text) * 10
                             + double.Parse(_cmbSizeAKo.Text);
                // VB6 バグ再現: 下サイズの小数点以下も _cmbSizeBIchi を参照
                double sizeB = double.Parse(_cmbSizeBJu.Text) * 100
                             + double.Parse(_cmbSizeBIchi.Text) * 10
                             + double.Parse(_cmbSizeBIchi.Text); // VB6 バグ: サイズ下一 × 2

                double bSize = sizeA > sizeB ? sizeA : sizeB;
                double sSize = sizeA > sizeB ? sizeB : sizeA;
                double tSize = sizeA + sizeB;

                string shapeCode = _cmbShapeView.Text.Substring(0, 1);
                string setStyleCode = _cmbSetStyle.Text.Substring(0, 1);

                wkSql = "";
                if (shapeCode != "3")
                {
                    // ボール以外: 石留スタイルで式を切り替える
                    if (setStyleCode == "2")
                    {
                        // 式2 (爪無し): AB2 補正あり
                        wkSql += "AND (( NOT (a LIKE 'AB2*') ";
                        wkSql += $"AND bigsize >= {D(bSize / 1.18)} ";
                        wkSql += $"AND bigsize <= {D(bSize / 0.75)} ";
                        wkSql += $"AND smallsize >= {D(sSize / 1.18)} ";
                        wkSql += $"AND smallsize <= {D(sSize / 0.75)} ";
                        wkSql += $"AND totalsize >= {D(tSize)} ";
                        wkSql += $"AND totalsize <= {D(tSize + 10)}) ";
                        wkSql += "OR ( a LIKE 'AB2*' ";
                        wkSql += $"AND bigsize >= {D(bSize / 1.062)} ";
                        wkSql += $"AND bigsize <= {D(bSize / 0.675)} ";
                        wkSql += $"AND smallsize >= {D(sSize / 1.062)} ";
                        wkSql += $"AND smallsize <= {D(sSize / 0.675)} ";
                        wkSql += $"AND totalsize >= {D(tSize / 0.9)} ";
                        wkSql += $"AND totalsize <= {D((tSize + 10) / 0.9)})) ";
                    }
                    else
                    {
                        // 式1 (爪留め): AB2 補正あり
                        wkSql += "AND (( NOT (a LIKE 'AB2*') ";
                        wkSql += $"AND bigsize >= {D(bSize / 1.18)} ";
                        wkSql += $"AND bigsize <= {D(bSize / 0.75)} ";
                        wkSql += $"AND smallsize >= {D(sSize / 1.18)} ";
                        wkSql += $"AND smallsize <= {D(sSize / 0.75)} ";
                        wkSql += $"AND totalsize >= {D(tSize - 5)} ";
                        wkSql += $"AND totalsize <= {D(tSize + 10)}) ";
                        wkSql += "OR ( a LIKE 'AB2*' ";
                        wkSql += $"AND bigsize >= {D(bSize / 1.062)} ";
                        wkSql += $"AND bigsize <= {D(bSize / 0.675)} ";
                        wkSql += $"AND smallsize >= {D(sSize / 1.062)} ";
                        wkSql += $"AND smallsize <= {D(sSize / 0.675)} ";
                        wkSql += $"AND totalsize >= {D((tSize - 5) / 0.9)} ";
                        wkSql += $"AND totalsize <= {D((tSize + 10) / 0.9)})) ";
                    }
                }
                else
                {
                    // ボール（球）: AB2 補正なし、totalsize 指定なし
                    wkSql += $"AND bigsize >= {D(bSize / 1.18)} ";
                    wkSql += $"AND bigsize <= {D(bSize / 0.75)} ";
                    wkSql += $"AND smallsize >= {D(sSize / 1.18)} ";
                    wkSql += $"AND smallsize <= {D(sSize / 0.75)} ";
                    wkSql += $"AND totalsize >= {D(tSize - 5)} ";
                    wkSql += $"AND totalsize <= {D(tSize + 10)} ";
                }

                sql += wkSql;
                strSqlKotei += wkSql;
            }

            // ---- リングサイズ条件 (アイテム "9" or "6" のみ) ----
            // VB6: chk_sql_m = strsql (リングサイズ追加前の SQL を保存)
            chkSqlM = "";
            if (itemCode == "9" || itemCode == "6")
            {
                chkSqlM = sql; // リングサイズ追加前のスナップショット
                int mSize = int.Parse(_cmbRingsizeJu.Text) * 10
                          + int.Parse(_cmbRingsizeIchi.Text);
                sql += $"AND ({mSize} BETWEEN val(w) AND val(x)) ";
            }

            return sql;
        }

        // ----------------------------------------------------------------
        // 結果なしの場合の原因診断 (VB6: recTB.EOF の分岐内の診断処理)
        // どの条件を加えたときに結果がゼロになるかを段階的に確認し、
        // ユーザーに "その条件を変えてください" と案内する。
        // ----------------------------------------------------------------
        private void DiagnoseNoResult(
            string strSqlKotei,
            string chkSqlD, string chkSqlE, string chkSqlF,
            string chkSqlG, string chkSqlI, string chkSqlJ,
            string chkSqlM)
        {
            string itemCode = _cmbItem.Text.Substring(0, 1);

            // ① 固定条件（アイテム・地金・サイズ種別・サイズ）だけで結果なし → 根本的な条件不一致
            if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
            {
                MessageBox.Show(
                    "㈠「固定条件」に該当する設計データがありません。",
                    "「固定条件」該当データなし",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ② ダイヤインの数
            strSqlKotei += chkSqlD;
            if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
            {
                MessageBox.Show(
                    "ⓑ「ダイヤインの数」を変えてください。",
                    "「ダイヤインの数」該当データなし",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ③ ダイヤインのスタイル (バチカン以外)
            if (itemCode != "0")
            {
                strSqlKotei += chkSqlD + chkSqlE;
                if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
                {
                    MessageBox.Show(
                        "ⓒ「ダイヤインのスタイル」を変えてください。",
                        "「ダイヤインのスタイル」該当データなし",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // ④ デザイン数量 (バチカン以外)
            if (itemCode != "0")
            {
                strSqlKotei += chkSqlD + chkSqlE + chkSqlF;
                if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
                {
                    MessageBox.Show(
                        "ⓓ「デザイン数量」を変えてください。",
                        "「デザイン数量」該当データなし",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // ⑤ 石留セットスタイル (バチカン以外)
            if (itemCode != "0")
            {
                strSqlKotei += chkSqlD + chkSqlE + chkSqlF + chkSqlG;
                if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
                {
                    MessageBox.Show(
                        "ⓔ「石留セットスタイル」を変えてください。",
                        "「石留セットスタイル」該当データなし",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // ⑥ 重視した見込予算
            if (itemCode != "0")
                strSqlKotei += chkSqlD + chkSqlE + chkSqlF + chkSqlG + chkSqlI;
            else
                strSqlKotei += chkSqlD + chkSqlI;
            if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
            {
                MessageBox.Show(
                    "ⓕ「重視した見込予算」を変えてください。",
                    "「重視した見込予算」該当データなし",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ⑦ 重視した方法
            if (itemCode != "0")
                strSqlKotei += chkSqlD + chkSqlE + chkSqlF + chkSqlG + chkSqlI + chkSqlJ;
            else
                strSqlKotei += chkSqlD + chkSqlI + chkSqlJ;
            // VB6: さらにもう一度全条件を付加（元コードの冗長記述を再現）
            strSqlKotei += chkSqlD + chkSqlE + chkSqlF + chkSqlG + chkSqlI + chkSqlJ;
            if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
            {
                MessageBox.Show(
                    "ⓖ「重視した方法」を変えてください。",
                    "「重視した方法」該当データなし",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ⑧ リングサイズ (アイテム "9" or "6" のみ)
            if (itemCode == "9" || itemCode == "6")
            {
                // VB6: chk_sql_m = strsql はリングサイズ追加前のスナップショット
                // "not eof" のときメッセージ → リングサイズ変更を促す
                if (AppState.Db.ExecuteQuery(chkSqlM).Rows.Count > 0)
                {
                    MessageBox.Show(
                        "ⓗ「リングサイズ」を変えてください。",
                        "「リングサイズ」該当データなし",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        // ----------------------------------------------------------------
        // 品番重複除去 (VB6: ワーク検索テーブルの重複削除処理)
        // index 順に並べ、直前の品番(a フィールド)と同じ行を削除する。
        // ----------------------------------------------------------------
        private void RemoveDuplicateHinban()
        {
            DataTable dt = AppState.Db.ExecuteQuery(
                "SELECT * FROM [ワーク検索テーブル] ORDER BY [index]");

            string prevHinban = "";
            foreach (DataRow row in dt.Rows)
            {
                string cur = row["a"].ToString();
                if (string.Compare(cur, prevHinban, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    AppState.Db.ExecuteNonQuery(
                        $"DELETE FROM [ワーク検索テーブル] WHERE [index] = {row["index"]}");
                }
                else
                {
                    prevHinban = cur;
                }
            }
        }

        // ----------------------------------------------------------------
        // ランダムに最大7件を検索結果テーブルへ格納 (VB6: 乱数選択ループ)
        // VB6 の「ランダム位置を決めて1件ずつ取り出す」アルゴリズムを
        // Fisher-Yates シャッフルで再現する。
        // ----------------------------------------------------------------
        private int SelectRandomResults()
        {
            // 乱数テーブルのカウントを消費（VB6 の RNG シード消費をシミュレート）
            try
            {
                DataTable seedDt = AppState.Db.ExecuteQuery("SELECT * FROM [乱数テーブル]");
                if (seedDt.Rows.Count > 0)
                {
                    int oldCount = Convert.ToInt32(seedDt.Rows[0]["count"]);
                    var rngSeed = new Random();
                    int newVal = 0;
                    for (int i = 0; i < oldCount + 1; i++)
                        newVal = rngSeed.Next(10);
                    AppState.Db.ExecuteNonQuery(
                        $"UPDATE [乱数テーブル] SET [count] = {newVal}");
                }
            }
            catch { /* 乱数テーブルが存在しない環境では無視 */ }

            // ワーク検索テーブルの全 index を取得
            DataTable wkDt = AppState.Db.ExecuteQuery(
                "SELECT [index] FROM [ワーク検索テーブル] ORDER BY [index]");

            var indices = new List<int>();
            foreach (DataRow row in wkDt.Rows)
                indices.Add(Convert.ToInt32(row["index"]));

            // Fisher-Yates シャッフルで VB6 の乱数選択を再現
            var rnd = new Random();
            for (int i = indices.Count - 1; i > 0; i--)
            {
                int j = rnd.Next(i + 1);
                int tmp = indices[i];
                indices[i] = indices[j];
                indices[j] = tmp;
            }

            // 先頭から最大7件を検索結果テーブルへ INSERT
            int count = Math.Min(7, indices.Count);
            for (int i = 0; i < count; i++)
            {
                AppState.Db.ExecuteNonQuery(
                    "INSERT INTO [検索結果テーブル] " +
                    "SELECT * FROM [ワーク検索テーブル] " +
                    $"WHERE [index] = {indices[i]}");
            }

            return count;
        }

        // ----------------------------------------------------------------
        // 入力チェック (VB6: check_control)
        // ----------------------------------------------------------------
        private bool CheckControl()
        {
            string itemCode = _cmbItem.Text.Substring(0, 1);

            // ---- サイズ入力チェック (バチカン以外) ----
            if (itemCode != "0")
            {
                bool sizeAZero = (_cmbSizeAJu.Text == "0" &&
                                  _cmbSizeAIchi.Text == "0" &&
                                  _cmbSizeAKo.Text == "0");
                bool sizeBZero = (_cmbSizeBJu.Text == "0" &&
                                  _cmbSizeBIchi.Text == "0" &&
                                  _cmbSizeBKo.Text == "0");
                if (sizeAZero || sizeBZero)
                {
                    MessageBox.Show(
                        "ⓙ「サイズ」を入力してください。",
                        "入力チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            // ---- 予算チェック ----
            string yosanCode = _cmbYosanType.Text.Substring(0, 1);
            bool yosanIsSpecified = (yosanCode == "1" || yosanCode == "2" || yosanCode == "3");
            bool yosanAmountZero = (_cmbYosanHyaku.Text == "0" &&
                                     _cmbYosanJu.Text == "0" &&
                                     _cmbYosanIchi.Text == "0");
            if (yosanIsSpecified && yosanAmountZero)
            {
                MessageBox.Show(
                    "ⓗ「重視した見込予算」に「第1」「第2」「第3」を選択した場合は、予算を入力してください。",
                    "入力チェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (yosanCode == "9" && !yosanAmountZero)
            {
                MessageBox.Show(
                    "ⓗ「重視した見込予算」が「その他」の場合、予算はゼロにしてください。",
                    "入力チェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // ---- リングサイズチェック (アイテム "9" or "6") ----
            if (itemCode == "9" || itemCode == "6")
            {
                if (_cmbRingsizeJu.Text == "0" && _cmbRingsizeIchi.Text == "0")
                {
                    MessageBox.Show(
                        "ⓙ「アイテム」が「その他（リング）」でリングで検索する場合は、商品サイズを入力してください。",
                        "入力チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            return true;
        }

        // ----------------------------------------------------------------
        // メニューへ戻るボタン (VB6: btn_back_Click)
        // ----------------------------------------------------------------
        private void BtnBack_Click(object sender, EventArgs e)
        {
            // VB6: Flag_Hinban = False; Flag_Mihon = False;
            //      form_search.Visible = False → form_menu.Visible = True
            AppState.FlagHinban = false;
            AppState.FlagMihon = false;

            foreach (Form f in Application.OpenForms)
            {
                if (f is FormMenu)
                {
                    f.Show();
                    break;
                }
            }
            this.Hide();
        }

        // ----------------------------------------------------------------
        // 小数点ロケール安全な数値文字列化
        // Access SQL に渡す際に小数点を "." にする (VB6: CDbl → CStr の英語版挙動を再現)
        // ----------------------------------------------------------------
        private static string D(double v) => v.ToString("G", CultureInfo.InvariantCulture);
    }
}
