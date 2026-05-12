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
        private ComboBox _cmbItem;         // アイテム
        private ComboBox _cmbShapeView;    // 横から見た場合の形
        private ComboBox _cmbJigane;       // 地金
        private ComboBox _cmbDiaNum;       // ダイヤインの数
        private ComboBox _cmbDiaStyle;     // ダイヤインのスタイル
        private ComboBox _cmbDesignQty;    // デザイン数量
        private ComboBox _cmbSetStyle;     // 石留セットスタイル

        // 重視した見込予算・方法
        private ComboBox _cmbYosanHyaku;   // 予算 百万の位
        private ComboBox _cmbYosanJu;      // 予算 十万の位
        private ComboBox _cmbYosanIchi;    // 予算 一万の位
        private ComboBox _cmbYosanType;    // 重視した見込予算
        private ComboBox _cmbOrderMethod;  // 重視した方法

        // サイズ入力コンボ
        private ComboBox _cmbSizeAJu;      // 上サイズ 十の位
        private ComboBox _cmbSizeAIchi;    // 上サイズ 一の位
        private ComboBox _cmbSizeAKo;      // 上サイズ 小数点以下
        private ComboBox _cmbSizeBJu;      // 下サイズ 十の位
        private ComboBox _cmbSizeBIchi;    // 下サイズ 一の位
        private ComboBox _cmbSizeBKo;      // 下サイズ 小数点以下
        private ComboBox _cmbSizeType;     // サイズ種別

        // リングサイズ・グレード
        private ComboBox _cmbRingsizeJu;   // リングサイズ 十の位
        private ComboBox _cmbRingsizeIchi; // リングサイズ 一の位
        private ComboBox _cmbGrade;        // グレード

        // ボタン
        private Button _btnSearch;
        private Button _btnBack;

        // ラベル
        private Label _lblTitle;
        private Label _lblItem;
        private Label _lblShapeView;
        private Label _lblJigane;
        private Label _lblDiaNum;
        private Label _lblDiaStyle;
        private Label _lblDesignQty;
        private Label _lblSetStyle;
        private Label _lblYosanBudget;
        private Label _lblYosanMan;
        private Label _lblOrderMethod;
        private Label _lblSize;
        private Label _lblSizeADot;
        private Label _lblSizeAMmX;
        private Label _lblSizeBDot;
        private Label _lblSizeBMm;
        private Label _lblRingSizeGrade;
        private Label _lblRingsizeNote;
        private Label _lblGrade;
        private Label _lblRingsizeHint;

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

            this.Text = "夢仕立て - 商品検索入力画面";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;

            var labelFont = new Font("ＭＳ Ｐゴシック", 14.25f, FontStyle.Bold);
            var comboFont = new Font("ＭＳ Ｐゴシック", 14.25f, FontStyle.Bold);
            var smallFont = new Font("ＭＳ Ｐゴシック", 12f, FontStyle.Bold);
            var titleFont = new Font("ＭＳ Ｐゴシック", 18f, FontStyle.Italic);

            // タイトル
            _lblTitle = new Label
            {
                Text = "商品検索",
                Font = titleFont,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(20, 10),
            };

            // ---- ヘルパ ----
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
                    Location = new Point(0, 0),
                };
            }

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

            // ---- コントロール生成 ----
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

            _lblYosanBudget = MkLabel("重視した見込予算", 250);
            _cmbYosanHyaku = MkDigit();
            _cmbYosanJu = MkDigit();
            _cmbYosanIchi = MkDigit();
            _lblYosanMan = MkSep("万円");
            _cmbYosanType = MkWide(280);

            _lblOrderMethod = MkLabel("重視した方法");
            _cmbOrderMethod = MkWide(300);

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
                Text = "1,2,3,4,5,9 はリングのみ有効",
                Font = smallFont,
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(0, 0),
            };

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
        // ----------------------------------------------------------------
        private void InitControlItems()
        {
            // アイテム
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
            _cmbItem.SelectedIndex = 3;

            // 横から見た場合の形
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
            _cmbShapeView.SelectedIndex = 6;

            // 地金
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
            _cmbJigane.SelectedIndex = 5;

            // ダイヤインの数
            _cmbDiaNum.Items.AddRange(new object[]
            {
                "1:シルバー",
                "2:テーパー（楕形）",
                "3:ファンシーカット（カクア、ペア他）",
                "4:ミックス",
                "5:その他",
                "9:その他（銀）",
            });
            _cmbDiaNum.SelectedIndex = 5;

            // ダイヤインのスタイル
            _cmbDiaStyle.Items.AddRange(new object[]
            {
                "1:石留あり",
                "2:サイド（脇石のみ）のみ",
                "3:爪なしのみ",
                "4:その他",
                "9:その他（銀）",
            });
            _cmbDiaStyle.SelectedIndex = 4;

            // デザイン数量
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
            _cmbDesignQty.SelectedIndex = 5;

            // 石留セットスタイル
            _cmbSetStyle.Items.AddRange(new object[]
            {
                "1:爪留め",
                "2:爪無し（ビール留め等）",
                "9:その他（銀）",
            });
            _cmbSetStyle.SelectedIndex = 2;

            // 予算桁
            _cmbYosanHyaku.SelectedIndex = 0;
            _cmbYosanJu.SelectedIndex = 0;
            _cmbYosanIchi.SelectedIndex = 0;

            // 重視した見込予算タイプ
            _cmbYosanType.Items.AddRange(new object[]
            {
                "1:第1重視　予算70%UP",
                "2:第2重視　予算50%UP",
                "3:第3重視　予算30%UP",
                "9:その他（指定なし）",
            });
            _cmbYosanType.SelectedIndex = 3;

            // 重視した方法
            _cmbOrderMethod.Items.AddRange(new object[]
            {
                "1:イージーオーダー",
                "2:カスタムオーダー",
                "9:その他（銀）",
            });
            _cmbOrderMethod.SelectedIndex = 2;

            // サイズ桁
            _cmbSizeAJu.SelectedIndex = 0;
            _cmbSizeAIchi.SelectedIndex = 0;
            _cmbSizeAKo.SelectedIndex = 0;
            _cmbSizeBJu.SelectedIndex = 0;
            _cmbSizeBIchi.SelectedIndex = 0;
            _cmbSizeBKo.SelectedIndex = 0;

            // サイズ種別
            _cmbSizeType.Items.AddRange(new object[]
            {
                "1:ファセットカット石の場合はボール",
                "2:カボッションセットは球",
            });
            _cmbSizeType.SelectedIndex = 0;

            // リングサイズ
            _cmbRingsizeJu.SelectedIndex = 0;
            _cmbRingsizeIchi.SelectedIndex = 0;

            // グレード
            _cmbGrade.Items.AddRange(new object[] { "A", "B", "C" });
            _cmbGrade.SelectedIndex = 0;
            _cmbGrade.Enabled = false;

            UpdateItemDependentControls();
        }

        // ----------------------------------------------------------------
        // コントロール中央配置
        // ----------------------------------------------------------------
        private void CenterControls()
        {
            if (this.ClientSize.Width == 0 || this.ClientSize.Height == 0) return;

            int cx = this.ClientSize.Width / 2;
            int lx = cx - 580;
            int rx = cx - 330;
            int ry = 55;
            int dy = 46;

            _lblTitle.Location = new Point(cx - 80, 10);

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

            _lblItem.SetBounds(lx, yItem + 2, 230, 26);
            _cmbItem.SetBounds(rx, yItem, 310, 27);

            _lblShapeView.SetBounds(lx, yShape + 2, 250, 26);
            _cmbShapeView.SetBounds(rx, yShape, 310, 27);

            _lblJigane.SetBounds(lx, yJigane + 2, 230, 26);
            _cmbJigane.SetBounds(rx, yJigane, 310, 27);

            _lblDiaNum.SetBounds(lx, yDiaNum + 2, 230, 26);
            _cmbDiaNum.SetBounds(rx, yDiaNum, 310, 27);

            _lblDiaStyle.SetBounds(lx, yDiaStyle + 2, 250, 26);
            _cmbDiaStyle.SetBounds(rx, yDiaStyle, 310, 27);

            _lblDesignQty.SetBounds(lx, yDesignQty + 2, 230, 26);
            _cmbDesignQty.SetBounds(rx, yDesignQty, 310, 27);

            _lblSetStyle.SetBounds(lx, ySetStyle + 2, 250, 26);
            _cmbSetStyle.SetBounds(rx, ySetStyle, 310, 27);

            _lblYosanBudget.SetBounds(lx, yYosan + 2, 250, 26);
            _cmbYosanHyaku.SetBounds(rx, yYosan, 50, 27);
            _cmbYosanJu.SetBounds(rx + 55, yYosan, 50, 27);
            _cmbYosanIchi.SetBounds(rx + 110, yYosan, 50, 27);
            _lblYosanMan.SetBounds(rx + 165, yYosan + 2, 50, 26);
            _cmbYosanType.SetBounds(rx + 215, yYosan, 290, 27);

            _lblOrderMethod.SetBounds(lx, yOrder + 2, 230, 26);
            _cmbOrderMethod.SetBounds(rx, yOrder, 310, 27);

            _lblSize.SetBounds(lx, ySize + 2, 230, 26);
            _cmbSizeAJu.SetBounds(rx, ySize, 50, 27);
            _cmbSizeAIchi.SetBounds(rx + 55, ySize, 50, 27);
            _lblSizeADot.SetBounds(rx + 107, ySize + 5, 15, 20);
            _cmbSizeAKo.SetBounds(rx + 125, ySize, 50, 27);   // +10
            _lblSizeAMmX.SetBounds(rx + 178, ySize + 2, 55, 26); // +10
            _cmbSizeBJu.SetBounds(rx + 240, ySize, 50, 27);   // +15
            _cmbSizeBIchi.SetBounds(rx + 295, ySize, 50, 27); // +15
            _lblSizeBDot.SetBounds(rx + 347, ySize + 5, 15, 20); // +15
            _cmbSizeBKo.SetBounds(rx + 365, ySize, 50, 27);   // +25
            _lblSizeBMm.SetBounds(rx + 418, ySize + 2, 40, 26);  // +25
            _cmbSizeType.SetBounds(rx + 465, ySize, 290, 27); // +30

            int rxRing = lx + 320; // ラベル幅310 + 余白10
            _lblRingSizeGrade.SetBounds(lx, yRing + 2, 310, 26);
            _cmbRingsizeJu.SetBounds(rxRing, yRing, 50, 27);
            _cmbRingsizeIchi.SetBounds(rxRing + 55, yRing, 50, 27);
            _lblRingsizeNote.SetBounds(rxRing + 110, yRing + 5, 60, 20);
            _lblGrade.SetBounds(rxRing + 175, yRing + 2, 80, 26);
            _cmbGrade.SetBounds(rxRing + 260, yRing, 100, 27);
            _lblRingsizeHint.SetBounds(rxRing + 375, yRing + 5, 300, 20);

            _btnSearch.SetBounds(cx - 160, yBtn, 140, 60);
            _btnBack.SetBounds(cx + 20, yBtn, 140, 60);
        }

        // ----------------------------------------------------------------
        // アイテム変更イベント (VB6: アイテム_Click)
        // ----------------------------------------------------------------
        private void CmbItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateItemDependentControls();
        }

        private void UpdateItemDependentControls()
        {
            if (_cmbItem.SelectedIndex < 0) return;
            string itemCode = _cmbItem.Text.Substring(0, 1);

            bool isRingLike = (itemCode == "9" || itemCode == "6");
            _cmbRingsizeJu.Enabled = isRingLike;
            _cmbRingsizeIchi.Enabled = isRingLike;
            _cmbGrade.Enabled = !isRingLike;

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
            if (!CheckControl()) return;

            MessageBox.Show(
                "検索中以上のデータから御要望のデザインを探します。",
                "検索開始",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            AppState.Db.ExecuteNonQuery("DELETE * FROM [ワーク検索テーブル]");
            AppState.Db.ExecuteNonQuery("DELETE * FROM [検索結果テーブル]");

            string strSql = BuildSearchSql(
                out string strSqlKotei,
                out string chkSqlD, out string chkSqlE, out string chkSqlF,
                out string chkSqlG, out string chkSqlI, out string chkSqlJ,
                out string chkSqlM);

            DataTable dt = AppState.Db.ExecuteQuery(strSql);

            if (dt.Rows.Count == 0)
            {
                DiagnoseNoResult(strSqlKotei, chkSqlD, chkSqlE, chkSqlF,
                                 chkSqlG, chkSqlI, chkSqlJ, chkSqlM);
                return;
            }

            AppState.Db.ExecuteNonQuery("INSERT INTO [ワーク検索テーブル] " + strSql);

            RemoveDuplicateHinban();

            int resultCnt = SelectRandomResults();

            MessageBox.Show(
                $"御要望のデザインが {resultCnt} 種類見つかりました。",
                "検索結果件数表示",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            this.Hide();
            var movie = new FormMovie2();
            movie.FormClosed += (s2, e2) => this.Show();
            movie.Show();
        }

        // ----------------------------------------------------------------
        // 検索SQL構築 (VB6: btn_search_Click 内の SQL 構築部分)
        // ※ LIKE条件は .NET OleDb に合わせて % ワイルドカードを使用
        // ----------------------------------------------------------------
        private string BuildSearchSql(
            out string strSqlKotei,
            out string chkSqlD, out string chkSqlE, out string chkSqlF,
            out string chkSqlG, out string chkSqlI, out string chkSqlJ,
            out string chkSqlM)
        {
            string itemCode = _cmbItem.Text.Substring(0, 1);

            // アイテムコード変換 (T→1, P→2, E→3)
            string wkStr;
            if (itemCode == "T") wkStr = "1";
            else if (itemCode == "P") wkStr = "2";
            else if (itemCode == "E") wkStr = "3";
            else wkStr = itemCode;

            // 基本条件
            string sql = "SELECT * FROM [検索テーブル] WHERE [b] LIKE '%" + wkStr + "%' ";

            if (itemCode == "0")
                sql += "AND [l] = 'XXX' AND [m] = 'XXX' ";

            // 横から見た場合の形 (c フィールド)
            if (itemCode != "0")
                sql += "AND [c] LIKE '%" + _cmbShapeView.Text.Substring(0, 1) + "%' ";

            // 地金 (d フィールド)
            string jiganeCode = _cmbJigane.Text.Substring(0, 1);
            if (jiganeCode == "3" || jiganeCode == "6")
                sql += "AND [d] LIKE '%2%' ";
            else if (jiganeCode == "9")
                sql += "AND ([d] LIKE '%9%' AND [d] NOT LIKE '%3%' AND [d] NOT LIKE '%5%' AND [d] NOT LIKE '%6%') ";
            else
                sql += "AND [d] LIKE '%" + jiganeCode + "%' ";

            strSqlKotei = sql;

            // ダイヤインの数 (e フィールド)
            string wkSql = "AND [e] LIKE '%" + _cmbDiaNum.Text.Substring(0, 1) + "%' ";
            sql += wkSql;
            chkSqlD = wkSql;

            // ダイヤインのスタイル (f フィールド)
            wkSql = "";
            if (itemCode != "0")
            {
                wkSql = "AND [f] LIKE '%" + _cmbDiaStyle.Text.Substring(0, 1) + "%' ";
                sql += wkSql;
            }
            chkSqlE = wkSql;

            // デザイン数量 (g フィールド)
            wkSql = "";
            if (itemCode != "0")
            {
                wkSql = "AND [g] LIKE '%" + _cmbDesignQty.Text.Substring(0, 1) + "%' ";
                sql += wkSql;
            }
            chkSqlF = wkSql;

            // 石留セットスタイル (h フィールド)
            wkSql = "";
            if (itemCode != "0")
            {
                wkSql = "AND [h] LIKE '%" + _cmbSetStyle.Text.Substring(0, 1) + "%' ";
                sql += wkSql;
            }
            chkSqlG = wkSql;

            // 重視した見込予算 (price フィールド)
            long yosan = (long.Parse(_cmbYosanHyaku.Text) * 100
                        + long.Parse(_cmbYosanJu.Text) * 10
                        + long.Parse(_cmbYosanIchi.Text)) * 10000L;
            wkSql = "";
            string priceCode = _cmbYosanType.Text.Substring(0, 1);
            if (priceCode == "1") wkSql = $"AND [price] <= {D(yosan * 1.7)} AND [price] >= {yosan} ";
            else if (priceCode == "2") wkSql = $"AND [price] <= {D(yosan * 1.5)} AND [price] >= {yosan} ";
            else if (priceCode == "3") wkSql = $"AND [price] <= {D(yosan * 1.3)} AND [price] >= {yosan} ";
            sql += wkSql;
            chkSqlI = wkSql;

            // 重視した方法 (j フィールド)
            wkSql = "AND [j] LIKE '%" + _cmbOrderMethod.Text.Substring(0, 1) + "%' ";
            sql += wkSql;
            chkSqlJ = wkSql;

            // サイズ種別 (k フィールド)
            wkSql = "AND [k] LIKE '%" + _cmbSizeType.Text.Substring(0, 1) + "%' ";
            sql += wkSql;
            strSqlKotei += wkSql;

            // サイズ条件 (bigsize/smallsize/totalsize)
            if (itemCode != "0")
            {
                double sizeA = double.Parse(_cmbSizeAJu.Text) * 100
                             + double.Parse(_cmbSizeAIchi.Text) * 10
                             + double.Parse(_cmbSizeAKo.Text);
                // VB6バグ再現: 下サイズ小数点以下も _cmbSizeBIchi を参照
                double sizeB = double.Parse(_cmbSizeBJu.Text) * 100
                             + double.Parse(_cmbSizeBIchi.Text) * 10
                             + double.Parse(_cmbSizeBIchi.Text);

                double bSize = sizeA > sizeB ? sizeA : sizeB;
                double sSize = sizeA > sizeB ? sizeB : sizeA;
                double tSize = sizeA + sizeB;

                string shapeCode = _cmbShapeView.Text.Substring(0, 1);
                string setStyleCode = _cmbSetStyle.Text.Substring(0, 1);

                wkSql = "";
                if (shapeCode != "3")
                {
                    if (setStyleCode == "2")
                    {
                        // 式2: 爪無し・AB2補正あり
                        wkSql += "AND (( NOT ([a] LIKE 'AB2%') ";
                        wkSql += $"AND [bigsize] >= {D(bSize / 1.18)} ";
                        wkSql += $"AND [bigsize] <= {D(bSize / 0.75)} ";
                        wkSql += $"AND [smallsize] >= {D(sSize / 1.18)} ";
                        wkSql += $"AND [smallsize] <= {D(sSize / 0.75)} ";
                        wkSql += $"AND [totalsize] >= {D(tSize)} ";
                        wkSql += $"AND [totalsize] <= {D(tSize + 10)}) ";
                        wkSql += "OR ( [a] LIKE 'AB2%' ";
                        wkSql += $"AND [bigsize] >= {D(bSize / 1.062)} ";
                        wkSql += $"AND [bigsize] <= {D(bSize / 0.675)} ";
                        wkSql += $"AND [smallsize] >= {D(sSize / 1.062)} ";
                        wkSql += $"AND [smallsize] <= {D(sSize / 0.675)} ";
                        wkSql += $"AND [totalsize] >= {D(tSize / 0.9)} ";
                        wkSql += $"AND [totalsize] <= {D((tSize + 10) / 0.9)})) ";
                    }
                    else
                    {
                        // 式1: 爪留め・AB2補正あり
                        wkSql += "AND (( NOT ([a] LIKE 'AB2%') ";
                        wkSql += $"AND [bigsize] >= {D(bSize / 1.18)} ";
                        wkSql += $"AND [bigsize] <= {D(bSize / 0.75)} ";
                        wkSql += $"AND [smallsize] >= {D(sSize / 1.18)} ";
                        wkSql += $"AND [smallsize] <= {D(sSize / 0.75)} ";
                        wkSql += $"AND [totalsize] >= {D(tSize - 5)} ";
                        wkSql += $"AND [totalsize] <= {D(tSize + 10)}) ";
                        wkSql += "OR ( [a] LIKE 'AB2%' ";
                        wkSql += $"AND [bigsize] >= {D(bSize / 1.062)} ";
                        wkSql += $"AND [bigsize] <= {D(bSize / 0.675)} ";
                        wkSql += $"AND [smallsize] >= {D(sSize / 1.062)} ";
                        wkSql += $"AND [smallsize] <= {D(sSize / 0.675)} ";
                        wkSql += $"AND [totalsize] >= {D((tSize - 5) / 0.9)} ";
                        wkSql += $"AND [totalsize] <= {D((tSize + 10) / 0.9)})) ";
                    }
                }
                else
                {
                    // ボール（球）: AB2補正なし
                    wkSql += $"AND [bigsize] >= {D(bSize / 1.18)} ";
                    wkSql += $"AND [bigsize] <= {D(bSize / 0.75)} ";
                    wkSql += $"AND [smallsize] >= {D(sSize / 1.18)} ";
                    wkSql += $"AND [smallsize] <= {D(sSize / 0.75)} ";
                    wkSql += $"AND [totalsize] >= {D(tSize - 5)} ";
                    wkSql += $"AND [totalsize] <= {D(tSize + 10)} ";
                }

                sql += wkSql;
                strSqlKotei += wkSql;
            }

            // リングサイズ条件 (アイテム "9" or "6" のみ)
            chkSqlM = "";
            if (itemCode == "9" || itemCode == "6")
            {
                chkSqlM = sql; // リングサイズ追加前のスナップショット
                int mSize = int.Parse(_cmbRingsizeJu.Text) * 10
                          + int.Parse(_cmbRingsizeIchi.Text);
                sql += $"AND ({mSize} BETWEEN CLng([w]) AND CLng([x])) ";
            }

            return sql;
        }

        // ----------------------------------------------------------------
        // 結果なし時の診断 (VB6: 段階的原因診断)
        // ----------------------------------------------------------------
        private void DiagnoseNoResult(
            string strSqlKotei,
            string chkSqlD, string chkSqlE, string chkSqlF,
            string chkSqlG, string chkSqlI, string chkSqlJ,
            string chkSqlM)
        {
           

            string itemCode = _cmbItem.Text.Substring(0, 1);

            // ① 固定条件のみ
            if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
            {
                MessageBox.Show("㈠「固定条件」に該当する設計データがありません。",
                    "「固定条件」該当データなし", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ② ダイヤインの数
            strSqlKotei += chkSqlD;
            if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
            {
                MessageBox.Show("ⓑ「ダイヤインの数」を変えてください。",
                    "「ダイヤインの数」該当データなし", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ③ ダイヤインのスタイル
            if (itemCode != "0")
            {
                strSqlKotei += chkSqlE;
                if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
                {
                    MessageBox.Show("ⓒ「ダイヤインのスタイル」を変えてください。",
                        "「ダイヤインのスタイル」該当データなし", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // ④ デザイン数量
            if (itemCode != "0")
            {
                strSqlKotei += chkSqlF;
                if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
                {
                    MessageBox.Show("ⓓ「デザイン数量」を変えてください。",
                        "「デザイン数量」該当データなし", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // ⑤ 石留セットスタイル
            if (itemCode != "0")
            {
                strSqlKotei += chkSqlG;
                if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
                {
                    MessageBox.Show("ⓔ「石留セットスタイル」を変えてください。",
                        "「石留セットスタイル」該当データなし", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // ⑥ 重視した見込予算
            strSqlKotei += chkSqlI;
            if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
            {
                MessageBox.Show("ⓕ「重視した見込予算」を変えてください。",
                    "「重視した見込予算」該当データなし", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ⑦ 重視した方法
            strSqlKotei += chkSqlJ;
            if (AppState.Db.ExecuteQuery(strSqlKotei).Rows.Count == 0)
            {
                MessageBox.Show("ⓖ「重視した方法」を変えてください。",
                    "「重視した方法」該当データなし", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ⑧ リングサイズ
            if (itemCode == "9" || itemCode == "6")
            {
                if (!string.IsNullOrEmpty(chkSqlM) &&
                    AppState.Db.ExecuteQuery(chkSqlM).Rows.Count > 0)
                {
                    MessageBox.Show("ⓗ「リングサイズ」を変えてください。",
                        "「リングサイズ」該当データなし", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        // ----------------------------------------------------------------
        // 品番重複除去
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
        // ランダムに最大7件を検索結果テーブルへ格納
        // ----------------------------------------------------------------
        private int SelectRandomResults()
        {
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

            DataTable wkDt = AppState.Db.ExecuteQuery(
                "SELECT [index] FROM [ワーク検索テーブル] ORDER BY [index]");

            var indices = new List<int>();
            foreach (DataRow row in wkDt.Rows)
                indices.Add(Convert.ToInt32(row["index"]));

            var rnd = new Random();
            for (int i = indices.Count - 1; i > 0; i--)
            {
                int j = rnd.Next(i + 1);
                int tmp = indices[i];
                indices[i] = indices[j];
                indices[j] = tmp;
            }

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
                    MessageBox.Show("ⓙ「サイズ」を入力してください。",
                        "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            string yosanCode = _cmbYosanType.Text.Substring(0, 1);
            bool yosanIsSpecified = (yosanCode == "1" || yosanCode == "2" || yosanCode == "3");
            bool yosanAmountZero = (_cmbYosanHyaku.Text == "0" &&
                                     _cmbYosanJu.Text == "0" &&
                                     _cmbYosanIchi.Text == "0");

            if (yosanIsSpecified && yosanAmountZero)
            {
                MessageBox.Show(
                    "ⓗ「重視した見込予算」に「第1」「第2」「第3」を選択した場合は、予算を入力してください。",
                    "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (yosanCode == "9" && !yosanAmountZero)
            {
                MessageBox.Show(
                    "ⓗ「重視した見込予算」が「その他」の場合、予算はゼロにしてください。",
                    "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (itemCode == "9" || itemCode == "6")
            {
                if (_cmbRingsizeJu.Text == "0" && _cmbRingsizeIchi.Text == "0")
                {
                    MessageBox.Show(
                        "ⓙ「アイテム」が「その他（リング）」でリングで検索する場合は、商品サイズを入力してください。",
                        "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            return true;
        }

        // ----------------------------------------------------------------
        // メニューへ戻る (VB6: btn_back_Click)
        // ----------------------------------------------------------------
        private void BtnBack_Click(object sender, EventArgs e)
        {
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
        // ----------------------------------------------------------------
        private static string D(double v) =>
    ((long)Math.Round(v)).ToString();
    }
}