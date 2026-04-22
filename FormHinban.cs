using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows.Forms;

namespace YumejitateApp
{
    /// <summary>
    /// AB品番検索フォーム (VB6: form_hinban.frm の移植)
    ///
    /// 検索テーブルに対して各条件を AND 結合した SQL を発行し、
    /// 重複排除・ランダム絞り込み（最大7件）を行って
    /// 検索結果テーブルに格納した後、FormMovie へ遷移する。
    /// </summary>
    public class FormHinban : Form
    {
        // ----------------------------------------------------------------
        // コントロール
        // ----------------------------------------------------------------

        // --- 検索条件コンボ ---
        private ComboBox _cmbItem;           // アイテム
        private ComboBox _cmbStoneShape;     // 上から見た石の形
        private ComboBox _cmbSetStyle;       // 中石セットスタイル
        private ComboBox _cmbJigane;         // 地金

        // --- 品番コンボ (品番一〜七) ---
        private ComboBox _cmbHinban1;        // A/B/C/I/J/K/N/V
        private ComboBox _cmbHinban2;        // A/B/C/I/J/K/N/V
        private ComboBox _cmbHinban3;        // 0-9
        private ComboBox _cmbHinban4;        // 0-9
        private ComboBox _cmbHinban5;        // 0-9
        private ComboBox _cmbHinban6;        // 0-9
        private ComboBox _cmbHinban7;        // 0-9

        // --- サイズコンボ (上：xx.x mm, 下：xx.x mm) ---
        private ComboBox _cmbSizeUeJu;       // サイズ上十
        private ComboBox _cmbSizeUeIchi;     // サイズ上一
        private ComboBox _cmbSizeUeKo;       // サイズ上小
        private ComboBox _cmbSizeShitaJu;    // サイズ下十
        private ComboBox _cmbSizeShitaIchi;  // サイズ下一
        private ComboBox _cmbSizeShitaKo;    // サイズ下小

        // --- リングサイズ / グレード ---
        private ComboBox _cmbRingsizeJu;     // cmb_ringsize十 (アイテム=リング時有効)
        private ComboBox _cmbRingsizeIchi;   // cmb_ringsize一
        private ComboBox _cmbGrade;          // cmb_grade A/B/C (アイテム≠リング時有効)

        // --- ボタン ---
        private Button _btnSearch;           // 品番検索
        private Button _btnBack;             // メニュー

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------
        public FormHinban()
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
            // VB6: BackColor=&H00D8FFFF&, Caption="夢仕立て-品番検索画面"
            this.Text = "夢仕立て - 品番検索";
            this.BackColor = System.Drawing.Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.AutoScroll = true;
            this.MinimumSize = new System.Drawing.Size(900, 650);

            var fntNormal = new System.Drawing.Font("ＭＳ Ｐゴシック", 14.25f, System.Drawing.FontStyle.Bold);
            var fntTitle = new System.Drawing.Font("ＭＳ Ｐゴシック", 18f, System.Drawing.FontStyle.Italic);
            var fntBtn = new System.Drawing.Font("ＭＳ Ｐゴシック", 15.75f, System.Drawing.FontStyle.Bold);
            var cmbSz = new System.Drawing.Size(50, 27);

            // ---- タイトル (VB6: Label2 "品番検索", Left=6480twip≈432px, Top=240twip≈16px) ----
            var lblTitle = MakeLabel("品番検索", fntTitle, new System.Drawing.Point(432, 16), true);

            // ================================================================
            // 検索条件 GroupBox
            // ================================================================
            var grp = new GroupBox
            {
                Text = "検索条件",
                Font = fntNormal,
                Location = new System.Drawing.Point(16, 56),
                Size = new System.Drawing.Size(940, 440),
                BackColor = System.Drawing.Color.Transparent,
            };

            int lblX = 16, cmbX = 220;

            // --- アイテム (Top=1920twip≈128px → grp内 row=0, y≈16) ---
            int row0 = 16, rowH = 44;
            grp.Controls.Add(MakeLabel("アイテム", fntNormal, new System.Drawing.Point(lblX, row0 + 4), false));
            _cmbItem = MakeCombo(fntNormal, new System.Drawing.Point(cmbX, row0), new System.Drawing.Size(280, 27));
            _cmbItem.SelectedIndexChanged += CmbItem_SelectedIndexChanged;
            grp.Controls.Add(_cmbItem);

            // --- 上から見た石の形 ---
            int row1 = row0 + rowH;
            grp.Controls.Add(MakeLabel("中石形状", fntNormal, new System.Drawing.Point(lblX, row1 + 4), false));
            _cmbStoneShape = MakeCombo(fntNormal, new System.Drawing.Point(cmbX, row1), new System.Drawing.Size(280, 27));
            grp.Controls.Add(_cmbStoneShape);

            // --- 中石セットスタイル ---
            int row2 = row1 + rowH;
            grp.Controls.Add(MakeLabel("中石セットスタイル", fntNormal, new System.Drawing.Point(lblX, row2 + 4), false));
            _cmbSetStyle = MakeCombo(fntNormal, new System.Drawing.Point(cmbX, row2), new System.Drawing.Size(280, 27));
            grp.Controls.Add(_cmbSetStyle);

            // --- 地金 ---
            int row3 = row2 + rowH;
            grp.Controls.Add(MakeLabel("地　金", fntNormal, new System.Drawing.Point(lblX, row3 + 4), false));
            _cmbJigane = MakeCombo(fntNormal, new System.Drawing.Point(cmbX, row3), new System.Drawing.Size(280, 27));
            grp.Controls.Add(_cmbJigane);

            // --- 品番 (VB6: 品番一〜七, Top=4800twip≈320px) ---
            int row4 = row3 + rowH;
            grp.Controls.Add(MakeLabel("品番", fntNormal, new System.Drawing.Point(lblX, row4 + 4), false));
            int hinbanX = cmbX;
            _cmbHinban1 = MakeCombo(fntNormal, new System.Drawing.Point(hinbanX, row4), cmbSz); grp.Controls.Add(_cmbHinban1);
            _cmbHinban2 = MakeCombo(fntNormal, new System.Drawing.Point(hinbanX + 52, row4), cmbSz); grp.Controls.Add(_cmbHinban2);
            _cmbHinban3 = MakeCombo(fntNormal, new System.Drawing.Point(hinbanX + 104, row4), cmbSz); grp.Controls.Add(_cmbHinban3);
            _cmbHinban4 = MakeCombo(fntNormal, new System.Drawing.Point(hinbanX + 156, row4), cmbSz); grp.Controls.Add(_cmbHinban4);
            _cmbHinban5 = MakeCombo(fntNormal, new System.Drawing.Point(hinbanX + 208, row4), cmbSz); grp.Controls.Add(_cmbHinban5);
            _cmbHinban6 = MakeCombo(fntNormal, new System.Drawing.Point(hinbanX + 260, row4), cmbSz); grp.Controls.Add(_cmbHinban6);
            _cmbHinban7 = MakeCombo(fntNormal, new System.Drawing.Point(hinbanX + 312, row4), cmbSz); grp.Controls.Add(_cmbHinban7);

            // --- サイズ (VB6: Top=5520twip≈368px) ---
            // xx.x mm × xx.x mm の形式（上限 × 下限）
            int row5 = row4 + rowH;
            grp.Controls.Add(MakeLabel("サイズ", fntNormal, new System.Drawing.Point(lblX, row5 + 4), false));
            int szX = cmbX;
            _cmbSizeUeJu = MakeCombo(fntNormal, new System.Drawing.Point(szX, row5), cmbSz); grp.Controls.Add(_cmbSizeUeJu);
            _cmbSizeUeIchi = MakeCombo(fntNormal, new System.Drawing.Point(szX + 52, row5), cmbSz); grp.Controls.Add(_cmbSizeUeIchi);
            grp.Controls.Add(MakeLabel("．", fntNormal, new System.Drawing.Point(szX + 104, row5 + 2), false));
            _cmbSizeUeKo = MakeCombo(fntNormal, new System.Drawing.Point(szX + 116, row5), cmbSz); grp.Controls.Add(_cmbSizeUeKo);
            grp.Controls.Add(MakeLabel("mm×", fntNormal, new System.Drawing.Point(szX + 168, row5 + 2), false));
            _cmbSizeShitaJu = MakeCombo(fntNormal, new System.Drawing.Point(szX + 218, row5), cmbSz); grp.Controls.Add(_cmbSizeShitaJu);
            _cmbSizeShitaIchi = MakeCombo(fntNormal, new System.Drawing.Point(szX + 270, row5), cmbSz); grp.Controls.Add(_cmbSizeShitaIchi);
            grp.Controls.Add(MakeLabel("．", fntNormal, new System.Drawing.Point(szX + 322, row5 + 2), false));
            _cmbSizeShitaKo = MakeCombo(fntNormal, new System.Drawing.Point(szX + 334, row5), cmbSz); grp.Controls.Add(_cmbSizeShitaKo);
            grp.Controls.Add(MakeLabel("mm", fntNormal, new System.Drawing.Point(szX + 386, row5 + 2), false));

            // --- リングサイズ / グレード (VB6: Top=6240twip≈416px) ---
            int row6 = row5 + rowH;
            grp.Controls.Add(MakeLabel("リングサイズまたは石のグレード", fntNormal, new System.Drawing.Point(lblX, row6), false));
            grp.Controls.Add(MakeLabel("＃", fntNormal, new System.Drawing.Point(cmbX - 20, row6 + 4), false));
            _cmbRingsizeJu = MakeCombo(fntNormal, new System.Drawing.Point(cmbX, row6), cmbSz); grp.Controls.Add(_cmbRingsizeJu);
            _cmbRingsizeIchi = MakeCombo(fntNormal, new System.Drawing.Point(cmbX + 52, row6), cmbSz); grp.Controls.Add(_cmbRingsizeIchi);
            grp.Controls.Add(MakeLabel("グレード", fntNormal, new System.Drawing.Point(cmbX + 110, row6 + 4), false));
            _cmbGrade = MakeCombo(fntNormal, new System.Drawing.Point(cmbX + 210, row6), new System.Drawing.Size(70, 27));
            grp.Controls.Add(_cmbGrade);

            // グループの高さ調整
            grp.Height = row6 + rowH + 16;

            // ---- 品番検索ボタン (VB6: Left=6480twip≈432px, Top=7080twip≈472px) ----
            _btnSearch = new Button
            {
                Text = "品番検索",
                Font = fntBtn,
                Location = new System.Drawing.Point(400, grp.Bottom + 20),
                Size = new System.Drawing.Size(160, 57),
                BackColor = System.Drawing.Color.FromArgb(192, 224, 255),
                Cursor = Cursors.Hand,
            };
            _btnSearch.Click += BtnSearch_Click;

            // ---- メニューボタン (VB6: Left=11040twip≈736px, Top=8040twip≈536px) ----
            _btnBack = new Button
            {
                Text = "メニュー",
                Font = fntBtn,
                Location = new System.Drawing.Point(700, grp.Bottom + 20),
                Size = new System.Drawing.Size(160, 57),
                BackColor = System.Drawing.Color.FromArgb(255, 192, 192),
                Cursor = Cursors.Hand,
            };
            _btnBack.Click += BtnBack_Click;

            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                lblTitle, grp, _btnSearch, _btnBack,
            });

            this.ResumeLayout(false);
        }

        // ----------------------------------------------------------------
        // ラベル生成ヘルパー
        // ----------------------------------------------------------------
        private static System.Windows.Forms.Label MakeLabel(
            string text, System.Drawing.Font font,
            System.Drawing.Point loc, bool autoSize)
        {
            return new System.Windows.Forms.Label
            {
                Text = text,
                Font = font,
                BackColor = System.Drawing.Color.Transparent,
                AutoSize = autoSize,
                Location = loc,
            };
        }

        // ----------------------------------------------------------------
        // コンボボックス生成ヘルパー
        // ----------------------------------------------------------------
        private static ComboBox MakeCombo(
            System.Drawing.Font font,
            System.Drawing.Point loc,
            System.Drawing.Size size)
        {
            return new ComboBox
            {
                Font = font,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = loc,
                Size = size,
            };
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
            // ---- アイテム ----
            _cmbItem.Items.AddRange(new object[] {
                "リング", "ブローチ", "ペンダント", "タイタック",
                "ピアス", "イアリング", "その他（バチカン）"
            });
            _cmbItem.SelectedIndex = 0;

            // ---- 上から見た石の形 ----
            _cmbStoneShape.Items.AddRange(new object[] {
                "ラウンド（円）", "オーバル（楕円）", "ボール（球）",
                "エメラルド（四角）", "マーキース", "ドロップ"
            });
            _cmbStoneShape.SelectedIndex = 0;

            // ---- 中石セットスタイル ----
            _cmbSetStyle.Items.AddRange(new object[] { "爪留め", "爪無し（レール留め等）" });
            _cmbSetStyle.SelectedIndex = 0;

            // ---- 地金 ----
            _cmbJigane.Items.AddRange(new object[] {
                "プラチナ", "Ｋ１８ＹＧ", "Ｋ１８ＷＧ", "コンビ", "シルバー", "Ｋ１０"
            });
            _cmbJigane.SelectedIndex = 0;

            // ---- 品番一,二: A/B/C/I/J/K/N/V ----
            string[] alpha = { "A", "B", "C", "I", "J", "K", "N", "V" };
            _cmbHinban1.Items.AddRange(alpha); _cmbHinban1.SelectedIndex = 0;
            _cmbHinban2.Items.AddRange(alpha); _cmbHinban2.SelectedIndex = 0;

            // ---- 品番三〜七: 0-9 ----
            string[] digits = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            _cmbHinban3.Items.AddRange(digits); _cmbHinban3.SelectedIndex = 0;
            _cmbHinban4.Items.AddRange(digits); _cmbHinban4.SelectedIndex = 0;
            _cmbHinban5.Items.AddRange(digits); _cmbHinban5.SelectedIndex = 0;
            _cmbHinban6.Items.AddRange(digits); _cmbHinban6.SelectedIndex = 0;
            _cmbHinban7.Items.AddRange(digits); _cmbHinban7.SelectedIndex = 0;

            // ---- サイズコンボ: 0-9 ----
            foreach (var cmb in new[] {
                _cmbSizeUeJu, _cmbSizeUeIchi, _cmbSizeUeKo,
                _cmbSizeShitaJu, _cmbSizeShitaIchi, _cmbSizeShitaKo })
            {
                cmb.Items.AddRange(digits);
                cmb.SelectedIndex = 0;
            }

            // ---- リングサイズコンボ: 0-9 ----
            _cmbRingsizeJu.Items.AddRange(digits); _cmbRingsizeJu.SelectedIndex = 0;
            _cmbRingsizeIchi.Items.AddRange(digits); _cmbRingsizeIchi.SelectedIndex = 0;

            // ---- グレード: A/B/C (初期無効) ----
            _cmbGrade.Items.AddRange(new object[] { "A", "B", "C" });
            _cmbGrade.SelectedIndex = 0;
            _cmbGrade.Enabled = false; // VB6: cmb_grade.Enabled = False

            // アイテム初期状態の反映
            UpdateItemDependent();
        }

        // ----------------------------------------------------------------
        // アイテム変更ハンドラ (VB6: アイテム_Click)
        // アイテムに応じてコントロールの有効/無効を切り替える
        // ----------------------------------------------------------------
        private void CmbItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateItemDependent();
        }

        private void UpdateItemDependent()
        {
            string item = _cmbItem.Text;

            // リング → リングサイズ有効, グレード無効
            bool isRing = item == "リング";
            bool isVatican = item == "その他（バチカン）";

            _cmbRingsizeJu.Enabled = isRing;
            _cmbRingsizeIchi.Enabled = isRing;
            _cmbGrade.Enabled = !isRing; // VB6: リング以外でグレード有効

            // バチカン → 石形状・セットスタイル・サイズ 無効
            bool sizeEnabled = !isVatican;
            _cmbStoneShape.Enabled = sizeEnabled;
            _cmbSetStyle.Enabled = sizeEnabled;
            _cmbSizeUeJu.Enabled = sizeEnabled;
            _cmbSizeUeIchi.Enabled = sizeEnabled;
            _cmbSizeUeKo.Enabled = sizeEnabled;
            _cmbSizeShitaJu.Enabled = sizeEnabled;
            _cmbSizeShitaIchi.Enabled = sizeEnabled;
            _cmbSizeShitaKo.Enabled = sizeEnabled;
        }

        // ----------------------------------------------------------------
        // 画面入力情報チェック (VB6: check_control → Boolean)
        // ----------------------------------------------------------------
        private bool CheckControl()
        {
            string item = _cmbItem.Text;

            // サイズ入力チェック（バチカン以外）
            if (item != "その他（バチカン）")
            {
                bool ueZero = _cmbSizeUeJu.Text == "0"
                            && _cmbSizeUeIchi.Text == "0"
                            && _cmbSizeUeKo.Text == "0";
                bool shitaZero = _cmbSizeShitaJu.Text == "0"
                              && _cmbSizeShitaIchi.Text == "0"
                              && _cmbSizeShitaKo.Text == "0";
                if (ueZero || shitaZero)
                {
                    MessageBox.Show("「サイズ」を入力して下さい。", "入力チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            // リングサイズチェック
            if (item == "リング")
            {
                if (_cmbRingsizeJu.Text == "0" && _cmbRingsizeIchi.Text == "0")
                {
                    MessageBox.Show(
                        "「アイテム」を「リング」で検索する場合、＃（サイズ）を入力して下さい。",
                        "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            return true;
        }

        // ----------------------------------------------------------------
        // 品番検索ボタン (VB6: btn_search_Click)
        // ----------------------------------------------------------------
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            // 入力チェック
            if (!CheckControl()) return;

            // 検索開始メッセージ
            MessageBox.Show(
                "数千点以上のデータからご希望のデザインをお探しします。",
                "検索開始", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // ワークテーブル・検索結果テーブル初期化
            // VB6: dao_database.Execute "delete * from ワーク検索テーブル"
            AppState.Db.ExecuteNonQuery("DELETE FROM [ワーク検索テーブル]");
            AppState.Db.ExecuteNonQuery("DELETE FROM [検索結果テーブル]");

            // ---- SQL 構築 ----
            string strsql = BuildSearchSql();

            // ---- 検索実行 ----
            DataTable dtResult = AppState.Db.ExecuteQuery(strsql);
            if (dtResult == null || dtResult.Rows.Count == 0)
            {
                MessageBox.Show(
                    "誠に申し訳ございませんが\nご希望のデザインはございません。",
                    "「固定情報」該当データ無し", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ---- ワーク検索テーブルに INSERT ----
            // VB6: INSERT INTO ワーク検索テーブル SELECT * FROM 検索テーブル WHERE ...
            string insertSql = "INSERT INTO [ワーク検索テーブル] " + strsql;
            AppState.Db.ExecuteNonQuery(insertSql);

            // ---- 品番の重複を除去 ----
            // VB6: ワーク検索テーブルを index 順にループして同じ品番(a)の2件目以降を削除
            RemoveDuplicateHinban();

            // ---- ランダムに最大7件を検索結果テーブルへ ----
            int searchCnt = SelectRandomResults();

            // ---- 結果件数表示 ----
            MessageBox.Show(
                $"ご希望のデザインは {searchCnt} 種類ございます。",
                "検索結果件数表示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            if (searchCnt == 0) return;

            // ---- 検索結果画面へ遷移 (VB6: form_hinban.Visible=False → form_movie.Show) ----
            this.Hide();
            var movieForm = new FormMovie();
            movieForm.FormClosed += (s2, e2) =>
            {
                // FormMovie が閉じたらこのフォームも閉じる
                this.Close();
            };
            movieForm.Show();
        }

        // ----------------------------------------------------------------
        // 検索 SQL 構築 (VB6: btn_search_Click 内の SQL 組み立て部分)
        //
        // 検索テーブルのカラム:
        //   a = 品番, b = アイテムコード, c = 石形状コード
        //   d = 地金コード, h = セットスタイルコード
        //   bigsize, smallsize, totalsize = サイズ情報
        //   w, x = リングサイズ範囲（下限、上限）
        //
        // Access LIKE では * がワイルドカード
        // ----------------------------------------------------------------
        private string BuildSearchSql()
        {
            var sb = new System.Text.StringBuilder();
            string item = _cmbItem.Text;

            sb.Append("SELECT * FROM [検索テーブル] WHERE ");

            // ---- アイテム → b like '*N*' ----
            string itemCode;
            switch (item)
            {
                case "リング": itemCode = "6"; break;
                case "ブローチ": itemCode = "8"; break;
                case "ペンダント": itemCode = "7"; break;
                case "タイタック": itemCode = "1"; break;
                case "ピアス": itemCode = "2"; break;
                case "イアリング": itemCode = "3"; break;
                case "その他（バチカン）": itemCode = "0"; break;
                default: itemCode = "6"; break;
            }
            sb.Append($"b LIKE '*{itemCode}*' ");

            // バチカンの場合は石形状なし: l='XXX' and m='XXX'
            if (item == "その他（バチカン）")
            {
                sb.Append("AND l = 'XXX' AND m = 'XXX' ");
            }
            else
            {
                // ---- 上から見た石の形 → c like '*N*' ----
                string shapeCode;
                switch (_cmbStoneShape.Text)
                {
                    case "ラウンド（円）": shapeCode = "1"; break;
                    case "オーバル（楕円）": shapeCode = "2"; break;
                    case "ボール（球）": shapeCode = "3"; break;
                    case "エメラルド（四角）": shapeCode = "4"; break;
                    case "マーキース": shapeCode = "5"; break;
                    case "ドロップ": shapeCode = "6"; break;
                    default: shapeCode = ""; break;
                }
                if (shapeCode != "")
                    sb.Append($"AND c LIKE '*{shapeCode}*' ");

                // ---- 中石セットスタイル → h like '*N*' ----
                string styleCode;
                switch (_cmbSetStyle.Text)
                {
                    case "爪留め": styleCode = "1"; break;
                    case "爪無し（レール留め等）": styleCode = "2"; break;
                    default: styleCode = ""; break;
                }
                if (styleCode != "")
                    sb.Append($"AND h LIKE '*{styleCode}*' ");
            }

            // ---- 地金 → d like '*N*' ----
            string jiganeCode;
            switch (_cmbJigane.Text)
            {
                case "プラチナ": jiganeCode = "1"; break;
                case "Ｋ１８ＹＧ": jiganeCode = "2"; break;
                case "Ｋ１８ＷＧ": jiganeCode = "3"; break;
                case "コンビ": jiganeCode = "4"; break;
                case "シルバー": jiganeCode = "5"; break;
                case "Ｋ１０": jiganeCode = "6"; break;
                default: jiganeCode = ""; break;
            }
            if (jiganeCode != "")
                sb.Append($"AND d LIKE '*{jiganeCode}*' ");

            // ---- 品番 → a = '品番一〜七連結' ----
            string hinban = _cmbHinban1.Text + _cmbHinban2.Text
                          + _cmbHinban3.Text + _cmbHinban4.Text
                          + _cmbHinban5.Text + _cmbHinban6.Text
                          + _cmbHinban7.Text;
            sb.Append($"AND a = '{hinban}' ");

            // ---- サイズ条件（バチカン以外）----
            if (item != "その他（バチカン）")
            {
                // VB6: size_a = サイズ上十*100 + サイズ上一*10 + サイズ上小
                double sizeA = ToDigit(_cmbSizeUeJu) * 100
                             + ToDigit(_cmbSizeUeIchi) * 10
                             + ToDigit(_cmbSizeUeKo);
                double sizeB = ToDigit(_cmbSizeShitaJu) * 100
                             + ToDigit(_cmbSizeShitaIchi) * 10
                             + ToDigit(_cmbSizeShitaKo);

                // 大きい方を bSize、小さい方を sSize
                double bSize = Math.Max(sizeA, sizeB);
                double sSize = Math.Min(sizeA, sizeB);
                double tSize = sizeA + sizeB;

                string wkSql;
                if (_cmbStoneShape.Text != "ボール（球）")
                {
                    if (_cmbSetStyle.Text == "爪無し（レール留め等）")
                    {
                        // VB6: 式２・式４
                        wkSql = BuildSizeCondition4(bSize, sSize, tSize);
                    }
                    else
                    {
                        // VB6: 式１・式３
                        wkSql = BuildSizeCondition3(bSize, sSize, tSize);
                    }
                }
                else
                {
                    // VB6: ボール（球）→ 式１ のみ
                    wkSql = BuildSizeConditionBall(bSize, sSize, tSize);
                }
                sb.Append(wkSql);
            }

            // ---- リングサイズ (VB6: m_size between val(w) and val(x)) ----
            if (item == "リング")
            {
                int mSize = ToDigit(_cmbRingsizeJu) * 10 + ToDigit(_cmbRingsizeIchi);
                sb.Append($"AND ({mSize} BETWEEN val(w) AND val(x)) ");
            }

            return sb.ToString();
        }

        // ----------------------------------------------------------------
        // サイズ条件: 式１・式３（爪留めの場合）
        // VB6: 非AB2品番式 + AB2品番式 を OR
        // ----------------------------------------------------------------
        private static string BuildSizeCondition3(double bSize, double sSize, double tSize)
        {
            // 式１: not(a like 'AB2*') → 通常条件
            // 式３: a like 'AB2*'     → AB2補正条件
            return "AND (("
                + "NOT (a LIKE 'AB2*') "
                + $"AND bigsize >= {D(bSize / 1.18)} "
                + $"AND bigsize <= {D(bSize / 0.75)} "
                + $"AND smallsize >= {D(sSize / 1.18)} "
                + $"AND smallsize <= {D(sSize / 0.75)} "
                + $"AND totalsize >= {D(tSize - 5)} "
                + $"AND totalsize <= {D(tSize + 10)}) "
                + "OR (a LIKE 'AB2*' "
                + $"AND bigsize >= {D(bSize / 1.062)} "
                + $"AND bigsize <= {D(bSize / 0.675)} "
                + $"AND smallsize >= {D(sSize / 1.062)} "
                + $"AND smallsize <= {D(sSize / 0.675)} "
                + $"AND totalsize >= {D((tSize - 5) / 0.9)} "
                + $"AND totalsize <= {D((tSize + 10) / 0.9)})) ";
        }

        // ----------------------------------------------------------------
        // サイズ条件: 式２・式４（爪無しの場合）
        // ----------------------------------------------------------------
        private static string BuildSizeCondition4(double bSize, double sSize, double tSize)
        {
            // 式２: 通常, 式４: AB2補正（totalsize の除算が異なる）
            return "AND (("
                + "NOT (a LIKE 'AB2*') "
                + $"AND bigsize >= {D(bSize / 1.18)} "
                + $"AND bigsize <= {D(bSize / 0.75)} "
                + $"AND smallsize >= {D(sSize / 1.18)} "
                + $"AND smallsize <= {D(sSize / 0.75)} "
                + $"AND totalsize >= {D(tSize)} "
                + $"AND totalsize <= {D(tSize + 10)}) "
                + "OR (a LIKE 'AB2*' "
                + $"AND bigsize >= {D(bSize / 1.062)} "
                + $"AND bigsize <= {D(bSize / 0.675)} "
                + $"AND smallsize >= {D(sSize / 1.062)} "
                + $"AND smallsize <= {D(sSize / 0.675)} "
                + $"AND totalsize >= {D(tSize / 0.9)} "
                + $"AND totalsize <= {D((tSize + 10) / 0.9)})) ";
        }

        // ----------------------------------------------------------------
        // サイズ条件: ボール（球）の場合（分岐なし）
        // ----------------------------------------------------------------
        private static string BuildSizeConditionBall(double bSize, double sSize, double tSize)
        {
            return $"AND bigsize >= {D(bSize / 1.18)} "
                 + $"AND bigsize <= {D(bSize / 0.75)} "
                 + $"AND smallsize >= {D(sSize / 1.18)} "
                 + $"AND smallsize <= {D(sSize / 0.75)} "
                 + $"AND totalsize >= {D(tSize - 5)} "
                 + $"AND totalsize <= {D(tSize + 10)} ";
        }

        // ----------------------------------------------------------------
        // 品番の重複排除 (VB6: preHinban/aftHinban ループで重複削除)
        // ワーク検索テーブルを index 順に取得し、同じ品番(a)が
        // 隣り合っていたら後の方を削除する
        // ----------------------------------------------------------------
        private void RemoveDuplicateHinban()
        {
            DataTable dt = AppState.Db.ExecuteQuery(
                "SELECT * FROM [ワーク検索テーブル] ORDER BY [index]");
            if (dt == null || dt.Rows.Count == 0) return;

            string prevHinban = "";
            foreach (DataRow row in dt.Rows)
            {
                string curHinban = row["a"].ToString();
                if (string.Compare(curHinban, prevHinban, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    // 同じ品番 → 削除
                    AppState.Db.ExecuteNonQuery(
                        $"DELETE FROM [ワーク検索テーブル] WHERE [index] = {row["index"]}");
                }
                else
                {
                    prevHinban = curHinban;
                }
            }
        }

        // ----------------------------------------------------------------
        // ランダムに最大7件を検索結果テーブルへ INSERT
        // VB6: 乱数テーブルを使った RNG シード → C# では標準 Random で代替
        //
        // アルゴリズム:
        //   1. ワークテーブルから全 index を取得してリスト化
        //   2. Fisher-Yates シャッフル
        //   3. 先頭から最大7件を 検索結果テーブルへ INSERT
        //   4. 実際に INSERT できた件数を返す
        // ----------------------------------------------------------------
        private int SelectRandomResults()
        {
            DataTable dtWork = AppState.Db.ExecuteQuery(
                "SELECT * FROM [ワーク検索テーブル] ORDER BY [index]");
            if (dtWork == null || dtWork.Rows.Count == 0) return 0;

            // index 一覧を取得
            var indexList = new List<object>();
            foreach (DataRow row in dtWork.Rows)
                indexList.Add(row["index"]);

            // Fisher-Yates シャッフル (VB6 の Rnd ベース乱数選択を再現)
            var rnd = new Random();
            for (int i = indexList.Count - 1; i > 0; i--)
            {
                int j = rnd.Next(0, i + 1);
                var tmp = indexList[i];
                indexList[i] = indexList[j];
                indexList[j] = tmp;
            }

            // 先頭から最大7件を 検索結果テーブルへ
            int count = 0;
            int maxCount = Math.Min(7, indexList.Count);
            for (int i = 0; i < maxCount; i++)
            {
                string insertSql =
                    "INSERT INTO [検索結果テーブル] "
                    + "SELECT * FROM [ワーク検索テーブル] "
                    + $"WHERE [index] = {indexList[i]}";
                AppState.Db.ExecuteNonQuery(insertSql);
                count++;
            }
            return count;
        }

        // ----------------------------------------------------------------
        // メニューボタン (VB6: btn_back_Click)
        // VB6: Flag_Hinban = False → form_hinban.Visible = False → form_menu.Visible = True
        // ----------------------------------------------------------------
        private void BtnBack_Click(object sender, EventArgs e)
        {
            AppState.FlagHinban = false;

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
        // ユーティリティ
        // ----------------------------------------------------------------

        /// <summary>コンボボックスの選択値を整数で返す</summary>
        private static int ToDigit(ComboBox cmb)
        {
            if (int.TryParse(cmb.Text, out int v)) return v;
            return 0;
        }

        /// <summary>
        /// double を SQL 文字列化（ロケール非依存）
        /// VB6 の CDbl → CStr 相当
        /// </summary>
        private static string D(double v) =>
            v.ToString("G", CultureInfo.InvariantCulture);
    }
}
