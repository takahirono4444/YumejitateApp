using System;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Windows.Forms;

// ================================================================
// FormDispPicture.cs
// VB6: form_disp_picture.frm → C# + WinForms 移行
// 機能: 画像表示画面
//       商品画像・カメラ画像・合成画像を切り替え表示し、
//       商品データを下部に表示する。印刷・削除・オーダーメイド連携も担う。
//
// LEADTools 置き換え:
//   LEAD1.Load()   → Image.FromFile() + PictureBox.Image
//   LEAD1.Size()   → PictureBox.SizeMode = Zoom（自動縮拡）
//   LEAD2.Rotate() → Image.RotateFlip(Rotate90FlipNone)
//   LEAD2.Render() → Graphics.DrawImage() in PrintDocument.PrintPage
//
// 対象フレームワーク: .NET Framework 4.8
// C#バージョン: 7.3
// ================================================================

namespace YumejitateApp
{
    public class FormDispPicture : Form
    {
        // ----------------------------------------------------------------
        // 画像ファイル名テーブル（VB6の Option_picture(Index) に対応）
        // Index 0-3   : 商品画像①-④ (syouhin_1.jpg - syouhin_4.jpg)
        // Index 4-7   : カメラ画像①-④ (camera_1.jpg - camera_4.jpg)
        // Index 8-11  : 合成画像①-④  (gousei_1.jpg  - gousei_4.jpg)
        // Index 12-13 : 商品画像⑤-⑥ (syouhin_5.jpg - syouhin_6.jpg)
        // Index 14-15 : カメラ画像⑤-⑥ (camera_5.jpg - camera_6.jpg)
        // Index 16-17 : 合成画像⑤-⑥  (gousei_5.jpg  - gousei_6.jpg)
        // ----------------------------------------------------------------
        private static readonly string[] ImageFiles = new string[]
        {
            "syouhin_1.jpg", "syouhin_2.jpg", "syouhin_3.jpg", "syouhin_4.jpg", // 0-3
            "camera_1.jpg",  "camera_2.jpg",  "camera_3.jpg",  "camera_4.jpg",  // 4-7
            "gousei_1.jpg",  "gousei_2.jpg",  "gousei_3.jpg",  "gousei_4.jpg",  // 8-11
            "syouhin_5.jpg", "syouhin_6.jpg",                                   // 12-13
            "camera_5.jpg",  "camera_6.jpg",                                    // 14-15
            "gousei_5.jpg",  "gousei_6.jpg",                                    // 16-17
        };

        // 商品画像インデックス（SetSyouhinData を呼ぶ対象）
        // Index 0-3, 12-13 が商品画像に対応
        // DB の index フィールド値（1-6）との対応
        private static readonly int[] SyouhinDbIndex = new int[]
        {
            1, 2, 3, 4, -1, -1, -1, -1, -1, -1, -1, -1, 5, 6, -1, -1, -1, -1
        };

        // ----------------------------------------------------------------
        // UIコントロール
        // ----------------------------------------------------------------

        // メイン画像表示（LEAD1 の代替）
        private PictureBox _picMain;

        // オプションボタン（18個・VB6の Option_picture(0)-(17) に対応）
        private RadioButton[] _optPicture = new RadioButton[18];

        // ------- 商品データ表示ラベル（フィールド名は数字） -------
        // 価格帯バー: フィールド "1"-"9","101"（横並び10個）
        private Label[] _lblPrice = new Label[10]; // 1,2,3,4,5,6,7,8,9,101

        private Label _lblField11;   // フィールド "11" 品番ラベル行
        private Label _lblField12;   // フィールド "12" FD
        private Label _lblField13;   // フィールド "13" wt
        private Label _lblField14;   // フィールド "14"
        private Label _lblField15;   // フィールド "15" 品番
        private Label _lblField18;   // フィールド "18" 加工コード
        private Label _lblField22;   // フィールド "22" 地金種別
        private Label _lblField23;   // フィールド "23"
        private Label _lblField24;   // フィールド "24" 警告（赤字）
        private Label _lblField25;   // フィールド "25"
        private Label _lblField26;   // フィールド "26"
        private Label _lblField27;   // フィールド "27" 重量
        private Label _lblField102;  // フィールド "102"
        private Label _lblField30;   // フィールド "30" 税込価格
        // 固定ラベル
        private Label _lblWt;        // "wt．約"
        private Label _lblMd;        // "MD．約"
        private Label _lblFd;        // "FD．約"

        // ボタン
        private Button _btnBack;     // メニュー
        private Button _btnPrint;    // 画像印刷（VB6では Visible=False）
        private Button _btnDelete;   // 全て削除
        private Button _btnOrder;    // オーダーメイド

        // ----------------------------------------------------------------
        // 印刷用バッファ（LEAD2 の代替）
        // ----------------------------------------------------------------
        private Image _printImage;

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------
        public FormDispPicture()
        {
            InitializeComponent();
        }

        // ----------------------------------------------------------------
        // フォームデザイン初期化
        // ----------------------------------------------------------------
        private void InitializeComponent()
        {
            this.Text = "夢仕立て-画像表示画面";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(1100, 780);
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF); // VB6: &H00D8FFFF
            this.Font = new Font("メイリオ", 10F);
            this.Load += new EventHandler(FormDispPicture_Load);
            this.FormClosing += new FormClosingEventHandler(FormDispPicture_FormClosing);

            // ================================================================
            // タイトルラベル
            // ================================================================
            var lblTitle = new Label();
            lblTitle.Text = "画像表示";
            lblTitle.Font = new Font("メイリオ", 18F, FontStyle.Italic);
            lblTitle.AutoSize = true;
            lblTitle.BackColor = Color.Transparent;
            lblTitle.Location = new Point(480, 10);
            this.Controls.Add(lblTitle);

            // ================================================================
            // 中央: メイン画像表示 PictureBox（LEAD1 の代替）
            // VB6: Left=2760 Top=600 Width=9615 Height=6615 (twips)
            //      ≒ 約184px, 40px, 641px, 441px (@15twip/px)
            // ================================================================
            _picMain = new PictureBox();
            _picMain.Location = new Point(30, 50);
            _picMain.Size = new Size(780, 530);
            _picMain.SizeMode = PictureBoxSizeMode.Zoom; // LEADTools.Size() に相当（アスペクト比維持）
            _picMain.BackColor = Color.Black;
            _picMain.BorderStyle = BorderStyle.FixedSingle;
            this.Controls.Add(_picMain);

            // ================================================================
            // 右パネル: オプションボタン（LEAD1 右側, VB6 Left≈13080 twips ≈ 872px）
            // 3グループ: 商品画像(0-3,12-13), カメラ画像(4-7,14-15), 合成画像(8-11,16-17)
            // ================================================================
            var pnlOptions = new Panel();
            pnlOptions.Location = new Point(820, 40);
            pnlOptions.Size = new Size(250, 570);
            pnlOptions.BackColor = Color.Transparent;
            this.Controls.Add(pnlOptions);

            // 各グループのラベルとオプションボタンを生成
            BuildOptionButtons(pnlOptions);

            // ================================================================
            // 下部: 商品データ表示エリア
            // VB6: Top ≈ 7680-9240 twips ≈ 512-616px → フォーム下部
            // ================================================================
            var pnlData = new Panel();
            pnlData.Location = new Point(0, 590);
            pnlData.Size = new Size(1100, 130);
            pnlData.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.Controls.Add(pnlData);

            BuildDataLabels(pnlData);

            // ================================================================
            // ボタン行（フォーム最下部, VB6 Top=9720 twips ≈ 648px）
            // ================================================================
            var btnY = 725;

            _btnBack = new Button();
            _btnBack.Text = "メニュー";
            _btnBack.Location = new Point(700, btnY);
            _btnBack.Size = new Size(130, 45);
            _btnBack.Font = new Font("メイリオ", 11F, FontStyle.Bold);
            _btnBack.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            _btnBack.Click += new EventHandler(BtnBack_Click);
            this.Controls.Add(_btnBack);

            _btnPrint = new Button();
            _btnPrint.Text = "画像印刷";
            _btnPrint.Location = new Point(120, btnY);
            _btnPrint.Size = new Size(130, 45);
            _btnPrint.Font = new Font("メイリオ", 11F, FontStyle.Bold);
            _btnPrint.Visible = false; // VB6では Visible=False
            _btnPrint.Click += new EventHandler(BtnPrint_Click);
            this.Controls.Add(_btnPrint);

            _btnDelete = new Button();
            _btnDelete.Text = "全て削除";
            _btnDelete.Location = new Point(840, btnY);
            _btnDelete.Size = new Size(130, 45);
            _btnDelete.Font = new Font("メイリオ", 11F, FontStyle.Bold);
            _btnDelete.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            _btnDelete.Click += new EventHandler(BtnDelete_Click);
            this.Controls.Add(_btnDelete);

            _btnOrder = new Button();
            _btnOrder.Text = "  オーダー   メイド";
            _btnOrder.Location = new Point(445, btnY);
            _btnOrder.Size = new Size(130, 45);
            _btnOrder.Font = new Font("メイリオ", 11F, FontStyle.Bold);
            _btnOrder.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            _btnOrder.Click += new EventHandler(BtnOrder_Click);
            this.Controls.Add(_btnOrder);
        }

        // ================================================================
        // オプションボタン18個をグループ別に生成
        // ================================================================
        private void BuildOptionButtons(Panel parent)
        {
            int y = 0;
            int btnH = 28;
            int rowGap = 2;

            // グループ定義: (ラベルテキスト, 開始Index, 個数)
            var groups = new (string label, int start, int count)[]
            {
                ("商品画像",   0,  4),   // Index 0-3
                ("",           12, 2),   // Index 12-13 (商品画像⑤⑥, 連続)
                ("カメラ画像",  4,  4),   // Index 4-7
                ("",           14, 2),   // Index 14-15
                ("合成画像",    8,  4),   // Index 8-11
                ("",           16, 2),   // Index 16-17
            };

            // 全Index順の配置リスト（トップから）
            // 商品画像 0,1,2,3,12,13 / カメラ 4,5,6,7,14,15 / 合成 8,9,10,11,16,17
            var layout = new int[][]
            {
                new int[]{ 0, 1, 2, 3, 12, 13 },
                new int[]{ 4, 5, 6, 7, 14, 15 },
                new int[]{ 8, 9, 10, 11, 16, 17 },
            };
            string[] groupNames = { "商品画像", "カメラ画像", "合成画像" };

            for (int g = 0; g < 3; g++)
            {
                // グループラベル
                var lbl = new Label();
                lbl.Text = groupNames[g];
                lbl.Location = new Point(0, y);
                lbl.AutoSize = true;
                lbl.Font = new Font("メイリオ", 10F, FontStyle.Bold);
                lbl.BackColor = Color.Transparent;
                parent.Controls.Add(lbl);
                y += 22;

                // 2列でオプションボタンを配置（①② / ③④ / ⑤⑥）
                int[] indices = layout[g];
                for (int row = 0; row < 3; row++)
                {
                    for (int col = 0; col < 2; col++)
                    {
                        int idx = indices[row * 2 + col];
                        var rb = new RadioButton();
                        string num = GetCircledNumber(row * 2 + col + 1);
                        rb.Text = num;
                        rb.Location = new Point(col * 110, y);
                        rb.Size = new Size(100, btnH);
                        rb.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
                        rb.Font = new Font("メイリオ", 11F, FontStyle.Bold);
                        rb.Tag = idx; // オプションボタンのIndex番号を Tag に保持
                        rb.CheckedChanged += new EventHandler(OptionPicture_CheckedChanged);
                        parent.Controls.Add(rb);
                        _optPicture[idx] = rb;
                    }
                    y += btnH + rowGap;
                }
                y += 6; // グループ間スペース
            }
        }

        private string GetCircledNumber(int n)
        {
            // ①②③④⑤⑥
            char[] circled = { '①', '②', '③', '④', '⑤', '⑥' };
            if (n >= 1 && n <= 6) return circled[n - 1].ToString();
            return n.ToString();
        }

        // ================================================================
        // 商品データ表示ラベルを生成（フォーム下部パネル内）
        // VB6: Label_1〜Label_9, Label_101（価格バー）
        //      label_11/12/13/14/15/18/22〜27/102/30 など
        // ================================================================
        private void BuildDataLabels(Panel parent)
        {
            int fontSz = 9;
            var font = new Font("メイリオ", fontSz, FontStyle.Bold);

            // --- 価格バー: Label_1〜Label_9, Label_101（横10個, フィールド"1"-"9","101"）---
            // VB6: 枠付きラベル（BorderStyle=Fixed Single）が横並びに並ぶ価格帯表示
            int bx = 470, by = 5, bw = 36, bh = 30;
            for (int i = 0; i < 10; i++)
            {
                _lblPrice[i] = new Label();
                _lblPrice[i].Location = new Point(bx + i * (bw + 2), by);
                _lblPrice[i].Size = new Size(bw, bh);
                _lblPrice[i].BorderStyle = BorderStyle.FixedSingle;
                _lblPrice[i].TextAlign = ContentAlignment.MiddleCenter;
                _lblPrice[i].Font = new Font("メイリオ", fontSz, FontStyle.Bold);
                _lblPrice[i].BackColor = SystemColors.Window;
                _lblPrice[i].Visible = false;
                parent.Controls.Add(_lblPrice[i]);
            }

            // --- 左列（地金・加工コード・品番等）---
            _lblField18 = MakeDataLabel(parent, "", 10, 5, 250, font);   // 加工コード (field "18")
            _lblField22 = MakeDataLabel(parent, "", 10, 35, 200, font);  // 地金種別 (field "22")
            _lblField23 = MakeDataLabel(parent, "", 10, 65, 200, font);  // field "23"
            _lblField24 = MakeDataLabel(parent, "", 10, 95, 350, font);  // 赤字警告 (field "24")
            _lblField24.ForeColor = Color.Red;

            // --- 中左列（MD/FD/wt固定ラベル + 値）---
            _lblMd = MakeFixedLabel(parent, "MD．約", 270, 35, font);
            _lblFd = MakeFixedLabel(parent, "FD．約", 270, 65, font);
            _lblWt = MakeFixedLabel(parent, "wt．約", 270, 95, font);

            _lblField25 = MakeDataLabel(parent, "", 340, 35, 90, font);   // field "25" MD値
            _lblField26 = MakeDataLabel(parent, "", 340, 65, 90, font);   // field "26" FD値
            _lblField27 = MakeDataLabel(parent, "", 340, 95, 90, font);   // field "27" 重量

            _lblField15 = MakeDataLabel(parent, "", 10, 5, 450, font);    // 品番 (field "15") – 上書き位置
            // field 15 は field 18 と同じ行かもしれないので少し右にずらす
            _lblField15.Location = new Point(10, 5);
            _lblField15.Size = new Size(450, 25);

            // 実際にはfield18が加工コード行: "加工：XX-YY" 形式
            // field15が品番行: "品番：AB1234..." 形式
            // VB6では label_15 は label_18 の下の行なので y位置を調整
            _lblField15.Location = new Point(10, 5);
            _lblField18.Location = new Point(10, 30);
            _lblField22.Location = new Point(10, 55);
            _lblField23.Location = new Point(10, 80);

            // --- 右列（税込価格等）---
            _lblField11 = MakeDataLabel(parent, "", 700, 5, 370, font);   // field "11"
            _lblField12 = MakeDataLabel(parent, "", 700, 35, 370, font);  // field "12"
            _lblField13 = MakeDataLabel(parent, "", 700, 65, 370, font);  // field "13"
            _lblField14 = MakeDataLabel(parent, "", 700, 95, 370, font);  // field "14"
            _lblField102 = MakeDataLabel(parent, "", 700, 5, 370, font);  // field "102"

            // 税込価格（Label_30）: フォームの左に目立つように
            _lblField30 = MakeDataLabel(parent, "", 160, 95, 250, font);  // field "30" 税込価格
            _lblField30.Font = new Font("メイリオ", 11F, FontStyle.Bold);
            _lblField30.ForeColor = Color.DarkBlue;
        }

        private Label MakeDataLabel(Panel parent, string text, int x, int y, int width, Font font)
        {
            var lbl = new Label();
            lbl.Text = text;
            lbl.Location = new Point(x, y);
            lbl.Size = new Size(width, 25);
            lbl.Font = font;
            lbl.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            lbl.Visible = false;
            parent.Controls.Add(lbl);
            return lbl;
        }

        private Label MakeFixedLabel(Panel parent, string text, int x, int y, Font font)
        {
            var lbl = new Label();
            lbl.Text = text;
            lbl.Location = new Point(x, y);
            lbl.AutoSize = true;
            lbl.Font = font;
            lbl.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            lbl.Visible = false;
            parent.Controls.Add(lbl);
            return lbl;
        }

        // ================================================================
        // Form_Load
        // VB6: 最大化 → syouhin_1.jpg を LEAD1 に読み込み →
        //       Init_Control → SetSyouhinData → Option_picture(0) を選択色に
        // ================================================================
        private void FormDispPicture_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;

            // 初期画像読み込み（商品画像①）
            LoadImage(0);

            // 商品データラベル初期化
            InitControl();

            // 商品データ表示（index=1 のレコード）
            SetSyouhinData(1);

            // オプションボタン(0)を選択状態にし、選択色（ピンク）を設定
            if (_optPicture[0] != null)
            {
                _optPicture[0].Checked = true;
                UpdateOptionColors();
            }
        }

        private void FormDispPicture_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 印刷バッファを解放
            DisposePrintImage();
        }

        // ================================================================
        // 画像読み込み
        // VB6: LEAD1.Load App.Path + "\xxx.jpg", 0, 0, 1
        // C# : Image.FromFile() を使用し PictureBox にセット
        //      SizeMode=Zoom のため LEAD1.Size() 相当の縮拡は自動
        // ================================================================
        private void LoadImage(int index)
        {
            string path = GetImagePath(index);
            if (!File.Exists(path)) return;

            try
            {
                // 以前の Image を解放してからロード（メモリリーク防止）
                var prev = _picMain.Image;
                _picMain.Image = Image.FromFile(path);
                if (prev != null) prev.Dispose();
            }
            catch
            {
                // ファイルが壊れている・読み取れない場合は何もしない
            }
        }

        private string GetImagePath(int index)
        {
            return Path.Combine(Application.StartupPath, ImageFiles[index]);
        }

        // ================================================================
        // オプションボタン変更イベント
        // VB6: Option_picture_Click(Index As Integer)
        // ================================================================
        private void OptionPicture_CheckedChanged(object sender, EventArgs e)
        {
            var rb = (RadioButton)sender;
            if (!rb.Checked) return; // 選択解除イベントは無視

            int idx = (int)rb.Tag;

            // 画像を一旦非表示にして切り替え
            _picMain.Visible = false;

            // ラベルをクリア
            InitControl();

            // 画像を読み込む
            LoadImage(idx);

            // 商品画像の場合はデータも表示
            int dbIdx = SyouhinDbIndex[idx];
            if (dbIdx > 0)
            {
                SetSyouhinData(dbIdx);
            }

            _picMain.Visible = true;

            // オプションボタンの選択色を更新
            UpdateOptionColors();
        }

        /// <summary>
        /// 選択中のオプションボタンをピンク、非選択を水色に設定
        /// VB6: Option_picture(i).BackColor = &HFFC0C0 (selected) / &HD8FFFF (unselected)
        /// </summary>
        private void UpdateOptionColors()
        {
            Color selected = Color.FromArgb(0xFF, 0xC0, 0xC0); // ピンク
            Color unselected = Color.FromArgb(0xD8, 0xFF, 0xFF); // 水色

            for (int i = 0; i < _optPicture.Length; i++)
            {
                if (_optPicture[i] != null)
                    _optPicture[i].BackColor = _optPicture[i].Checked ? selected : unselected;
            }
        }

        // ================================================================
        // 商品データ表示ラベルをクリア・非表示にする
        // VB6: Init_Control() - 全ラベルのCaption=""かつ Visible=False
        // ================================================================
        private void InitControl()
        {
            // 価格バー
            for (int i = 0; i < _lblPrice.Length; i++)
            {
                _lblPrice[i].Text = "";
                _lblPrice[i].Visible = false;
            }

            // データラベル
            Label[] allDataLabels = new Label[]
            {
                _lblField11, _lblField12, _lblField13, _lblField14, _lblField15,
                _lblField18, _lblField22, _lblField23, _lblField24,
                _lblField25, _lblField26, _lblField27,
                _lblField102, _lblField30
            };
            foreach (var lbl in allDataLabels)
            {
                if (lbl != null) { lbl.Text = ""; lbl.Visible = false; }
            }

            // 固定ラベル
            if (_lblWt != null) _lblWt.Visible = false;
            if (_lblMd != null) _lblMd.Visible = false;
            if (_lblFd != null) _lblFd.Visible = false;
        }

        // ================================================================
        // 商品データをDBから読み込んでラベルに表示
        // VB6: SetSyouhinData() - 商品データテーブル WHERE index = 'dbIndex'
        // フィールド名: "1"-"9","11"-"15","18","22"-"27","101","102","30"
        // ================================================================
        private void SetSyouhinData(int dbIndex)
        {
            try
            {
                var dt = AppState.Db.ExecuteQuery(
                    "SELECT * FROM [商品データテーブル] WHERE [index] = '" + dbIndex + "'");

                if (dt.Rows.Count == 0) return;

                var row = dt.Rows[0];

                // 価格バー (フィールド "1"-"9","101")
                string[] priceFields = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "101" };
                for (int i = 0; i < priceFields.Length; i++)
                {
                    _lblPrice[i].Text = GetField(row, priceFields[i]);
                    _lblPrice[i].Visible = true;
                }

                // データラベル
                SetLabelField(_lblField11, row, "11");
                SetLabelField(_lblField12, row, "12");
                SetLabelField(_lblField13, row, "13");
                SetLabelField(_lblField14, row, "14");
                SetLabelField(_lblField15, row, "15");
                SetLabelField(_lblField18, row, "18");
                SetLabelField(_lblField22, row, "22");
                SetLabelField(_lblField23, row, "23");
                SetLabelField(_lblField24, row, "24");
                SetLabelField(_lblField25, row, "25");
                SetLabelField(_lblField26, row, "26");
                SetLabelField(_lblField27, row, "27");
                SetLabelField(_lblField102, row, "102");
                SetLabelField(_lblField30, row, "30");

                // 固定ラベル
                _lblWt.Visible = true;
                _lblMd.Visible = true;
                _lblFd.Visible = true;
            }
            catch
            {
                // DBエラーはサイレントで無視（データが無い場合もある）
            }
        }

        private void SetLabelField(Label lbl, DataRow row, string fieldName)
        {
            if (lbl == null) return;
            lbl.Text = GetField(row, fieldName);
            lbl.Visible = true;
        }

        private string GetField(DataRow row, string fieldName)
        {
            try { return row[fieldName].ToString(); }
            catch { return ""; }
        }

        // ================================================================
        // 画像印刷ボタン
        // VB6: btn_print_Click
        //   LEAD2 に LEAD1 のビットマップをコピー → 90°回転 → 輝度調整 →
        //   1ページに 2×2 計4枚印刷
        //
        // C# 移植:
        //   Image.Clone() → RotateFlip(Rotate90FlipNone) →
        //   PrintDocument.PrintPage で DrawImage × 4
        // ================================================================
        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("画像を印刷します。よろしいですか？", "画像印刷確認",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            if (_picMain.Image == null)
            {
                MessageBox.Show("表示中の画像がありません。", "印刷エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 印刷バッファを解放してから作成
                DisposePrintImage();

                // 画像を複製して 90°回転（LEAD2.Rotate 9000 に相当）
                _printImage = (Image)_picMain.Image.Clone();
                _printImage.RotateFlip(RotateFlipType.Rotate90FlipNone);

                // 印刷ダイアログ経由で印刷
                var pd = new PrintDocument();
                pd.PrintPage += new PrintPageEventHandler(PrintDocument_PrintPage);

                var dlg = new PrintDialog();
                dlg.Document = pd;
                if (dlg.ShowDialog() == DialogResult.OK)
                    pd.Print();
            }
            catch (Exception ex)
            {
                MessageBox.Show("印刷中にエラーが発生しました。\n" + ex.Message,
                    "印刷エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 印刷ページ描画
        /// VB6: 1ページに同じ画像を2列×2行（計4枚）印刷
        ///      LEAD2.Render Printer.hdc, left, top, width, height
        /// </summary>
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (_printImage == null) return;

            Graphics g = e.Graphics;

            // 印刷可能領域
            int pageWidth = e.PageBounds.Width;
            int pageHeight = e.PageBounds.Height;

            // VB6: 上半分をページの半分のエリアに2列印刷、下半分も同様
            // MaxImageHeight = (UsableHeight - 12行分) / 2
            // C# では単純にページを4分割して各象限に印刷
            int imgW = pageWidth / 2 - 20;
            int imgH = pageHeight / 2 - 40;

            // アスペクト比を維持したサイズ算出
            double srcW = _printImage.Width;
            double srcH = _printImage.Height;
            double scaleX = imgW / srcW;
            double scaleY = imgH / srcH;
            double scale = Math.Min(scaleX, scaleY);
            int drawW = (int)(srcW * scale);
            int drawH = (int)(srcH * scale);

            // 2列×2行 = 4枚
            int leftA = 420 / 100;           // VB6: PrintLeft = 420 → ≈ 1px
            int leftB = 2750 / 100;          // VB6: PrintLeft = 2750
            int[] xs = new int[] { 30, drawW + 40 };
            int[] ys = new int[] { 30, drawH + 60 };

            foreach (int row in new int[] { 0, 1 })
            {
                foreach (int col in new int[] { 0, 1 })
                {
                    var destRect = new Rectangle(xs[col], ys[row], drawW, drawH);
                    g.DrawImage(_printImage, destRect,
                        new Rectangle(0, 0, _printImage.Width, _printImage.Height),
                        GraphicsUnit.Pixel);
                }
            }

            e.HasMorePages = false;
        }

        private void DisposePrintImage()
        {
            if (_printImage != null)
            {
                _printImage.Dispose();
                _printImage = null;
            }
        }

        // ================================================================
        // 全て削除ボタン
        // VB6: btn_delete_Click
        //   18個の jpg を Kill（削除）→ データ無し.JPG をコピー →
        //   現在のオプションに応じた画像を再ロード →
        //   商品データテーブルを全削除 → Init_Control
        // ================================================================
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("全ての画像とデータを削除します。よろしいですか？",
                "画像削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                return;

            string appPath = Application.StartupPath;
            string noDataSrc = Path.Combine(appPath, "データ無し.JPG");

            // --- 18個の画像ファイルを削除してプレースホルダーに置換 ---
            foreach (string fname in ImageFiles)
            {
                string fullPath = Path.Combine(appPath, fname);
                try
                {
                    if (File.Exists(fullPath)) File.Delete(fullPath);
                    if (File.Exists(noDataSrc)) File.Copy(noDataSrc, fullPath, true);
                }
                catch { /* 削除・コピーに失敗しても処理を継続 */ }
            }

            // --- 現在選択中の画像を再ロード ---
            int curIdx = GetCheckedIndex();
            if (curIdx >= 0) LoadImage(curIdx);

            // --- 商品データテーブルを全削除 ---
            try
            {
                AppState.Db.ExecuteNonQuery("DELETE * FROM [商品データテーブル]");
            }
            catch { }

            // --- ラベルをクリア ---
            InitControl();

            MessageBox.Show("画像とデータの削除が完了しました。",
                "画像削除完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>現在チェックされているオプションボタンのインデックスを返す</summary>
        private int GetCheckedIndex()
        {
            for (int i = 0; i < _optPicture.Length; i++)
                if (_optPicture[i] != null && _optPicture[i].Checked) return i;
            return 0;
        }

        // ================================================================
        // メニューボタン（戻る）
        // VB6: Flag_Disp_Picture = False; Unload form_disp_picture; form_menu.Visible = True
        // ================================================================
        private void BtnBack_Click(object sender, EventArgs e)
        {
            AppState.FlagDispPicture = false;
            DisposePrintImage();
            this.Close();
        }

        // ================================================================
        // オーダーメイドボタン
        // VB6: btn_order_Click
        //   商品画像選択時のみデータ抽出 →
        //   AppState（OrderSample_*）にセット →
        //   form_tedukuri に受け渡して画面遷移
        // ================================================================
        private void BtnOrder_Click(object sender, EventArgs e)
        {
            // AppState のオーダーサンプル情報を初期化
            AppState.OrderSamplePt900 = 0;
            AppState.OrderSampleK18 = 0;
            AppState.OrderSampleWgPg = 0;
            AppState.OrderSampleK10 = 0;
            AppState.OrderSampleCode = 0;
            AppState.OrderSamplePrice = 0;
            AppState.OrderSampleHinban = "";

            int idx = GetCheckedIndex();
            bool isSyouhin = (SyouhinDbIndex[idx] > 0); // 商品画像を選択中か

            if (isSyouhin)
            {
                // --- 品番を取得（label_15 から "品番：" を除去）---
                string hinbanStr = _lblField15 != null ? _lblField15.Text : "";
                if (!string.IsNullOrEmpty(hinbanStr))
                {
                    // VB6: Replace(wkSTR, "品番：", "") → InStr("  ＃") or InStr("  石のグレード")
                    hinbanStr = hinbanStr.Replace("品番：", "");
                    int pos = hinbanStr.IndexOf("  ＃", StringComparison.Ordinal);
                    if (pos >= 0) hinbanStr = hinbanStr.Substring(0, pos);
                    pos = hinbanStr.IndexOf("  石のグレード", StringComparison.Ordinal);
                    if (pos >= 0) hinbanStr = hinbanStr.Substring(0, pos);
                    AppState.OrderSampleHinban = hinbanStr.Trim();
                }

                // --- 加工コードを取得（label_18 の "-" 以降）---
                string codeStr = _lblField18 != null ? _lblField18.Text : "";
                int dashPos = codeStr.IndexOf('-');
                if (dashPos >= 0)
                {
                    int code;
                    if (int.TryParse(codeStr.Substring(dashPos + 1), out code))
                        AppState.OrderSampleCode = code;
                }

                // --- 税込価格を取得（Label_30 から "税込","¥","," を除去）---
                string priceStr = _lblField30 != null ? _lblField30.Text : "";
                if (!string.IsNullOrEmpty(priceStr))
                {
                    priceStr = priceStr.Replace("税込", "")
                                       .Replace(" ", "")
                                       .Replace("　", "")
                                       .Replace("\\", "")
                                       .Replace("¥", "")
                                       .Replace(",", "");
                    double price;
                    if (double.TryParse(priceStr, out price))
                        AppState.OrderSamplePrice = price;
                }

                // --- 地金重量を取得（label_22 で地金種別判定, label_27 で重量値）---
                string metalStr = _lblField22 != null ? _lblField22.Text : "";
                string weightStr = _lblField27 != null ? _lblField27.Text.Replace("g", "") : "0";
                double weight = 0;
                double.TryParse(weightStr, out weight);

                if (metalStr.Contains("Pt"))
                    AppState.OrderSamplePt900 = weight;
                else if (metalStr.Contains("WG") || metalStr.Contains("PG"))
                    AppState.OrderSampleWgPg = weight;
                else if (metalStr.Contains("K18"))
                    AppState.OrderSampleK18 = weight;
                else if (metalStr.Contains("K10"))
                    AppState.OrderSampleK10 = weight * 0.6; // VB6: wkSTR2 * 0.6
            }

            // --- オーダーメイドフォームへ画面遷移 ---
            // VB6: Unload form_disp_picture → form_tedukuri.Visible = True → フィールドに値セット
            var formTedukuri = new FormTedukuri();
            formTedukuri.SetSampleData(
                AppState.OrderSampleHinban,
                AppState.OrderSamplePt900,
                AppState.OrderSampleK18,
                AppState.OrderSampleWgPg,
                AppState.OrderSampleK10,
                AppState.OrderSampleCode,
                AppState.OrderSamplePrice);

            DisposePrintImage();
            AppState.FlagDispPicture = false;

            this.Hide();
            formTedukuri.FormClosed += (s, args) => this.Close();
            formTedukuri.Show();
        }
    }
}

    // ================================================================
    // FormTedukuri: オーダーメイドフォーム（スタブ）
    // VB6: form_tedukuri.frm に相当
    // 実装時に form_tedukuri.frm を読み込んで置き換えること
    // ================================================================
   