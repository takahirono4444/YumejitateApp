using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Windows.Forms;

namespace YumejitateApp
{
    /// <summary>
    /// 画像印刷フォーム (VB6: form_print.frm の移植)
    /// LEADTools OCX の代わりに System.Drawing + PrintDocument を使用。
    /// </summary>
    public class FormPrint : Form
    {
        // ----------------------------------------------------------------
        // 印刷レイアウトの種類（VB6 の複数プリンタ別サブルーチンを統合）
        // ----------------------------------------------------------------
        private enum PrintLayout
        {
            EpsonPx105,   // btn_print: Epson PX-105
            CanonIp100,   // btn_print_new: Canon iP100
            MimosaEpson,  // btn_print_2: ミモザ用紙 Epson
            MimosaCanon   // btn_print_2_ip: ミモザ用紙 Canon
        }

        // ----------------------------------------------------------------
        // フィールド
        // ----------------------------------------------------------------
        // array_lead[0..17]: CheckBox 番号 n → array_lead[n-1] = 画像ファイル相対パス
        private readonly string[] _arrayLead = new string[18]
        {
            @"\syouhin_1.jpg",  // Check1  → LEAD1
            @"\syouhin_2.jpg",  // Check2  → LEAD2
            @"\syouhin_3.jpg",  // Check3  → LEAD3
            @"\syouhin_4.jpg",  // Check4  → LEAD4  (初期非表示)
            @"\camera_1.jpg",   // Check5  → LEAD5
            @"\camera_2.jpg",   // Check6  → LEAD6
            @"\camera_3.jpg",   // Check7  → LEAD7
            @"\camera_4.jpg",   // Check8  → LEAD8  (初期非表示)
            @"\gousei_1.jpg",   // Check9  → LEAD9
            @"\gousei_2.jpg",   // Check10 → LEAD10
            @"\gousei_3.jpg",   // Check11 → LEAD11
            @"\gousei_4.jpg",   // Check12 → LEAD12 (初期非表示)
            @"\syouhin_5.jpg",  // Check13 → LEAD13 (初期非表示)
            @"\syouhin_6.jpg",  // Check14 → LEAD14 (初期非表示)
            @"\camera_5.jpg",   // Check15 → LEAD15 (初期非表示)
            @"\camera_6.jpg",   // Check16 → LEAD16 (初期非表示)
            @"\gousei_5.jpg",   // Check17 → LEAD17 (初期非表示)
            @"\gousei_6.jpg",   // Check18 → LEAD18 (初期非表示)
        };

        // 現在の印刷ジョブ用パラメータ
        private int _printCount;               // チェック済み画像枚数
        private string[] _printIndices;             // チェック済み Check 番号配列
        private PrintLayout _currentPrintLayout;    // 使用する印刷レイアウト

        // コントロール
        private PictureBox[] _pictureBoxes;         // LEAD1-LEAD18 相当 (index 0=LEAD1)
        private CheckBox[] _checkBoxes;           // Check1-Check18 (index 0=Check1)
        private RadioButton[] _radioOptions;        // Option1(0-3) 相当

        private Button _btnPrint;          // Epson PX-105 印刷
        private Button _btnPrintNew;       // Canon iP100 印刷
        private Button _btnPrint2;         // ミモザ用紙 Epson 印刷
        private Button _btnPrint2Ip;       // ミモザ用紙 Canon 印刷
        private Button _btn123;            // ①②③ 表示切替
        private Button _btn456;            // ④⑤⑥ 表示切替
        private Button _btnBack;           // メニューへ戻る

        private Label _lblTitle;
        private Label _lblPrintMode;
        private GroupBox _grpPrintMode;
        private Panel _panelImages;

        // チェックされた画像の色定義
        private static readonly Color ColorChecked = Color.FromArgb(0xFF, 0xC0, 0xC0); // &HFFC0C0 ピンク
        private static readonly Color ColorUnchecked = Color.FromArgb(0xD8, 0xFF, 0xFF); // &HD8FFFF 水色
        private static readonly Color ColorSelected = Color.FromArgb(0xFF, 0xC0, 0xC0); // 選択中ラジオ

        // ----------------------------------------------------------------
        // コンストラクタ
        // ----------------------------------------------------------------
        public FormPrint()
        {
            InitializeComponent();
            LoadPreviewImages();
        }

        // ----------------------------------------------------------------
        // InitializeComponent: フォームとコントロールの構築
        // ----------------------------------------------------------------
        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォーム基本設定
            this.Text = "画像印刷";
            this.BackColor = Color.White;
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;

            // ---- ラベル（タイトル） ----
            _lblTitle = new Label
            {
                Text = "画像印刷",
                Font = new Font("ＭＳ Ｐゴシック", 18f, FontStyle.Italic),
                AutoSize = true,
                Location = new Point(700, 10),
            };

            // ---- 印刷方法 GroupBox ----
            _grpPrintMode = new GroupBox
            {
                Text = "印刷方法",
                Font = new Font("メイリオ", 9f),
                Location = new Point(10, 10),
                Size = new Size(680, 60),
            };

            string[] optionLabels = { "４分割１画像", "４分割複数画像", "分割無し１画像", "２分割１画像" };
            _radioOptions = new RadioButton[4];
            for (int i = 0; i < 4; i++)
            {
                int idx = i; // capture
                _radioOptions[i] = new RadioButton
                {
                    Text = optionLabels[i],
                    Font = new Font("メイリオ", 9f),
                    Location = new Point(10 + i * 165, 20),
                    Size = new Size(160, 30),
                    BackColor = (i == 0) ? ColorSelected : ColorUnchecked,
                    Checked = (i == 0),
                };
                _radioOptions[i].CheckedChanged += (s, e) => UpdateRadioColors();
                _grpPrintMode.Controls.Add(_radioOptions[i]);
            }

            // ---- 画像プレビューパネル ----
            _panelImages = new Panel
            {
                Location = new Point(10, 80),
                Size = new Size(860, 750),
                AutoScroll = true,
                BackColor = Color.FromArgb(240, 240, 240),
            };

            // PictureBox(LEAD1-LEAD18) と CheckBox(Check1-Check18) の生成
            // 配置: 3列 × 3行（商品/カメラ/合成の3グループ × 3枚）= 9枚 + 拡張9枚
            _pictureBoxes = new PictureBox[18];
            _checkBoxes = new CheckBox[18];

            // 表示ラベル（行ヘッダー）
            string[] rowLabels = { "商品", "カメラ", "合成" };
            for (int r = 0; r < 3; r++)
            {
                var lbl = new Label
                {
                    Text = rowLabels[r],
                    Font = new Font("メイリオ", 9f, FontStyle.Bold),
                    Location = new Point(0, 10 + r * 250),
                    Size = new Size(55, 240),
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = Color.LightGray,
                };
                _panelImages.Controls.Add(lbl);
            }

            // 3行×3列 = 9枚（①②③グループ）
            // インデックス対応: Check1-3=商品①②③, Check5-7=カメラ①②③, Check9-11=合成①②③
            int[] group123CheckNums = { 1, 2, 3, 5, 6, 7, 9, 10, 11 };
            for (int idx = 0; idx < 9; idx++)
            {
                int checkNum = group123CheckNums[idx];  // 1-based Check 番号
                int leadIdx = checkNum - 1;             // 0-based Lead インデックス
                int col = idx % 3;                  // 0,1,2
                int row = idx / 3;                  // 0,1,2
                int pbX = 60 + col * 270;
                int pbY = 10 + row * 250;

                _pictureBoxes[leadIdx] = new PictureBox
                {
                    Location = new Point(pbX, pbY),
                    Size = new Size(240, 210),
                    SizeMode = PictureBoxSizeMode.Zoom,
                    BackColor = Color.Black,
                    BorderStyle = BorderStyle.FixedSingle,
                };

                string[] captions = { "商品①", "商品②", "商品③", "カメラ①", "カメラ②", "カメラ③", "合成①", "合成②", "合成③" };
                _checkBoxes[leadIdx] = new CheckBox
                {
                    Text = captions[idx],
                    Font = new Font("メイリオ", 8f),
                    Location = new Point(pbX, pbY + 212),
                    Size = new Size(240, 25),
                    BackColor = ColorUnchecked,
                };
                int cn = checkNum;
                _checkBoxes[leadIdx].CheckedChanged += (s, e) =>
                    ((CheckBox)s).BackColor = ((CheckBox)s).Checked ? ColorChecked : ColorUnchecked;

                _panelImages.Controls.Add(_pictureBoxes[leadIdx]);
                _panelImages.Controls.Add(_checkBoxes[leadIdx]);
            }

            // 拡張9枚（④⑤⑥グループ）: Check4,8,12,13,14,15,16,17,18
            // 初期非表示; btn_456 クリックで表示
            int[] group456CheckNums = { 4, 8, 12, 13, 14, 15, 16, 17, 18 };
            string[] captions456 = { "商品④", "カメラ④", "合成④", "商品⑤", "商品⑥", "カメラ⑤", "カメラ⑥", "合成⑤", "合成⑥" };
            for (int idx = 0; idx < 9; idx++)
            {
                int checkNum = group456CheckNums[idx];
                int leadIdx = checkNum - 1;
                int col = idx % 3;
                int row = idx / 3;
                int pbX = 60 + col * 270;
                int pbY = 10 + row * 250;

                _pictureBoxes[leadIdx] = new PictureBox
                {
                    Location = new Point(pbX, pbY),
                    Size = new Size(240, 210),
                    SizeMode = PictureBoxSizeMode.Zoom,
                    BackColor = Color.Black,
                    BorderStyle = BorderStyle.FixedSingle,
                    Visible = false,
                };

                _checkBoxes[leadIdx] = new CheckBox
                {
                    Text = captions456[idx],
                    Font = new Font("メイリオ", 8f),
                    Location = new Point(pbX, pbY + 212),
                    Size = new Size(240, 25),
                    BackColor = ColorUnchecked,
                    Visible = false,
                };
                _checkBoxes[leadIdx].CheckedChanged += (s, e) =>
                    ((CheckBox)s).BackColor = ((CheckBox)s).Checked ? ColorChecked : ColorUnchecked;

                _panelImages.Controls.Add(_pictureBoxes[leadIdx]);
                _panelImages.Controls.Add(_checkBoxes[leadIdx]);
            }

            // ---- ボタン類 ----
            int btnX = 880;
            int btnY = 80;
            int btnW = 160;
            int btnH = 45;
            int btnGap = 10;

            _btnPrint = new Button
            {
                Text = "Epson PX-105\n印刷",
                Font = new Font("メイリオ", 9f, FontStyle.Bold),
                Location = new Point(btnX, btnY),
                Size = new Size(btnW, btnH),
                BackColor = Color.FromArgb(192, 224, 255),
            };
            _btnPrint.Click += (s, e) => BtnPrint_Click(PrintLayout.EpsonPx105);

            _btnPrintNew = new Button
            {
                Text = "Canon iP100\n印刷",
                Font = new Font("メイリオ", 9f, FontStyle.Bold),
                Location = new Point(btnX, btnY + (btnH + btnGap)),
                Size = new Size(btnW, btnH),
                BackColor = Color.FromArgb(192, 224, 255),
            };
            _btnPrintNew.Click += (s, e) => BtnPrint_Click(PrintLayout.CanonIp100);

            _btnPrint2 = new Button
            {
                Text = "ミモザ用紙\nEpson印刷",
                Font = new Font("メイリオ", 9f, FontStyle.Bold),
                Location = new Point(btnX, btnY + 2 * (btnH + btnGap)),
                Size = new Size(btnW, btnH),
                BackColor = Color.FromArgb(192, 255, 224),
            };
            _btnPrint2.Click += (s, e) => BtnPrint_Click(PrintLayout.MimosaEpson);

            _btnPrint2Ip = new Button
            {
                Text = "ミモザ用紙\nCanon印刷",
                Font = new Font("メイリオ", 9f, FontStyle.Bold),
                Location = new Point(btnX, btnY + 3 * (btnH + btnGap)),
                Size = new Size(btnW, btnH),
                BackColor = Color.FromArgb(192, 255, 224),
            };
            _btnPrint2Ip.Click += (s, e) => BtnPrint_Click(PrintLayout.MimosaCanon);

            _btn123 = new Button
            {
                Text = "①②③",
                Font = new Font("メイリオ", 10f, FontStyle.Bold),
                Location = new Point(btnX, btnY + 4 * (btnH + btnGap) + 20),
                Size = new Size(btnW, btnH),
                BackColor = Color.LightYellow,
                Enabled = false,   // VB6: Form_Load で無効化
            };
            _btn123.Click += Btn123_Click;

            _btn456 = new Button
            {
                Text = "④⑤⑥",
                Font = new Font("メイリオ", 10f, FontStyle.Bold),
                Location = new Point(btnX, btnY + 5 * (btnH + btnGap) + 20),
                Size = new Size(btnW, btnH),
                BackColor = Color.LightYellow,
            };
            _btn456.Click += Btn456_Click;

            _btnBack = new Button
            {
                Text = "メニュー",
                Font = new Font("メイリオ", 10f, FontStyle.Bold),
                Location = new Point(btnX, btnY + 7 * (btnH + btnGap) + 20),
                Size = new Size(btnW, btnH),
                BackColor = Color.FromArgb(255, 200, 200),
            };
            _btnBack.Click += BtnBack_Click;

            // ---- コントロールをフォームに追加 ----
            this.Controls.Add(_lblTitle);
            this.Controls.Add(_grpPrintMode);
            this.Controls.Add(_panelImages);
            this.Controls.Add(_btnPrint);
            this.Controls.Add(_btnPrintNew);
            this.Controls.Add(_btnPrint2);
            this.Controls.Add(_btnPrint2Ip);
            this.Controls.Add(_btn123);
            this.Controls.Add(_btn456);
            this.Controls.Add(_btnBack);

            this.ResumeLayout(false);
        }

        // ----------------------------------------------------------------
        // フォームロード後の処理: プレビュー画像を PictureBox に読み込む
        // ----------------------------------------------------------------
        private void LoadPreviewImages()
        {
            string appPath = Application.StartupPath;
            for (int i = 0; i < 18; i++)
            {
                string path = appPath + _arrayLead[i];
                if (_pictureBoxes[i] != null && File.Exists(path))
                {
                    try
                    {
                        // ファイルロックを避けるためコピーして読み込む
                        using (var tmp = Image.FromFile(path))
                        {
                            _pictureBoxes[i].Image = new Bitmap(tmp);
                        }
                    }
                    catch
                    {
                        // 画像が読み込めない場合は無視
                    }
                }
            }
        }

        // ----------------------------------------------------------------
        // ラジオボタン色の更新 (VB6: Option1_Click)
        // ----------------------------------------------------------------
        private void UpdateRadioColors()
        {
            for (int i = 0; i < 4; i++)
            {
                _radioOptions[i].BackColor = _radioOptions[i].Checked ? ColorSelected : ColorUnchecked;
            }
        }

        // ----------------------------------------------------------------
        // 印刷モードの取得 (選択中の RadioButton インデックス)
        // ----------------------------------------------------------------
        private int GetSelectedOption()
        {
            for (int i = 0; i < 4; i++)
                if (_radioOptions[i].Checked) return i;
            return 0;
        }

        // ----------------------------------------------------------------
        // 印刷ボタン共通ハンドラ (VB6: btn_print_Click / btn_print_new_Click etc.)
        // ----------------------------------------------------------------
        private void BtnPrint_Click(PrintLayout layout)
        {
            // チェック済み画像カウント（1回目のパス）
            int checkCnt = 0;
            for (int i = 0; i < 18; i++)
                if (_checkBoxes[i] != null && _checkBoxes[i].Checked)
                    checkCnt++;

            if (checkCnt == 0)
            {
                MessageBox.Show("印刷する画像をチェックして下さい。", "画像印刷チェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 印刷モード別チェック数バリデーション
            int opt = GetSelectedOption();
            if ((opt == 0 || opt == 2 || opt == 3) && checkCnt > 1)
            {
                string modeName = opt == 0 ? "４分割１画像" : opt == 2 ? "分割無し１画像" : "２分割１画像";
                MessageBox.Show($"{modeName}で印刷する場合、チェックできる画像は１つだけです。",
                    "画像印刷チェック", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (opt == 1 && checkCnt > 4)
            {
                MessageBox.Show("４分割複数画像で印刷する場合、チェックできる画像は４つまでです。",
                    "画像印刷チェック", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 印刷確認ダイアログ
            string[] modeNames = { "４分割１画像", "４分割複数画像", "分割無し１画像", "２分割１画像" };
            var confirm = MessageBox.Show(
                $"チェックした画像を{modeNames[opt]}で印刷します。よろしいですか？",
                "画像印刷確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm == DialogResult.No) return;

            // チェック済み画像インデックスを収集（2回目のパス: Check 番号 1-18）
            _printIndices = new string[4];
            _printCount = 0;
            for (int i = 0; i < 18; i++)
            {
                if (_checkBoxes[i] != null && _checkBoxes[i].Checked)
                {
                    if (_printCount < 4)
                        _printIndices[_printCount] = (i + 1).ToString(); // Check番号(1-based)
                    _printCount++;
                }
            }

            _currentPrintLayout = layout;

            // PrintDocument を構成して印刷
            using (var pd = new PrintDocument())
            {
                // プリンタ選択ダイアログ
                using (var dlg = new PrintDialog { Document = pd, UseEXDialog = true })
                {
                    if (dlg.ShowDialog() != DialogResult.OK) return;
                }

                pd.PrintPage += PrintDocument_PrintPage;
                try
                {
                    pd.Print();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("印刷エラーが発生しました。\n" + ex.Message,
                        "印刷エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // ----------------------------------------------------------------
        // PrintDocument.PrintPage ハンドラ: 全印刷モードを処理
        // ----------------------------------------------------------------
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            int opt = GetSelectedOption();

            switch (opt)
            {
                case 0: // ４分割１画像: 1枚の画像を2×2=4コマに同じ画像を印刷
                    PrintMode_4Div1Image(e);
                    break;
                case 1: // ４分割複数画像: 最大4枚を2×2に配置
                    PrintMode_4DivMultiple(e);
                    break;
                case 2: // 分割無し１画像: 1枚を大きく印刷
                    PrintMode_NoDiv1Image(e);
                    break;
                case 3: // ２分割１画像: 1枚を左右2コマに同じ画像を印刷
                    PrintMode_2Div1Image(e);
                    break;
            }

            e.HasMorePages = false;
        }

        // ----------------------------------------------------------------
        // 印刷ヘルパー: 画像ファイルを読み込み、90°回転して返す
        // VB6: LEAD0.Load + LEAD0.Rotate 9000
        // ----------------------------------------------------------------
        private Image LoadAndRotateImage(string checkNumStr)
        {
            if (!int.TryParse(checkNumStr, out int checkNum)) return null;
            int leadIdx = checkNum - 1; // 0-based
            if (leadIdx < 0 || leadIdx >= 18) return null;

            string path = Application.StartupPath + _arrayLead[leadIdx];
            if (!File.Exists(path)) return null;

            try
            {
                var img = Image.FromFile(path);
                img.RotateFlip(RotateFlipType.Rotate90FlipNone); // VB6: Rotate 9000
                return img;
            }
            catch
            {
                return null;
            }
        }

        // ----------------------------------------------------------------
        // アスペクト比を維持して矩形を計算するヘルパー
        // ----------------------------------------------------------------
        private static RectangleF CalcAspectRect(float destH, float imgW, float imgH, float left, float top)
        {
            float destW = (imgH > 0) ? (destH * imgW / imgH) : destH;
            return new RectangleF(left, top, destW, destH);
        }

        // ----------------------------------------------------------------
        // 印刷モード: ４分割１画像 (VB6: Print_LEAD_1X4_*)
        // 1枚の画像を2行×2列の4コマに同一画像を配置
        // ----------------------------------------------------------------
        private void PrintMode_4Div1Image(PrintPageEventArgs e)
        {
            if (_printCount == 0 || _printIndices[0] == null) return;

            using (Image img = LoadAndRotateImage(_printIndices[0]))
            {
                if (img == null) return;

                RectangleF[] cells = GetCells4Div(e, _currentPrintLayout);
                float imgW = img.Width, imgH = img.Height;

                // 4コマすべてに同じ画像を描画
                foreach (var cell in cells)
                {
                    var rect = CalcAspectRect(cell.Height, imgW, imgH, cell.X, cell.Y);
                    e.Graphics.DrawImage(img, rect);
                }
            }
        }

        // ----------------------------------------------------------------
        // 印刷モード: ４分割複数画像 (VB6: Print_LEAD_X_*)
        // チェックされた最大4枚を2行×2列に個別配置
        // ----------------------------------------------------------------
        private void PrintMode_4DivMultiple(PrintPageEventArgs e)
        {
            RectangleF[] cells = GetCells4Div(e, _currentPrintLayout);

            for (int i = 0; i < Math.Min(_printCount, 4); i++)
            {
                using (Image img = LoadAndRotateImage(_printIndices[i]))
                {
                    if (img == null) continue;
                    var rect = CalcAspectRect(cells[i].Height, img.Width, img.Height,
                        cells[i].X, cells[i].Y);
                    e.Graphics.DrawImage(img, rect);
                }
            }
        }

        // ----------------------------------------------------------------
        // 印刷モード: 分割無し１画像 (VB6: Print_LEAD_1X1_*)
        // 1枚の画像をページいっぱいに大きく印刷
        // ----------------------------------------------------------------
        private void PrintMode_NoDiv1Image(PrintPageEventArgs e)
        {
            if (_printCount == 0 || _printIndices[0] == null) return;

            using (Image img = LoadAndRotateImage(_printIndices[0]))
            {
                if (img == null) return;

                float pageW = e.PageBounds.Width;
                float pageH = e.PageBounds.Height;
                float marginX = pageW * 0.05f;
                float marginY = pageH * 0.05f;
                float availW = pageW - 2 * marginX;
                float availH = pageH - 2 * marginY;

                // アスペクト比維持で最大サイズに
                float scaleW = availW / img.Width;
                float scaleH = availH / img.Height;
                float scale = Math.Min(scaleW, scaleH);
                float dstW = img.Width * scale;
                float dstH = img.Height * scale;
                float dstX = marginX + (availW - dstW) / 2f;
                float dstY = marginY + (availH - dstH) / 2f;

                e.Graphics.DrawImage(img, new RectangleF(dstX, dstY, dstW, dstH));
            }
        }

        // ----------------------------------------------------------------
        // 印刷モード: ２分割１画像 (VB6: Print_LEAD_2X1_*)
        // 1枚の画像を左右2コマに同一画像を配置（上段のみ）
        // ----------------------------------------------------------------
        private void PrintMode_2Div1Image(PrintPageEventArgs e)
        {
            if (_printCount == 0 || _printIndices[0] == null) return;

            using (Image img = LoadAndRotateImage(_printIndices[0]))
            {
                if (img == null) return;

                // 上段2コマのセルを取得
                RectangleF[] cells = GetCells4Div(e, _currentPrintLayout);
                float imgW = img.Width, imgH = img.Height;

                // セル[0]=左上, セル[1]=右上 の2コマのみ描画
                for (int i = 0; i < 2; i++)
                {
                    var rect = CalcAspectRect(cells[i].Height, imgW, imgH, cells[i].X, cells[i].Y);
                    e.Graphics.DrawImage(img, rect);
                }
            }
        }

        // ----------------------------------------------------------------
        // 4分割レイアウトのセル矩形を返す [左上, 右上, 左下, 右下]
        // VB6の各プリンタ別座標をページ比率に変換して統合
        // ----------------------------------------------------------------
        private static RectangleF[] GetCells4Div(PrintPageEventArgs e, PrintLayout layout)
        {
            float pageW = e.PageBounds.Width;
            float pageH = e.PageBounds.Height;

            // レイアウト別マージン・比率パラメータ
            // VB6の印刷座標(pixel)から比率に変換:
            //   epson/mimosa_epson: PrintLeft=400/2900, TopOffset=1000-2000
            //   canon: PrintLeft=210/2850, TopOffset=1200
            //   mimosa_canon: PrintLeft=420/2800, height=2850(fixed pixel)
            //
            // 各プリンタで左カラム位置 ≈ ページ幅の7-10%
            //          右カラム位置 ≈ ページ幅の49-52%
            //          上マージン  ≈ ページ高の7-12%
            //          セル高さ   ≈ ページ高の40%

            float col1X, col2X, topY, cellH;

            switch (layout)
            {
                case PrintLayout.EpsonPx105:
                    col1X = pageW * 0.07f;
                    col2X = pageW * 0.50f;
                    topY = pageH * 0.07f;
                    cellH = pageH * 0.40f;
                    break;
                case PrintLayout.CanonIp100:
                    col1X = pageW * 0.04f;
                    col2X = pageW * 0.50f;
                    topY = pageH * 0.09f;
                    cellH = pageH * 0.40f;
                    break;
                case PrintLayout.MimosaEpson:
                    col1X = pageW * 0.07f;
                    col2X = pageW * 0.50f;
                    topY = pageH * 0.12f;
                    cellH = pageH * 0.38f;
                    break;
                case PrintLayout.MimosaCanon:
                default:
                    col1X = pageW * 0.07f;
                    col2X = pageW * 0.49f;
                    topY = pageH * 0.12f;
                    cellH = pageH * 0.38f;
                    break;
            }

            float rowGap = pageH * 0.03f;
            float row2Y = topY + cellH + rowGap;

            return new RectangleF[]
            {
                new RectangleF(col1X, topY,  cellH * 1.33f, cellH), // 左上 (画像幅は高さ×4:3想定)
                new RectangleF(col2X, topY,  cellH * 1.33f, cellH), // 右上
                new RectangleF(col1X, row2Y, cellH * 1.33f, cellH), // 左下
                new RectangleF(col2X, row2Y, cellH * 1.33f, cellH), // 右下
            };
        }

        // ----------------------------------------------------------------
        // btn_123_Click: ①②③グループ表示 (VB6: btn_123_Click)
        // ----------------------------------------------------------------
        private void Btn123_Click(object sender, EventArgs e)
        {
            // ①②③グループ: Check1-3,5-7,9-11 → PictureBox/CheckBox インデックス 0-2,4-6,8-10
            int[] show123 = { 0, 1, 2, 4, 5, 6, 8, 9, 10 };
            // ④⑤⑥グループ: Check4,8,12,13,14,15,16,17,18 → idx 3,7,11,12,13,14,15,16,17
            int[] hide456 = { 3, 7, 11, 12, 13, 14, 15, 16, 17 };

            foreach (int i in show123)
            {
                if (_pictureBoxes[i] != null) _pictureBoxes[i].Visible = true;
                if (_checkBoxes[i] != null) _checkBoxes[i].Visible = true;
            }
            foreach (int i in hide456)
            {
                if (_pictureBoxes[i] != null) _pictureBoxes[i].Visible = false;
                if (_checkBoxes[i] != null) _checkBoxes[i].Visible = false;
            }

            _btn123.Enabled = false;
            _btn456.Enabled = true;
        }

        // ----------------------------------------------------------------
        // btn_456_Click: ④⑤⑥グループ表示 (VB6: btn_456_Click)
        // ----------------------------------------------------------------
        private void Btn456_Click(object sender, EventArgs e)
        {
            int[] hide123 = { 0, 1, 2, 4, 5, 6, 8, 9, 10 };
            int[] show456 = { 3, 7, 11, 12, 13, 14, 15, 16, 17 };

            foreach (int i in hide123)
            {
                if (_pictureBoxes[i] != null) _pictureBoxes[i].Visible = false;
                if (_checkBoxes[i] != null) _checkBoxes[i].Visible = false;
            }
            foreach (int i in show456)
            {
                if (_pictureBoxes[i] != null) _pictureBoxes[i].Visible = true;
                if (_checkBoxes[i] != null) _checkBoxes[i].Visible = true;
            }

            _btn123.Enabled = true;
            _btn456.Enabled = false;
        }

        // ----------------------------------------------------------------
        // btn_back_Click: メニューへ戻る (VB6: btn_back_Click)
        // VB6: Unload Form_Print → form_menu.Visible = True
        // ----------------------------------------------------------------
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
        // フォームクローズ時のリソース解放
        // ----------------------------------------------------------------
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);

            // PictureBox の画像を解放
            if (_pictureBoxes != null)
            {
                foreach (var pb in _pictureBoxes)
                {
                    if (pb?.Image != null)
                    {
                        pb.Image.Dispose();
                        pb.Image = null;
                    }
                }
            }
        }
    }
}
