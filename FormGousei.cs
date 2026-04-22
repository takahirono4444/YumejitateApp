// ================================================================
// FormGousei.cs
// VB6: form_gousei.frm → C# + WinForms 移行
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.IO;
using System.Windows.Forms;

// ================================================================
// FormGousei.cs
// VB6: form_gousei.frm → C# + WinForms 移行
// 機能: 画像合成画面
//       サブ画像の範囲を切り取り、メイン画像へドラッグ合成する。
//       回転・ズーム・貼り付けモード・合成履歴・印刷・保存に対応。
//
// LEADTools 置き換え:
//   LEAD1（サブ画像）  → PictureBox + ラバーバンド選択（MouseDown/Move/Up）
//   LEAD2（メイン画像）→ GouseiCanvas（カスタム Panel：floater overlay 描画）
//   LEAD3（切り取りバッファ）→ Image _cropImage
//   LEAD4（印刷バッファ）    → Image _printImage（RotateFlip後）
//   LEAD5（ズーム作業）      → Image _workImage
//   LEAD2.Combine()          → Graphics.DrawImage()
//   LEAD3.Rotate()           → RotateFlip() + Matrix
//   LEAD3.Size()             → new Bitmap(w,h) + DrawImage
//   Floater ドラッグ         → MouseDown/Move/Up + _floaterRect
//
// 対象フレームワーク: .NET Framework 4.8
// C#バージョン: 7.3
// ================================================================

namespace YumejitateApp
{
    public class FormGousei : Form
    {
        // ----------------------------------------------------------------
        // 画像ファイル名テーブル（VB6: Option_ld1(Index), Option_ld2(Index)）
        // 0-3  : syouhin_1-4.jpg
        // 4-7  : camera_1-4.jpg
        // 8-11 : gousei_1-4.jpg
        // 12-13: syouhin_5-6.jpg
        // 14-15: camera_5-6.jpg
        // 16-17: gousei_5-6.jpg
        // ----------------------------------------------------------------
        private static readonly string[] ImageFiles = new string[]
        {
            "syouhin_1.jpg", "syouhin_2.jpg", "syouhin_3.jpg", "syouhin_4.jpg",
            "camera_1.jpg",  "camera_2.jpg",  "camera_3.jpg",  "camera_4.jpg",
            "gousei_1.jpg",  "gousei_2.jpg",  "gousei_3.jpg",  "gousei_4.jpg",
            "syouhin_5.jpg", "syouhin_6.jpg",
            "camera_5.jpg",  "camera_6.jpg",
            "gousei_5.jpg",  "gousei_6.jpg",
        };

        // ----------------------------------------------------------------
        // 画像バッファ（LEADTools の各コントロールに相当）
        // ----------------------------------------------------------------
        private Image _subImage;    // LEAD1: サブ画像（選択元）
        private Image _mainImage;   // LEAD2: メイン画像（合成先）
        private Image _cropImage;   // LEAD3: 切り取り画像バッファ
        private Image _printImage;  // LEAD4: 印刷バッファ（90°回転済み）
        private Image _workImage;   // LEAD5: ズーム作業バッファ

        // ----------------------------------------------------------------
        // ラバーバンド選択（LEAD1 のリージョン選択に相当）
        // ----------------------------------------------------------------
        private bool _selecting;     // マウスドラッグ中
        private Point _selStart;     // 選択開始座標（サブ画像コントロール座標）
        private Point _selEnd;       // 選択終了座標
        private bool _useEllipse;    // True=楕円選択（Option_camera(1)）/ False=矩形

        // ----------------------------------------------------------------
        // フロータ（LEAD2.Floater に相当：ドラッグ合成用オーバーレイ）
        // ----------------------------------------------------------------
        private bool _floaterVisible; // FloaterVisible
        private RectangleF _floaterRect; // SetFloaterDstRect の内容
        private bool _readyToDrag;   // ReadyToDrag
        private PointF _dragStart;   // drag_StartX/Y
        private PointF _floaterPos;  // FloaterX/FloaterY

        // ----------------------------------------------------------------
        // 貼り付けモードフラグ（VB6: PasteFlag / PasteFlag2 / PasteFlag3）
        // ----------------------------------------------------------------
        private bool _pasteFlag;     // フロータ表示中（ドラッグ可能状態）
        private bool _pasteFlag2;    // 指定貼り付けモード（クリック位置に中心配置）
        private bool _pasteFlag3;    // 選択貼り付けモード（矩形/楕円ドラッグで配置）

        // PasteFlag3 用: MouseDown で開始点を記録
        private PointF _paste3Start;

        // ----------------------------------------------------------------
        // 回転フラグ（btn_kaiten → btn_zoom の再適用に使う）
        // ----------------------------------------------------------------
        private bool _kaiten;

        // ----------------------------------------------------------------
        // 合成履歴（VB6: gousei_max, gousei_current）
        // gousei\ フォルダに gousei_1.jpg, gousei_2.jpg, ... を保存
        // ----------------------------------------------------------------
        private int _gouseiMax;
        private int _gouseiCurrent;

        // ----------------------------------------------------------------
        // UI コントロール
        // ----------------------------------------------------------------

        // サブ画像表示（LEAD1 相当）
        private PictureBox _picSub;
        // メイン画像合成キャンバス（LEAD2 相当）
        private GouseiCanvas _canvasMain;

        // オプションボタン群
        private RadioButton[] _optLd1 = new RadioButton[18]; // サブ画像選択
        private RadioButton[] _optLd2 = new RadioButton[18]; // メイン画像選択
        private RadioButton[] _optCamera = new RadioButton[2];  // 四角/楕円切り取り
        private RadioButton[] _optKaiten = new RadioButton[2];  // 左/右回転方向
        private RadioButton[] _optMainSentaku = new RadioButton[2];  // 選択貼り付け形状
        private RadioButton[] _optZoom = new RadioButton[2];  // 縮小/拡大

        // 回転・ズームコンボ（VB6: 百の位・十の位・一の位を結合して角度/率を得る）
        private ComboBox _cmbKaiten1, _cmbKaiten2, _cmbKaiten3; // 0-3/0-9/0-9 → 例"090"→90°
        private ComboBox _cmbZoom1, _cmbZoom2, _cmbZoom3;   // 1-9/0-9/0-9 → 例"100"→100%

        // ボタン
        private Button _btnKaiten;        // 回転実行
        private Button _btnZoom;          // ズーム実行
        private Button _btnInitKiritori;  // 切り取り初期化
        private Button _btnPaste2;        // 指定貼り付け
        private Button _btnPaste3;        // 選択貼り付け
        private Button _btnFix;           // 固定（合成確定）
        private Button _btnInitSyouhin;   // メイン画像リセット
        private Button _btnSaveJpeg;      // gousei.jpg 保存
        private Button _btnPrint;         // 印刷
        private Button _btnBack;          // メニューへ戻る
        private Button _btnLeft;          // 履歴：前へ
        private Button _btnRight;         // 履歴：次へ

        // ================================================================
        // コンストラクタ
        // ================================================================
        public FormGousei()
        {
            InitializeComponent();
        }

        // ================================================================
        // フォームデザイン初期化
        // ================================================================
        private void InitializeComponent()
        {
            this.Text = "夢仕立て-画像合成";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(1200, 820);
            this.Font = new Font("メイリオ", 9F);
            this.Load += new EventHandler(FormGousei_Load);
            this.FormClosing += new FormClosingEventHandler(FormGousei_FormClosing);

            // ============================================================
            // 上部ツールエリア（左列: サブ画像, 右列: メイン画像の操作）
            // ============================================================

            // --- サブ画像表示（LEAD1 相当）---
            // VB6: Left=240 Top=1320 Width=5295 Height=3615 twips ≈ 350×240 px
            _picSub = new PictureBox();
            _picSub.Location = new Point(10, 160);
            _picSub.Size = new Size(370, 260);
            _picSub.BackColor = Color.Black;
            _picSub.BorderStyle = BorderStyle.FixedSingle;
            _picSub.SizeMode = PictureBoxSizeMode.Zoom;
            _picSub.Paint += new PaintEventHandler(PicSub_Paint);
            _picSub.MouseDown += new MouseEventHandler(PicSub_MouseDown);
            _picSub.MouseMove += new MouseEventHandler(PicSub_MouseMove);
            _picSub.MouseUp += new MouseEventHandler(PicSub_MouseUp);
            this.Controls.Add(_picSub);

            // --- メイン画像合成キャンバス（LEAD2 相当）---
            // VB6: Left=5640 Top=1320 Width=9615 Height=6615 twips ≈ 640×440 px
            _canvasMain = new GouseiCanvas(this);
            _canvasMain.Location = new Point(390, 160);
            _canvasMain.Size = new Size(790, 550);
            _canvasMain.BackColor = Color.Black;
            _canvasMain.BorderStyle = BorderStyle.FixedSingle;
            _canvasMain.MouseDown += new MouseEventHandler(Canvas_MouseDown);
            _canvasMain.MouseMove += new MouseEventHandler(Canvas_MouseMove);
            _canvasMain.MouseUp += new MouseEventHandler(Canvas_MouseUp);
            this.Controls.Add(_canvasMain);

            // ============================================================
            // 左コントロールパネル（サブ画像操作）
            // ============================================================
            int lx = 10, ly = 10;

            // --- Option_ld1: サブ画像選択（18個）---
            var lblSub = NewLabel("【サブ画像】", lx, ly, 180, bold: true); ly += 20;
            string[] subLabels = BuildImageLabels();
            for (int i = 0; i < 18; i++)
            {
                _optLd1[i] = new RadioButton();
                _optLd1[i].Text = subLabels[i];
                _optLd1[i].Location = new Point(lx + (i % 3) * 125, ly + (i / 3) * 20);
                _optLd1[i].Size = new Size(120, 20);
                _optLd1[i].BackColor = Color.Transparent;
                _optLd1[i].Tag = i;
                _optLd1[i].CheckedChanged += new EventHandler(OptLd1_CheckedChanged);
                this.Controls.Add(_optLd1[i]);
            }
            ly += (18 / 3) * 20 + 5;

            // --- Option_camera: 切り取り形状（四角/楕円）---
            var lblCam = NewLabel("切取形状:", lx, ly, 80);
            _optCamera[0] = NewRadio("四角", lx + 70, ly, 60); _optCamera[0].Checked = true;
            _optCamera[1] = NewRadio("楕円", lx + 130, ly, 60);
            for (int i = 0; i < 2; i++) { _optCamera[i].CheckedChanged += new EventHandler(OptCamera_CheckedChanged); this.Controls.Add(_optCamera[i]); }
            this.Controls.Add(lblCam);
            ly += 22;

            // --- 回転コントロール ---
            var lblKa = NewLabel("回転角:", lx, ly, 65);
            _cmbKaiten1 = NewDigitCombo(lx + 65, ly, 40, 0, 3);  // 百の位
            _cmbKaiten2 = NewDigitCombo(lx + 107, ly, 40, 0, 9); // 十の位
            _cmbKaiten3 = NewDigitCombo(lx + 149, ly, 40, 0, 9); // 一の位
            this.Controls.Add(lblKa);
            ly += 22;
            _optKaiten[0] = NewRadio("左回転", lx, ly, 70); _optKaiten[0].Checked = true;
            _optKaiten[1] = NewRadio("右回転", lx + 75, ly, 70);
            for (int i = 0; i < 2; i++) { _optKaiten[i].CheckedChanged += new EventHandler(OptKaiten_CheckedChanged); this.Controls.Add(_optKaiten[i]); }
            _btnKaiten = NewButton("回転", lx + 155, ly - 2, 60, 20);
            _btnKaiten.Click += new EventHandler(BtnKaiten_Click);
            ly += 25;

            // --- ズームコントロール ---
            var lblZm = NewLabel("ズーム%:", lx, ly, 65);
            _cmbZoom1 = NewDigitCombo(lx + 65, ly, 40, 1, 9);   // 百の位（1始まり）
            _cmbZoom2 = NewDigitCombo(lx + 107, ly, 40, 0, 9);  // 十の位
            _cmbZoom3 = NewDigitCombo(lx + 149, ly, 40, 0, 9);  // 一の位
            this.Controls.Add(lblZm);
            ly += 22;
            _optZoom[0] = NewRadio("縮小", lx, ly, 60); _optZoom[0].Checked = true;
            _optZoom[1] = NewRadio("拡大", lx + 65, ly, 60);
            for (int i = 0; i < 2; i++) { _optZoom[i].CheckedChanged += new EventHandler(OptZoom_CheckedChanged); this.Controls.Add(_optZoom[i]); }
            _btnZoom = NewButton("ズーム", lx + 135, ly - 2, 65, 20);
            _btnZoom.Click += new EventHandler(BtnZoom_Click);
            ly += 25;

            // --- 切り取り操作ボタン ---
            _btnInitKiritori = NewButton("切取初期化", lx, ly, 90, 22);
            _btnInitKiritori.Click += new EventHandler(BtnInitKiritori_Click);
            ly += 28;

            // ============================================================
            // 右コントロールパネル（メイン画像操作）
            // ============================================================
            int rx = 395, ry = 10;

            // --- Option_ld2: メイン画像選択（18個）---
            var lblMain = NewLabel("【メイン画像】", rx, ry, 180, bold: true); ry += 20;
            for (int i = 0; i < 18; i++)
            {
                _optLd2[i] = new RadioButton();
                _optLd2[i].Text = subLabels[i];
                _optLd2[i].Location = new Point(rx + (i % 6) * 130, ry + (i / 6) * 20);
                _optLd2[i].Size = new Size(125, 20);
                _optLd2[i].BackColor = Color.Transparent;
                _optLd2[i].Tag = i;
                _optLd2[i].CheckedChanged += new EventHandler(OptLd2_CheckedChanged);
                this.Controls.Add(_optLd2[i]);
            }
            ry += (18 / 6) * 20 + 5;

            // --- Option_main_sentaku: 選択貼り付け形状 ---
            var lblMs = NewLabel("選択形状:", rx, ry, 75);
            _optMainSentaku[0] = NewRadio("四角", rx + 80, ry, 60); _optMainSentaku[0].Checked = true;
            _optMainSentaku[1] = NewRadio("楕円", rx + 145, ry, 60);
            for (int i = 0; i < 2; i++) { _optMainSentaku[i].CheckedChanged += new EventHandler(OptMainSentaku_CheckedChanged); this.Controls.Add(_optMainSentaku[i]); }
            this.Controls.Add(lblMs);
            ry += 22;

            // --- 貼り付け・固定ボタン ---
            _btnPaste2 = NewButton("指定貼り付け", rx, ry, 100, 22);
            _btnPaste2.Click += new EventHandler(BtnPaste2_Click);
            _btnPaste3 = NewButton("選択貼り付け", rx + 108, ry, 100, 22);
            _btnPaste3.Click += new EventHandler(BtnPaste3_Click);
            _btnFix = NewButton("固  定", rx + 216, ry, 80, 22);
            _btnFix.BackColor = Color.FromArgb(0xFF, 0xFF, 0xC0);
            _btnFix.Click += new EventHandler(BtnFix_Click);
            ry += 28;

            // --- メイン画像リセット ---
            _btnInitSyouhin = NewButton("元に戻す", rx, ry, 90, 22);
            _btnInitSyouhin.Click += new EventHandler(BtnInitSyouhin_Click);

            // --- 履歴ボタン ---
            _btnLeft = NewButton("≪", rx + 100, ry, 50, 22);
            _btnLeft.Click += new EventHandler(BtnLeft_Click);
            _btnRight = NewButton("≫", rx + 157, ry, 50, 22);
            _btnRight.Click += new EventHandler(BtnRight_Click);

            // --- 保存・印刷・戻るボタン ---
            _btnSaveJpeg = NewButton("画像保存", rx + 218, ry, 80, 22);
            _btnSaveJpeg.BackColor = Color.LightGreen;
            _btnSaveJpeg.Click += new EventHandler(BtnSaveJpeg_Click);
            ry += 28;

            _btnPrint = NewButton("印　刷", rx, ry, 80, 22);
            _btnPrint.Click += new EventHandler(BtnPrint_Click);
            _btnBack = NewButton("メニュー", rx + 90, ry, 80, 22);
            _btnBack.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            _btnBack.Click += new EventHandler(BtnBack_Click);
        }

        // ================================================================
        // Form_Load
        // VB6: camera_1.jpg→LEAD1 / syouhin_1.jpg→LEAD2 / コンボ初期化 / Resize
        // ================================================================
        private void FormGousei_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;

            // 初期画像読み込み
            LoadSubImage(4);   // camera_1.jpg (index=4)
            LoadMainImage(0);  // syouhin_1.jpg (index=0)

            // オプションボタン初期選択色を設定
            if (_optLd1.Length > 4 && _optLd1[4] != null) { _optLd1[4].Checked = true; }
            if (_optLd2.Length > 0 && _optLd2[0] != null) { _optLd2[0].Checked = true; }

            // コンボボックス初期値設定
            _cmbKaiten1.SelectedIndex = 0;
            _cmbKaiten2.SelectedIndex = 0;
            _cmbKaiten3.SelectedIndex = 0;
            _cmbZoom1.SelectedIndex = 0;
            _cmbZoom2.SelectedIndex = 0;
            _cmbZoom3.SelectedIndex = 0;

            // 状態フラグ初期化
            _kaiten = false;
            _pasteFlag = false;
            _pasteFlag2 = false;
            _pasteFlag3 = false;

            // 合成履歴フォルダ確認・初期化
            string histDir = GouseiHistoryPath();
            Directory.CreateDirectory(histDir);
            string src = GetImagePath(0);
            string dst = GetHistoryPath(1);
            if (File.Exists(src)) File.Copy(src, dst, true);

            _gouseiMax = 1;
            _gouseiCurrent = 1;
            _btnLeft.Enabled = false;
            _btnRight.Enabled = false;
        }

        private void FormGousei_FormClosing(object sender, FormClosingEventArgs e)
        {
            DisposeImages();
        }

        // ================================================================
        // サブ画像の読み込み・リサイズ
        // VB6: LEAD1.Load + Resize_LEAD1（5295×3615 twips に収まるよう縮小）
        // C# : Image.FromFile → FitImageToBox に相当のサイズで _subImage に保持
        // ================================================================
        private void LoadSubImage(int index)
        {
            string path = GetImagePath(index);
            if (!File.Exists(path)) return;
            var prev = _subImage;
            _subImage = FitImageToSize(Image.FromFile(path), _picSub.Width, _picSub.Height);
            if (prev != null) prev.Dispose();
            _picSub.Image = _subImage;
            _cropImage = null; // 切り取りバッファもリセット
        }

        // ================================================================
        // メイン画像の読み込み・リサイズ
        // VB6: LEAD2.Load + Resize_LEAD2（9615×6615 twips に収まるよう縮小）
        // C# : Image.FromFile → FitImageToBox
        // ================================================================
        private void LoadMainImage(int index)
        {
            string path = GetImagePath(index);
            if (!File.Exists(path)) return;
            var prev = _mainImage;
            _mainImage = FitImageToSize(Image.FromFile(path), _canvasMain.Width, _canvasMain.Height);
            if (prev != null) prev.Dispose();
            // FloaterVisible = False 相当
            _floaterVisible = false;
            _pasteFlag = false;
            _pasteFlag2 = false;
            _pasteFlag3 = false;
            _canvasMain.Invalidate();
        }

        // ================================================================
        // 画像を指定サイズに収まるようリサイズ（アスペクト比維持）
        // VB6: LEAD.Size ImageWidth, ImageHeight, RESIZE_RESAMPLE 相当
        // ================================================================
        private Image FitImageToSize(Image src, int maxW, int maxH)
        {
            if (src.Width <= maxW && src.Height <= maxH) return src;

            double sx = (double)maxW / src.Width;
            double sy = (double)maxH / src.Height;
            double scale = Math.Min(sx, sy);
            int w = Math.Max(1, (int)(src.Width * scale));
            int h = Math.Max(1, (int)(src.Height * scale));

            var bmp = new Bitmap(w, h);
            using (var g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(src, 0, 0, w, h);
            }
            src.Dispose();
            return bmp;
        }

        // ================================================================
        // サブ画像のラバーバンド描画（Paint イベント）
        // VB6: LEAD1 の RgnMarkingMode（RGNMARK_RECT/RGNMARK_ELLIPSE）
        // ================================================================
        private void PicSub_Paint(object sender, PaintEventArgs e)
        {
            if (!_selecting) return;
            var rect = GetSelectionRect();
            using (var pen = new Pen(Color.Red, 1))
            {
                pen.DashStyle = DashStyle.Dash;
                if (_useEllipse)
                    e.Graphics.DrawEllipse(pen, rect);
                else
                    e.Graphics.DrawRectangle(pen, rect);
            }
        }

        // ================================================================
        // サブ画像 マウスダウン（選択開始）
        // ================================================================
        private void PicSub_MouseDown(object sender, MouseEventArgs e)
        {
            _selecting = true;
            _selStart = e.Location;
            _selEnd = e.Location;
        }

        // ================================================================
        // サブ画像 マウス移動（ラバーバンド更新）
        // ================================================================
        private void PicSub_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_selecting) return;
            _selEnd = e.Location;
            _picSub.Invalidate(); // Paint を呼んでラバーバンドを再描画
        }

        // ================================================================
        // サブ画像 マウスアップ（切り取り実行）
        // VB6: LEAD1_MouseUp → LEAD1.Floater = LEAD1.Bitmap → LEAD3.Bitmap = LEAD1.Floater
        // C# : 選択範囲を _cropImage としてクロップ
        // ================================================================
        private void PicSub_MouseUp(object sender, MouseEventArgs e)
        {
            _selEnd = e.Location;
            _selecting = false;
            _picSub.Invalidate();
            _kaiten = false;

            if (_subImage == null) return;

            var rect = GetSelectionRect();
            if (rect.Width < 2 || rect.Height < 2)
            {
                // 選択なし → 画像全体を使用
                var prev = _cropImage;
                _cropImage = (Image)_subImage.Clone();
                if (prev != null) prev.Dispose();
                return;
            }

            // 選択範囲を _cropImage としてクロップ
            // （PictureBox.SizeMode=Zoom の場合はオフセット補正が必要）
            Rectangle srcRect = ScalePicBoxRectToImage(rect, _picSub, _subImage);
            srcRect.Intersect(new Rectangle(0, 0, _subImage.Width, _subImage.Height));
            if (srcRect.Width < 1 || srcRect.Height < 1) return;

            var crop = new Bitmap(srcRect.Width, srcRect.Height);
            using (var g = Graphics.FromImage(crop))
            {
                if (_useEllipse)
                {
                    // 楕円クリップ（RGNMARK_ELLIPSE 相当）
                    var gp = new GraphicsPath();
                    gp.AddEllipse(0, 0, srcRect.Width, srcRect.Height);
                    g.SetClip(gp);
                }
                g.DrawImage(_subImage, new Rectangle(0, 0, srcRect.Width, srcRect.Height),
                            srcRect, GraphicsUnit.Pixel);
            }
            var prevCrop = _cropImage;
            _cropImage = crop;
            if (prevCrop != null) prevCrop.Dispose();
        }

        // ================================================================
        // メイン画像キャンバス マウスダウン
        // VB6: LEAD2_MouseDown → PasteFlag3/PasteFlag2/PasteFlag 処理
        // ================================================================
        private void Canvas_MouseDown(object sender, MouseEventArgs e)
        {
            // --- 選択貼り付けモード（PasteFlag3）---
            // クリック→ドラッグ後マウスアップで貼り付けサイズを決定
            if (_pasteFlag3)
            {
                if (_cropImage == null)
                {
                    MessageBox.Show("サブ画像が選択されていません。", "メイン画像編集チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                _floaterVisible = false;
                _paste3Start = e.Location;
                _canvasMain.Invalidate();
                return;
            }

            // --- 指定貼り付けモード（PasteFlag2）: クリック位置を中心に配置 ---
            if (_pasteFlag2)
            {
                if (_cropImage == null)
                {
                    MessageBox.Show("サブ画像が選択されていません。", "メイン画像編集チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                _floaterVisible = false;

                // VB6: LEAD5.Bitmap = LEAD3.Floater → LEAD2.Floater = LEAD5.Floater
                // → SetFloaterDstRect x - width/2, y - height/2
                float fw = _cropImage.Width;
                float fh = _cropImage.Height;
                _floaterRect = new RectangleF(e.X - fw / 2, e.Y - fh / 2, fw, fh);
                _floaterVisible = true;
                _floaterPos = _floaterRect.Location;

                _pasteFlag2 = false;
                _pasteFlag = true;
                _canvasMain.Invalidate();
                return;
            }

            // --- 通常フロータドラッグ（PasteFlag = True）---
            if (!_pasteFlag) return;
            _readyToDrag = true;
            if (_floaterRect.Contains(e.Location))
            {
                _dragStart = e.Location;
                _floaterPos = _floaterRect.Location;
            }
        }

        // ================================================================
        // メイン画像キャンバス マウス移動
        // VB6: LEAD2_MouseMove → フロータをドラッグ移動
        // ================================================================
        private void Canvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_readyToDrag || !_pasteFlag) return;
            if (!_floaterRect.Contains(e.Location) && !_readyToDrag) return;

            // フロータ位置を移動量だけ更新
            float dx = e.X - _dragStart.X;
            float dy = e.Y - _dragStart.Y;

            _floaterRect = new RectangleF(
                _floaterPos.X + dx,
                _floaterPos.Y + dy,
                _floaterRect.Width,
                _floaterRect.Height);
            _floaterPos = _floaterRect.Location;
            _dragStart = e.Location;

            _canvasMain.Invalidate();
        }

        // ================================================================
        // メイン画像キャンバス マウスアップ
        // VB6: LEAD2_MouseUp → PasteFlag3 処理（選択範囲のサイズで貼り付け）
        // ================================================================
        private void Canvas_MouseUp(object sender, MouseEventArgs e)
        {
            if (_pasteFlag3)
            {
                if (_cropImage == null)
                {
                    MessageBox.Show("サブ画像が選択されていません。", "メイン画像編集チェック",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _readyToDrag = false;
                    return;
                }

                float wk_x = Math.Abs(e.X - _paste3Start.X);
                float wk_y = Math.Abs(e.Y - _paste3Start.Y);
                if (wk_x < 1 || wk_y < 1)
                {
                    MessageBox.Show("選択貼り付けの場合、マウスで動かしながら範囲を指定して下さい。",
                        "メイン画像編集チェック", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _readyToDrag = false;
                    return;
                }

                // VB6: LEAD5.Size wk_x, wk_y → 選択サイズに拡縮
                var prev = _workImage;
                _workImage = ResizeImage(_cropImage, (int)wk_x, (int)wk_y);
                if (prev != null) prev.Dispose();

                // 選択左上に配置
                float nx = Math.Min(e.X, _paste3Start.X);
                float ny = Math.Min(e.Y, _paste3Start.Y);
                _floaterRect = new RectangleF(nx, ny, _workImage.Width, _workImage.Height);
                _floaterVisible = true;

                _pasteFlag3 = false;
                _pasteFlag = true;
                _canvasMain.Invalidate();
            }

            _readyToDrag = false;
        }

        // ================================================================
        // 回転ボタン
        // VB6: btn_kaiten_Click → 角度 = Combo_kaiten1&2&3 → LEAD3.Rotate
        // C# : _cropImage を Matrix で回転
        // ================================================================
        private void BtnKaiten_Click(object sender, EventArgs e)
        {
            if (_cropImage == null)
            {
                MessageBox.Show("サブ画像を選択して下さい。", "回転ボタンチェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 角度取得（百・十・一の桁を結合）
            float angle = GetRotationAngle();

            // 切り取り画像をリセットして回転適用
            if (_subImage != null)
            {
                var prev = _cropImage;
                _cropImage = (Image)_subImage.Clone();
                if (prev != null) prev.Dispose();
            }

            var rotated = RotateImage(_cropImage, angle);
            var prevCrop = _cropImage;
            _cropImage = rotated;
            if (prevCrop != null) prevCrop.Dispose();

            _kaiten = true;

            // ズームコンボをリセット
            _cmbZoom1.SelectedIndex = 0;
            _cmbZoom2.SelectedIndex = 0;
            _cmbZoom3.SelectedIndex = 0;
            _optZoom[0].Checked = true;
        }

        // ================================================================
        // ズームボタン
        // VB6: btn_zoom_Click → 率 = Combo_zoom1&2&3 → LEAD3.Size
        // C# : _cropImage をリサイズ（縮小モードは 200-rate%）
        // ================================================================
        private void BtnZoom_Click(object sender, EventArgs e)
        {
            if (_cropImage == null)
            {
                MessageBox.Show("サブ画像を選択して下さい。", "ズームボタンチェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // VB6: 縮小モードは Option_zoom(0)=True → Ookisa = 200 - Ookisa
            double rate = GetZoomRate();
            if (_optZoom[0].Checked && rate >= 200)
            {
                MessageBox.Show("縮小の範囲は199%までです。", "ズームボタンチェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (_optZoom[0].Checked) rate = 200 - rate; // 縮小

            // 切り取り画像を最初からやり直して回転→ズームを適用
            System.Drawing.Image baseImg = (_subImage != null) ? (Image)_subImage.Clone() : (Image)_cropImage.Clone();
            if (_kaiten)
            {
                float angle = GetRotationAngle();
                var rotated = RotateImage(baseImg, angle);
                baseImg.Dispose();
                baseImg = rotated;
            }

            int newW = Math.Max(1, (int)(baseImg.Width * rate / 100.0));
            int newH = Math.Max(1, (int)(baseImg.Height * rate / 100.0));
            var zoomed = ResizeImage(baseImg, newW, newH);
            baseImg.Dispose();

            var prev = _cropImage;
            _cropImage = zoomed;
            if (prev != null) prev.Dispose();
        }

        // ================================================================
        // 切り取り初期化ボタン
        // VB6: btn_init_kiritori_Click → LEAD3.Bitmap = LEAD1.Floater（全画像に戻す）
        // ================================================================
        private void BtnInitKiritori_Click(object sender, EventArgs e)
        {
            if (_cropImage == null)
            {
                MessageBox.Show("サブ画像が選択されていません。", "元に戻すボタンチェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (_subImage != null)
            {
                var prev = _cropImage;
                _cropImage = (Image)_subImage.Clone();
                if (prev != null) prev.Dispose();
            }
            _kaiten = false;
        }

        // ================================================================
        // 指定貼り付けボタン
        // VB6: btn_paste2_Click → PasteFlag2 = True（次クリック位置に中心配置）
        // ================================================================
        private void BtnPaste2_Click(object sender, EventArgs e)
        {
            if (_cropImage == null)
            {
                MessageBox.Show("サブ画像が選択されていません。", "メイン画像編集チェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            MessageBox.Show("貼り付けたい場所の中心をクリックして下さい。", "指定貼り付け",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            _pasteFlag2 = true;
            _pasteFlag3 = false;
        }

        // ================================================================
        // 選択貼り付けボタン
        // VB6: btn_paste3_Click → PasteFlag3 = True（ドラッグ範囲に拡縮して貼り付け）
        // ================================================================
        private void BtnPaste3_Click(object sender, EventArgs e)
        {
            if (_cropImage == null)
            {
                MessageBox.Show("サブ画像が選択されていません。", "メイン画像編集チェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            MessageBox.Show("貼り付けたい場所を選択して下さい。", "選択貼り付け",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            _pasteFlag3 = true;
            _pasteFlag2 = false;
        }

        // ================================================================
        // 固定ボタン（合成確定）
        // VB6: btn_fix_Click → LEAD2.Combine → 履歴に保存
        // C# : _mainImage へ _cropImage を DrawImage で合成 → history 保存
        // ================================================================
        private void BtnFix_Click(object sender, EventArgs e)
        {
            if (_cropImage == null)
            {
                MessageBox.Show("サブ画像が選択されていません。", "メイン画像編集チェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!_pasteFlag)
            {
                MessageBox.Show("切り取り画像が貼り付けられていません。", "メイン画像編集チェック",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // VB6: LEAD2.Combine → フロータをメイン画像に合成
            // C# : _mainImage に _workImage（or _cropImage）を DrawImage
            Image srcImg = (_workImage != null) ? _workImage : _cropImage;
            var newMain = new Bitmap(_mainImage.Width, _mainImage.Height);
            using (var g = Graphics.FromImage(newMain))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                // ベース画像を描画
                g.DrawImage(_mainImage, 0, 0, _mainImage.Width, _mainImage.Height);
                // フロータ画像を FloaterRect 位置に描画
                var destRectF = _floaterRect;
                g.DrawImage(srcImg, destRectF.X, destRectF.Y, destRectF.Width, destRectF.Height);
            }

            var prevMain = _mainImage;
            _mainImage = newMain;
            if (prevMain != null) prevMain.Dispose();
            _floaterVisible = false;
            _pasteFlag = false;
            _pasteFlag2 = false;
            _pasteFlag3 = false;

            // 合成履歴に保存
            _gouseiCurrent++;
            _gouseiMax = _gouseiCurrent;
            string histPath = GetHistoryPath(_gouseiCurrent);
            SaveImageAsJpeg(_mainImage, histPath);

            _btnLeft.Enabled = true;
            _btnRight.Enabled = false;
            _canvasMain.Invalidate();
        }

        // ================================================================
        // メイン画像リセット
        // VB6: btn_init_syouhin_Click → Option_ld2 の選択に戻す
        // ================================================================
        private void BtnInitSyouhin_Click(object sender, EventArgs e)
        {
            int idx = GetCheckedIndex(_optLd2);
            LoadMainImage(idx >= 0 ? idx : 0);

            string src = GetImagePath(idx >= 0 ? idx : 0);
            if (File.Exists(src))
                File.Copy(src, GetHistoryPath(1), true);

            _gouseiMax = 1;
            _gouseiCurrent = 1;
            _btnLeft.Enabled = false;
            _btnRight.Enabled = false;
        }

        // ================================================================
        // 合成履歴：前へ
        // VB6: btn_left_Click → gousei_current-- → LEAD2.Load history
        // ================================================================
        private void BtnLeft_Click(object sender, EventArgs e)
        {
            _gouseiCurrent--;
            _btnRight.Enabled = true;
            if (_gouseiCurrent <= 1) _btnLeft.Enabled = false;
            LoadMainImageFromFile(GetHistoryPath(_gouseiCurrent));
        }

        // ================================================================
        // 合成履歴：次へ
        // VB6: btn_right_Click → gousei_current++ → LEAD2.Load history
        // ================================================================
        private void BtnRight_Click(object sender, EventArgs e)
        {
            _gouseiCurrent++;
            _btnLeft.Enabled = true;
            if (_gouseiCurrent >= _gouseiMax) _btnRight.Enabled = false;
            LoadMainImageFromFile(GetHistoryPath(_gouseiCurrent));
        }

        // ================================================================
        // 画像保存
        // VB6: btn_save_jpeg_Click → LEAD2.Save "gousei.jpg" → form_save_gousei.Show
        // ================================================================
        private void BtnSaveJpeg_Click(object sender, EventArgs e)
        {
            if (_mainImage == null) return;
            string savePath = Path.Combine(Application.StartupPath, "gousei.jpg");
            SaveImageAsJpeg(_mainImage, savePath);

            // 保存先選択（gousei_1.jpg ～ gousei_6.jpg）
            var dlg = new FormSaveGousei();
            this.Hide();
            dlg.FormClosed += (s, args) => this.Show();
            dlg.Show();
        }

        // ================================================================
        // 印刷
        // VB6: btn_print_Click → LEAD4 = LEAD2.Rotate90 → 2×2 印刷
        // ================================================================
        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("メイン画像を印刷します。よろしいですか？", "メイン画像印刷確認",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            if (_mainImage == null) return;

            // 印刷バッファを作成して 90°回転（LEAD4.Rotate 9000）
            var prev = _printImage;
            _printImage = (Image)_mainImage.Clone();
            _printImage.RotateFlip(RotateFlipType.Rotate90FlipNone);
            if (prev != null) prev.Dispose();

            var pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(PrintPage);
            var dlg = new PrintDialog { Document = pd };
            if (dlg.ShowDialog() == DialogResult.OK) pd.Print();
        }

        /// <summary>1ページに 2列×2行（計4枚）印刷</summary>
        private void PrintPage(object sender, PrintPageEventArgs e)
        {
            if (_printImage == null) return;
            int pageW = e.PageBounds.Width;
            int pageH = e.PageBounds.Height;
            int drawW = pageW / 2 - 30;
            int drawH = pageH / 2 - 50;
            double sx = (double)drawW / _printImage.Width;
            double sy = (double)drawH / _printImage.Height;
            double sc = Math.Min(sx, sy);
            int iw = (int)(_printImage.Width * sc);
            int ih = (int)(_printImage.Height * sc);
            int[] xs = { 20, iw + 40 };
            int[] ys = { 20, ih + 50 };
            foreach (int row in new[] { 0, 1 })
                foreach (int col in new[] { 0, 1 })
                    e.Graphics.DrawImage(_printImage, xs[col], ys[row], iw, ih);
            e.HasMorePages = false;
        }

        // ================================================================
        // メニューへ戻る
        // VB6: btn_back_Click → form_gousei.Visible = False; form_menu.Visible = True
        // ================================================================
        private void BtnBack_Click(object sender, EventArgs e)
        {
            DisposeImages();
            this.Close();
        }

        // ================================================================
        // Option_ld1 変更（サブ画像切り替え）
        // ================================================================
        private void OptLd1_CheckedChanged(object sender, EventArgs e)
        {
            var rb = (RadioButton)sender;
            if (!rb.Checked) return;
            int idx = (int)rb.Tag;
            _picSub.Visible = false;
            LoadSubImage(idx);
            _picSub.Visible = true;
            UpdateOptionColors(_optLd1);
        }

        // ================================================================
        // Option_ld2 変更（メイン画像切り替え）
        // ================================================================
        private void OptLd2_CheckedChanged(object sender, EventArgs e)
        {
            var rb = (RadioButton)sender;
            if (!rb.Checked) return;
            int idx = (int)rb.Tag;
            LoadMainImage(idx);
            // 履歴リセット
            string src = GetImagePath(idx);
            if (File.Exists(src)) File.Copy(src, GetHistoryPath(1), true);
            _gouseiMax = 1;
            _gouseiCurrent = 1;
            _btnLeft.Enabled = false;
            _btnRight.Enabled = false;
            UpdateOptionColors(_optLd2);
        }

        // ================================================================
        // Option_camera 変更（四角/楕円選択モード切り替え）
        // VB6: LEAD1.RgnMarkingMode = RGNMARK_RECT/ELLIPSE
        // ================================================================
        private void OptCamera_CheckedChanged(object sender, EventArgs e)
        {
            _useEllipse = _optCamera[1].Checked;
            UpdateOptionColors(_optCamera);
        }

        private void OptKaiten_CheckedChanged(object sender, EventArgs e) => UpdateOptionColors(_optKaiten);
        private void OptZoom_CheckedChanged(object sender, EventArgs e) => UpdateOptionColors(_optZoom);
        private void OptMainSentaku_CheckedChanged(object sender, EventArgs e) => UpdateOptionColors(_optMainSentaku);

        // ================================================================
        // ユーティリティ
        // ================================================================

        private float GetRotationAngle()
        {
            // VB6: CLng(Combo_kaiten1 & Combo_kaiten2 & Combo_kaiten3) → 100倍してLEAD3.Rotate
            int digits = (_cmbKaiten1.SelectedIndex * 100
                        + _cmbKaiten2.SelectedIndex * 10
                        + _cmbKaiten3.SelectedIndex);
            float angle = digits;
            // VB6: Option_kaiten(0)=左回転 → 負の角度
            if (_optKaiten[0].Checked) angle = -angle;
            return angle;
        }

        private double GetZoomRate()
        {
            // VB6: CDbl(Combo_zoom1 & Combo_zoom2 & Combo_zoom3)
            return _cmbZoom1.SelectedIndex * 100
                 + _cmbZoom2.SelectedIndex * 10
                 + _cmbZoom3.SelectedIndex;
        }

        /// <summary>画像を任意角度で回転（背景=青 RGB(0,0,255) VB6と同仕様）</summary>
        private Image RotateImage(Image src, float angle)
        {
            double rad = angle * Math.PI / 180.0;
            double cos = Math.Abs(Math.Cos(rad));
            double sin = Math.Abs(Math.Sin(rad));
            int newW = (int)(src.Width * cos + src.Height * sin);
            int newH = (int)(src.Width * sin + src.Height * cos);

            var bmp = new Bitmap(newW, newH);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Blue); // LEAD3.Rotate ... RGB(0,0,255) と同仕様
                g.TranslateTransform(newW / 2f, newH / 2f);
                g.RotateTransform(angle);
                g.TranslateTransform(-src.Width / 2f, -src.Height / 2f);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(src, 0, 0, src.Width, src.Height);
            }
            return bmp;
        }

        /// <summary>画像をリサイズ（LEAD3.Size 相当）</summary>
        private Image ResizeImage(Image src, int w, int h)
        {
            var bmp = new Bitmap(w, h);
            using (var g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(src, 0, 0, w, h);
            }
            return bmp;
        }

        /// <summary>JPEG 保存</summary>
        private void SaveImageAsJpeg(Image img, string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            var codec = GetJpegCodec();
            var encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
            encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                System.Drawing.Imaging.Encoder.Quality, 90L);
            img.Save(path, codec, encoderParams);
        }

        private System.Drawing.Imaging.ImageCodecInfo GetJpegCodec()
        {
            foreach (var c in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders())
                if (c.MimeType == "image/jpeg") return c;
            return null;
        }

        private void LoadMainImageFromFile(string path)
        {
            if (!File.Exists(path)) return;
            var prev = _mainImage;
            _mainImage = FitImageToSize(Image.FromFile(path), _canvasMain.Width, _canvasMain.Height);
            if (prev != null) prev.Dispose();
            _floaterVisible = false;
            _canvasMain.Invalidate();
        }

        private string GetImagePath(int index) => Path.Combine(Application.StartupPath, ImageFiles[index]);
        private string GouseiHistoryPath() => Path.Combine(Application.StartupPath, "gousei");
        private string GetHistoryPath(int n) => Path.Combine(GouseiHistoryPath(), "gousei_" + n + ".jpg");

        private Rectangle GetSelectionRect()
        {
            return new Rectangle(
                Math.Min(_selStart.X, _selEnd.X),
                Math.Min(_selStart.Y, _selEnd.Y),
                Math.Abs(_selEnd.X - _selStart.X),
                Math.Abs(_selEnd.Y - _selStart.Y));
        }

        /// <summary>
        /// PictureBox(SizeMode=Zoom)上の座標を実画像座標に変換する
        /// </summary>
        private Rectangle ScalePicBoxRectToImage(Rectangle picRect, PictureBox pb, Image img)
        {
            double sx = (double)img.Width / pb.Width;
            double sy = (double)img.Height / pb.Height;
            // Zoom では長辺を基準に均等スケールするためオフセットが生じる
            double scale = Math.Max(sx, sy);
            // 表示領域オフセット
            int dispW = (int)(img.Width / scale);
            int dispH = (int)(img.Height / scale);
            int offX = (pb.Width - dispW) / 2;
            int offY = (pb.Height - dispH) / 2;
            double imgScale = (double)img.Width / dispW;

            return new Rectangle(
                (int)((picRect.X - offX) * imgScale),
                (int)((picRect.Y - offY) * imgScale),
                (int)(picRect.Width * imgScale),
                (int)(picRect.Height * imgScale));
        }

        private int GetCheckedIndex(RadioButton[] arr)
        {
            for (int i = 0; i < arr.Length; i++)
                if (arr[i] != null && arr[i].Checked) return (int)arr[i].Tag;
            return 0;
        }

        private void UpdateOptionColors(RadioButton[] arr)
        {
            var sel = Color.FromArgb(0xFF, 0xC0, 0xC0);
            var unsel = Color.FromArgb(0xD8, 0xFF, 0xFF);
            foreach (var rb in arr)
                if (rb != null) rb.BackColor = rb.Checked ? sel : unsel;
        }

        private void DisposeImages()
        {
            foreach (var img in new Image[] { _subImage, _mainImage, _cropImage, _printImage, _workImage })
                img?.Dispose();
        }

        // ================================================================
        // コントロール生成ヘルパー
        // ================================================================
        private Label NewLabel(string text, int x, int y, int w, bool bold = false)
        {
            var lbl = new Label { Text = text, Location = new Point(x, y), Size = new Size(w, 18), BackColor = Color.Transparent };
            if (bold) lbl.Font = new Font("メイリオ", 9F, FontStyle.Bold);
            this.Controls.Add(lbl);
            return lbl;
        }

        private RadioButton NewRadio(string text, int x, int y, int w)
        {
            return new RadioButton { Text = text, Location = new Point(x, y), Size = new Size(w, 20), BackColor = Color.Transparent };
        }

        private Button NewButton(string text, int x, int y, int w, int h)
        {
            var btn = new Button { Text = text, Location = new Point(x, y), Size = new Size(w, h) };
            this.Controls.Add(btn);
            return btn;
        }

        private ComboBox NewDigitCombo(int x, int y, int w, int start, int end)
        {
            var cmb = new ComboBox { Location = new Point(x, y), Size = new Size(w, 22), DropDownStyle = ComboBoxStyle.DropDownList };
            for (int i = start; i <= end; i++) cmb.Items.Add(i.ToString());
            cmb.SelectedIndex = 0;
            this.Controls.Add(cmb);
            return cmb;
        }

        private string[] BuildImageLabels()
        {
            return new string[]
            {
                "商品①", "商品②", "商品③", "商品④",
                "カメラ①", "カメラ②", "カメラ③", "カメラ④",
                "合成①", "合成②", "合成③", "合成④",
                "商品⑤", "商品⑥", "カメラ⑤", "カメラ⑥", "合成⑤", "合成⑥",
            };
        }

        // ================================================================
        // GouseiCanvas: メイン画像 + フロータオーバーレイを描画するカスタムPanel
        // VB6: LEAD2 に相当（フロータ機能はカスタムPaintで実装）
        // ================================================================
        private class GouseiCanvas : Panel
        {
            private readonly FormGousei _owner;

            public GouseiCanvas(FormGousei owner)
            {
                _owner = owner;
                this.DoubleBuffered = true;
                this.ResizeRedraw = true;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                var g = e.Graphics;

                // ベース画像（LEAD2.Bitmap）を描画
                if (_owner._mainImage != null)
                    g.DrawImage(_owner._mainImage, 0, 0, _owner._mainImage.Width, _owner._mainImage.Height);

                // フロータオーバーレイ（FloaterVisible = True のとき）
                if (_owner._floaterVisible)
                {
                    Image floatSrc = _owner._workImage ?? _owner._cropImage;
                    if (floatSrc != null)
                    {
                        var rect = _owner._floaterRect;
                        // 半透明で描画（アルファ 180/255 程度）
                        using (var attr = new System.Drawing.Imaging.ImageAttributes())
                        {
                            float alpha = 0.7f;
                            var cm = new System.Drawing.Imaging.ColorMatrix();
                            cm.Matrix33 = alpha;
                            attr.SetColorMatrix(cm);
                            g.DrawImage(floatSrc,
                                new Rectangle((int)rect.X, (int)rect.Y, (int)rect.Width, (int)rect.Height),
                                0, 0, floatSrc.Width, floatSrc.Height,
                                GraphicsUnit.Pixel, attr);
                        }
                        // フロータ枠線
                        using (var pen = new Pen(Color.Yellow, 1) { DashStyle = DashStyle.Dash })
                            g.DrawRectangle(pen, rect.X, rect.Y, rect.Width, rect.Height);
                    }
                }
            }
        }
    }

    // ================================================================
    // FormSaveGousei: 合成画像の保存先選択フォーム（スタブ）
    // VB6: form_save_gousei.frm に相当
    // ================================================================
    public class FormSaveGousei : Form
    {
        public FormSaveGousei()
        {
            this.Text = "合成画像の保存（form_save_gousei の移植先）";
            this.Size = new Size(500, 300);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;

            var lbl = new Label();
            lbl.Text = "保存先を選択してください（gousei_1.jpg ～ gousei_6.jpg）";
            lbl.Location = new Point(20, 40);
            lbl.AutoSize = true;
            this.Controls.Add(lbl);

            for (int i = 1; i <= 6; i++)
            {
                int idx = i;
                var btn = new Button();
                btn.Text = "gousei_" + i + ".jpg";
                btn.Location = new Point(20 + (i - 1) % 3 * 150, 90 + (i - 1) / 3 * 50);
                btn.Size = new Size(140, 40);
                btn.Click += (s, e) =>
                {
                    string src = Path.Combine(Application.StartupPath, "gousei.jpg");
                    string dest = Path.Combine(Application.StartupPath, "gousei_" + idx + ".jpg");
                    if (File.Exists(src)) File.Copy(src, dest, true);
                    MessageBox.Show("gousei_" + idx + ".jpg に保存しました。", "保存完了",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                };
                this.Controls.Add(btn);
            }

            var btnClose = new Button();
            btnClose.Text = "閉じる";
            btnClose.Location = new Point(190, 210);
            btnClose.Size = new Size(120, 40);
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);
        }
    }
}
