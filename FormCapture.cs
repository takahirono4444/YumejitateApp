using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

namespace YumejitateApp
{
    /// <summary>
    /// デジタルカメラ画像読み込みフォーム (VB6: form_capture3.frm の移植)
    /// デジカメ画像を取り込んでJPEGとして保存し、form_save_dejicame へ遷移する。
    ///
    /// VB6との主な相違点:
    ///   - VB6: EzScan.OCX (TWAIN スキャナ制御 32bit OCX) を使用
    ///     → x64環境では EzScan.OCX を使用できないため、スキャン機能を無効化し
    ///       「ファイルから読込」ボタンによる OpenFileDialog での代替を提供する。
    ///   - VB6: Image1 コントロール (Stretch=True) で画像を伸縮表示
    ///     → C#: PictureBox (SizeMode=Zoom) で代替
    ///   - VB6: EzScan1.SaveImageFile(App.Path & "\capture_3.jpg")
    ///     → C#: Image.Save(Application.StartupPath + "\capture_3.jpg")
    /// </summary>
    public class FormCapture : Form
    {
        // ----------------------------------------------------------------
        // 保存先パス (VB6: App.Path & "\capture_3.jpg")
        // ----------------------------------------------------------------
        private static readonly string SavePath =
            Path.Combine(Application.StartupPath, "capture_3.jpg");

        // ----------------------------------------------------------------
        // コントロール
        // ----------------------------------------------------------------
        private PictureBox _picImage;       // VB6: Image1 (Stretch=True)
        private Label _lblTitle;       // VB6: Label2 "デジタルカメラ画像　読み込み"
        private Label _lblInstTitle;   // VB6: Label3 "操作説明"
        private Label _lblInst;        // VB6: label1 (スキャナ操作説明)
        private Label _lblX64Notice;   // C#追加: x64非対応の通知ラベル
        private Button _btnSelectScan;  // VB6: pbSelectScan "スキャナ選択"
        private Button _btnScan;        // VB6: pbScan "スキャン"
        private Button _btnLoadFile;    // C#追加: ファイルから画像を読み込む
        private Button _btnSave;        // VB6: btn_save "画像保存"
        private Button _btnMenu;        // VB6: btn_menu "メニュー"

        // ----------------------------------------------------------------
        // 状態
        // ----------------------------------------------------------------
        private Image _loadedImage = null;  // VB6: Image1.Picture に相当

        // ================================================================
        // コンストラクタ
        // ================================================================
        public FormCapture()
        {
            InitializeComponent();
        }

        // ================================================================
        // InitializeComponent
        // ================================================================
        private void InitializeComponent()
        {
            this.SuspendLayout();

            // ---- フォーム基本設定 ----
            // VB6: BackColor=&H00D8FFFF&, Caption="夢仕立て-デジタルカメラ画像　読み込み"
            this.Text = "夢仕立て - デジタルカメラ画像　読み込み";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;

            // ---- タイトルラベル (VB6: Label2, Italic, Size=18) ----
            // VB6: "デジタルカメラ画像　読み込み", Left=5280, Top=120 twips → 352, 8px
            _lblTitle = new Label
            {
                Text = "デジタルカメラ画像　読み込み",
                Font = new Font("ＭＳ Ｐゴシック", 18f, FontStyle.Italic),
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(350, 8),
            };
            this.Controls.Add(_lblTitle);

            // ---- 画像表示エリア (VB6: Image1, Stretch=True) ----
            // VB6: Left=2880, Top=600, Width=9615, Height=6615 twips → 192,40,641,441px
            _picImage = new PictureBox
            {
                Location = new Point(192, 40),
                Size = new Size(641, 441),
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
            };
            this.Controls.Add(_picImage);

            // ---- x64非対応通知ラベル (C#追加) ----
            // VB6の EzScan.OCX は32bitコンポーネントのためx64環境では使用不可
            _lblX64Notice = new Label
            {
                Text = "【x64環境のご案内】\r\n" +
                            "スキャン機能（EzScan.OCX）は32bit専用のため、\r\n" +
                            "このx64版アプリでは使用できません。\r\n" +
                            "「ファイルから読込」ボタンで画像ファイルを直接指定してください。",
                Font = new Font("ＭＳ Ｐゴシック", 11f),
                BackColor = Color.FromArgb(0xFF, 0xFF, 0xCC),
                ForeColor = Color.DarkRed,
                BorderStyle = BorderStyle.FixedSingle,
                AutoSize = false,
                Size = new Size(420, 100),
                Location = new Point(12, 50),
                TextAlign = ContentAlignment.MiddleLeft,
            };
            this.Controls.Add(_lblX64Notice);

            // ---- 操作説明ラベル (VB6: Label3 "操作説明") ----
            // VB6: Left=1200, Top=8400 twips → 80, 560px (スクロールバー分を考慮して相対配置)
            _lblInstTitle = new Label
            {
                Text = "操作説明",
                Font = new Font("ＭＳ Ｐゴシック", 15.75f, FontStyle.Bold),
                BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF),
                AutoSize = true,
                Location = new Point(80, 496),
            };
            this.Controls.Add(_lblInstTitle);

            // ---- スキャナ操作説明ラベル (VB6: label1) ----
            // VB6: "　スキャナ選択に「スキャン」ボタンが出ます"
            _lblInst = new Label
            {
                Text = "　スキャナ選択に「スキャン」ボタンが出ます",
                Font = new Font("ＭＳ Ｐゴシック", 15.75f, FontStyle.Bold),
                BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF),
                AutoSize = true,
                Location = new Point(104, 520),
            };
            this.Controls.Add(_lblInst);

            // ---- スキャナ選択ボタン (VB6: pbSelectScan "スキャナ選択") ----
            // VB6: Left=6600, Top=7560 twips → 440, 504px
            // x64環境では使用不可 → クリック時にメッセージを表示
            _btnSelectScan = new Button
            {
                Text = "スキャナ選択",
                Font = new Font("ＭＳ Ｐゴシック", 13f, FontStyle.Bold),
                Size = new Size(150, 60),
                Location = new Point(440, 490),
                BackColor = Color.FromArgb(0xD0, 0xD0, 0xD0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Enabled = true,  // x64では機能しないがボタン自体は表示する
            };
            _btnSelectScan.FlatAppearance.BorderColor = Color.Gray;
            _btnSelectScan.FlatAppearance.BorderSize = 2;
            _btnSelectScan.Click += BtnSelectScan_Click;
            this.Controls.Add(_btnSelectScan);

            // ---- スキャンボタン (VB6: pbScan "スキャン") ----
            // VB6: Left=6600, Top=8640 twips → 440, 576px
            _btnScan = new Button
            {
                Text = "スキャン",
                Font = new Font("ＭＳ Ｐゴシック", 13f, FontStyle.Bold),
                Size = new Size(150, 60),
                Location = new Point(440, 560),
                BackColor = Color.FromArgb(0xD0, 0xD0, 0xD0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Enabled = true,
            };
            _btnScan.FlatAppearance.BorderColor = Color.Gray;
            _btnScan.FlatAppearance.BorderSize = 2;
            _btnScan.Click += BtnScan_Click;
            this.Controls.Add(_btnScan);

            // ---- ファイルから読込ボタン (C#追加: EzScan OCX の代替) ----
            // OpenFileDialog で画像ファイルを選択して PictureBox に表示する
            _btnLoadFile = new Button
            {
                Text = "ファイルから読込",
                Font = new Font("ＭＳ Ｐゴシック", 13f, FontStyle.Bold),
                Size = new Size(200, 60),
                Location = new Point(600, 560),
                BackColor = Color.FromArgb(0xC0, 0xFF, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnLoadFile.FlatAppearance.BorderColor = Color.SeaGreen;
            _btnLoadFile.FlatAppearance.BorderSize = 2;
            _btnLoadFile.Click += BtnLoadFile_Click;
            this.Controls.Add(_btnLoadFile);

            // ---- 画像保存ボタン (VB6: btn_save "画像保存") ----
            // VB6: Left=6600, Top=9720 twips → 440, 648px
            // → 画像を capture_3.jpg として保存後 form_save_dejicame へ遷移
            _btnSave = new Button
            {
                Text = "画像保存",
                Font = new Font("ＭＳ Ｐゴシック", 14f, FontStyle.Bold),
                Size = new Size(150, 60),
                Location = new Point(440, 630),
                BackColor = Color.FromArgb(0xFF, 0xFF, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnSave.FlatAppearance.BorderColor = Color.Goldenrod;
            _btnSave.FlatAppearance.BorderSize = 2;
            _btnSave.Click += BtnSave_Click;
            this.Controls.Add(_btnSave);

            // ---- メニューボタン (VB6: btn_menu "メニュー") ----
            // VB6: Left=10560, Top=9720 twips → 704, 648px
            _btnMenu = new Button
            {
                Text = "メニュー",
                Font = new Font("ＭＳ Ｐゴシック", 14f, FontStyle.Bold),
                Size = new Size(150, 60),
                BackColor = Color.FromArgb(0xFF, 0xC0, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            };
            _btnMenu.FlatAppearance.BorderColor = Color.Crimson;
            _btnMenu.FlatAppearance.BorderSize = 2;
            _btnMenu.Click += BtnMenu_Click;
            this.Controls.Add(_btnMenu);

            this.ResumeLayout(false);
        }

        // ================================================================
        // OnLoad (VB6: Form_Load)
        // ================================================================
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.WindowState = FormWindowState.Maximized;
            CenterControls();

            // VB6: txImageDir.Text = "" / txFileName.Text = ""
            // → C#: 対応するテキストボックスなし (非表示だったため省略)

            // VB6: Call EzScan1.SaveImageFile(App.Path & "\capture_3.jpg")
            //   → EzScan OCX が NULL の場合にファイルダイアログを開く動作
            //   → x64では EzScan.OCX を使用できないため、x64通知のみ表示
            // (通知ラベルは InitializeComponent で常に表示)
        }

        // ================================================================
        // OnResize
        // ================================================================
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            CenterControls();
        }

        // ================================================================
        // コントロールを画面サイズに合わせて配置する
        // VB6では固定座標だったが、C#では最大化対応のため動的配置する。
        // ================================================================
        private void CenterControls()
        {
            if (this.ClientSize.Width == 0 || this.ClientSize.Height == 0) return;

            int w = this.ClientSize.Width;
            int h = this.ClientSize.Height;

            // タイトルラベル (上部中央)
            if (_lblTitle != null)
                _lblTitle.Location = new Point((w - _lblTitle.Width) / 2, 8);

            // 画像表示エリア (中央上部, 幅の約2/3)
            if (_picImage != null)
            {
                int picW = (int)(w * 0.65);
                int picH = (int)(h * 0.60);
                int picX = (w - picW) / 2 + 60;
                int picY = 35;
                _picImage.SetBounds(picX, picY, picW, picH);
            }

            // x64通知ラベル (左上)
            if (_lblX64Notice != null)
                _lblX64Notice.Location = new Point(8, 35);

            // 操作ボタンとラベル (下部)
            int btnAreaY = h - 220;
            if (btnAreaY < 500) btnAreaY = 500;

            if (_lblInstTitle != null)
                _lblInstTitle.Location = new Point(12, btnAreaY);

            if (_lblInst != null)
                _lblInst.Location = new Point(12, btnAreaY + 30);

            // スキャナボタン群 (左下エリア)
            if (_btnSelectScan != null)
                _btnSelectScan.Location = new Point(12, btnAreaY + 65);

            if (_btnScan != null)
                _btnScan.Location = new Point(175, btnAreaY + 65);

            if (_btnLoadFile != null)
                _btnLoadFile.Location = new Point(340, btnAreaY + 65);

            // 保存ボタン
            if (_btnSave != null)
                _btnSave.Location = new Point(12, btnAreaY + 145);

            // メニューボタン (右下)
            if (_btnMenu != null)
                _btnMenu.Location = new Point(w - 170, h - 80);
        }

        // ================================================================
        // イベントハンドラ
        // ================================================================

        /// <summary>
        /// スキャナ選択ボタン (VB6: pbSelectScan_Click → EzScan1.SelectScanner)
        /// x64環境では EzScan.OCX を使用できないためメッセージを表示してスキップする。
        /// </summary>
        private void BtnSelectScan_Click(object sender, EventArgs e)
        {
            ShowX64Notice("スキャナ選択");
        }

        /// <summary>
        /// スキャンボタン (VB6: pbScan_Click → EzScan1.startScan(0))
        /// x64環境では EzScan.OCX を使用できないためメッセージを表示してスキップする。
        /// VB6: EzScan1_EndScan イベントで Image1.Picture = EzScan1.Picture
        /// </summary>
        private void BtnScan_Click(object sender, EventArgs e)
        {
            ShowX64Notice("スキャン");
        }

        /// <summary>
        /// x64非対応の通知メッセージを表示する共通メソッド。
        /// 「ファイルから読込」ボタンの使用を案内する。
        /// </summary>
        private static void ShowX64Notice(string featureName)
        {
            MessageBox.Show(
                $"【x64環境のご案内】\r\n\r\n" +
                $"「{featureName}」機能（EzScan.OCX）は32bitコンポーネントのため、\r\n" +
                $"このx64版アプリでは使用できません。\r\n\r\n" +
                $"代わりに「ファイルから読込」ボタンを使用して、\r\n" +
                $"デジタルカメラから取り込んだ画像ファイルを直接指定してください。",
                "機能制限のご案内",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        /// <summary>
        /// ファイルから読込ボタン (C#追加: EzScan OCX の代替)
        /// OpenFileDialog で画像ファイルを選択して PictureBox に表示する。
        /// VB6: pbLoad_Click → EzScan1.LoadImageFile / Image1.Picture = EzScan1.Picture
        /// </summary>
        private void BtnLoadFile_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "デジタルカメラ画像を選択してください";
                dlg.Filter = "画像ファイル (*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff)|" +
                                  "*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff|" +
                                  "すべてのファイル (*.*)|*.*";
                dlg.FilterIndex = 1;

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // 既存の画像を解放してから新しい画像を読み込む
                        _loadedImage?.Dispose();
                        _loadedImage = Image.FromFile(dlg.FileName);
                        _picImage.Image = _loadedImage;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            "画像の読み込みに失敗しました。\r\n" + ex.Message,
                            "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// 画像保存ボタン (VB6: btn_save_Click)
        /// 読み込まれた画像を capture_3.jpg として保存し form_save_dejicame へ遷移する。
        ///
        /// VB6:
        ///   If CDbl(Image1.Picture.Handle) = 0 Then MsgBox "画像が読み込まれていません"
        ///   Else
        ///     EzScan1.SaveImageFile(App.Path &amp; "\capture_3.jpg")
        ///     Unload form_capture3 / form_save_dejicame.Show
        ///   End If
        /// </summary>
        private void BtnSave_Click(object sender, EventArgs e)
        {
            // VB6: CDbl(Image1.Picture.Handle) = 0 → 画像未読み込みチェック
            if (_loadedImage == null || _picImage.Image == null)
            {
                MessageBox.Show("画像が読み込まれていません",
                    "確認", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // VB6: EzScan1.SaveImageFile(App.Path & "\capture_3.jpg")
            try
            {
                // JPEG 品質 90% で保存 (EzScan の JPEG 保存に相当)
                SaveAsJpeg(_loadedImage, SavePath, 90L);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "画像の保存に失敗しました。\r\n" + ex.Message,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // VB6: Unload form_capture3 / form_save_dejicame.Show
            this.Close();

            try
            {
                var saveDejicame = new FormSaveDejicame();
                saveDejicame.Show();
            }
            catch
            {
                // FormSaveDejicame が未実装の場合はメニューに戻る
                foreach (Form f in Application.OpenForms)
                {
                    if (f is FormMenu) { f.Show(); break; }
                }
            }
        }

        /// <summary>
        /// メニューへ戻るボタン (VB6: btn_menu_Click → Unload form_capture3 / form_menu.Show)
        /// </summary>
        private void BtnMenu_Click(object sender, EventArgs e)
        {
            foreach (Form f in Application.OpenForms)
            {
                if (f is FormMenu) { f.Show(); break; }
            }
            this.Close();
        }

        // ================================================================
        // フォームクローズ時のクリーンアップ
        // ================================================================
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _loadedImage?.Dispose();
            _loadedImage = null;
            base.OnFormClosed(e);
        }

        // ================================================================
        // JPEG保存ヘルパー (System.Drawing で品質指定保存)
        // VB6: EzScan1.SaveImageFile(path) の代替
        // ================================================================

        /// <summary>
        /// Image を指定品質 (0〜100) の JPEG ファイルとして保存する。
        /// </summary>
        private static void SaveAsJpeg(Image image, string path, long quality)
        {
            // JPEG エンコーダを取得
            ImageCodecInfo jpegCodec = GetJpegCodec();
            if (jpegCodec == null)
            {
                // フォールバック: 品質指定なしで保存
                image.Save(path, ImageFormat.Jpeg);
                return;
            }

            var encoderParams = new EncoderParameters(1);
            encoderParams.Param[0] = new EncoderParameter(
                Encoder.Quality, quality);

            image.Save(path, jpegCodec, encoderParams);
        }

        /// <summary>JPEG ImageCodecInfo を取得する。</summary>
        private static ImageCodecInfo GetJpegCodec()
        {
            foreach (ImageCodecInfo codec in ImageCodecInfo.GetImageEncoders())
            {
                if (codec.MimeType == "image/jpeg")
                    return codec;
            }
            return null;
        }
    }
}
