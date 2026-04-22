using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace YumejitateApp
{
    /// <summary>
    /// CCDカメラモニターフォーム (VB6: frmVideo.frm の移植)
    /// Sentech USB3.0 CCDカメラ (StCamD.dll 64bit SDK) を使用してライブプレビューを表示する。
    ///
    /// VB6 との主な相違点:
    ///   - VB6: StCam_CreatePreviewWindowA が独立フローティングウィンドウを生成
    ///     → C#: PictureBox を親ウィンドウとして子ウィンドウ化し、フォームに埋め込む
    ///   - VB6: SetWindowLong + AddressOf WindowProc でWndProc をフック
    ///     → C#: WndProc をオーバーライドして WM_STCAM_* メッセージを処理
    ///   - VB6: Long (32bit) でカメラハンドルを保持
    ///     → C#: IntPtr (64bit) を使用
    /// </summary>
    public class FormVideo : Form
    {
        // ================================================================
        // StCamD.dll DLLパス (64bit SDK)
        // ================================================================
        private const string STCAM_DLL =
            @"C:\Users\yumejitate\Desktop\ソフト\センテックCCDカメラ\USB3.0シーモス" +
            @"\StCamUSBPack_JP_160325\3_SDK\StandardSDK(v3.10)\Bin\x64\StCamD.dll";

        // ================================================================
        // StCamD.dll P/Invoke 宣言
        // VB6: Long → C#: IntPtr (ハンドル), uint (DWORD数値), ushort (WORD数値)
        // ================================================================

        // --- 初期化 ---
        /// <summary>VB6: StCam_Open(dwInstance) As Long → カメラハンドルを取得</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern IntPtr StCam_Open(uint dwInstance);

        /// <summary>VB6: StCam_Close(hCamera)</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern void StCam_Close(IntPtr hCamera);

        /// <summary>VB6: StCam_GetLastError(hCamera) As Long</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern uint StCam_GetLastError(IntPtr hCamera);

        /// <summary>VB6: StCam_SetReceiveMsgWindow(hCamera, hWnd) As Long
        /// WM_STCAM_* メッセージを受信するウィンドウを設定する。</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern uint StCam_SetReceiveMsgWindow(IntPtr hCamera, IntPtr hWnd);

        // --- プレビュー ---
        /// <summary>VB6: StCam_StartTransfer(hCamera) As Long → 転送開始</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern uint StCam_StartTransfer(IntPtr hCamera);

        /// <summary>VB6: StCam_StopTransfer(hCamera) As Long → 転送停止</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern uint StCam_StopTransfer(IntPtr hCamera);

        /// <summary>
        /// VB6: StCam_CreatePreviewWindowA(...) As Long
        /// プレビューウィンドウを作成する。
        /// hWndParent に PictureBox.Handle を渡して子ウィンドウとして埋め込む。
        /// </summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi)]
        private static extern uint StCam_CreatePreviewWindowA(
            IntPtr hCamera,
            [MarshalAs(UnmanagedType.LPStr)] string pszWindowName,
            uint dwStyle,
            int lngPositionX,
            int lngPositionY,
            uint dwWidth,
            uint dwHeight,
            IntPtr hWndParent,
            IntPtr hMenu,
            int bCloseEnable);

        /// <summary>VB6: StCam_DestroyPreviewWindow(hCamera) As Long</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern uint StCam_DestroyPreviewWindow(IntPtr hCamera);

        /// <summary>VB6: StCam_SetPreviewWindowSize(hCamera, x, y, w, h) As Long
        /// プレビューウィンドウのサイズ・位置を変更する。</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern uint StCam_SetPreviewWindowSize(
            IntPtr hCamera, int lngPositionX, int lngPositionY, uint dwWidth, uint dwHeight);

        /// <summary>VB6: StCam_SetAspectMode(hCamera, byteAspectMode) As Boolean</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StCam_SetAspectMode(IntPtr hCamera, byte byteAspectMode);

        // --- シャッター / ゲイン制御 ---
        /// <summary>VB6: StCam_SetGain(hCamera, wGain As Integer) As Boolean
        /// ゲイン (輝度) を設定する。VB6 Integer = 16bit → C# ushort</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StCam_SetGain(IntPtr hCamera, ushort wGain);

        // --- 画像取得 ---
        /// <summary>VB6: StCam_GetImageSize(...) As Boolean → 画像サイズを取得する</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StCam_GetImageSize(
            IntPtr hCamera,
            ref uint pdwImageSizeMode,
            ref short pwScanMode,
            ref uint pdwOffsetX,
            ref uint pdwOffsetY,
            ref uint pdwWidth,
            ref uint pdwHeight);

        /// <summary>VB6: StCam_GetPreviewPixelFormat(hCamera, dwFormat) As Boolean</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StCam_GetPreviewPixelFormat(
            IntPtr hCamera, ref uint pdwPreviewPixelFormat);

        /// <summary>VB6: StCam_TakePreviewSnapShot(...) As Boolean → スナップショット取得</summary>
        [DllImport(STCAM_DLL, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StCam_TakePreviewSnapShot(
            IntPtr hCamera,
            [In, Out] byte[] pbyteBuffer,
            uint dwBufferSize,
            ref uint pdwNumberOfByteTrans,
            ref uint pdwFrameNo,
            uint dwMilliseconds);

        // ================================================================
        // Win32 API (プレビューウィンドウのリサイズ制御)
        // ================================================================
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(
            IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;

        // ================================================================
        // WM_STCAM_* ウィンドウメッセージ定数 (StCamD.Bas より)
        // ================================================================
        private const int WM_STCAM_TRANSFER_START = 0xB001;
        private const int WM_STCAM_TRANSFER_FINISH = 0xB002;
        private const int WM_STCAM_PREVIEW_WINDOW_CREATE = 0xB003;
        private const int WM_STCAM_PREVIEW_WINDOW_CLOSE = 0xB004;
        private const int WM_STCAM_AVI_FILE_START = 0xB008;
        private const int WM_STCAM_AVI_FILE_FINISH = 0xB009;

        // ================================================================
        // StCamD 定数
        // ================================================================
        // アスペクトモード (StCamD.Bas: STCAM_ASPECT_MODE_FIXED=0)
        private const byte STCAM_ASPECT_MODE_FIXED = 0;

        // ピクセルフォーマット (StCamD.Bas より)
        private const uint STCAM_PIXEL_FORMAT_08_MONO_OR_RAW = 0x01;
        private const uint STCAM_PIXEL_FORMAT_24_BGR = 0x04;
        private const uint STCAM_PIXEL_FORMAT_32_BGR = 0x08;

        // ウィンドウスタイル (WS_CHILD | WS_VISIBLE = 子ウィンドウとして埋め込む)
        private const uint WS_CHILD = 0x40000000;
        private const uint WS_VISIBLE = 0x10000000;

        // ================================================================
        // フィールド
        // ================================================================
        private IntPtr _hCamera = IntPtr.Zero;   // カメラハンドル
        private bool _statusTransfer = false;          // 転送中フラグ
        private bool _statusPreviewWnd = false;          // プレビューウィンドウ生成済みフラグ
        private bool _statusAviFile = false;          // AVI録画中フラグ
        private int _gain = 128;            // VB6: Int_Light

        // ================================================================
        // コントロール
        // ================================================================
        private PictureBox _picPreview;   // カメラプレビュー表示領域
        private Button _btnPreview;   // VB6: cmdPreview "開始/停止"
        private Button _btnSnap;      // VB6: cmdSnap    "キャプチャ"
        private Button _btnBack;      // VB6: btn_back   "メニュー"
        private Button _btnReset;     // VB6: Command3   "元に戻す"
        private HScrollBar _hScroll;      // VB6: HScroll1   ゲイン調整スクロールバー
        private Label _lblTitle;     // VB6: Label2     "カメラ画像取り込み"
        private Label _lblGainHint;  // VB6: Label1     "▼暗　　明△"

        // ================================================================
        // コンストラクタ
        // ================================================================
        public FormVideo()
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
            // VB6: BackColor=&H00D8FFFF&, Caption="ccd"
            this.Text = "夢仕立て - カメラモニター";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;

            // ---- プレビュー PictureBox ----
            // VB6では StCam_CreatePreviewWindowA が独立ウィンドウを作成したが、
            // C#では PictureBox を親ウィンドウとして子ウィンドウ化して埋め込む。
            _picPreview = new PictureBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Black,
                SizeMode = PictureBoxSizeMode.Normal,
            };
            _picPreview.SizeChanged += PicPreview_SizeChanged;
            this.Controls.Add(_picPreview);

            // ---- コントロールバー (下部固定) ----
            BuildControlBar();

            this.ResumeLayout(false);
        }

        // ================================================================
        // コントロールバー構築 (VB6: frmVideo の上部ボタン群を下部に配置)
        // ================================================================
        private void BuildControlBar()
        {
            var bar = new Panel
            {
                Height = 68,
                Dock = DockStyle.Bottom,
                BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF),
            };

            // タイトルラベル (VB6: Label2 "カメラ画像取り込み", Italic, Size=18)
            _lblTitle = new Label
            {
                Text = "カメラ画像取り込み",
                Font = new Font("ＭＳ Ｐゴシック", 14f, FontStyle.Italic),
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(8, 22),
            };
            bar.Controls.Add(_lblTitle);

            // 元に戻すボタン (VB6: Command3 "元に戻す")
            _btnReset = new Button
            {
                Text = "元に戻す",
                Font = new Font("ＭＳ Ｐゴシック", 9f, FontStyle.Bold),
                Size = new Size(72, 54),
                Location = new Point(220, 7),
                BackColor = Color.FromArgb(0xE0, 0xE0, 0xE0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnReset.Click += BtnReset_Click;
            bar.Controls.Add(_btnReset);

            // ゲインヒントラベル (VB6: Label1 "▼暗　　明△", Size=11.25 Bold)
            _lblGainHint = new Label
            {
                Text = "▼暗　　明△",
                Font = new Font("ＭＳ Ｐゴシック", 11.25f, FontStyle.Bold),
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(300, 24),
            };
            bar.Controls.Add(_lblGainHint);

            // ゲイン調整スクロールバー (VB6: HScroll1, 範囲 0～32767)
            // 初期値 16384 (ゲイン=128)
            _hScroll = new HScrollBar
            {
                Minimum = 0,
                Maximum = 32767,
                Value = 16384,
                SmallChange = 100,
                LargeChange = 1000,
                Size = new Size(160, 20),
                Location = new Point(460, 28),
            };
            _hScroll.ValueChanged += HScroll_ValueChanged;
            bar.Controls.Add(_hScroll);

            // 開始/停止ボタン (VB6: cmdPreview "開始")
            _btnPreview = new Button
            {
                Text = "＞　開始",
                Font = new Font("ＭＳ Ｐゴシック", 13f, FontStyle.Bold),
                Size = new Size(150, 54),
                Location = new Point(640, 7),
                BackColor = Color.FromArgb(0xC0, 0xFF, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnPreview.FlatAppearance.BorderColor = Color.SeaGreen;
            _btnPreview.FlatAppearance.BorderSize = 2;
            _btnPreview.Click += BtnPreview_Click;
            bar.Controls.Add(_btnPreview);

            // キャプチャボタン (VB6: cmdSnap "キャプチャ")
            _btnSnap = new Button
            {
                Text = "キャプチャ",
                Font = new Font("ＭＳ Ｐゴシック", 13f, FontStyle.Bold),
                Size = new Size(150, 54),
                Location = new Point(800, 7),
                BackColor = Color.FromArgb(0xFF, 0xFF, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            _btnSnap.FlatAppearance.BorderColor = Color.Goldenrod;
            _btnSnap.FlatAppearance.BorderSize = 2;
            _btnSnap.Click += BtnSnap_Click;
            bar.Controls.Add(_btnSnap);

            // メニューボタン (VB6: btn_back "メニュー", 右端)
            _btnBack = new Button
            {
                Text = "メニュー",
                Font = new Font("ＭＳ Ｐゴシック", 13f, FontStyle.Bold),
                Size = new Size(150, 54),
                BackColor = Color.FromArgb(0xFF, 0xC0, 0xC0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            };
            _btnBack.FlatAppearance.BorderColor = Color.Crimson;
            _btnBack.FlatAppearance.BorderSize = 2;
            _btnBack.Click += BtnBack_Click;
            bar.Controls.Add(_btnBack);

            bar.SizeChanged += (s, e) =>
                _btnBack.Location = new Point(bar.ClientSize.Width - 160, 7);

            this.Controls.Add(bar);
        }

        // ================================================================
        // OnLoad (VB6: Form_Load)
        // ================================================================
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.WindowState = FormWindowState.Maximized;
            FormLoad();
        }

        // ================================================================
        // VB6: Form_Load
        // カメラを開き、プレビューを自動起動する。
        // ================================================================
        private void FormLoad()
        {
            // --- カメラオープン (VB6: m_hCamera = StCam_Open(0)) ---
            try
            {
                _hCamera = StCam_Open(0);
            }
            catch (Exception ex) when (ex is BadImageFormatException || ex is DllNotFoundException)
            {
                MessageBox.Show(
                    "カメラモニタはx64ビルドでのみ使用できます。\n" +
                    "現在はx86ビルドのため起動できません。",
                    "カメラ未対応",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                this.Close();
                return;
            }
            if (_hCamera == IntPtr.Zero)
            {
                ShowCameraError(0);
                return;
            }

            _statusTransfer = false;
            _statusAviFile = false;
            _statusPreviewWnd = false;

            // --- メッセージ受信ウィンドウ設定 ---
            // VB6: m_lpPreviewWndProc = SetWindowLong(GWL_WNDPROC, AddressOf WindowProc)
            //      → C#: WndProc オーバーライドで対応するため SetWindowLong 不要
            //      StCam_SetReceiveMsgWindow で WM_STCAM_* の送信先を自フォームに指定
            uint ret = StCam_SetReceiveMsgWindow(_hCamera, this.Handle);
            if (ret == 0)
                ShowCameraError(StCam_GetLastError(_hCamera));

            // --- プレビュー開始 (Form_Load と同じ START 分岐) ---
            StartPreview();

            // --- ゲイン初期化 (VB6: Int_Light=128, HScroll1.Value=16384) ---
            _gain = 128;
            _hScroll.Value = 16384;
            StCam_SetGain(_hCamera, (ushort)_gain);

            StatusChanged();
        }

        // ================================================================
        // WndProc オーバーライド (VB6: mdlCommon.bas の WindowProc に相当)
        // WM_STCAM_* メッセージを受けてカメラ状態フラグを更新する。
        // ================================================================
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_STCAM_TRANSFER_START:
                    // VB6: frmVideo.m_bStatusTransfer = True / StatusChanged
                    _statusTransfer = true;
                    StatusChanged();
                    return;

                case WM_STCAM_TRANSFER_FINISH:
                    // VB6: m_bStatusTransfer = False / エラーがあれば表示 / StatusChanged
                    _statusTransfer = false;
                    if (m.LParam != IntPtr.Zero)
                        ShowCameraError((uint)m.LParam.ToInt64());
                    StatusChanged();
                    return;

                case WM_STCAM_PREVIEW_WINDOW_CREATE:
                    // VB6: m_bStatusPreviewWnd = True / StatusChanged
                    _statusPreviewWnd = true;
                    StatusChanged();
                    return;

                case WM_STCAM_PREVIEW_WINDOW_CLOSE:
                    // VB6: m_bStatusPreviewWnd = False / エラーがあれば表示 / StatusChanged
                    _statusPreviewWnd = false;
                    if (m.LParam != IntPtr.Zero)
                        ShowCameraError((uint)m.LParam.ToInt64());
                    StatusChanged();
                    return;

                case WM_STCAM_AVI_FILE_START:
                    _statusAviFile = true;
                    StatusChanged();
                    return;

                case WM_STCAM_AVI_FILE_FINISH:
                    _statusAviFile = false;
                    if (m.LParam != IntPtr.Zero)
                        ShowCameraError((uint)m.LParam.ToInt64());
                    StatusChanged();
                    return;
            }

            base.WndProc(ref m);
        }

        // ================================================================
        // VB6: StatusChanged
        // ボタンキャプションを転送状態に応じて切り替える。
        // ================================================================
        public void StatusChanged()
        {
            // VB6: If m_bStatusPreviewWnd And m_bStatusTransfer Then "‖　停止" Else "＞　開始"
            if (_statusPreviewWnd && _statusTransfer)
                _btnPreview.Text = "‖　停止";
            else
                _btnPreview.Text = "＞　開始";
        }

        // ================================================================
        // プレビュー開始ヘルパー (Form_Load と cmdPreview_Click の START 分岐を共通化)
        // VB6: StCam_CreatePreviewWindowA → StCam_SetAspectMode → StCam_SetGain
        //       → StCam_StartTransfer
        // ================================================================
        private void StartPreview()
        {
            if (!_statusPreviewWnd)
            {
                // プレビューウィンドウをまだ生成していない場合のみ作成
                // VB6: StCam_CreatePreviewWindowA(hCamera, "Preview",
                //          WS_OVERLAPPEDWINDOW Or WS_VISIBLE, 0,0,0,0, 0,0, clngTrue)
                // C#: WS_CHILD | WS_VISIBLE + 親=PictureBox.Handle で埋め込む
                uint w = (uint)Math.Max(_picPreview.Width, 1);
                uint h = (uint)Math.Max(_picPreview.Height, 1);
                uint ret = StCam_CreatePreviewWindowA(
                    _hCamera,
                    "Preview",
                    WS_CHILD | WS_VISIBLE,  // 子ウィンドウとして埋め込む
                    0, 0, w, h,
                    _picPreview.Handle,     // 親ウィンドウ = PictureBox
                    IntPtr.Zero,
                    0);                     // bCloseEnable=0 (子なので閉じる制御不要)

                if (ret == 0)
                {
                    ShowCameraError(StCam_GetLastError(_hCamera));
                    return;
                }

                // アスペクトモード設定 (VB6: StCam_SetAspectMode(m_hCamera, 1))
                // 1 = STCAM_ASPECT_MODE_KEEP_ASPECT だが VB6 コードは定数=1 を直接指定
                StCam_SetAspectMode(_hCamera, STCAM_ASPECT_MODE_FIXED);

                // ゲイン設定 (VB6: StCam_SetGain(m_hCamera, 100))
                StCam_SetGain(_hCamera, 100);
            }

            if (!_statusTransfer)
            {
                // 転送開始 (VB6: StCam_StartTransfer)
                uint ret = StCam_StartTransfer(_hCamera);
                if (ret == 0)
                    ShowCameraError(StCam_GetLastError(_hCamera));
            }
        }

        // ================================================================
        // イベントハンドラ
        // ================================================================

        /// <summary>
        /// 開始/停止ボタン (VB6: cmdPreview_Click)
        /// 転送中なら停止、停止中なら起動する。
        /// </summary>
        private void BtnPreview_Click(object sender, EventArgs e)
        {
            // アスペクトモード設定 (VB6: cmdPreview_Click の先頭)
            if (_hCamera != IntPtr.Zero)
                StCam_SetAspectMode(_hCamera, STCAM_ASPECT_MODE_FIXED);

            if (_statusPreviewWnd && _statusTransfer)
            {
                // --- STOP ---
                // VB6: StCam_StopTransfer(m_hCamera)
                uint ret = StCam_StopTransfer(_hCamera);
                if (ret == 0)
                    ShowCameraError(StCam_GetLastError(_hCamera));
            }
            else
            {
                // --- START ---
                StartPreview();
            }
        }

        /// <summary>
        /// キャプチャボタン (VB6: cmdSnap_Click)
        /// スナップショットを取得してFormSnapで処理し、form_save_camera へ遷移する。
        /// </summary>
        private void BtnSnap_Click(object sender, EventArgs e)
        {
            if (_hCamera == IntPtr.Zero)
            {
                MessageBox.Show("カメラが接続されていません。", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // VB6: form_wait.Show / Me.Visible = False
            this.Hide();

            uint dwLastErrorNo = 0;
            bool bReval = false;

            // GetImageSize
            uint dwImageSizeMode = 0;
            short wScanMode = 0;
            uint dwOffsetX = 0;
            uint dwOffsetY = 0;
            uint dwWidth = 0;
            uint dwHeight = 0;

            // GetPreviewPixelFormat
            uint dwPreviewPixelFormat = 0;

            // バッファ
            byte[] pbyteImageBuffer = null;
            uint dwBufferSize = 0;
            uint dwNumberOfByteTrans = 0;
            uint dwFrameNo = 0;

            bool success = false;

            do
            {
                // VB6: bReval = StCam_GetImageSize(...)
                bReval = StCam_GetImageSize(
                    _hCamera, ref dwImageSizeMode, ref wScanMode,
                    ref dwOffsetX, ref dwOffsetY, ref dwWidth, ref dwHeight);
                if (!bReval) { dwLastErrorNo = StCam_GetLastError(_hCamera); break; }

                // VB6: bReval = StCam_GetPreviewPixelFormat(...)
                bReval = StCam_GetPreviewPixelFormat(_hCamera, ref dwPreviewPixelFormat);
                if (!bReval) { dwLastErrorNo = StCam_GetLastError(_hCamera); break; }

                // バッファサイズ計算 (VB6: Select Case dwPreviewPixelFormat)
                dwBufferSize = dwWidth * dwHeight;
                if (dwPreviewPixelFormat == STCAM_PIXEL_FORMAT_24_BGR)
                    dwBufferSize *= 3;
                else if (dwPreviewPixelFormat == STCAM_PIXEL_FORMAT_32_BGR)
                    dwBufferSize *= 4;
                // MONO/RAW: ×1 のまま

                // VB6: ReDim pbyteImageBuffer(dwBufferSize)
                pbyteImageBuffer = new byte[dwBufferSize];

                // スナップショット取得 (VB6: dwMilliseconds=1000)
                bReval = StCam_TakePreviewSnapShot(
                    _hCamera, pbyteImageBuffer,
                    dwBufferSize, ref dwNumberOfByteTrans, ref dwFrameNo,
                    1000);
                if (!bReval) { dwLastErrorNo = StCam_GetLastError(_hCamera); break; }

                success = true;
            } while (false);

            if (!success)
            {
                // VB6: ShowErrorMsg / form_wait Unload / Me.Visible = True
                ShowCameraError(dwLastErrorNo);
                this.Show();
                return;
            }

            // ================================================================
            // VB6: frmSnap.form_init() / frmSnap.bUpdateSnapShot(...) / frmSnap.Save_Image()
            //      → C# では FormSnap に相当するクラスを呼び出す
            //      FormSnap は別途作成が必要。暫定的にインラインで処理する。
            // ================================================================
            try
            {
                // FormSnap の処理を呼び出す (FormSnap は別ファイルで実装)
                // VB6: bReval = frmSnap.form_init()
                //      bReval = frmSnap.bUpdateSnapShot(m_hCamera, dwWidth, dwHeight,
                //                   dwPreviewPixelFormat, pbyteImageBuffer, dwLastErrorNo)
                //      bReval = frmSnap.Save_Image()
                using (var snapForm = new FormSnap())
                {
                    snapForm.FormInit();
                    snapForm.UpdateSnapShot(
                        _hCamera, dwWidth, dwHeight,
                        dwPreviewPixelFormat, pbyteImageBuffer,
                        ref dwLastErrorNo);
                    snapForm.SaveImage();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("スナップショット処理エラー：" + ex.Message,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Show();
                return;
            }

            // VB6: Unload Me / Unload frmSnap / Unload form_wait
            //       form_save_camera.Show
            CloseCamera();
            this.Close();

            // form_save_camera.Show に相当 (FormSaveCamera は別ファイルで実装)
            var saveForm = new FormSaveCamera();
            saveForm.Show();
        }

        /// <summary>
        /// 元に戻すボタン (VB6: Command3_Click)
        /// ゲインをデフォルト値 128 にリセットし、スクロールバーを中央に戻す。
        /// </summary>
        private void BtnReset_Click(object sender, EventArgs e)
        {
            // VB6: Int_Light = 128 / HScroll1.Value = 16384
            _gain = 128;
            _hScroll.Value = 16384;

            if (_hCamera != IntPtr.Zero)
                StCam_SetGain(_hCamera, (ushort)_gain);
        }

        /// <summary>
        /// ゲイン調整スクロールバー (VB6: HScroll1_Change)
        /// Value (0～32767) をゲイン (0～255) に変換して StCam_SetGain を呼び出す。
        /// VB6: Int_Light = HScroll1.Value / 32767 * 255
        /// </summary>
        private void HScroll_ValueChanged(object sender, EventArgs e)
        {
            // VB6: / は浮動小数演算 → double で計算して切り捨て
            _gain = (int)((double)_hScroll.Value / 32767.0 * 255.0);

            if (_hCamera != IntPtr.Zero)
                StCam_SetGain(_hCamera, (ushort)_gain);
        }

        /// <summary>
        /// メニューへ戻るボタン (VB6: btn_back_Click → Unload Me / form_menu.Visible=True)
        /// </summary>
        private void BtnBack_Click(object sender, EventArgs e)
        {
            CloseCamera();

            // VB6: form_menu.Visible = True
            foreach (Form f in Application.OpenForms)
            {
                if (f is FormMenu) { f.Show(); break; }
            }

            this.Close();
        }

        /// <summary>
        /// PictureBox サイズ変更時にプレビューウィンドウをリサイズする。
        /// VB6では独立ウィンドウだったため不要だったが、子埋め込みでは必要。
        /// </summary>
        private void PicPreview_SizeChanged(object sender, EventArgs e)
        {
            ResizePreviewWindow();
        }

        /// <summary>
        /// プレビューウィンドウのサイズを PictureBox に合わせる。
        /// StCam_SetPreviewWindowSize または SetWindowPos で制御する。
        /// </summary>
        private void ResizePreviewWindow()
        {
            if (_hCamera == IntPtr.Zero || !_statusPreviewWnd) return;

            int w = Math.Max(_picPreview.Width, 1);
            int h = Math.Max(_picPreview.Height, 1);

            // プレビューウィンドウサイズを更新
            StCam_SetPreviewWindowSize(_hCamera, 0, 0, (uint)w, (uint)h);
        }

        // ================================================================
        // フォームクローズ処理 (VB6: Form_QueryUnload / Form_Unload)
        // ================================================================
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // VB6: Form_QueryUnload → SetWindowLong で WndProc を元に戻す
            //      (C#では WndProc オーバーライドのため不要)
            // VB6: Form_Unload → If m_hCamera Then StCam_Close(m_hCamera)
            CloseCamera();
            base.OnFormClosing(e);
        }

        /// <summary>カメラハンドルを閉じる共通処理。</summary>
        private void CloseCamera()
        {
            if (_hCamera == IntPtr.Zero) return;

            // 転送中なら停止
            if (_statusTransfer)
                StCam_StopTransfer(_hCamera);

            // プレビューウィンドウ破棄
            if (_statusPreviewWnd)
                StCam_DestroyPreviewWindow(_hCamera);

            StCam_Close(_hCamera);
            _hCamera = IntPtr.Zero;
        }

        // ================================================================
        // エラー表示ヘルパー (VB6: ShowErrorMsg)
        // StCamMsg.dll からエラーメッセージを取得して表示する。
        // ================================================================
        private static void ShowCameraError(uint dwErrorCode)
        {
            if (dwErrorCode == 0) return;

            // 既知エラーコードの簡易マッピング
            string msg;
            switch (dwErrorCode)
            {
                case 0xE0000001: msg = "カメラが見つかりません。"; break;
                case 0xE0000002: msg = "すべてのカメラが既に開かれています。"; break;
                case 0xE0000003: msg = "カメラハンドルが無効です。"; break;
                case 0xE0000006: msg = "プレビューウィンドウは既に作成されています。"; break;
                case 0xE0000007: msg = "プレビューウィンドウが存在しません。"; break;
                default: msg = $"エラーコード: 0x{dwErrorCode:X8}"; break;
            }

            MessageBox.Show(
                msg + "\r\nCCDカメラのＵＳＢケーブルがパソコンにつながっているか確認して下さい。",
                "ＣＣＤカメラ接続エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
