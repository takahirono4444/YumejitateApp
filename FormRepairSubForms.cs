using System;
using System.Drawing;
using System.Windows.Forms;

// ================================================================
// FormRepairSubForms.cs
// VB6: form_sizenaoshi / form_ishidome / form_shintsume /
//      form_henkei / form_brand_new の C# 移植
// 修理見積もりサブフォーム群（独立実装・抽象基底クラスなし）
// ================================================================

namespace YumejitateApp
{
    // ============================================================
    // RepairConst: 選択肢定数＋共通価格計算（static）
    // ============================================================
    internal static class RepairConst
    {
        internal static readonly string[] Atumi =
        {
            "-0.9mm", "〜1.5mm", "1.6-2.2mm",
            "2.3mm〜(別途お見積もり)", "不明"
        };
        internal static readonly string[] Haba =
        {
            "-1.8mm", "〜2.5mm", "2.6-3.2mm", "3.3-3.9mm",
            "4.0-5.0mm", "5.1-6.0mm", "6.1mm〜(別途お見積もり)", "不明"
        };
        internal static readonly string[] Kinzoku =
        {
            "イエローゴールド", "プラチナ", "ホワイトゴールド",
            "ピンクゴールド", "コンビ", "シルバー", "その他", "不明"
        };
        internal static readonly string[] Ishitsuki =
        {
            "付いていません",
            "付いている：ダイヤ、ルビー、サファイヤ",
            "付いている：ダイヤ、ルビー、サファイヤ　の大粒石（約２CtUP）",
            "付いている：エメラルド、オパール、有機石、カボッション石及びダメージを受け易い石",
            "その他の石"
        };
        internal static readonly string[] Brand =
        {
            "ノンブランド（国内）製品", "ノンブランド（海外）製品",
            "国内ブランド製品", "海外ブランド製品"
        };
        internal static readonly string[] Ude = { "セットされていない", "セットされている" };
        internal static readonly string[] Chojobu = { "ロー", "ハイ" };
        internal static readonly string[] Hyomen = { "施されていない", "施されている" };
        internal static readonly string[] Uchigawa = { "ない", "ある" };

        // --------------------------------------------------------
        // 共通価格計算（VB6: 各フォームの Btn_Gaisan_Click 共通部分）
        // --------------------------------------------------------
        internal static bool CalcCommon(
            double baseRate,
            string atumi, string haba, string kinzoku,
            string ishitsuki, string brand,
            string ude, string chojobu, string hyomen, string uchigawa,
            out double total, out bool gaisan, out string error)
        {
            double c = baseRate;
            double bk = 3500.0;
            double ic = 0;
            double br = 0;
            gaisan = false;
            error = "";

            switch (atumi)
            {
                case "〜1.5mm":
                    c *= 1.3; bk *= 1.3; break;
                case "1.6-2.2mm":
                    c *= 1.69; bk *= 1.3; break;
                case "2.3mm〜(別途お見積もり)":
                    c *= 1.69; bk *= 1.3; gaisan = true; break;
                case "不明":
                    error = "指輪の厚みが不明の為、お見積り出来ません。";
                    total = 0; return false;
            }

            switch (haba)
            {
                case "〜2.5mm": c *= 1.2; bk *= 1.2; break;
                case "2.6-3.2mm": c *= 1.4; bk *= 1.4; break;
                case "3.3-3.9mm": c *= 1.6; bk *= 1.6; break;
                case "4.0-5.0mm": c *= 1.8; bk *= 1.8; break;
                case "5.1-6.0mm": c *= 2.0; bk *= 2.0; break;
                case "6.1mm〜(別途お見積もり)":
                    c *= 2.7; bk *= 2.7; gaisan = true; break;
                case "不明":
                    error = "指輪の幅が不明の為、お見積り出来ません。";
                    total = 0; return false;
            }

            if (kinzoku == "プラチナ" || kinzoku == "ホワイトゴールド" ||
                kinzoku == "ピンクゴールド" || kinzoku == "その他")
            { c *= 1.2; bk *= 1.2; }
            else if (kinzoku == "コンビ")
            { ic += bk * 0.3; c *= 1.2; bk *= 1.2; }
            else if (kinzoku == "不明")
            { error = "金属が不明の為、お見積り出来ません。"; total = 0; return false; }

            if (brand == "海外ブランド製品" || brand == "国内ブランド製品" ||
                brand == "ノンブランド（海外）製品")
                br = bk * 0.6;

            if (ishitsuki == "付いている：ダイヤ、ルビー、サファイヤ")
                ic += bk * 0.3;
            else if (ishitsuki ==
                "付いている：エメラルド、オパール、有機石、カボッション石及びダメージを受け易い石")
                ic += bk * 1.2;
            else if (ishitsuki == "付いている：ダイヤ、ルビー、サファイヤ　の大粒石（約２CtUP）" ||
                     ishitsuki == "その他の石")
                ic += bk * 0.6;

            if (ude != null && ude != "セットされていない") br += bk * 0.6;
            if (chojobu != null && chojobu == "ハイ") br += bk * 0.6;
            if (hyomen != null && hyomen == "施されている") ic += bk * 0.3;
            if (uchigawa != null && uchigawa == "ある") c += 3500;

            c += ic + br;
            total = RoundUp100(c);
            return true;
        }

        internal static double RoundUp100(double v)
        {
            long r = (long)Math.Round(v) % 100;
            return (r != 0) ? (long)Math.Round(v) - r + 100 : Math.Round(v);
        }

        internal static string CommonAttention()
        {
            return "\r\n\r\n※最終的な加工可能か否かの判断は、職人が致します。" +
                   "\r\n\r\n※加工不可能と判断した場合は、お戻しさせて頂きます。ご了承下さい。" +
                   "\r\n\r\n※サシメッキ等の特殊メッキは別途見積もりとなります。";
        }

        internal static string BuildAttention(string atumi, string haba)
        {
            string s = "";
            if (atumi.Contains("別途"))
                s += "\r\n※リングの厚さが２．２ｍｍ以上の場合は、別途お見積もりとなります。";
            if (haba.Contains("別途"))
                s += "\r\n※リングの幅が６．０ｍｍ以上の場合は、別途お見積もりとなります。";
            return s;
        }

        // --------------------------------------------------------
        // UI生成ヘルパー
        // --------------------------------------------------------
        internal static GroupBox MakeGrp(Font fnt, string title, int x, int y, int w, int h)
        {
            var g = new GroupBox();
            g.Text = title;
            g.Font = fnt;
            g.BackColor = Color.Transparent;
            g.Location = new Point(x, y);
            g.Size = new Size(w, h);
            return g;
        }

        internal static void Row(GroupBox grp, Font fnt, string text,
            int lx, int cx, ref int row, int rh, out ComboBox cmb,
            string[] items, int w = 300, int sel = 0)
        {
            var lbl = new Label();
            lbl.Text = text;
            lbl.Font = fnt;
            lbl.BackColor = Color.Transparent;
            lbl.Location = new Point(lx, row + 4);
            lbl.Size = new Size(cx - lx - 8, 26);
            lbl.AutoSize = false;
            grp.Controls.Add(lbl);

            cmb = new ComboBox();
            cmb.Font = fnt;
            cmb.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb.Location = new Point(cx, row);
            cmb.Size = new Size(w, 27);
            cmb.Items.AddRange(items);
            if (sel < cmb.Items.Count) cmb.SelectedIndex = sel;
            grp.Controls.Add(cmb);

            row += rh;
        }
    }

    // ============================================================
    // FormSizanaoshi: サイズ直し
    // ============================================================
    public class FormSizanaoshi : Form
    {
        private ComboBox _cmbCurrentSize, _cmbHopeSize;
        private ComboBox _cmbAtumi, _cmbHaba, _cmbKinzoku, _cmbIshitsuki;
        private ComboBox _cmbStoneShape, _cmbSetMethod, _cmbBrand;
        private Label _lblResult;
        private TextBox _txtHelp;

        public FormSizanaoshi()
        {
            this.Text = "夢仕立て - サイズ直し";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            var fnt = new Font("ＭＳ Ｐゴシック", 12f, FontStyle.Bold);
            var fntS = new Font("ＭＳ Ｐゴシック", 10f);

            var grp = RepairConst.MakeGrp(fnt, "入力項目", 16, 50, 760, 20);
            this.Controls.Add(grp);
            int lx = 10, cx = 270, row = 20, rh = 42;

            string[] sizeItems = BuildSizeItems();
            RepairConst.Row(grp, fnt, "現在のリングサイズ", lx, cx, ref row, rh, out _cmbCurrentSize, sizeItems);
            RepairConst.Row(grp, fnt, "ご希望のリングサイズ", lx, cx, ref row, rh, out _cmbHopeSize, sizeItems);
            RepairConst.Row(grp, fnt, "リング厚み", lx, cx, ref row, rh, out _cmbAtumi, RepairConst.Atumi);
            RepairConst.Row(grp, fnt, "リング幅", lx, cx, ref row, rh, out _cmbHaba, RepairConst.Haba);
            RepairConst.Row(grp, fnt, "お手持ちのリングの金属", lx, cx, ref row, rh, out _cmbKinzoku, RepairConst.Kinzoku);
            RepairConst.Row(grp, fnt, "リング石付き", lx, cx, ref row, rh, out _cmbIshitsuki, RepairConst.Ishitsuki, 480);
            RepairConst.Row(grp, fnt, "メインストーンの形状", lx, cx, ref row, rh, out _cmbStoneShape,
                new string[] { "珠", "ラウンド", "オーバル", "エメラルド", "マーキース", "ドロップ", "その他" });
            RepairConst.Row(grp, fnt, "メインストーンの石留方法", lx, cx, ref row, rh, out _cmbSetMethod,
                new string[] { "爪留め", "ビール留め", "接着：珠のみ", "その他" });
            RepairConst.Row(grp, fnt, "ブランド", lx, cx, ref row, rh, out _cmbBrand, RepairConst.Brand);
            grp.Height = row + 16;

            int oy = grp.Bottom + 10;
            _lblResult = new Label();
            _lblResult.Font = new Font("ＭＳ Ｐゴシック", 14f, FontStyle.Bold);
            _lblResult.ForeColor = Color.DarkBlue;
            _lblResult.BackColor = Color.Transparent;
            _lblResult.AutoSize = true;
            _lblResult.Location = new Point(16, oy);
            this.Controls.Add(_lblResult);

            _txtHelp = new TextBox();
            _txtHelp.Multiline = true;
            _txtHelp.ScrollBars = ScrollBars.Vertical;
            _txtHelp.ReadOnly = true;
            _txtHelp.Font = fntS;
            _txtHelp.BackColor = Color.White;
            _txtHelp.Location = new Point(16, oy + 40);
            _txtHelp.Size = new Size(900, 180);
            this.Controls.Add(_txtHelp);

            var btnCalc = new Button();
            btnCalc.Text = "見積もり計算";
            btnCalc.Font = fnt;
            btnCalc.Location = new Point(16, oy + 230);
            btnCalc.Size = new Size(180, 50);
            btnCalc.BackColor = Color.LightSkyBlue;
            btnCalc.Click += BtnCalc_Click;
            this.Controls.Add(btnCalc);

            var btnBack = new Button();
            btnBack.Text = "戻る";
            btnBack.Font = fnt;
            btnBack.Location = new Point(210, oy + 230);
            btnBack.Size = new Size(130, 50);
            btnBack.BackColor = Color.FromArgb(255, 192, 192);
            btnBack.Click += (s, ev) => this.Close();
            this.Controls.Add(btnBack);
        }

        private string[] BuildSizeItems()
        {
            var list = new System.Collections.Generic.List<string>();
            list.Add("不明");
            for (int i = 1; i <= 30; i++) list.Add(i.ToString());
            for (int i = 1; i <= 30; i++) list.Add(i + ".5");
            return list.ToArray();
        }

        private void BtnCalc_Click(object sender, EventArgs e)
        {
            _txtHelp.Text = ""; _lblResult.Text = "";

            if (_cmbStoneShape.Text != "珠" && _cmbSetMethod.Text == "接着：珠のみ")
            {
                _txtHelp.ForeColor = Color.Red;
                _txtHelp.Text = "メインストーンの石留方法が「接着」の場合は、メインストーンの形状が「珠」の場合となります。";
                return;
            }

            double currentSize = 0, hopeSize = 0;
            double.TryParse(_cmbCurrentSize.Text, out currentSize);
            double.TryParse(_cmbHopeSize.Text, out hopeSize);

            double repairSize = (_cmbCurrentSize.Text == "不明" || _cmbHopeSize.Text == "不明")
                ? 1 : hopeSize - currentSize;

            if (repairSize == 0)
            {
                _txtHelp.ForeColor = Color.Red;
                _txtHelp.Text = "サイズの変更がありません。";
                return;
            }

            bool isUp = (repairSize > 0);
            double chijime = isUp ? 3500 * 1.2 : 3500;
            double nobashi = 945;
            bool gaisan = false;
            string error = "";

            switch (_cmbAtumi.Text)
            {
                case "〜1.5mm": chijime *= 1.3; nobashi *= 1.2; break;
                case "1.6-2.2mm": chijime *= 1.69; nobashi *= 1.44; break;
                case "2.3mm〜(別途お見積もり)": chijime *= 1.69; nobashi *= 1.44; gaisan = true; break;
                case "不明": error = "指輪の厚みが不明の為、お見積り出来ません。"; break;
            }
            if (error != "") { _txtHelp.ForeColor = Color.Red; _txtHelp.Text = error; return; }

            switch (_cmbHaba.Text)
            {
                case "〜2.5mm": chijime *= 1.2; nobashi *= 1.5; break;
                case "2.6-3.2mm": chijime *= 1.4; nobashi *= 1.8; break;
                case "3.3-3.9mm": chijime *= 1.6; nobashi *= 2.1; break;
                case "4.0-5.0mm": chijime *= 1.8; nobashi *= 2.5; break;
                case "5.1-6.0mm": chijime *= 2.0; nobashi *= 3.0; break;
                case "6.1mm〜(別途お見積もり)": chijime *= 2.7; nobashi *= 4.5; gaisan = true; break;
                case "不明": error = "指輪の幅が不明の為、お見積り出来ません。"; break;
            }
            if (error != "") { _txtHelp.ForeColor = Color.Red; _txtHelp.Text = error; return; }

            double ic = 0;
            string kin = _cmbKinzoku.Text;
            if (kin == "プラチナ" || kin == "ホワイトゴールド" ||
                kin == "ピンクゴールド" || kin == "その他")
            { chijime *= 1.2; nobashi *= 1.2; }
            else if (kin == "コンビ")
            { ic += chijime * 0.3; chijime *= 1.2; nobashi *= 1.2; }
            else if (kin == "不明")
            { _txtHelp.ForeColor = Color.Red; _txtHelp.Text = "金属が不明の為、お見積り出来ません。"; return; }

            string ishi = _cmbIshitsuki.Text;
            if (isUp)
            {
                if (ishi == "付いている：ダイヤ、ルビー、サファイヤ")
                    ic += chijime * 0.4;
                else if (ishi ==
                    "付いている：エメラルド、オパール、有機石、カボッション石及びダメージを受け易い石")
                    ic += chijime * 1.2;
                else if (ishi != "付いていません") ic += chijime * 0.8;
            }
            else
            {
                if (ishi == "付いている：ダイヤ、ルビー、サファイヤ")
                    ic += chijime * 0.3;
                else if (ishi ==
                    "付いている：エメラルド、オパール、有機石、カボッション石及びダメージを受け易い石")
                    ic += chijime * 0.9;
                else if (ishi != "付いていません") ic += chijime * 0.6;
            }

            string br_str = _cmbBrand.Text;
            double br = 0;
            if (br_str == "海外ブランド製品" || br_str == "国内ブランド製品" ||
                br_str == "ノンブランド（海外）製品")
                br = chijime * 0.6;

            double total = RepairConst.RoundUp100(chijime + ic + br);
            string attention = RepairConst.BuildAttention(_cmbAtumi.Text, _cmbHaba.Text);
            bool hb = attention.Contains("別途お見積もり");
            _txtHelp.ForeColor = hb ? Color.Red : Color.Black;
            if (hb) attention += "\r\n\r\n※特殊加工に付き、加工不可能な場合がございます。";
            _txtHelp.Text = attention + RepairConst.CommonAttention();
            _lblResult.Text = (gaisan ? "概算：" : "") + "¥" + total.ToString("#,##0") + "（税別）";
        }
    }

    // ============================================================
    // FormIshidome: 石留め/石外し
    // ============================================================
    public class FormIshidome : Form
    {
        private ComboBox _cmbAtumi, _cmbHaba, _cmbKinzoku, _cmbIshitsuki, _cmbBrand;
        private ComboBox _cmbUde, _cmbChojobu, _cmbHyomen, _cmbUchigawa;
        private ComboBox _cmbStoneSize, _cmbStoneShape, _cmbStoneKosuu, _cmbTomekata;
        private Label _lblResult;
        private TextBox _txtHelp;

        public FormIshidome()
        {
            this.Text = "夢仕立て - 石留め/石外し";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            var fnt = new Font("ＭＳ Ｐゴシック", 12f, FontStyle.Bold);
            var fntS = new Font("ＭＳ Ｐゴシック", 10f);

            var grp = RepairConst.MakeGrp(fnt, "入力項目", 16, 50, 760, 20);
            this.Controls.Add(grp);
            int lx = 10, cx = 270, row = 20, rh = 42;

            RepairConst.Row(grp, fnt, "リング厚み", lx, cx, ref row, rh, out _cmbAtumi, RepairConst.Atumi);
            RepairConst.Row(grp, fnt, "リング幅", lx, cx, ref row, rh, out _cmbHaba, RepairConst.Haba);
            RepairConst.Row(grp, fnt, "お手持ちのリングの金属", lx, cx, ref row, rh, out _cmbKinzoku, RepairConst.Kinzoku);
            RepairConst.Row(grp, fnt, "リング石付き（既存）", lx, cx, ref row, rh, out _cmbIshitsuki, RepairConst.Ishitsuki, 480);
            RepairConst.Row(grp, fnt, "ブランド", lx, cx, ref row, rh, out _cmbBrand, RepairConst.Brand);
            RepairConst.Row(grp, fnt, "腕三分の一の石飾り", lx, cx, ref row, rh, out _cmbUde, RepairConst.Ude);
            RepairConst.Row(grp, fnt, "頂上部の形", lx, cx, ref row, rh, out _cmbChojobu, RepairConst.Chojobu);
            RepairConst.Row(grp, fnt, "表面模様", lx, cx, ref row, rh, out _cmbHyomen, RepairConst.Hyomen);
            RepairConst.Row(grp, fnt, "リング内側彫刻", lx, cx, ref row, rh, out _cmbUchigawa, RepairConst.Uchigawa);
            RepairConst.Row(grp, fnt, "お石の大きさ", lx, cx, ref row, rh, out _cmbStoneSize,
                new string[] { "0.1Ct以下", "0.2Ct以下", "0.3Ct以下", "0.5Ct以下", "0.7Ct以下", "1.0Ct以下", "1.5Ct以下", "2.0Ct以下" });
            RepairConst.Row(grp, fnt, "石の形状", lx, cx, ref row, rh, out _cmbStoneShape,
                new string[] { "ラウンド", "オーバル", "エメラルド", "マーキース", "ドロップ", "珠", "その他" });
            RepairConst.Row(grp, fnt, "石の個数", lx, cx, ref row, rh, out _cmbStoneKosuu,
                new string[] { "1個", "2個", "3個", "4個", "5個", "6個", "7個", "8個", "9個", "10個以上" });
            RepairConst.Row(grp, fnt, "石の留め方", lx, cx, ref row, rh, out _cmbTomekata,
                new string[] { "通常タイプ", "接着タイプ" });
            grp.Height = row + 16;

            BuildOutputControls(fnt, fntS, grp.Bottom - 32);
        }

        private void BuildOutputControls(Font fnt, Font fntS, int oy)
        {
            _lblResult = new Label();
            _lblResult.Font = new Font("ＭＳ Ｐゴシック", 14f, FontStyle.Bold);
            _lblResult.ForeColor = Color.DarkBlue;
            _lblResult.BackColor = Color.Transparent;
            _lblResult.AutoSize = true;
            _lblResult.Location = new Point(16, oy);
            this.Controls.Add(_lblResult);

            _txtHelp = new TextBox();
            _txtHelp.Multiline = true;
            _txtHelp.ScrollBars = ScrollBars.Vertical;
            _txtHelp.ReadOnly = true;
            _txtHelp.Font = fntS;
            _txtHelp.BackColor = Color.White;
            _txtHelp.Location = new Point(16, oy + 40);
            _txtHelp.Size = new Size(900, 180);
            this.Controls.Add(_txtHelp);

            var btnCalc = new Button();
            btnCalc.Text = "見積もり計算";
            btnCalc.Font = fnt;
            btnCalc.Location = new Point(16, oy + 230);
            btnCalc.Size = new Size(180, 50);
            btnCalc.BackColor = Color.LightSkyBlue;
            btnCalc.Click += BtnCalc_Click;
            this.Controls.Add(btnCalc);

            var btnBack = new Button();
            btnBack.Text = "戻る";
            btnBack.Font = fnt;
            btnBack.Location = new Point(210, oy + 230);
            btnBack.Size = new Size(130, 50);
            btnBack.BackColor = Color.FromArgb(255, 192, 192);
            btnBack.Click += (s, ev) => this.Close();
            this.Controls.Add(btnBack);
        }

        private void BtnCalc_Click(object sender, EventArgs e)
        {
            _txtHelp.Text = ""; _lblResult.Text = "";

            bool ok = RepairConst.CalcCommon(3500.0 * 0.8,
                _cmbAtumi.Text, _cmbHaba.Text, _cmbKinzoku.Text,
                _cmbIshitsuki.Text, _cmbBrand.Text,
                _cmbUde.Text, _cmbChojobu.Text, _cmbHyomen.Text, _cmbUchigawa.Text,
                out double total, out bool gaisan, out string error);
            if (!ok) { _txtHelp.ForeColor = Color.Red; _txtHelp.Text = error; return; }

            double stoneCost = 0;
            try
            {
                string sizeStr = _cmbStoneSize.Text.Replace("Ct以下", "");
                double.TryParse(sizeStr, out double ctSize);
                var dt = AppState.Db.ExecuteQuery(
                    "SELECT * FROM [石留め石外しテーブル] WHERE [大きさ]=" + ctSize.ToString() +
                    " AND [形状]='" + _cmbStoneShape.Text + "'");
                if (dt.Rows.Count > 0)
                    double.TryParse(dt.Rows[0]["値段"].ToString(), out stoneCost);
            }
            catch { }

            string kosuuStr = _cmbStoneKosuu.Text.Replace("個", "").Replace("以上", "");
            int.TryParse(kosuuStr, out int kosuu);
            if (_cmbStoneKosuu.Text == "10個以上") kosuu = 10;

            if (kosuu >= 10)
            {
                _txtHelp.ForeColor = Color.Red;
                _txtHelp.Text = "石の個数が１０個以上の為、別途見積もりとなります。";
                return;
            }

            double add = (_cmbTomekata.Text == "接着タイプ")
                ? stoneCost * kosuu * 0.5
                : stoneCost * kosuu;
            total = RepairConst.RoundUp100(total + add);

            string attention = RepairConst.BuildAttention(_cmbAtumi.Text, _cmbHaba.Text);
            attention += "\r\n\r\n※当社でご用意する石の代金は含まれておりません。別途見積もりとなります。";
            bool hb = attention.Contains("別途お見積もり");
            _txtHelp.ForeColor = hb ? Color.Red : Color.Black;
            if (hb) attention += "\r\n\r\n※特殊加工に付き、加工不可能な場合がございます。";
            _txtHelp.Text = attention + RepairConst.CommonAttention();
            _lblResult.Text = (gaisan ? "概算：" : "") + "¥" + total.ToString("#,##0") + "（税別）";
        }
    }

    // ============================================================
    // FormShintsume: 芯爪立て替えと石留め
    // ============================================================
    public class FormShintsume : Form
    {
        private ComboBox _cmbAtumi, _cmbHaba, _cmbKinzoku, _cmbIshitsuki, _cmbBrand;
        private ComboBox _cmbUde, _cmbChojobu, _cmbHyomen, _cmbUchigawa;
        private ComboBox _cmbStoneSize, _cmbTsumeHonsu, _cmbTsumeMethod;
        private Label _lblResult;
        private TextBox _txtHelp;

        public FormShintsume()
        {
            this.Text = "夢仕立て - 芯爪立て替えと石留め";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            var fnt = new Font("ＭＳ Ｐゴシック", 12f, FontStyle.Bold);
            var fntS = new Font("ＭＳ Ｐゴシック", 10f);

            var grp = RepairConst.MakeGrp(fnt, "入力項目", 16, 50, 760, 20);
            this.Controls.Add(grp);
            int lx = 10, cx = 270, row = 20, rh = 42;

            RepairConst.Row(grp, fnt, "リング厚み", lx, cx, ref row, rh, out _cmbAtumi, RepairConst.Atumi);
            RepairConst.Row(grp, fnt, "リング幅", lx, cx, ref row, rh, out _cmbHaba, RepairConst.Haba);
            RepairConst.Row(grp, fnt, "お手持ちのリングの金属", lx, cx, ref row, rh, out _cmbKinzoku, RepairConst.Kinzoku);
            RepairConst.Row(grp, fnt, "リング石付き", lx, cx, ref row, rh, out _cmbIshitsuki, RepairConst.Ishitsuki, 480);
            RepairConst.Row(grp, fnt, "ブランド", lx, cx, ref row, rh, out _cmbBrand, RepairConst.Brand);
            RepairConst.Row(grp, fnt, "腕三分の一の石飾り", lx, cx, ref row, rh, out _cmbUde, RepairConst.Ude);
            RepairConst.Row(grp, fnt, "頂上部の形", lx, cx, ref row, rh, out _cmbChojobu, RepairConst.Chojobu);
            RepairConst.Row(grp, fnt, "表面模様", lx, cx, ref row, rh, out _cmbHyomen, RepairConst.Hyomen);
            RepairConst.Row(grp, fnt, "リング内側彫刻", lx, cx, ref row, rh, out _cmbUchigawa, RepairConst.Uchigawa);
            RepairConst.Row(grp, fnt, "お石の大きさ", lx, cx, ref row, rh, out _cmbStoneSize,
                new string[] { "0.1Ct以下", "0.2Ct以下", "0.3Ct以下", "0.5Ct以下", "0.7Ct以下", "1.0Ct以下", "1.5Ct以下", "2.0Ct以下" });
            RepairConst.Row(grp, fnt, "芯爪の本数", lx, cx, ref row, rh, out _cmbTsumeHonsu,
                new string[] { "1本", "2本", "3本", "4本", "5本", "6本", "7本", "8本", "9本", "10本以上" });
            RepairConst.Row(grp, fnt, "芯爪の接合方法", lx, cx, ref row, rh, out _cmbTsumeMethod,
                new string[] { "通常ロー付け", "レーザーロー付け" });
            grp.Height = row + 16;

            int oy = grp.Bottom + 10;
            _lblResult = new Label();
            _lblResult.Font = new Font("ＭＳ Ｐゴシック", 14f, FontStyle.Bold);
            _lblResult.ForeColor = Color.DarkBlue;
            _lblResult.BackColor = Color.Transparent;
            _lblResult.AutoSize = true;
            _lblResult.Location = new Point(16, oy);
            this.Controls.Add(_lblResult);

            _txtHelp = new TextBox();
            _txtHelp.Multiline = true;
            _txtHelp.ScrollBars = ScrollBars.Vertical;
            _txtHelp.ReadOnly = true;
            _txtHelp.Font = fntS;
            _txtHelp.BackColor = Color.White;
            _txtHelp.Location = new Point(16, oy + 40);
            _txtHelp.Size = new Size(900, 180);
            this.Controls.Add(_txtHelp);

            var btnCalc = new Button();
            btnCalc.Text = "見積もり計算";
            btnCalc.Font = fnt;
            btnCalc.Location = new Point(16, oy + 230);
            btnCalc.Size = new Size(180, 50);
            btnCalc.BackColor = Color.LightSkyBlue;
            btnCalc.Click += BtnCalc_Click;
            this.Controls.Add(btnCalc);

            var btnBack = new Button();
            btnBack.Text = "戻る";
            btnBack.Font = fnt;
            btnBack.Location = new Point(210, oy + 230);
            btnBack.Size = new Size(130, 50);
            btnBack.BackColor = Color.FromArgb(255, 192, 192);
            btnBack.Click += (s, ev) => this.Close();
            this.Controls.Add(btnBack);
        }

        private void BtnCalc_Click(object sender, EventArgs e)
        {
            _txtHelp.Text = ""; _lblResult.Text = "";

            bool ok = RepairConst.CalcCommon(3500.0 * 0.8,
                _cmbAtumi.Text, _cmbHaba.Text, _cmbKinzoku.Text,
                _cmbIshitsuki.Text, _cmbBrand.Text,
                _cmbUde.Text, _cmbChojobu.Text, _cmbHyomen.Text, _cmbUchigawa.Text,
                out double total, out bool gaisan, out string error);
            if (!ok) { _txtHelp.ForeColor = Color.Red; _txtHelp.Text = error; return; }

            string kinshitsu = "Pt";
            switch (_cmbKinzoku.Text)
            {
                case "プラチナ": kinshitsu = "Pt"; break;
                case "ホワイトゴールド": kinshitsu = "WG"; break;
                case "イエローゴールド": kinshitsu = "K18"; break;
                case "ピンクゴールド": kinshitsu = "Pt"; break;
                case "コンビ": kinshitsu = "Pt"; break;
                case "シルバー": kinshitsu = "K18"; break;
            }

            double tsumeCost = 0;
            try
            {
                string sizeStr = _cmbStoneSize.Text.Replace("Ct以下", "");
                double.TryParse(sizeStr, out double ctSize);
                var dt = AppState.Db.ExecuteQuery(
                    "SELECT * FROM [芯立て替えと石留めテーブル] WHERE [大きさ]=" + ctSize.ToString() +
                    " AND [金質]='" + kinshitsu + "'");
                if (dt.Rows.Count > 0)
                    double.TryParse(dt.Rows[0]["値段"].ToString(), out tsumeCost);
            }
            catch { }

            string honsuStr = _cmbTsumeHonsu.Text.Replace("本", "").Replace("以上", "");
            int.TryParse(honsuStr, out int honsu);
            if (_cmbTsumeHonsu.Text == "10本以上") honsu = 10;

            if (honsu >= 10)
            {
                _txtHelp.ForeColor = Color.Red;
                _txtHelp.Text = "爪の本数が１０本以上の為、別途見積もりとなります。";
                return;
            }

            double add = (_cmbTsumeMethod.Text == "レーザーロー付け")
                ? tsumeCost * honsu * 1.5
                : tsumeCost * honsu;
            total = RepairConst.RoundUp100(total + add);

            string attention = RepairConst.BuildAttention(_cmbAtumi.Text, _cmbHaba.Text);
            bool hb = attention.Contains("別途お見積もり");
            _txtHelp.ForeColor = hb ? Color.Red : Color.Black;
            if (hb) attention += "\r\n\r\n※特殊加工に付き、加工不可能な場合がございます。";
            _txtHelp.Text = attention + RepairConst.CommonAttention();
            _lblResult.Text = (gaisan ? "概算：" : "") + "¥" + total.ToString("#,##0") + "（税別）";
        }
    }

    // ============================================================
    // FormHenkei: 変形修理（基本料金: 3500 × 2 × 0.8）
    // ============================================================
    public class FormHenkei : Form
    {
        private ComboBox _cmbAtumi, _cmbHaba, _cmbKinzoku, _cmbIshitsuki, _cmbBrand;
        private ComboBox _cmbUde, _cmbChojobu, _cmbHyomen, _cmbUchigawa;
        private Label _lblResult;
        private TextBox _txtHelp;

        public FormHenkei()
        {
            this.Text = "夢仕立て - 変形修理";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            var fnt = new Font("ＭＳ Ｐゴシック", 12f, FontStyle.Bold);
            var fntS = new Font("ＭＳ Ｐゴシック", 10f);

            var grp = RepairConst.MakeGrp(fnt, "入力項目", 16, 50, 760, 20);
            this.Controls.Add(grp);
            int lx = 10, cx = 270, row = 20, rh = 42;

            RepairConst.Row(grp, fnt, "リング厚み", lx, cx, ref row, rh, out _cmbAtumi, RepairConst.Atumi);
            RepairConst.Row(grp, fnt, "リング幅", lx, cx, ref row, rh, out _cmbHaba, RepairConst.Haba);
            RepairConst.Row(grp, fnt, "お手持ちのリングの金属", lx, cx, ref row, rh, out _cmbKinzoku, RepairConst.Kinzoku);
            RepairConst.Row(grp, fnt, "リング石付き", lx, cx, ref row, rh, out _cmbIshitsuki, RepairConst.Ishitsuki, 480);
            RepairConst.Row(grp, fnt, "ブランド", lx, cx, ref row, rh, out _cmbBrand, RepairConst.Brand);
            RepairConst.Row(grp, fnt, "腕三分の一の石飾り", lx, cx, ref row, rh, out _cmbUde, RepairConst.Ude);
            RepairConst.Row(grp, fnt, "頂上部の形", lx, cx, ref row, rh, out _cmbChojobu, RepairConst.Chojobu);
            RepairConst.Row(grp, fnt, "表面模様", lx, cx, ref row, rh, out _cmbHyomen, RepairConst.Hyomen);
            RepairConst.Row(grp, fnt, "リング内側彫刻", lx, cx, ref row, rh, out _cmbUchigawa, RepairConst.Uchigawa);
            grp.Height = row + 16;

            int oy = grp.Bottom + 10;
            _lblResult = new Label();
            _lblResult.Font = new Font("ＭＳ Ｐゴシック", 14f, FontStyle.Bold);
            _lblResult.ForeColor = Color.DarkBlue;
            _lblResult.BackColor = Color.Transparent;
            _lblResult.AutoSize = true;
            _lblResult.Location = new Point(16, oy);
            this.Controls.Add(_lblResult);

            _txtHelp = new TextBox();
            _txtHelp.Multiline = true;
            _txtHelp.ScrollBars = ScrollBars.Vertical;
            _txtHelp.ReadOnly = true;
            _txtHelp.Font = fntS;
            _txtHelp.BackColor = Color.White;
            _txtHelp.Location = new Point(16, oy + 40);
            _txtHelp.Size = new Size(900, 180);
            this.Controls.Add(_txtHelp);

            var btnCalc = new Button();
            btnCalc.Text = "見積もり計算";
            btnCalc.Font = fnt;
            btnCalc.Location = new Point(16, oy + 230);
            btnCalc.Size = new Size(180, 50);
            btnCalc.BackColor = Color.LightSkyBlue;
            btnCalc.Click += BtnCalc_Click;
            this.Controls.Add(btnCalc);

            var btnBack = new Button();
            btnBack.Text = "戻る";
            btnBack.Font = fnt;
            btnBack.Location = new Point(210, oy + 230);
            btnBack.Size = new Size(130, 50);
            btnBack.BackColor = Color.FromArgb(255, 192, 192);
            btnBack.Click += (s, ev) => this.Close();
            this.Controls.Add(btnBack);
        }

        private void BtnCalc_Click(object sender, EventArgs e)
        {
            _txtHelp.Text = ""; _lblResult.Text = "";
            bool ok = RepairConst.CalcCommon(3500.0 * 2.0 * 0.8,
                _cmbAtumi.Text, _cmbHaba.Text, _cmbKinzoku.Text,
                _cmbIshitsuki.Text, _cmbBrand.Text,
                _cmbUde.Text, _cmbChojobu.Text, _cmbHyomen.Text, _cmbUchigawa.Text,
                out double total, out bool gaisan, out string error);
            if (!ok) { _txtHelp.ForeColor = Color.Red; _txtHelp.Text = error; return; }

            string attention = RepairConst.BuildAttention(_cmbAtumi.Text, _cmbHaba.Text);
            bool hb = attention.Contains("別途お見積もり");
            _txtHelp.ForeColor = hb ? Color.Red : Color.Black;
            if (hb) attention += "\r\n\r\n※特殊加工に付き、加工不可能な場合がございます。";
            _txtHelp.Text = attention + RepairConst.CommonAttention();
            _lblResult.Text = (gaisan ? "概算：" : "") + "¥" + total.ToString("#,##0") + "（税別）";
        }
    }

    // ============================================================
    // FormBrandNew: 新品仕上げ（基本料金: 3500 × 0.8）
    // ============================================================
    public class FormBrandNew : Form
    {
        private ComboBox _cmbAtumi, _cmbHaba, _cmbKinzoku, _cmbIshitsuki, _cmbBrand;
        private ComboBox _cmbUde, _cmbChojobu, _cmbHyomen, _cmbUchigawa;
        private Label _lblResult;
        private TextBox _txtHelp;

        public FormBrandNew()
        {
            this.Text = "夢仕立て - 新品仕上げ";
            this.BackColor = Color.FromArgb(0xD8, 0xFF, 0xFF);
            this.WindowState = FormWindowState.Maximized;
            var fnt = new Font("ＭＳ Ｐゴシック", 12f, FontStyle.Bold);
            var fntS = new Font("ＭＳ Ｐゴシック", 10f);

            var grp = RepairConst.MakeGrp(fnt, "入力項目", 16, 50, 760, 20);
            this.Controls.Add(grp);
            int lx = 10, cx = 270, row = 20, rh = 42;

            RepairConst.Row(grp, fnt, "リング厚み", lx, cx, ref row, rh, out _cmbAtumi, RepairConst.Atumi);
            RepairConst.Row(grp, fnt, "リング幅", lx, cx, ref row, rh, out _cmbHaba, RepairConst.Haba);
            RepairConst.Row(grp, fnt, "お手持ちのリングの金属", lx, cx, ref row, rh, out _cmbKinzoku, RepairConst.Kinzoku);
            RepairConst.Row(grp, fnt, "リング石付き", lx, cx, ref row, rh, out _cmbIshitsuki, RepairConst.Ishitsuki, 480);
            RepairConst.Row(grp, fnt, "ブランド", lx, cx, ref row, rh, out _cmbBrand, RepairConst.Brand);
            RepairConst.Row(grp, fnt, "腕三分の一の石飾り", lx, cx, ref row, rh, out _cmbUde, RepairConst.Ude);
            RepairConst.Row(grp, fnt, "頂上部の形", lx, cx, ref row, rh, out _cmbChojobu, RepairConst.Chojobu);
            RepairConst.Row(grp, fnt, "表面模様", lx, cx, ref row, rh, out _cmbHyomen, RepairConst.Hyomen);
            RepairConst.Row(grp, fnt, "リング内側彫刻", lx, cx, ref row, rh, out _cmbUchigawa, RepairConst.Uchigawa);
            grp.Height = row + 16;

            int oy = grp.Bottom + 10;
            _lblResult = new Label();
            _lblResult.Font = new Font("ＭＳ Ｐゴシック", 14f, FontStyle.Bold);
            _lblResult.ForeColor = Color.DarkBlue;
            _lblResult.BackColor = Color.Transparent;
            _lblResult.AutoSize = true;
            _lblResult.Location = new Point(16, oy);
            this.Controls.Add(_lblResult);

            _txtHelp = new TextBox();
            _txtHelp.Multiline = true;
            _txtHelp.ScrollBars = ScrollBars.Vertical;
            _txtHelp.ReadOnly = true;
            _txtHelp.Font = fntS;
            _txtHelp.BackColor = Color.White;
            _txtHelp.Location = new Point(16, oy + 40);
            _txtHelp.Size = new Size(900, 200);
            this.Controls.Add(_txtHelp);

            var btnCalc = new Button();
            btnCalc.Text = "見積もり計算";
            btnCalc.Font = fnt;
            btnCalc.Location = new Point(16, oy + 250);
            btnCalc.Size = new Size(180, 50);
            btnCalc.BackColor = Color.LightSkyBlue;
            btnCalc.Click += BtnCalc_Click;
            this.Controls.Add(btnCalc);

            var btnBack = new Button();
            btnBack.Text = "戻る";
            btnBack.Font = fnt;
            btnBack.Location = new Point(210, oy + 250);
            btnBack.Size = new Size(130, 50);
            btnBack.BackColor = Color.FromArgb(255, 192, 192);
            btnBack.Click += (s, ev) => this.Close();
            this.Controls.Add(btnBack);
        }

        private void BtnCalc_Click(object sender, EventArgs e)
        {
            _txtHelp.Text = ""; _lblResult.Text = "";
            bool ok = RepairConst.CalcCommon(3500.0 * 0.8,
                _cmbAtumi.Text, _cmbHaba.Text, _cmbKinzoku.Text,
                _cmbIshitsuki.Text, _cmbBrand.Text,
                _cmbUde.Text, _cmbChojobu.Text, _cmbHyomen.Text, _cmbUchigawa.Text,
                out double total, out bool gaisan, out string error);
            if (!ok) { _txtHelp.ForeColor = Color.Red; _txtHelp.Text = error; return; }

            string attention = RepairConst.BuildAttention(_cmbAtumi.Text, _cmbHaba.Text);
            attention += "\r\n新品仕上げの場合、職人の判断でメッキを併用する場合があります。";
            attention += "\r\n\r\nメッキを併用しない場合は仕上時間が多くかかります。従ってメッキを併用しない場合でも、価格の変更はありません（同じ価格です）。";
            attention += "\r\n\r\nシルバーの場合でメッキをご希望されない場合でも、価格は同じです。";

            bool hb = attention.Contains("別途お見積もり");
            _txtHelp.ForeColor = hb ? Color.Red : Color.Black;
            if (hb) attention += "\r\n\r\n※特殊加工に付き、加工不可能な場合がございます。";
            _txtHelp.Text = attention + RepairConst.CommonAttention();
            _lblResult.Text = (gaisan ? "概算：" : "") + "¥" + total.ToString("#,##0") + "（税別）";
        }
    }
}