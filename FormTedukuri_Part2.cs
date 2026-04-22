using System;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Text;
using System.Windows.Forms;

namespace YumejitateApp
{
    // ================================================================
    // FormTedukuri_Part2.cs
    // VB6: form_tedukuri.frm → C# + WinForms 移植（Part 2 / 2）
    // オーダーメイドお見積り画面 — ボタンハンドラ・計算メソッド・DB操作
    // ================================================================
    public partial class FormTedukuri : Form
    {
        // ============================================================
        // イベントハンドラ登録（InitializeComponent から呼ばれる）
        // ============================================================
        private void WireEvents()
        {
            _btnKeisan.Click += (s, e) => BtnKeisan_Click();
            _btnClear.Click += (s, e) => BtnClear_Click();
            _btnPrint.Click += (s, e) => BtnPrint_Click();
            _btnBack.Click += (s, e) => BtnBack_Click();
            _btnMenu.Click += (s, e) => BtnMenu_Click();
            _btnSave.Click += (s, e) => BtnSave_Click();
        }

        // ============================================================
        // ■ 計算ボタン（VB6: btn_keisan_Click）
        // ============================================================
        private void BtnKeisan_Click()
        {
            double dispPrice, dispPt900, dispK18, dispWgpg;
            double zentaiKakeritu, dblIngo = 0;

            // 累計変数リセット
            _chkFlgKeisan = false;
            _orderPrice = 0;
            _orderPt900 = 0;
            _orderK18 = 0;
            _orderWgpg = 0;

            // 結果ラベルをデフォルト色に戻す
            _lblPrice.BackColor = CDefault;
            _lblPt900.BackColor = CDefault;
            _lblK18.BackColor = CDefault;
            _lblWgpg.BackColor = CDefault;
            _lblIngo.BackColor = CDefault;
            _cKakouA1.BackColor = CDefault;
            _cKakouB1.BackColor = CDefault;

            // 15 個の計算サブルーチンを順に呼ぶ
            UdeKeisan1Rtn();
            UdeKeisan2Rtn();
            ItaSenzai1Rtn();
            ItaSenzai2Rtn();
            Ishiza1Rtn();
            Ishiza2Rtn();
            Ishiza3Rtn();
            PaipuShaton1Rtn();
            PaipuShaton2Rtn();
            PaipuShaton3Rtn();
            PaipuShaton4Rtn();
            Mere1Rtn();
            Mere2Rtn();
            Mere3Rtn();
            Mere4Rtn();
            MenRowRtn();
            TenRowRtn();
            KakouGradeRtn();

            if (_chkFlgKeisan) return;

            // 参照金額（lbl_sample_price）を加算
            double.TryParse(_lblSmpPrice.Text.Replace(",", "").Replace("¥", "").Trim(), out double smpPrice);
            dispPrice = _orderPrice + smpPrice;

            // ■ 全体掛け率取得
            zentaiKakeritu = GetZentaiKakeritu();
            if (zentaiKakeritu == 0)
            {
                MessageBox.Show("システムエラー", "加工グレード計算：掛け率テーブル取得",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            zentaiKakeritu = zentaiKakeritu / 10.0;

            // ■ 加工難易度（VB6: 加工グレードA1）
            string nanidoText = _cKakouA1.Text;
            if (nanidoText == "部品合計")
            {
                // 加算式なし
            }
            else if (nanidoText == "96・95：部品集合＆整形")
            {
                dblIngo = (5250 + 0.08 * dispPrice) * zentaiKakeritu;
                dispPrice += dblIngo;
                _cKakouA1.BackColor = CCalc;
            }
            else if (nanidoText == "94・93：部分手作り")
            {
                dblIngo = (10500 + 0.08 * dispPrice) * zentaiKakeritu;
                dispPrice += dblIngo;
                _cKakouA1.BackColor = CCalc;
            }
            else if (nanidoText == "92：手作り(難易度1)")
            {
                dblIngo = (21000 + 0.08 * dispPrice) * zentaiKakeritu;
                dispPrice += dblIngo;
                _cKakouA1.BackColor = CCalc;
            }
            else if (nanidoText == "91：手作り(難易度2)")
            {
                dblIngo = (31500 + 0.08 * dispPrice) * zentaiKakeritu;
                dispPrice += dblIngo;
                _cKakouA1.BackColor = CCalc;
            }
            else if (nanidoText == "90：手作り(難易度3)")
            {
                dblIngo = (47250 + 0.08 * dispPrice) * zentaiKakeritu;
                dispPrice += dblIngo;
                _cKakouA1.BackColor = CCalc;
            }
            else if (nanidoText == "89：手作り(特注)")
            {
                dblIngo = (73500 + 0.08 * dispPrice) * zentaiKakeritu;
                dispPrice += dblIngo;
                _cKakouA1.BackColor = CCalc;
            }
            else
            {
                MessageBox.Show("システムエラー：加工グレード取得");
            }

            // ■ 加工グレード（VB6: 加工グレードB1）
            string gradeText = _cKakouB1.Text;
            if (gradeText == "A")
            {
                // 掛け率なし
            }
            else if (gradeText == "SS") { dblIngo *= 1.24; dispPrice *= 1.24; }
            else if (gradeText == "S") { dblIngo *= 1.12; dispPrice *= 1.12; }
            else if (gradeText == "AS") { dblIngo *= 1.06; dispPrice *= 1.06; }
            else if (gradeText == "AB") { dblIngo *= 0.97; dispPrice *= 0.97; }
            else if (gradeText == "B") { dblIngo *= 0.93; dispPrice *= 0.93; }
            else
            {
                MessageBox.Show("システムエラー加工グレード計算");
                return;
            }

            // ■ 地金重量の集計（参照データ分を加算）
            double.TryParse(_lblSmpPt900.Text.Replace("Pt900:", "").Replace("g", "").Trim(), out double smpPt900);
            double.TryParse(_lblSmpK18.Text.Replace("K18:", "").Replace("g", "").Trim(), out double smpK18);
            double.TryParse(_lblSmpWgpg.Text.Replace("WG/PG:", "").Replace("g", "").Trim(), out double smpWgpg);

            dispPt900 = _orderPt900 + smpPt900;
            dispK18 = _orderK18 + smpK18;
            dispWgpg = _orderWgpg + smpWgpg;

            // 部品合計 以外は ×1.08
            if (nanidoText != "部品合計")
            {
                dispPt900 *= 1.08;
                dispK18 *= 1.08;
                dispWgpg *= 1.08;
            }

            // ■ 結果ラベルに表示（消費税 1.1 倍）
            _lblPrice.Text = ((long)(dispPrice * 1.1)).ToString("#,##0");
            _lblPt900.Text = dispPt900.ToString("#0.00");
            _lblK18.Text = dispK18.ToString("#0.00");
            _lblWgpg.Text = dispWgpg.ToString("#0.00");
            // 隠語（VB6: Right(CStr(Rnd(1)),1) & "LK" & Format(dblIngo/10000,"#0")）
            _lblIngo.Text = new Random().Next(0, 10).ToString()
                           + "LK" + Math.Round(dblIngo / 10000, 0).ToString("F0");

            // ■ カラーリング
            if (dispPrice != 0)
            {
                _lblPrice.BackColor = CCalc;
                _lblIngo.BackColor = CCalc;
                _cKakouB1.BackColor = CCalc;
            }
            if (dispPt900 != 0) _lblPt900.BackColor = CCalc;
            if (dispK18 != 0) _lblK18.BackColor = CCalc;
            if (dispWgpg != 0) _lblWgpg.BackColor = CCalc;
        }

        // ============================================================
        // ■ クリアボタン（VB6: btn_clear_Click）
        // ============================================================
        private void BtnClear_Click()
        {
            // 数値ドロップダウンをすべてリセット（先頭要素 = "0"）
            ResetIndex(_cU1TJu, _cU1TJi, _cU1TJs);
            ResetIndex(_cU1TAJu, _cU1TAJi, _cU1TAJs);
            ResetIndex(_cU1BJu, _cU1BJi, _cU1BJs);
            ResetIndex(_cU1BAJu, _cU1BAJi, _cU1BAJs);
            ResetIndex(_cU1SzJu, _cU1SzJi, _cU1Pm, _cU1F);

            ResetIndex(_cU2TJu, _cU2TJi, _cU2TJs);
            ResetIndex(_cU2TAJu, _cU2TAJi, _cU2TAJs);
            ResetIndex(_cU2BJu, _cU2BJi, _cU2BJs);
            ResetIndex(_cU2BAJu, _cU2BAJi, _cU2BAJs);
            ResetIndex(_cU2SzJu, _cU2SzJi, _cU2Pm, _cU2F);

            ResetIndex(_cIt1LJu, _cIt1LJi, _cIt1LJs);
            ResetIndex(_cIt1SJu, _cIt1SJi, _cIt1SJs);
            ResetIndex(_cIt1AJu, _cIt1AJi, _cIt1AJs, _cIt1F);

            ResetIndex(_cIt2LJu, _cIt2LJi, _cIt2LJs);
            ResetIndex(_cIt2SJu, _cIt2SJi, _cIt2SJs);
            ResetIndex(_cIt2AJu, _cIt2AJi, _cIt2AJs, _cIt2F);

            ResetIndex(_cIsz1LJu, _cIsz1LJi, _cIsz1LJs);
            ResetIndex(_cIsz1SJu, _cIsz1SJi, _cIsz1SJs, _cIsz1F);
            ResetIndex(_cIsz2LJu, _cIsz2LJi, _cIsz2LJs);
            ResetIndex(_cIsz2SJu, _cIsz2SJi, _cIsz2SJs, _cIsz2F);
            ResetIndex(_cIsz3LJu, _cIsz3LJi, _cIsz3LJs);
            ResetIndex(_cIsz3SJu, _cIsz3SJi, _cIsz3SJs, _cIsz3F);

            ResetIndex(_cIsm1EJu, _cIsm1EJi);
            ResetIndex(_cIsm2EJu, _cIsm2EJi);
            ResetIndex(_cIsm3EJu, _cIsm3EJi);
            ResetIndex(_cIsm4EJu, _cIsm4EJi);

            ResetIndex(_cDia1Ju, _cDia1Ji);
            ResetIndex(_cDia2Ju, _cDia2Ji);
            ResetIndex(_cDia3Ju, _cDia3Ji);
            ResetIndex(_cDia4Ju, _cDia4Ji);

            ResetIndex(_cRoB1Ju, _cRoB1Ji);
            ResetIndex(_cRoD1Ju, _cRoD1Ji);

            // 種別コンボをリセット
            ResetIndex(_cU1A, _cU1B, _cU1D);
            ResetIndex(_cU2A, _cU2B, _cU2D);

            ResetIndex(_cIt1A, _cIt1B, _cIt1C);
            ResetIndex(_cIt2A, _cIt2B, _cIt2C);

            ResetIndex(_cIsz1A, _cIsz1B, _cIsz1D, _cIsz1E, _cIsz1Pm);
            ResetIndex(_cIsz2A, _cIsz2B, _cIsz2D, _cIsz2E, _cIsz2Pm);
            ResetIndex(_cIsz3A, _cIsz3B, _cIsz3D, _cIsz3E, _cIsz3Pm);

            ResetIndex(_cIsm1A, _cIsm1B, _cIsm1C, _cIsm1D);
            ResetIndex(_cIsm2A, _cIsm2B, _cIsm2C, _cIsm2D);
            ResetIndex(_cIsm3A, _cIsm3B, _cIsm3C, _cIsm3D);
            ResetIndex(_cIsm4A, _cIsm4B, _cIsm4C, _cIsm4D);

            ResetIndex(_cDia1A, _cDia1B, _cDia1Kigo);
            ResetIndex(_cDia2A, _cDia2B, _cDia2Kigo);
            ResetIndex(_cDia3A, _cDia3B, _cDia3Kigo);
            ResetIndex(_cDia4A, _cDia4B, _cDia4Kigo);

            ResetIndex(_cRoA1, _cRoC1);
            ResetIndex(_cKakouA1, _cKakouB1);

            // 結果ラベル初期化
            _lblPt900.Text = "0.00";
            _lblK18.Text = "0.00";
            _lblWgpg.Text = "0.00";
            _lblPrice.Text = "0";
            _lblIngo.Text = new Random().Next(0, 10).ToString() + "LK0";

            BtnKeisan_Click();
        }

        // ============================================================
        // ■ 保存ボタン（VB6: btn_save_Click）
        // ============================================================
        private void BtnSave_Click()
        {
            // まず計算を実行
            BtnKeisan_Click();
            if (_chkFlgKeisan) return;

            // 作業書番号を組み立て
            string sagyousyoNum =
                _cmb保1.Text + _cmb保2.Text + _cmb保3.Text +
                _cmb保4.Text + _cmb保5.Text + _cmb保6.Text + _cmb保7.Text;

            // 重複チェック
            var dtCheck = AppState.Db.ExecuteQuery(
                "SELECT * FROM [オーダーメイドテーブル] WHERE 作業書番号 = '" + sagyousyoNum + "'");
            if (dtCheck.Rows.Count > 0)
            {
                MessageBox.Show(
                    "作業書番号『" + sagyousyoNum + "』は既に保存されています。",
                    "オーダーメイド保存処理エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show(
                "作業書番号『" + sagyousyoNum + "』で、画面データを保存します。よろしいですか？",
                "データ作成確認", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            // INSERT 文を構築（VB6 の動作を忠実に再現）
            var sb = new StringBuilder();
            sb.Append("INSERT INTO [オーダーメイドテーブル] (");
            sb.Append("作業書番号,参照品番,参照PT900,参照K18,参照WGPG,参照CODE,参照金額,");
            sb.Append("腕地金種類1,腕形状1,腕天幅1,腕天厚1,腕底幅1,腕底厚1,腕特殊加工1,腕リングサイズ1,腕プラマイ1,腕本数1,");
            sb.Append("腕地金種類2,腕形状2,腕天幅2,腕天厚2,腕底幅2,腕底厚2,腕特殊加工2,腕リングサイズ2,腕プラマイ2,腕本数2,");
            sb.Append("板線材部材種類1,板線材地金種類1,板線材形状1,板線材長径1,板線材短径1,板線材厚さ1,板線材個数1,");
            sb.Append("板線材部材種類2,板線材地金種類2,板線材形状2,板線材長径2,板線材短径2,板線材厚さ2,板線材個数2,");
            sb.Append("石座地金種類1,石座石の形状1,石座石の長径1,石座石の短径1,石座主石座の種類1,石座腰高1,石座プラマイ1,石座個数1,");
            sb.Append("石座地金種類2,石座石の形状2,石座石の長径2,石座石の短径2,石座主石座の種類2,石座腰高2,石座プラマイ2,石座個数2,");
            sb.Append("石座地金種類3,石座石の形状3,石座石の長径3,石座石の短径3,石座主石座の種類3,石座腰高3,石座プラマイ3,石座個数3,");
            sb.Append("石留め留め方法1,石留め地金種類1,石留め石の形状1,石留め石のサイズ1,石留め個数1,");
            sb.Append("石留め留め方法2,石留め地金種類2,石留め石の形状2,石留め石のサイズ2,石留め個数2,");
            sb.Append("石留め留め方法3,石留め地金種類3,石留め石の形状3,石留め石のサイズ3,石留め個数3,");
            sb.Append("石留め留め方法4,石留め地金種類4,石留め石の形状4,石留め石のサイズ4,石留め個数4,");
            sb.Append("ダイヤグレード1,ダイヤサイズ1,ダイヤプラマイ1,ダイヤ個数1,");
            sb.Append("ダイヤグレード2,ダイヤサイズ2,ダイヤプラマイ2,ダイヤ個数2,");
            sb.Append("ダイヤグレード3,ダイヤサイズ3,ダイヤプラマイ3,ダイヤ個数3,");
            sb.Append("ダイヤグレード4,ダイヤサイズ4,ダイヤプラマイ4,ダイヤ個数4,");
            sb.Append("ロー付け種類1,ロー付け個数1,ロー付け種類2,ロー付け個数2,");
            sb.Append("加工難易度,加工グレード,合計PT900,合計K18,合計WGPG,合計金額,隠語");
            sb.Append(") VALUES (");

            // 値の追加（文字列は ' でエスケープ）
            sb.Append(Q(sagyousyoNum));
            // 参照データ
            sb.Append(Q(_lblSmpHinban.Text));
            sb.Append(Q(_lblSmpPt900.Text));
            sb.Append(Q(_lblSmpK18.Text));
            sb.Append(Q(_lblSmpWgpg.Text));
            sb.Append(Q(_lblSmpCode.Text));
            sb.Append(Q(_lblSmpPrice.Text));
            // 腕1
            sb.Append(Q(_cU1A.Text));
            sb.Append(Q(_cU1B.Text));
            sb.Append(Q(D3(_cU1TJu, _cU1TJi, _cU1TJs)));
            sb.Append(Q(D3(_cU1TAJu, _cU1TAJi, _cU1TAJs)));
            sb.Append(Q(D3(_cU1BJu, _cU1BJi, _cU1BJs)));
            sb.Append(Q(D3(_cU1BAJu, _cU1BAJi, _cU1BAJs)));
            sb.Append(Q(_cU1D.Text));
            sb.Append(Q(_cU1SzJu.Text + _cU1SzJi.Text));
            sb.Append(Q(_cU1Pm.Text));
            sb.Append(Q(_cU1F.Text));
            // 腕2
            sb.Append(Q(_cU2A.Text));
            sb.Append(Q(_cU2B.Text));
            sb.Append(Q(D3(_cU2TJu, _cU2TJi, _cU2TJs)));
            sb.Append(Q(D3(_cU2TAJu, _cU2TAJi, _cU2TAJs)));
            sb.Append(Q(D3(_cU2BJu, _cU2BJi, _cU2BJs)));
            sb.Append(Q(D3(_cU2BAJu, _cU2BAJi, _cU2BAJs)));
            sb.Append(Q(_cU2D.Text));
            sb.Append(Q(_cU2SzJu.Text + _cU2SzJi.Text));
            sb.Append(Q(_cU2Pm.Text));
            sb.Append(Q(_cU2F.Text));
            // 板線材1
            sb.Append(Q(_cIt1A.Text));
            sb.Append(Q(_cIt1B.Text));
            sb.Append(Q(_cIt1C.Text));
            sb.Append(Q(D3(_cIt1LJu, _cIt1LJi, _cIt1LJs)));
            sb.Append(Q(D3(_cIt1SJu, _cIt1SJi, _cIt1SJs)));
            sb.Append(Q(D3(_cIt1AJu, _cIt1AJi, _cIt1AJs)));
            sb.Append(Q(_cIt1F.Text));
            // 板線材2
            sb.Append(Q(_cIt2A.Text));
            sb.Append(Q(_cIt2B.Text));
            sb.Append(Q(_cIt2C.Text));
            sb.Append(Q(D3(_cIt2LJu, _cIt2LJi, _cIt2LJs)));
            sb.Append(Q(D3(_cIt2SJu, _cIt2SJi, _cIt2SJs)));
            sb.Append(Q(D3(_cIt2AJu, _cIt2AJi, _cIt2AJs)));
            sb.Append(Q(_cIt2F.Text));
            // 石座1
            sb.Append(Q(_cIsz1A.Text));
            sb.Append(Q(_cIsz1B.Text));
            sb.Append(Q(D3(_cIsz1LJu, _cIsz1LJi, _cIsz1LJs)));
            sb.Append(Q(D3(_cIsz1SJu, _cIsz1SJi, _cIsz1SJs)));
            sb.Append(Q(_cIsz1D.Text));
            sb.Append(Q(_cIsz1E.Text));
            sb.Append(Q(_cIsz1Pm.Text));
            sb.Append(Q(_cIsz1F.Text));
            // 石座2
            sb.Append(Q(_cIsz2A.Text));
            sb.Append(Q(_cIsz2B.Text));
            sb.Append(Q(D3(_cIsz2LJu, _cIsz2LJi, _cIsz2LJs)));
            sb.Append(Q(D3(_cIsz2SJu, _cIsz2SJi, _cIsz2SJs)));
            sb.Append(Q(_cIsz2D.Text));
            sb.Append(Q(_cIsz2E.Text));
            sb.Append(Q(_cIsz2Pm.Text));
            sb.Append(Q(_cIsz2F.Text));
            // 石座3
            sb.Append(Q(_cIsz3A.Text));
            sb.Append(Q(_cIsz3B.Text));
            sb.Append(Q(D3(_cIsz3LJu, _cIsz3LJi, _cIsz3LJs)));
            sb.Append(Q(D3(_cIsz3SJu, _cIsz3SJi, _cIsz3SJs)));
            sb.Append(Q(_cIsz3D.Text));
            sb.Append(Q(_cIsz3E.Text));
            sb.Append(Q(_cIsz3Pm.Text));
            sb.Append(Q(_cIsz3F.Text));
            // 石留め1〜4
            AppendIsm(sb, _cIsm1D, _cIsm1A, _cIsm1B, _cIsm1C, _cIsm1EJu, _cIsm1EJi);
            AppendIsm(sb, _cIsm2D, _cIsm2A, _cIsm2B, _cIsm2C, _cIsm2EJu, _cIsm2EJi);
            AppendIsm(sb, _cIsm3D, _cIsm3A, _cIsm3B, _cIsm3C, _cIsm3EJu, _cIsm3EJi);
            AppendIsm(sb, _cIsm4D, _cIsm4A, _cIsm4B, _cIsm4C, _cIsm4EJu, _cIsm4EJi, last: true);
            // ダイヤ1〜4
            sb.Append(Q(_cDia1A.Text)); sb.Append(Q(_cDia1B.Text)); sb.Append(Q(_cDia1Kigo.Text)); sb.Append(Q(_cDia1Ju.Text + _cDia1Ji.Text));
            sb.Append(Q(_cDia2A.Text)); sb.Append(Q(_cDia2B.Text)); sb.Append(Q(_cDia2Kigo.Text)); sb.Append(Q(_cDia2Ju.Text + _cDia2Ji.Text));
            sb.Append(Q(_cDia3A.Text)); sb.Append(Q(_cDia3B.Text)); sb.Append(Q(_cDia3Kigo.Text)); sb.Append(Q(_cDia3Ju.Text + _cDia3Ji.Text));
            sb.Append(Q(_cDia4A.Text)); sb.Append(Q(_cDia4B.Text)); sb.Append(Q(_cDia4Kigo.Text)); sb.Append(Q(_cDia4Ju.Text + _cDia4Ji.Text));
            // ロー付け
            sb.Append(Q(_cRoA1.Text)); sb.Append(Q(_cRoB1Ju.Text + _cRoB1Ji.Text));
            sb.Append(Q(_cRoC1.Text)); sb.Append(Q(_cRoD1Ju.Text + _cRoD1Ji.Text));
            // 加工難易度・グレード
            sb.Append(Q(_cKakouA1.Text));
            sb.Append(Q(_cKakouB1.Text));
            // 合計データ（最後はカンマなし）
            sb.Append(Q(_lblPt900.Text));
            sb.Append(Q(_lblK18.Text));
            sb.Append(Q(_lblWgpg.Text));
            sb.Append(Q(_lblPrice.Text));
            sb.Append("'" + _lblIngo.Text.Replace("'", "''") + "'");
            sb.Append(")");

            AppState.Db.ExecuteNonQuery(sb.ToString());
            MessageBox.Show("保存しました。", "オーダーメイドデータ保存完了",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ============================================================
        // ■ 印刷ボタン（VB6: btn_print_Click）
        // ============================================================
        private void BtnPrint_Click()
        {
            // A4 縦向き確認
            var r1 = MessageBox.Show(
                "プリンターの設定をA4縦向きに設定していますか。",
                "プリンター設定確認", MessageBoxButtons.YesNoCancel);
            if (r1 == DialogResult.Cancel) return;
            // No の場合はプリンター設定ダイアログ（省略）

            if (MessageBox.Show("印刷しますか？", "印刷確認", MessageBoxButtons.YesNo)
                == DialogResult.No) return;

            DoPrint();
        }

        // ■ 実際の印刷処理（VB6: Form_Print → PrintDocument に置き換え）
        private void DoPrint()
        {
            var lines = BuildPrintLines();

            using (var pd = new PrintDocument())
            {
                pd.DefaultPageSettings.PaperSize =
                    new PaperSize("A4", 827, 1169); // 1/100インチ単位
                pd.DefaultPageSettings.Landscape = false;

                int lineIndex = 0;
                float yPos = 0;
                float margin = 40f;
                float lineH = 0;

                pd.PrintPage += (s, e) =>
                {
                    var g = e.Graphics;
                    var fontTitle = new Font("MS PGothic", 20, FontStyle.Bold);
                    var fontBody = new Font("MS PGothic", 11, FontStyle.Regular);
                    yPos = margin;
                    lineH = fontBody.GetHeight(g);

                    foreach (var line in lines)
                    {
                        Font f = line.StartsWith("■") || line.StartsWith("オーダーメイド")
                            ? fontTitle : fontBody;
                        g.DrawString(line, f, Brushes.Black, margin, yPos);
                        yPos += f.GetHeight(g) + 2;
                    }

                    e.HasMorePages = false;
                    fontTitle.Dispose();
                    fontBody.Dispose();
                };

                pd.Print();
            }
        }

        // ■ 印刷行リストを構築（VB6: Form_Print の Printer.Print 相当）
        private System.Collections.Generic.List<string> BuildPrintLines()
        {
            var L = new System.Collections.Generic.List<string>();
            L.Add("オーダーメイドお見積り");
            L.Add("");
            L.Add("■参照データ----------------------");
            L.Add("");

            if (_lblSmpHinban.BackColor == CCalc) L.Add("品番 : " + _lblSmpHinban.Text);
            if (_lblSmpPt900.BackColor == CCalc) L.Add("pt900 : " + _lblSmpPt900.Text + " g");
            if (_lblSmpK18.BackColor == CCalc) L.Add("K18 : " + _lblSmpK18.Text + " g");
            if (_lblSmpWgpg.BackColor == CCalc) L.Add("WG/PG : " + _lblSmpWgpg.Text + " g");
            if (_lblSmpCode.BackColor == CCalc) L.Add("Code : " + _lblSmpCode.Text);
            if (_lblSmpPrice.BackColor == CCalc) L.Add("金額 \\ " + _lblSmpPrice.Text);

            // 腕1
            if (_cU1F.BackColor == CCalc)
            {
                L.Add(""); L.Add("■腕１----------------------");
                L.Add($"地金の種類 : {_cU1A.Text} , 腕の形状 : {_cU1B.Text} , " +
                      $"天幅 : {D3(_cU1TJu, _cU1TJi, _cU1TJs)} mm, " +
                      $"天厚 : {D3(_cU1TAJu, _cU1TAJi, _cU1TAJs)} mm, " +
                      $"底幅 : {D3(_cU1BJu, _cU1BJi, _cU1BJs)} mm, " +
                      $"底厚 : {D3(_cU1BAJu, _cU1BAJi, _cU1BAJs)} mm,");
                L.Add($"特殊加工 : {_cU1D.Text} , " +
                      $"リングサイズ : {_cU1SzJu.Text}{_cU1SzJi.Text} 号 , " +
                      $"個数 : {_cU1Pm.Text} {_cU1F.Text}");
            }
            // 腕2
            if (_cU2F.BackColor == CCalc)
            {
                L.Add(""); L.Add("■腕２----------------------");
                L.Add($"地金の種類 : {_cU2A.Text} , 腕の形状 : {_cU2B.Text} , " +
                      $"天幅 : {D3(_cU2TJu, _cU2TJi, _cU2TJs)} mm, " +
                      $"天厚 : {D3(_cU2TAJu, _cU2TAJi, _cU2TAJs)} mm, " +
                      $"底幅 : {D3(_cU2BJu, _cU2BJi, _cU2BJs)} mm, " +
                      $"底厚 : {D3(_cU2BAJu, _cU2BAJi, _cU2BAJs)} mm,");
                L.Add($"特殊加工 : {_cU2D.Text} , " +
                      $"リングサイズ : {_cU2SzJu.Text}{_cU2SzJi.Text} 号 , " +
                      $"個数 : {_cU2Pm.Text} {_cU2F.Text}");
            }
            // 板線材1
            if (_cIt1F.BackColor == CCalc)
            {
                L.Add(""); L.Add("■板／線材１----------------------");
                L.Add($"部材の種類 : {_cIt1A.Text} , 地金の種類 : {_cIt1B.Text} , " +
                      $"形状 : {_cIt1C.Text} , " +
                      $"長径 : {D3(_cIt1LJu, _cIt1LJi, _cIt1LJs)} mm, " +
                      $"短径 : {D3(_cIt1SJu, _cIt1SJi, _cIt1SJs)} mm, " +
                      $"厚み、長さ : {D3(_cIt1AJu, _cIt1AJi, _cIt1AJs)} mm, 個数 : {_cIt1F.Text}");
            }
            // 板線材2
            if (_cIt2F.BackColor == CCalc)
            {
                L.Add(""); L.Add("■板／線材２----------------------");
                L.Add($"部材の種類 : {_cIt2A.Text} , 地金の種類 : {_cIt2B.Text} , " +
                      $"形状 : {_cIt2C.Text} , " +
                      $"長径 : {D3(_cIt2LJu, _cIt2LJi, _cIt2LJs)} mm, " +
                      $"短径 : {D3(_cIt2SJu, _cIt2SJi, _cIt2SJs)} mm, " +
                      $"厚み、長さ : {D3(_cIt2AJu, _cIt2AJi, _cIt2AJs)} mm, 個数 : {_cIt2F.Text}");
            }
            // 石座1〜3
            AppendIszPrint(L, "石座１", _cIsz1F, _cIsz1A, _cIsz1B,
                _cIsz1LJu, _cIsz1LJi, _cIsz1LJs, _cIsz1SJu, _cIsz1SJi, _cIsz1SJs,
                _cIsz1D, _cIsz1E, _cIsz1Pm);
            AppendIszPrint(L, "石座２", _cIsz2F, _cIsz2A, _cIsz2B,
                _cIsz2LJu, _cIsz2LJi, _cIsz2LJs, _cIsz2SJu, _cIsz2SJi, _cIsz2SJs,
                _cIsz2D, _cIsz2E, _cIsz2Pm);
            AppendIszPrint(L, "石座３", _cIsz3F, _cIsz3A, _cIsz3B,
                _cIsz3LJu, _cIsz3LJi, _cIsz3LJs, _cIsz3SJu, _cIsz3SJi, _cIsz3SJs,
                _cIsz3D, _cIsz3E, _cIsz3Pm);
            // 石留め1〜4
            AppendIsmPrint(L, "石留め１", _cIsm1D, _cIsm1A, _cIsm1B, _cIsm1C, _cIsm1EJu, _cIsm1EJi);
            AppendIsmPrint(L, "石留め２", _cIsm2D, _cIsm2A, _cIsm2B, _cIsm2C, _cIsm2EJu, _cIsm2EJi);
            AppendIsmPrint(L, "石留め３", _cIsm3D, _cIsm3A, _cIsm3B, _cIsm3C, _cIsm3EJu, _cIsm3EJi);
            AppendIsmPrint(L, "石留め４", _cIsm4D, _cIsm4A, _cIsm4B, _cIsm4C, _cIsm4EJu, _cIsm4EJi);
            // ダイヤ1〜4
            AppendDiaPrint(L, "ダイヤ１", _cDia1A, _cDia1B, _cDia1Kigo, _cDia1Ju, _cDia1Ji);
            AppendDiaPrint(L, "ダイヤ２", _cDia2A, _cDia2B, _cDia2Kigo, _cDia2Ju, _cDia2Ji);
            AppendDiaPrint(L, "ダイヤ３", _cDia3A, _cDia3B, _cDia3Kigo, _cDia3Ju, _cDia3Ji);
            AppendDiaPrint(L, "ダイヤ４", _cDia4A, _cDia4B, _cDia4Kigo, _cDia4Ju, _cDia4Ji);
            // 面ロー
            if (_cRoA1.BackColor == CCalc)
            {
                L.Add(""); L.Add("■面ロー----------------------");
                L.Add($"ロー付け方法 : {_cRoA1.Text} , 個数 : {_cRoB1Ju.Text}{_cRoB1Ji.Text}");
            }
            // 点ロー
            if (_cRoC1.BackColor == CCalc)
            {
                L.Add(""); L.Add("■点ロー----------------------");
                L.Add($"ロー付け方法 : {_cRoC1.Text} , 個数 : {_cRoD1Ju.Text}{_cRoD1Ji.Text}");
            }
            // 加工難易度
            if (_cKakouA1.BackColor == CCalc)
            {
                L.Add(""); L.Add("■加工難易度----------------------");
                L.Add("加工難易度 : " + _cKakouA1.Text);
            }
            // 加工グレード
            if (_cKakouB1.BackColor == CCalc)
            {
                L.Add(""); L.Add("■加工グレード----------------------");
                L.Add("加工グレード : " + _cKakouB1.Text);
            }
            // 合計
            if (_lblPt900.BackColor == CCalc) { L.Add(""); L.Add("■Pt900 合計重量----------------------"); L.Add(_lblPt900.Text + " g"); }
            if (_lblK18.BackColor == CCalc) { L.Add(""); L.Add("■K18 合計重量----------------------"); L.Add(_lblK18.Text + " g"); }
            if (_lblWgpg.BackColor == CCalc) { L.Add(""); L.Add("■WG/PG 合計重量----------------------"); L.Add(_lblWgpg.Text + " g"); }
            if (_lblPrice.BackColor == CCalc)
            {
                L.Add(""); L.Add("■合計金額----------------------");
                L.Add(" \\ " + _lblPrice.Text + "（税抜き）");
                L.Add(""); L.Add(_lblIngo.Text);
            }
            return L;
        }

        // ============================================================
        // ■ 戻るボタン（VB6: btn_back_Click）
        // ============================================================
        private void BtnBack_Click()
        {
            _forceClose = true;
            if (AppState.FlagDispPicture)
            {
                // FormDispPicture を OpenForms から探して表示
                foreach (Form f in Application.OpenForms)
                {
                    if (f is FormDispPicture)
                    {
                        f.Show();
                        break;
                    }
                }
            }
            else
            {
                // FormMenu を OpenForms から探して表示
                foreach (Form f in Application.OpenForms)
                {
                    if (f is FormMenu)
                    {
                        f.Show();
                        break;
                    }
                }
            }
            this.Close();
        }

        // ============================================================
        // ■ メニューボタン（VB6: btn_menu_Click）
        // ============================================================
        private void BtnMenu_Click()
        {
            AppState.FlagDispPicture = false;
            _forceClose = true;
            foreach (Form f in Application.OpenForms)
            {
                if (f is FormMenu) { f.Show(); break; }
            }
            this.Close();
        }

        // ============================================================
        // ■ DB ヘルパーメソッド
        // ============================================================

        /// <summary>地金売り相場テーブルから Pt1000 と K24 の現在価格を取得</summary>
        private bool GetCurrentPrices(out double nowPt1000, out double nowK24)
        {
            nowPt1000 = 0; nowK24 = 0;
            var dt = AppState.Db.ExecuteQuery("SELECT * FROM [地金売り相場テーブル]");
            if (dt.Rows.Count == 0) return false;
            var row = dt.Rows[0];
            string ptStr = row["pt万"].ToString() + row["pt千"].ToString() +
                            row["pt百"].ToString() + row["pt十"].ToString() + row["pt一"].ToString();
            string k18Str = row["k18万"].ToString() + row["k18千"].ToString() +
                            row["k18百"].ToString() + row["k18十"].ToString() + row["k18一"].ToString();
            double.TryParse(ptStr, out nowPt1000);
            double.TryParse(k18Str, out nowK24);
            return true;
        }

        /// <summary>掛け率テーブルから全体掛け率1 を取得（÷10前の値）</summary>
        private double GetZentaiKakeritu()
        {
            var dt = AppState.Db.ExecuteQuery("SELECT * FROM [掛け率テーブル]");
            if (dt.Rows.Count == 0) return 0;
            return Convert.ToDouble(dt.Rows[0]["掛け率1"]);
        }

        /// <summary>加工コードテーブルから指定コードの掛け率を取得</summary>
        private double GetKakouCd(int code)
        {
            var dt = AppState.Db.ExecuteQuery(
                "SELECT * FROM [加工コード] WHERE コード = " + code);
            if (dt.Rows.Count == 0) return 0;
            return Convert.ToDouble(dt.Rows[0]["掛け率"]);
        }

        // ============================================================
        // ■■■ 計算サブルーチン群 ■■■
        // ============================================================

        // ──────────────────────────────────────────────────────────
        // 腕１計算（VB6: Ude_Keisan_1_Rtn）
        // ──────────────────────────────────────────────────────────
        private void UdeKeisan1Rtn()
        {
            // ホワイトニング
            SetWhite(_cU1A, _cU1B, _cU1TJu, _cU1TJi, _cU1TJs,
                     _cU1TAJu, _cU1TAJi, _cU1TAJs,
                     _cU1BJu, _cU1BJi, _cU1BJs,
                     _cU1BAJu, _cU1BAJi, _cU1BAJs,
                     _cU1D, _cU1SzJu, _cU1SzJi, _cU1F, _cU1Pm);

            if (_cU1F.Text == "0") return;

            double tenhaba = ToD3(_cU1TJu, _cU1TJi, _cU1TJs);
            double sokohaba = ToD3(_cU1BJu, _cU1BJi, _cU1BJs);
            double tenatsu = ToD3(_cU1TAJu, _cU1TAJi, _cU1TAJs);
            double sokoatsu = ToD3(_cU1BAJu, _cU1BAJi, _cU1BAJs);
            if (tenhaba == 0 || sokohaba == 0 || tenatsu == 0 || sokoatsu == 0) return;

            double danmen;
            if (_cU1B.Text == "平打ち") danmen = (0.8 * tenhaba * tenatsu + 0.2 * sokohaba * sokoatsu) / 100;
            else if (_cU1B.Text == "甲丸") danmen = 0.95 * (0.8 * tenhaba * tenatsu + 0.2 * sokohaba * sokoatsu) / 100;
            else { MessageBox.Show("システムエラー", "腕計算１：断面積取得"); return; }
            if (danmen == 0) return;

            int ringSize = int.Parse(_cU1SzJu.Text) * 10 + int.Parse(_cU1SzJi.Text);
            if (ringSize == 0) return;

            var dtSz = AppState.Db.ExecuteQuery(
                "SELECT * FROM [リングサイズ円周率] WHERE サイズ = " + ringSize);
            if (dtSz.Rows.Count == 0) { MessageBox.Show("システムエラー", "腕計算１：理論ウェイト取得"); return; }
            double ringEnsyu = Convert.ToDouble(dtSz.Rows[0]["円周"]);

            double hijuu = GetHijuu(_cU1A.Text);
            if (hijuu == 0) { MessageBox.Show("システムエラー", "腕計算１：理論ウェイト：比重取得"); return; }
            double rironWT = 1.1 * danmen * ringEnsyu * hijuu / 10;
            if (rironWT == 0) return;

            // 特殊加工（中抜き）
            if (_cU1D.Text == "中浅抜き") rironWT *= 0.9;
            else if (_cU1D.Text == "中深抜き") rironWT *= 0.8;
            else if (_cU1D.Text == "通常") { }
            else { MessageBox.Show("システムエラー：腕計算１：中抜き地金計算"); }

            // 重量累計
            AccumWeight(_cU1A.Text, _cU1Pm.Text, rironWT, int.Parse(_cU1F.Text));

            if (!GetCurrentPrices(out double nowPt1000, out double nowK24))
            { MessageBox.Show("システムエラーです。", "地金売り相場テーブル該当データ無し"); return; }

            double jiganeGenka = CalcJiganeGenka(_cU1A.Text, rironWT, nowPt1000, nowK24);
            if (jiganeGenka == 0) return;

            // 中抜き地金原価
            double nakaJigane;
            if (_cU1D.Text == "中浅抜き") nakaJigane = jiganeGenka * 0.9;
            else if (_cU1D.Text == "中深抜き") nakaJigane = jiganeGenka * 0.8;
            else nakaJigane = jiganeGenka;

            // 新枠工賃
            double sinwakuKoutin = GetSinwakuKoutin(_cU1A.Text);
            if (sinwakuKoutin == 0) { MessageBox.Show("システムエラー", "腕計算１：新枠工賃取得"); return; }

            double nakaSinwaku = nakaJigane + sinwakuKoutin;
            double kakouCd = GetKakouCd(95);
            double wakuGedai = nakaSinwaku * kakouCd;
            if (wakuGedai == 0) return;

            double zentaiKakeritu = GetZentaiKakeritu();
            double zeiNuki = wakuGedai * zentaiKakeritu / 10;

            double lastPrice = (_cU1Pm.Text == "＋")
                ? zeiNuki * int.Parse(_cU1F.Text)
                : zeiNuki * int.Parse(_cU1F.Text) * 0.5 * -1;
            if (lastPrice == 0) { MessageBox.Show("システムエラー", "腕１：最終税抜き価格取得取得"); return; }

            _orderPrice += lastPrice;

            SetCalc(_cU1A, _cU1B, _cU1TJu, _cU1TJi, _cU1TJs,
                    _cU1TAJu, _cU1TAJi, _cU1TAJs,
                    _cU1BJu, _cU1BJi, _cU1BJs,
                    _cU1BAJu, _cU1BAJi, _cU1BAJs,
                    _cU1D, _cU1SzJu, _cU1SzJi, _cU1F, _cU1Pm);
        }

        // ──────────────────────────────────────────────────────────
        // 腕２計算（VB6: Ude_Keisan_2_Rtn）
        // ──────────────────────────────────────────────────────────
        private void UdeKeisan2Rtn()
        {
            SetWhite(_cU2A, _cU2B, _cU2TJu, _cU2TJi, _cU2TJs,
                     _cU2TAJu, _cU2TAJi, _cU2TAJs,
                     _cU2BJu, _cU2BJi, _cU2BJs,
                     _cU2BAJu, _cU2BAJi, _cU2BAJs,
                     _cU2D, _cU2SzJu, _cU2SzJi, _cU2F, _cU2Pm);

            if (_cU2F.Text == "0") return;

            double tenhaba = ToD3(_cU2TJu, _cU2TJi, _cU2TJs);
            double sokohaba = ToD3(_cU2BJu, _cU2BJi, _cU2BJs);
            double tenatsu = ToD3(_cU2TAJu, _cU2TAJi, _cU2TAJs);
            double sokoatsu = ToD3(_cU2BAJu, _cU2BAJi, _cU2BAJs);
            if (tenhaba == 0 || sokohaba == 0 || tenatsu == 0 || sokoatsu == 0) return;

            double danmen;
            if (_cU2B.Text == "平打ち") danmen = (0.8 * tenhaba * tenatsu + 0.2 * sokohaba * sokoatsu) / 100;
            else if (_cU2B.Text == "甲丸") danmen = 0.95 * (0.8 * tenhaba * tenatsu + 0.2 * sokohaba * sokoatsu) / 100;
            else { MessageBox.Show("システムエラー", "腕計算２：断面積取得"); return; }
            if (danmen == 0) return;

            int ringSize = int.Parse(_cU2SzJu.Text) * 10 + int.Parse(_cU2SzJi.Text);
            if (ringSize == 0) return;

            var dtSz = AppState.Db.ExecuteQuery(
                "SELECT * FROM [リングサイズ円周率] WHERE サイズ = " + ringSize);
            if (dtSz.Rows.Count == 0) { MessageBox.Show("システムエラー", "腕計算２：理論ウェイト取得"); return; }
            double ringEnsyu = Convert.ToDouble(dtSz.Rows[0]["円周"]);

            double hijuu = GetHijuu(_cU2A.Text);
            if (hijuu == 0) { MessageBox.Show("システムエラー", "腕計算２：理論ウェイト：比重取得"); return; }
            double rironWT = 1.1 * danmen * ringEnsyu * hijuu / 10;
            if (rironWT == 0) return;

            if (_cU2D.Text == "中浅抜き") rironWT *= 0.9;
            else if (_cU2D.Text == "中深抜き") rironWT *= 0.8;
            else if (_cU2D.Text == "通常") { }
            else { MessageBox.Show("システムエラー：腕計算２：中抜き地金計算"); }

            AccumWeight(_cU2A.Text, _cU2Pm.Text, rironWT, int.Parse(_cU2F.Text));

            if (!GetCurrentPrices(out double nowPt1000, out double nowK24))
            { MessageBox.Show("システムエラーです。", "地金売り相場テーブル該当データ無し"); return; }

            double jiganeGenka = CalcJiganeGenka(_cU2A.Text, rironWT, nowPt1000, nowK24);
            if (jiganeGenka == 0) return;

            double nakaJigane;
            if (_cU2D.Text == "中浅抜き") nakaJigane = jiganeGenka * 0.9;
            else if (_cU2D.Text == "中深抜き") nakaJigane = jiganeGenka * 0.8;
            else nakaJigane = jiganeGenka;

            double sinwakuKoutin = GetSinwakuKoutin(_cU2A.Text);
            if (sinwakuKoutin == 0) { MessageBox.Show("システムエラー", "腕計算２：新枠工賃取得"); return; }

            double nakaSinwaku = nakaJigane + sinwakuKoutin;
            double kakouCd = GetKakouCd(95);
            double wakuGedai = nakaSinwaku * kakouCd;
            if (wakuGedai == 0) return;

            double zentaiKakeritu = GetZentaiKakeritu();
            double zeiNuki = wakuGedai * zentaiKakeritu / 10;

            double lastPrice = (_cU2Pm.Text == "＋")
                ? zeiNuki * int.Parse(_cU2F.Text)
                : zeiNuki * int.Parse(_cU2F.Text) * 0.5 * -1;
            if (lastPrice == 0) { MessageBox.Show("システムエラー", "腕２：最終税抜き価格取得取得"); return; }

            _orderPrice += lastPrice;

            SetCalc(_cU2A, _cU2B, _cU2TJu, _cU2TJi, _cU2TJs,
                    _cU2TAJu, _cU2TAJi, _cU2TAJs,
                    _cU2BJu, _cU2BJi, _cU2BJs,
                    _cU2BAJu, _cU2BAJi, _cU2BAJs,
                    _cU2D, _cU2SzJu, _cU2SzJi, _cU2F, _cU2Pm);
        }

        // ──────────────────────────────────────────────────────────
        // 板線材計算共通ロジック
        // ──────────────────────────────────────────────────────────
        private void ItaSenzaiRtn(
            ComboBox cA, ComboBox cB, ComboBox cC,
            ComboBox cLJu, ComboBox cLJi, ComboBox cLJs,
            ComboBox cSJu, ComboBox cSJi, ComboBox cSJs,
            ComboBox cAJu, ComboBox cAJi, ComboBox cAJs,
            ComboBox cF, string label)
        {
            SetWhite(cA, cB, cC, cLJu, cLJi, cLJs, cSJu, cSJi, cSJs, cAJu, cAJi, cAJs, cF);

            if (cF.Text == "0") return;

            double choKei = ToD3(cLJu, cLJi, cLJs);
            double tanKei = ToD3(cSJu, cSJi, cSJs);
            double atsumi = ToD3(cAJu, cAJi, cAJs);
            if (choKei == 0 || tanKei == 0 || atsumi == 0) return;

            double taiseki = 0;
            string buza = cA.Text;
            string katachi = cC.Text;
            if (katachi == "円、楕円") taiseki = 3.14 * (choKei + tanKei) * (choKei + tanKei) * atsumi / 16 / 1000;
            else if (katachi == "角") taiseki = choKei * tanKei * atsumi / 1000;
            else if (katachi == "ドロップ")
            {
                taiseki = (buza == "板")
                    ? 1.5 * choKei * tanKei * atsumi / 2000
                    : choKei * tanKei * atsumi / 2000;
            }
            else { MessageBox.Show("システムエラー：" + label + "：形状取得"); return; }

            if (taiseki == 0) return;

            double taisekiWt;
            if (buza == "板") taisekiWt = taiseki * 1.2;
            else if (buza == "棒") taisekiWt = taiseki * 2;
            else { MessageBox.Show("システムエラー", label + "：理論ウェイト取得"); return; }

            double jiganeKind = 0;
            string jigane = cB.Text;
            if (jigane == "Pt900") { double rw = taisekiWt * HijuuPt900; _orderPt900 += rw * double.Parse(cF.Text); jiganeKind = rw; }
            else if (jigane == "K18") { double rw = taisekiWt * HijuuK18; _orderK18 += rw * double.Parse(cF.Text); jiganeKind = rw; }
            else if (jigane == "WG/PG") { double rw = taisekiWt * HijuuWgpg; _orderWgpg += rw * double.Parse(cF.Text); jiganeKind = rw; }
            else { MessageBox.Show("システムエラー", label + "：理論ウェイト取得"); return; }

            if (jiganeKind == 0) return;

            if (!GetCurrentPrices(out double nowPt1000, out double nowK24))
            { MessageBox.Show("システムエラーです。", "地金売り相場テーブル該当データ無し"); return; }

            double jiganePrice = CalcJiganeGenka(jigane, jiganeKind, nowPt1000, nowK24);
            if (jiganePrice == 0) return;

            // 工賃
            double koutin = 0;
            if (jigane == "Pt900") koutin = (buza == "板") ? 7000 : 5000;
            else if (jigane == "K18") koutin = (buza == "板") ? 5800 : 4000;
            else if (jigane == "WG/PG") koutin = (buza == "板") ? 5800 * 1.2 : 4000 * 1.2;
            if (koutin == 0) { MessageBox.Show("システムエラー", label + "：工賃取得"); return; }

            double itaBouPrice = jiganePrice + koutin;
            double kakouCd = GetKakouCd(95);
            double itaBouGedai = itaBouPrice * kakouCd;
            if (itaBouGedai == 0) return;

            double zentaiKakeritu = GetZentaiKakeritu();
            double zeiNuki = itaBouGedai * zentaiKakeritu / 10;
            double lastPrice = zeiNuki * int.Parse(cF.Text);
            if (lastPrice == 0) { MessageBox.Show("システムエラー", label + "：最終税抜き価格取得取得"); return; }

            _orderPrice += lastPrice;
            SetCalc(cA, cB, cC, cLJu, cLJi, cLJs, cSJu, cSJi, cSJs, cAJu, cAJi, cAJs, cF);
        }

        private void ItaSenzai1Rtn() =>
            ItaSenzaiRtn(_cIt1A, _cIt1B, _cIt1C,
                         _cIt1LJu, _cIt1LJi, _cIt1LJs,
                         _cIt1SJu, _cIt1SJi, _cIt1SJs,
                         _cIt1AJu, _cIt1AJi, _cIt1AJs,
                         _cIt1F, "板線材計算１");

        private void ItaSenzai2Rtn() =>
            ItaSenzaiRtn(_cIt2A, _cIt2B, _cIt2C,
                         _cIt2LJu, _cIt2LJi, _cIt2LJs,
                         _cIt2SJu, _cIt2SJi, _cIt2SJs,
                         _cIt2AJu, _cIt2AJi, _cIt2AJs,
                         _cIt2F, "板線材計算２");

        // ──────────────────────────────────────────────────────────
        // 石座計算共通ロジック（VB6: Ishiza_1/2/3_rtn）
        // ──────────────────────────────────────────────────────────
        private void IshizaRtn(
            ComboBox cA, ComboBox cB,
            ComboBox cLJu, ComboBox cLJi, ComboBox cLJs,
            ComboBox cSJu, ComboBox cSJi, ComboBox cSJs,
            ComboBox cD, ComboBox cE, ComboBox cPm, ComboBox cF, string label)
        {
            SetWhite(cA, cB, cLJu, cLJi, cLJs, cSJu, cSJi, cSJs, cD, cE, cPm, cF);

            if (cF.Text == "0") return;

            double choKei = ToD3(cLJu, cLJi, cLJs);
            double tanKei = ToD3(cSJu, cSJi, cSJs);
            if (choKei == 0 || tanKei == 0) return;

            // MAX チェック
            double totKei = choKei + tanKei;
            string katachi = cB.Text;
            if (katachi == "円" && (totKei > 22 || totKei < 5.5)) { MessageBox.Show("入力チェック\n" + label + "\n石の形状が「円」の場合、合計サイズは5.5〜22.0mmです。"); return; }
            if (katachi == "楕円" && (totKei > 57 || totKei < 5.5)) { MessageBox.Show("入力チェック\n" + label + "\n石の形状が「楕円」の場合、合計サイズは5.5〜57.0mmです。"); return; }
            if (katachi == "角" && totKei > 35.5) { MessageBox.Show("入力チェック\n" + label + "\n石の形状が「角」の場合、合計サイズは35.5mmまでです。"); return; }
            if (katachi == "ドロップ" && totKei > 34) { MessageBox.Show("入力チェック\n" + label + "\n石の形状が「ドロップ」の場合、合計サイズは34.0mmまでです。"); return; }
            if (katachi == "0.1ct未満円" && totKei > 5.4) { MessageBox.Show("入力チェック\n" + label + "\n石の形状が「0.1ct未満円」の場合、合計サイズは5.4mmまでです。"); return; }
            if (katachi == "0.1ct未満ﾌｧﾝｼｰ" && totKei > 5.4) { MessageBox.Show("入力チェック\n" + label + "\n石の形状が「0.1ct未満ﾌｧﾝｼｰ」の場合、合計サイズは5.4mmまでです。"); return; }

            // 石座枠重量テーブルから weight 取得
            string jigane = cA.Text;
            string dbJigane = (jigane == "Pt900") ? "pt900" : "k18";
            var dtW = AppState.Db.ExecuteQuery(
                "SELECT * FROM [石座枠重量テーブル] WHERE 地金 = '" + dbJigane +
                "' AND 形状 = '" + katachi + "' ORDER BY 合計サイズ");
            if (dtW.Rows.Count == 0) { MessageBox.Show("システムエラーです。", label + "：石座枠重量テーブル取得"); return; }

            double ishizaWT = 0;
            foreach (DataRow r in dtW.Rows)
            {
                if (Convert.ToDouble(r["合計サイズ"]) / 10 >= totKei)
                { ishizaWT = Convert.ToDouble(r["ウェイト"]); break; }
            }
            if (ishizaWT == 0) { MessageBox.Show("システムエラーです。", label + "：石座枠重量テーブル取得２"); return; }

            // 腰高係数
            if (cE.Text == "深腰") ishizaWT *= 1.5;
            else if (cE.Text == "浅腰") ishizaWT *= 0.8;

            // 石座種類係数
            if (cD.Text == "覆輪") ishizaWT *= 1.5;
            else if (cD.Text == "チョコ" || cD.Text == "取り巻き") ishizaWT *= 2.0;
            else if (cD.Text == "通常爪") { }
            else { MessageBox.Show("システムエラー" + label + "：石座の種類取得"); return; }

            // 重量累計
            double qty = double.Parse(cF.Text);
            if (cPm.Text == "−" && jigane == "Pt900") _orderPt900 -= ishizaWT * qty;
            else if (cPm.Text == "＋" && jigane == "Pt900") _orderPt900 += ishizaWT * qty;
            else if (cPm.Text == "−" && jigane == "K18") _orderK18 -= ishizaWT * qty;
            else if (cPm.Text == "＋" && jigane == "K18") _orderK18 += ishizaWT * qty;
            else if (cPm.Text == "−" && jigane == "WG/PG") _orderWgpg -= ishizaWT * qty;
            else if (cPm.Text == "＋" && jigane == "WG/PG") _orderWgpg += ishizaWT * qty;

            // 必要地金原価
            if (!GetCurrentPrices(out double nowPt1000, out double nowK24))
            { MessageBox.Show("システムエラーです。", "地金売り相場テーブル該当データ無し"); return; }

            double jiganeGenka = 0;
            if (jigane == "Pt900") jiganeGenka = ishizaWT * nowPt1000 * 1.1 * 0.9 * 1.1;
            else if (jigane == "K18") jiganeGenka = ishizaWT * nowK24 * 1.1 * 0.75 * 1.1;
            else if (jigane == "WG/PG") jiganeGenka = ishizaWT * nowK24 * 1.1 * 0.75 * 1.1;
            if (jiganeGenka == 0) return;

            // 石座枠工賃テーブル
            var dtK = AppState.Db.ExecuteQuery(
                "SELECT * FROM [石座枠工賃テーブル] WHERE 地金 = '" + dbJigane +
                "' AND 形状 = '" + katachi + "' ORDER BY 合計サイズ");
            if (dtK.Rows.Count == 0) { MessageBox.Show("システムエラーです。", label + "：石座枠工賃テーブル取得"); return; }

            double ishiwakuKoutin = 0;
            foreach (DataRow r in dtK.Rows)
            {
                if (Convert.ToDouble(r["合計サイズ"]) / 10 >= totKei)
                {
                    ishiwakuKoutin = Convert.ToDouble(r["枠工賃"]);
                    if (jigane == "WG/PG") ishiwakuKoutin *= 1.5;
                    break;
                }
            }
            if (ishiwakuKoutin == 0) { MessageBox.Show("システムエラーです。", label + "：石座枠工賃テーブル取得２"); return; }

            double sinwakuGenka = ishiwakuKoutin + jiganeGenka;
            if (sinwakuGenka == 0) { MessageBox.Show("システムエラーです。", label + "：新枠原価取得"); return; }

            // 形状→加工コード
            int kakouCode = 96;
            if (katachi == "楕円") kakouCode = 94;
            else if (katachi == "角") kakouCode = 91;
            else if (katachi == "ドロップ") kakouCode = 93;

            double kakouCd = GetKakouCd(kakouCode);
            double wakuGedai = sinwakuGenka * kakouCd;
            if (wakuGedai == 0) return;

            // 石座留め工賃テーブル
            var dtT = AppState.Db.ExecuteQuery(
                "SELECT * FROM [石座留め工賃テーブル] WHERE 地金 = '" + dbJigane +
                "' AND 形状 = '" + katachi + "' ORDER BY 合計サイズ");
            if (dtT.Rows.Count == 0) { MessageBox.Show("システムエラーです。", label + "：石座留め工賃テーブル取得"); return; }

            double tomeKoutin = 0;
            foreach (DataRow r in dtT.Rows)
            {
                if (Convert.ToDouble(r["合計サイズ"]) / 10 >= totKei)
                { tomeKoutin = Convert.ToDouble(r["留め工賃"]); break; }
            }
            if (tomeKoutin == 0) { MessageBox.Show("システムエラーです。", label + "：石座留め工賃テーブル取得２"); return; }

            double tomeGedai = tomeKoutin * kakouCd;
            if (tomeGedai == 0) { MessageBox.Show("システムエラーです。", label + "：留め下代取得"); return; }

            double zentaiKakeritu = GetZentaiKakeritu();
            double tomeJoudai = (wakuGedai + tomeGedai) * zentaiKakeritu / 10;
            if (tomeJoudai == 0) { MessageBox.Show("システムエラー", label + "：留め上代取得"); return; }

            // 石座種類で掛ける
            if (cD.Text == "覆輪") tomeJoudai *= 1.5;
            else if (cD.Text == "チョコ" || cD.Text == "取り巻き") tomeJoudai *= 2.0;

            // 腰高で掛ける
            if (cE.Text == "深腰") tomeJoudai *= 1.5;
            else if (cE.Text == "浅腰") tomeJoudai *= 0.8;

            // プラマイ
            if (cPm.Text == "−") tomeJoudai = -0.5 * tomeJoudai;

            // 最終価格
            double lastPrice = (katachi == "0.1ct未満ﾌｧﾝｼｰ")
                ? tomeJoudai * int.Parse(cF.Text) * 2
                : tomeJoudai * int.Parse(cF.Text);
            if (lastPrice == 0) { MessageBox.Show("システムエラー", label + "：最終税抜き価格取得取得"); return; }

            _orderPrice += lastPrice;
            SetCalc(cA, cB, cLJu, cLJi, cLJs, cSJu, cSJi, cSJs, cD, cE, cPm, cF);
        }

        private void Ishiza1Rtn() =>
            IshizaRtn(_cIsz1A, _cIsz1B,
                      _cIsz1LJu, _cIsz1LJi, _cIsz1LJs,
                      _cIsz1SJu, _cIsz1SJi, _cIsz1SJs,
                      _cIsz1D, _cIsz1E, _cIsz1Pm, _cIsz1F, "石座１計算");
        private void Ishiza2Rtn() =>
            IshizaRtn(_cIsz2A, _cIsz2B,
                      _cIsz2LJu, _cIsz2LJi, _cIsz2LJs,
                      _cIsz2SJu, _cIsz2SJi, _cIsz2SJs,
                      _cIsz2D, _cIsz2E, _cIsz2Pm, _cIsz2F, "石座２計算");
        private void Ishiza3Rtn() =>
            IshizaRtn(_cIsz3A, _cIsz3B,
                      _cIsz3LJu, _cIsz3LJi, _cIsz3LJs,
                      _cIsz3SJu, _cIsz3SJi, _cIsz3SJs,
                      _cIsz3D, _cIsz3E, _cIsz3Pm, _cIsz3F, "石座３計算");

        // ──────────────────────────────────────────────────────────
        // パイプシャトン計算共通（VB6: Paipu_Shaton_1〜4_Rtn）
        // ──────────────────────────────────────────────────────────
        private void PaipuShatonRtn(
            ComboBox cD, ComboBox cA, ComboBox cB, ComboBox cC,
            ComboBox cEJu, ComboBox cEJi, string label)
        {
            SetWhite(cD, cA, cB, cC, cEJu, cEJi);

            int kosuu = int.Parse(cEJu.Text) * 10 + int.Parse(cEJi.Text);
            if (kosuu == 0) return;

            // 接着留/外 + ファンシー 系の入力チェック
            if (cD.Text == "接着留/外(石)" &&
                (cB.Text == "ﾌｧﾝｼｰ" || cB.Text == "ﾌｧﾝｼｰ寄留"))
            {
                MessageBox.Show(
                    "石留め方法が「接着留/外」の場合は、石の形状は「円楕円珠」もしくは「円楕円珠寄留」しか選べません。",
                    label + "：入力エラー");
                _chkFlgKeisan = true;
                return;
            }

            double joudai = 0;
            if (cD.Text == "芯爪建留(本)")
            {
                var dt = AppState.Db.ExecuteQuery(
                    "SELECT * FROM [パイプシャトンサイズ上代テーブル] WHERE " +
                    "地金 = '" + cA.Text + "' AND サイズ = '" + cC.Text + "' " +
                    "AND 石の形状 = '" + cB.Text + "'");
                if (dt.Rows.Count == 0) { MessageBox.Show("システムエラーです。\n" + label + "：芯爪建留：パイプシャトンサイズ上代テーブル取得"); return; }
                joudai = Convert.ToDouble(dt.Rows[0]["上代"]);
                if (joudai == 0) { MessageBox.Show("システムエラーです。\n" + label + "：芯爪建留：パイプシャトンサイズ上代ゼロ"); return; }
            }
            else
            {
                var dt = AppState.Db.ExecuteQuery(
                    "SELECT * FROM [石留め彫りサイズ上代テーブル] WHERE サイズ = '" + cC.Text + "'");
                if (dt.Rows.Count == 0) { MessageBox.Show("システムエラーです。\n" + label + "：留め外し：石留め彫りサイズ上代テーブル取得"); return; }
                joudai = Convert.ToDouble(dt.Rows[0]["上代"]);
                if (joudai == 0) { MessageBox.Show("システムエラーです。\n" + label + "：留め外し：石留め彫りサイズ上代テーブル上代ゼロ"); return; }

                if (cD.Text == "接着留/外(石)") joudai *= 0.5;

                if (cB.Text == "ﾌｧﾝｼｰ") joudai *= 1.5;
                else if (cB.Text == "円楕円珠寄留") joudai *= 1.5;
                else if (cB.Text == "ﾌｧﾝｼｰ寄留") joudai *= 1.5 * 1.5;
            }

            _orderPrice += joudai * kosuu;
            SetCalc(cD, cA, cB, cC, cEJu, cEJi);
        }

        private void PaipuShaton1Rtn() =>
            PaipuShatonRtn(_cIsm1D, _cIsm1A, _cIsm1B, _cIsm1C, _cIsm1EJu, _cIsm1EJi, "石留め等１");
        private void PaipuShaton2Rtn() =>
            PaipuShatonRtn(_cIsm2D, _cIsm2A, _cIsm2B, _cIsm2C, _cIsm2EJu, _cIsm2EJi, "石留め等２");
        private void PaipuShaton3Rtn() =>
            PaipuShatonRtn(_cIsm3D, _cIsm3A, _cIsm3B, _cIsm3C, _cIsm3EJu, _cIsm3EJi, "石留め等３");
        private void PaipuShaton4Rtn() =>
            PaipuShatonRtn(_cIsm4D, _cIsm4A, _cIsm4B, _cIsm4C, _cIsm4EJu, _cIsm4EJi, "石留め等４");

        // ──────────────────────────────────────────────────────────
        // メレー計算共通（VB6: Mere_1〜4_Rtn）
        // ──────────────────────────────────────────────────────────
        private void MereRtn(
            ComboBox cA, ComboBox cB, ComboBox cKigo,
            ComboBox cJu, ComboBox cJi, string label)
        {
            SetWhite(cA, cB, cKigo, cJu, cJi);

            int kosuu = int.Parse(cJu.Text) * 10 + int.Parse(cJi.Text);
            if (kosuu == 0) return;

            var dt = AppState.Db.ExecuteQuery(
                "SELECT * FROM [メレーサイズグレード上代テーブル] WHERE " +
                "グレード = '" + cA.Text + "' AND サイズ = '" + cB.Text + "'");
            if (dt.Rows.Count == 0) { MessageBox.Show("システムエラーです。\n" + label + "：メレーサイズグレード上代テーブル取得"); return; }
            double joudai = Convert.ToDouble(dt.Rows[0]["上代"]);
            if (joudai == 0) { MessageBox.Show("システムエラーです。\n" + label + "：メレーサイズグレード上代ゼロ"); return; }

            if (cKigo.Text == "＋") _orderPrice += joudai * kosuu;
            else if (cKigo.Text == "−") _orderPrice -= joudai * kosuu;
            else { MessageBox.Show("システムエラーです。\n" + label + "：記号取得"); }

            SetCalc(cA, cB, cKigo, cJu, cJi);
        }

        private void Mere1Rtn() => MereRtn(_cDia1A, _cDia1B, _cDia1Kigo, _cDia1Ju, _cDia1Ji, "メレー１計算");
        private void Mere2Rtn() => MereRtn(_cDia2A, _cDia2B, _cDia2Kigo, _cDia2Ju, _cDia2Ji, "メレー２計算");
        private void Mere3Rtn() => MereRtn(_cDia3A, _cDia3B, _cDia3Kigo, _cDia3Ju, _cDia3Ji, "メレー３計算");
        private void Mere4Rtn() => MereRtn(_cDia4A, _cDia4B, _cDia4Kigo, _cDia4Ju, _cDia4Ji, "メレー４計算");

        // ──────────────────────────────────────────────────────────
        // 面ロー計算（VB6: Men_Row_Rtn）
        // ──────────────────────────────────────────────────────────
        private void MenRowRtn()
        {
            SetWhite(_cRoA1, _cRoB1Ju, _cRoB1Ji);

            int kosuu = int.Parse(_cRoB1Ju.Text) * 10 + int.Parse(_cRoB1Ji.Text);
            if (kosuu == 0) return;

            var dt = AppState.Db.ExecuteQuery(
                "SELECT * FROM [ロー付け種類上代テーブル] WHERE 種類 = '" + _cRoA1.Text + "'");
            if (dt.Rows.Count == 0) { MessageBox.Show("システムエラーです。\nロー付け計算：ロー付け種類上代テーブル取得"); return; }
            double joudai = Convert.ToDouble(dt.Rows[0]["上代"]);
            if (joudai == 0) { MessageBox.Show("システムエラーです。\nロー付け計算：上代ゼロ"); return; }

            _orderPrice += joudai * kosuu;
            SetCalc(_cRoA1, _cRoB1Ju, _cRoB1Ji);
        }

        // ──────────────────────────────────────────────────────────
        // 点ロー計算（VB6: Ten_Row_Rtn）
        // ──────────────────────────────────────────────────────────
        private void TenRowRtn()
        {
            SetWhite(_cRoC1, _cRoD1Ju, _cRoD1Ji);

            int kosuu = int.Parse(_cRoD1Ju.Text) * 10 + int.Parse(_cRoD1Ji.Text);
            if (kosuu == 0) return;

            // まず個数一致で検索、なければ個数=21 でフォールバック
            var dt = AppState.Db.ExecuteQuery(
                "SELECT * FROM [ロー付け種類上代テーブル] WHERE " +
                "種類 = '" + _cRoC1.Text + "' AND 個数 = " + kosuu);
            if (dt.Rows.Count == 0)
            {
                dt = AppState.Db.ExecuteQuery(
                    "SELECT * FROM [ロー付け種類上代テーブル] WHERE " +
                    "種類 = '" + _cRoC1.Text + "' AND 個数 = 21");
                if (dt.Rows.Count == 0)
                { MessageBox.Show("システムエラーです。\nロー付け計算：ロー付け種類上代テーブル取得"); return; }
            }
            double joudai = Convert.ToDouble(dt.Rows[0]["上代"]);
            if (joudai == 0) { MessageBox.Show("システムエラーです。\nロー付け計算：上代ゼロ"); return; }

            _orderPrice += joudai * kosuu;
            SetCalc(_cRoC1, _cRoD1Ju, _cRoD1Ji);
        }

        // ──────────────────────────────────────────────────────────
        // 加工グレードチェック（VB6: Kakou_Grade_Rtn）
        // ──────────────────────────────────────────────────────────
        private void KakouGradeRtn()
        {
            _cKakouA1.BackColor = CDefault;
            _cKakouB1.BackColor = CDefault;

            if (_cKakouA1.Text == "部品合計" && _cKakouB1.Text == "B")
            {
                MessageBox.Show("加工難易度が「部品合計」の時は、加工グレードは「B」には出来ません。");
                _chkFlgKeisan = true;
                return;
            }
            SetCalc(_cKakouA1, _cKakouB1);
        }

        // ============================================================
        // ■ ユーティリティメソッド
        // ============================================================

        /// <summary>十・一・小 コンボから実数値を取得（例: 十=1,一=2,小=3 → 12.3）</summary>
        private static double ToD3(ComboBox ju, ComboBox ichi, ComboBox sho)
            => double.Parse(ju.Text) * 10 + double.Parse(ichi.Text)
             + double.Parse(sho.Text) * 0.1;

        /// <summary>十・一・小 コンボから文字列を生成（例: "12.3"）</summary>
        private static string D3(ComboBox ju, ComboBox ichi, ComboBox sho)
            => ju.Text + ichi.Text + "." + sho.Text;

        /// <summary>地金種別から比重定数を返す</summary>
        private static double GetHijuu(string jigane)
        {
            if (jigane == "Pt900") return HijuuPt900;
            if (jigane == "K18") return HijuuK18;
            if (jigane == "WG/PG") return HijuuWgpg;
            return 0;
        }

        /// <summary>地金原価を計算（Pt900 は nowPt1000 × 0.9 × 1.1、K18/WG/PG は nowK24 × 0.75 × 1.1）</summary>
        private static double CalcJiganeGenka(string jigane, double wt, double nowPt1000, double nowK24)
        {
            if (jigane == "Pt900") return wt * nowPt1000 * 1.1 * 0.9;
            if (jigane == "K18") return wt * nowK24 * 1.1 * 0.75;
            if (jigane == "WG/PG") return wt * nowK24 * 1.1 * 0.75;
            return 0;
        }

        /// <summary>新枠工賃（Pt900=9000, K18=6000, WG/PG=9000）</summary>
        private static double GetSinwakuKoutin(string jigane)
        {
            if (jigane == "Pt900") return 9000;
            if (jigane == "K18") return 6000;
            if (jigane == "WG/PG") return 9000;
            return 0;
        }

        /// <summary>地金重量の累計フィールドに加減算</summary>
        private void AccumWeight(string jigane, string pm, double wt, int count)
        {
            double delta = (pm == "＋") ? wt * count : -(wt * count);
            if (jigane == "Pt900") _orderPt900 += delta;
            else if (jigane == "K18") _orderK18 += delta;
            else if (jigane == "WG/PG") _orderWgpg += delta;
        }

        /// <summary>複数コンボの BackColor をデフォルト（白）に設定</summary>
        private static void SetWhite(params ComboBox[] controls)
        {
            foreach (var c in controls) if (c != null) c.BackColor = SystemColors.Window;
        }

        /// <summary>複数コンボの BackColor を計算済み（黄色）に設定</summary>
        private static void SetCalc(params ComboBox[] controls)
        {
            foreach (var c in controls) if (c != null) c.BackColor = Color.FromArgb(255, 255, 192);
        }

        /// <summary>複数コンボの SelectedIndex を 0 にリセット</summary>
        private static void ResetIndex(params ComboBox[] controls)
        {
            foreach (var c in controls)
                if (c != null && c.Items.Count > 0) c.SelectedIndex = 0;
        }

        /// <summary>文字列を SQL 用シングルクォート付き文字列に変換（カンマ付き）</summary>
        private static string Q(string val)
            => "'" + val.Replace("'", "''") + "',";

        /// <summary>INSERT 文に石留め行を追加</summary>
        private static void AppendIsm(StringBuilder sb,
            ComboBox cD, ComboBox cA, ComboBox cB, ComboBox cC,
            ComboBox cEJu, ComboBox cEJi, bool last = false)
        {
            sb.Append(Q(cD.Text));
            // 芯爪建留 の場合のみ地金種類を保存（それ以外はスペース）
            sb.Append(cD.Text == "芯爪建留(本)" ? Q(cA.Text) : "' ',");
            sb.Append(Q(cB.Text));
            sb.Append(Q(cC.Text));
            string kosuu = cEJu.Text + cEJi.Text;
            sb.Append(last ? ("'" + kosuu.Replace("'", "''") + "'") : Q(kosuu));
        }

        // ──── 印刷補助 ────

        private void AppendIszPrint(System.Collections.Generic.List<string> L,
            string title, ComboBox cF, ComboBox cA, ComboBox cB,
            ComboBox cLJu, ComboBox cLJi, ComboBox cLJs,
            ComboBox cSJu, ComboBox cSJi, ComboBox cSJs,
            ComboBox cD, ComboBox cE, ComboBox cPm)
        {
            if (cF.BackColor != CCalc) return;
            L.Add(""); L.Add($"■{title}----------------------");
            L.Add($"地金の種類 : {cA.Text} , 石の形状 : {cB.Text} , " +
                  $"石の長径 : {D3(cLJu, cLJi, cLJs)} mm, 石の短径 : {D3(cSJu, cSJi, cSJs)} mm,");
            L.Add($"主石座の種類 : {cD.Text} , 腰高 : {cE.Text} , 個数 : {cPm.Text}{cF.Text}");
        }

        private void AppendIsmPrint(System.Collections.Generic.List<string> L,
            string title, ComboBox cD, ComboBox cA, ComboBox cB,
            ComboBox cC, ComboBox cEJu, ComboBox cEJi)
        {
            if (cD.BackColor != CCalc) return;
            L.Add(""); L.Add($"■{title}----------------------");
            string kosuu = cEJu.Text + cEJi.Text;
            if (cD.Text == "芯爪建留(本)")
            {
                L.Add($"留め方法 : {cD.Text} , 地金の種類 : {cA.Text} ,");
                L.Add($"石の形状 : {cB.Text} , 石のサイズ : {cC.Text} , 本数 : {kosuu}");
            }
            else
            {
                L.Add($"留め方法 : {cD.Text} ,");
                L.Add($"石の形状 : {cB.Text} , 石のサイズ : {cC.Text} , 石数 : {kosuu}");
            }
        }

        private void AppendDiaPrint(System.Collections.Generic.List<string> L,
            string title, ComboBox cA, ComboBox cB, ComboBox cKigo,
            ComboBox cJu, ComboBox cJi)
        {
            if (cA.BackColor != CCalc) return;
            L.Add(""); L.Add($"■{title}----------------------");
            L.Add($"グレード : {cA.Text} , サイズ : {cB.Text} ct, " +
                  $"個数 : {cKigo.Text}{cJu.Text}{cJi.Text}");
        }
    }
}
