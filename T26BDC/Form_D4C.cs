using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace T26BDC
{
    public partial class Form_D4C : Form
    {
        // このﾌﾟﾛｸﾞﾗﾑのｳﾞｧｰｼﾞｮﾝ
        static System.Diagnostics.FileVersionInfo s_ver;

        // DLLのインポート
        [System.Runtime.InteropServices.DllImport("user32.dll",
                CharSet = System.Runtime.InteropServices.CharSet.Auto)]

        // 実行中のﾌﾟﾛｸﾞﾗﾑを調べる
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // プログラムが実行される場所
        static string s_apath = Directory.GetCurrentDirectory();

        // 共通変数
        static string s_msg0;     // ｴﾗｰﾒｯｾｰｼﾞ用0
        static string s_msg1;     // ｴﾗｰﾒｯｾｰｼﾞ用1
        static string s_msg2;     // ｴﾗｰﾒｯｾｰｼﾞ用2        

        static string s_idir;     // ｲﾝﾌﾟｯﾄのﾌｫﾙﾀﾞｰ
        static string s_odir;     // ｱｳﾄﾌﾟｯﾄのﾌｫﾙﾀﾞｰ
        static string s_ofile;    // ｱｳﾄﾌﾟｯﾄのﾌｧｲﾙ
        static string s_sheet;    // 対象シート
        static string s_keycol;   // 判断ｶﾗﾑ番号
        static int i_keycol;      // 判断ｶﾗﾑ番号
        static string s_keycode;  // 判断keyのｺｰﾄﾞ
        static string s_endcol;   // 最終ｶﾗﾑ番号
        static int i_endcol;      // 最終ｶﾗﾑ番号
        static string s_endword;  // 最終データ文字
        static int i_endline;     // 最終行
        static string s_password; // Unprotect password

        static string s_mes;      // 画面表示用メッセージ

        public Form_D4C()
        {
            InitializeComponent();  
        }

        private void Form_D4C_Load(object sender, EventArgs e)
        {
            // 起動　初期化処理

            string s_ans;

            s_ans = f1_ini();            // 初期化  

            label_v.Text = s_ver.FileVersion;

            s_ans = f3_setsumei();       // 起動　説明表示

            textBox_keycol.Text = "1";
            textBox_endcol.Text = "30";

        }

        private void button_idir_Click(object sender, EventArgs e)
        {
            // ==== 入力ﾌｫﾙﾀﾞの選択ﾎﾞﾀﾝ

            FolderBrowserDialog oFolderBD1 = new FolderBrowserDialog();
            oFolderBD1.Description = "入力ﾌｫﾙﾀﾞ設定";
            oFolderBD1.RootFolder = System.Environment.SpecialFolder.MyComputer;
            oFolderBD1.SelectedPath = @"c:\";
            if (oFolderBD1.ShowDialog() == DialogResult.OK)
            {
                textBox_idir.Text = oFolderBD1.SelectedPath;
            }
            oFolderBD1.Dispose();

        }

        private void button_ofile_Click(object sender, EventArgs e)
        {
            // ==== 出力ﾌｧｲﾙの選択ﾎﾞﾀﾝ
            FolderBrowserDialog oFileBD1 = new FolderBrowserDialog();
            oFileBD1.Description = "出力ﾌｧｲﾙの設定";            
            oFileBD1.SelectedPath = @"c:\";
            if (oFileBD1.ShowDialog() == DialogResult.OK)
            {
                textBox_odir.Text = oFileBD1.SelectedPath;
            }
            oFileBD1.Dispose();

        }

        private void button_run_Click(object sender, EventArgs e)
        {
            // ==== 結合処理実行のﾎﾞﾀﾝ

            string s_ans;       

            s_ans = f2_echeck();         // 起動　実行中ｴｸｾﾙの確認

            if (s_ans.Substring(0, 3) != "ERR")
            {
                s_ans = f4_settei();     // 起動　画面からの指定内容を設定
            }

            if (s_ans.Substring(0, 3) != "ERR")
            {
                s_ans = f5_ketsugou();   // 起動　結合処理
            }            

            s_mes += "\r\n b09 90 END " + s_ans + "\r\n ";

            textBox_mes.Text = s_mes;

            textBox_mes.SelectionStart = textBox_mes.Text.Length;
            textBox_mes.Focus();
            textBox_mes.ScrollToCaret();
        }



        static bool msgbox2(string mes, string atype)
        {
            if (atype == "W")
            {
                DialogResult dr = MessageBox.Show(
                    mes, "Warning",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);
                if (dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                MessageBox.Show(
                mes, "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                return false;
            }
        }

        static string f1_ini()
        {
            // ==== 初期設定

            // ﾌｧｲﾙﾊﾞｰｼﾞｮﾝをとる
            s_ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(
                System.Reflection.Assembly.GetExecutingAssembly().Location);

            return "END";
        }

        static string f2_echeck()
        {
            // ====  他のｴｸｾﾙﾌｧｲﾙのﾁｪｯｸ

            //ｴｸｾﾙが終了しているかを確認
            IntPtr hWnd = FindWindow("XLMAIN", null);
            if (hWnd.ToString() != "0")
            {
                s_msg2 = "関係しているｴｸｾﾙがあれば終了してください";
                s_msg2 += "\r\n関係していない場合は";
                s_msg2 += "\r\n「OK」で続行してください。";
                if (!msgbox2(s_msg2, "W"))
                {
                    return "ERROR";
                }
            }

            return "OK ";
        }


        public string f3_setsumei()
        {
            // ==== 説明

            s_mes = "準備の説明 \r\n\r\n";

            s_mes += "・合成したいシートのあるExcelﾌｧｲﾙを入力ホルダに入れます。（このﾌｫﾙﾀﾞには他のﾌｧｲﾙは入れない） \r\n";
            s_mes += "・各ファイルの合成したいシートの名前を同じにする。（シートの位置はどこでもよい） \r\n";
            s_mes += "・判断カラムを決め、判断カラムのタイトル名を同じにする。 \r\n";
            s_mes += "・判断カラムのタイトル名の下側の行が合成対象のデータ行になる。 \r\n";
            s_mes += "・この時判断カラムがブランクのデータは合成対象外。 \r\n";
            s_mes += "・データ行の終わりは１カラム目（A1の列）へ EndWord を入れる。 \r\n";
            s_mes += "・この EndWord も各シート統一する。（これをブランクにした場合、ブランクがあるとデータの終わり） \r\n";
            s_mes += "\r\n";

            s_mes += "画面入力の説明 \r\n\r\n";
            s_mes += "①参照ボタンから入力ホルダを指定。 \r\n";
            s_mes += "②保護解除パスワードを設定。(指定された場合のみ) \r\n";
            s_mes += "③シート名を指定。\r\n";
            s_mes += "④判断カラム位置とそのタイトル名を指定。 \r\n";
            s_mes += "⑤最終列を指定。（ここにファイル名が入る）\r\n";
            s_mes += "⑥データの終わりの EndWord を指定。 \r\n";            
            s_mes += "⑦出力ホルダを指定 。（入力ホルダとは別のホルダを指定） \r\n";
            s_mes += "⑧出力ファイルを指定。 \r\n";
            s_mes += "⑨結合ﾎﾞﾀﾝを押します。 \r\n";
            textBox_mes.Text = s_mes;

            textBox_mes.SelectionStart = textBox_mes.Text.Length;
            textBox_mes.Focus();
            textBox_mes.ScrollToCaret();

            return "OK ";
        }

        public string f4_settei()
        {
            // ==== 画面からの指定項目の取り込みとアウトプットの作成

            // 画面から

            try
            {

                s_idir = textBox_idir.Text;
                s_odir = textBox_odir.Text;
                s_sheet = textBox_sheet.Text;
                s_keycol = textBox_keycol.Text.Trim();
                s_keycode = textBox_keycode.Text.Trim();
                s_endcol = textBox_endcol.Text.Trim();
                s_password = textBox_PW.Text.Trim();
                s_endword = textBox_eword.Text.Trim();

                // 判断カラム位置のﾁｪｯｸ設定
                s_msg2 = "handan col ";

                if (System.Text.RegularExpressions.Regex.IsMatch(
                    s_keycol,
                    @"^[1-9]$",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                {
                    // １から９の数字
                    i_keycol = int.Parse(s_keycol);
                }
                else
                {
                    throw new Exception("エラー　判断カラム位置");
                }

                // 最終カラム位置のﾁｪｯｸ設定
                s_msg2 = "saishuu col ";

                if (System.Text.RegularExpressions.Regex.IsMatch(
                    s_endcol,
                    @"\d{1,2}",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                {
                    i_endcol = int.Parse(s_endcol);
                }
                else
                {
                    throw new Exception("エラー　最終カラム位置");
                }


                // 出力用のｴｸｾﾙ
                string s_bfile = s_apath + @"\T26xbase.xlsx";
                s_ofile = s_odir + @"\" + textBox_ofile.Text + ".xlsx";

                s_msg1 = "出力用Excelﾌｧｲﾙ ";
                if (File.Exists(s_ofile))
                {
                    // 出力用のｴｸｾﾙを開く前に同じファイルがあれば削除
                    File.Delete(s_ofile);
                }

                // 出力用のｴｸｾﾙ作成
                File.Copy(s_bfile, s_ofile);

                s_mes += "\r\n f04 01 F: " + s_ofile;
                s_mes += "\r\n f04 02 OUTPUT FILE 準備";
                textBox_mes.Text = s_mes;

                textBox_mes.SelectionStart = textBox_mes.Text.Length;
                textBox_mes.Focus();
                textBox_mes.ScrollToCaret();

                return "OK ";

            }

            catch (Exception ex)
            {
                s_msg0 = s_msg1 + s_msg2 + "\r\n ｴﾗｰ：" + ex.Message;
                if (!msgbox2(s_msg0, "E"))
                {
                    return "ERROR \r\n" + s_msg0 + "\r\n";
                }
                return "ERROR " + s_msg0;
            }
        }


        public string f5_ketsugou()
        {
            // ==== 結合処理の実行

            // ｴｸｾﾙ　準備
            s_msg2 = "Excel.Application ";
            Excel.Application oXls = new Excel.Application();
            oXls.Application.DisplayAlerts = false;

            // ｴｸｾﾙのｵﾌﾞｼﾞｪｸﾄ
            Excel.Workbook oWB_out = null;
            Excel.Worksheet oWS_out = null;
            Excel.Workbook oWB_in = null;
            Excel.Worksheet oWS_in = null;
            Excel.Range oRangeU = null;
            Excel.Range oRangeIN = null;
            Excel.Range oRangeOUT = null;

            //
            int i_loopc = 0;
            string s_fname;
            int i_readSR = 0;
            int i_readER = 0;
            int i_addR = 0;
            string s_cell_sakiS = "";
            string s_cell_sakiE = "";

            string s_cell_motoS;
            string s_cell_motoE;

            // ﾍﾟｰｽﾄ位置の設定
            int i_pastSR = 1;
            int i_pastER = 1;

            try
            {
                // 出力用のｴｸｾﾙを開く
                s_msg2 = "Workbooks.Open ";
                oWB_out = oXls.Workbooks.Open(s_ofile);

                // ｼｰﾄを開く
                //Excel.Worksheet oWS;
                s_msg2 = "Worksheet ";
                oWS_out = oWB_out.Sheets[1];
                oWS_out.Select();                

                // 出力準備完了
                int i_ketsugou = 0;

                // 入力ﾌｫﾙﾀﾞを設定
                s_msg1 = "入力用のﾎﾙﾀﾞｰ ";
                s_msg2 = "入力ﾌｧｲﾙをさがす ";
                string[] a_files = Directory.GetFiles(s_idir);
                s_mes += "\r\n f05 01 入力ﾌｫﾙﾀﾞから取出開始";
                textBox_mes.Text = s_mes;

                textBox_mes.SelectionStart = textBox_mes.Text.Length;
                textBox_mes.Focus();
                textBox_mes.ScrollToCaret();

                // 入力ﾌｫﾙﾀﾞからﾌｧｲﾙの取り出し
                foreach (string s_ifile in a_files)
                {
                    i_loopc++;
                    s_fname = Path.GetFileName(s_ifile);

                    if (Path.GetExtension(s_ifile) == ".xls" ||
                        Path.GetExtension(s_ifile) == ".xlsx")
                    {

                        // 入力用のｴｸｾﾙを開く
                        s_msg2 = "Workbooks.Open ";
                        oWB_in = oXls.Workbooks.Open(s_ifile);

                        s_mes += "\r\n f05 11 F: " + s_ifile;
                        
                        // ｼｰﾄを開く
                        s_msg2 = "Worksheetを開く ";

                        foreach (Excel.Worksheet oS in oWB_in.Sheets)
                        {

                            if (oS.Name == s_sheet)
                            {
                                // シートがあったとき
                                s_msg2 = "Worksheetがあったとき ";
                                oWS_in = oS;
                                oWS_in.Select();

                                s_mes += "\r\n f05 12 S=" + oS.Name;
                                textBox_mes.Text = s_mes;

                                textBox_mes.SelectionStart = textBox_mes.Text.Length;
                                textBox_mes.Focus();
                                textBox_mes.ScrollToCaret();

                                // 保護解除用のパスワードがあるとき
                                s_msg2 = "保護解除 ";
                                if (s_password != "")
                                {
                                    // 解除
                                    oWS_in.Unprotect(s_password);
                                }

                                // key位置のﾁｪｯｸとｺﾋﾟｰﾚﾝｼﾞの設定
                                s_msg2 = "for key位置のﾁｪｯｸ ";
                                i_readSR = 0;
                                i_readER = 0;

                                for (int i = 1; i < 1000000; i++)
                                {
                                    var o_cellsxv = oWS_in.Cells[i, i_keycol];
                                    var o_cells1v = oWS_in.Cells[i, 1];
                                    Excel.Range o_rx = oWS_in.Range[o_cellsxv, o_cellsxv];
                                    Excel.Range o_r1 = oWS_in.Range[o_cells1v, o_cells1v];

                                    i_endline = o_cellsxv.SpecialCells(11).Row;

                                    if (i_readSR == 0)
                                    {
                                        // 判断keyがまだ拾われていない行のとき
                                        s_msg2 = "入力シート 判断keyの前 ";

                                        if (o_rx.Value != null)
                                        {
                                            // 判断セルになにか入ってる
                                            if (o_rx.Value.GetType() == typeof(System.String))
                                            {
                                                // 内容は文字列
                                                if (o_rx.Value == s_keycode)
                                                {
                                                    // 判断keyのとき
                                                    i_readSR = i + 1;
                                                    s_mes += "\r\n f05 15 判断keyのLINE: " + i.ToString();
                                                    textBox_mes.Text = s_mes;

                                                    textBox_mes.SelectionStart = textBox_mes.Text.Length;
                                                    textBox_mes.Focus();
                                                    textBox_mes.ScrollToCaret();
                                                }
                                            }
                                        }
                                        if (i > 100)
                                        {
                                            throw new Exception("判断KEY確認不能");
                                        }
                                    }
                                    else
                                    {
                                        // 判断keyが拾われた後ろの行のとき
                                        s_msg2 = "入力シート 判断keyの後の行 ";
                                        if (oWS_in.Range[o_cellsxv, o_cellsxv].Value != null)
                                        {
                                            // 判断枠のデータあり
                                            // 結合対象のデータ
                                            // 最終列へファイル名をしまう
                                            oWS_in.Cells[i, i_endcol].Value = s_fname;
                                        }
                                        else
                                        {
                                            // 判断枠のデータなし
                                            // 1枠目（A列）をみる
                                            if (o_r1.Value == null)
                                            {
                                                // 内容はnull                                           
                                                if (s_endword == null)
                                                {
                                                    // 内容はnull  
                                                    // 終わりと判断
                                                    i_readER = i - 1;
                                                    break;
                                                }
                                            }
                                            else
                                            {
                                                // 1枠目（A列）内容は在る
                                                if (o_r1.Value.GetType() == typeof(System.String))
                                                {
                                                    // 内容は文字列                                           
                                                    if (o_r1.Value == s_endword)
                                                    {
                                                        // 内容は終了文字  
                                                        // 終わりと判断
                                                        i_readER = i - 1;
                                                        break;
                                                    }
                                                }
                                            }

                                            if (i >= i_endline)
                                            {
                                                // 全データの最終行  
                                                // 終わりと判断
                                                i_readER = i - 1;
                                                break;
                                            }

                                            // 次の行をみるのでここではブレークしない
                                            // 下記内容をいれておいて後で削除
                                            o_r1.Value = "::deleteLINE::";
                                        }
                                        
                                    }
                                }
                                i_addR = i_readER - i_readSR;
                                if (i_addR < 0)
                                {
                                    throw new Exception("入力Fﾃﾞｰﾀなし");
                                }

                                s_msg2 = "入力ファイル ｺﾋﾟｰﾚﾝｼﾞ ";
                                // ｶﾗﾑの右端の調査
                                oRangeU = oWS_in.UsedRange;

                                // 入力元の開始ｾﾙと終了ｾﾙの調査
                                s_cell_motoS = oWS_in.Cells[i_readSR, 1].Address;
                                s_cell_motoE = oWS_in.Cells[i_readER, i_endcol].Address;
                                // 入力元のﾚﾝｼﾞ確定
                                oRangeIN = oWS_in.get_Range(s_cell_motoS, s_cell_motoE);
                                // 出力元の終了行の調査
                                i_pastER = i_pastSR + i_addR;
                                // 出力元の開始ｾﾙと終了ｾﾙの調査
                                s_cell_sakiS = oWS_out.Cells[i_pastSR, 1].Address;
                                s_cell_sakiE = oWS_out.Cells[i_pastER, i_endcol].Address;
                                // 出力元のﾚﾝｼﾞ確定
                                oRangeOUT = oWS_out.get_Range(s_cell_sakiS, s_cell_sakiE);

                                // ｺﾋﾟｰﾍﾟｰｽﾄ
                                s_msg2 = "ｺﾋﾟｰﾍﾟｰｽﾄ ";
                                oRangeIN.Copy();
                                oRangeOUT.PasteSpecial(
                                    Excel.XlPasteType.xlPasteAll,
                                    Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                    Type.Missing, Type.Missing);

                                // 結合行数ｶｳﾝﾄ
                                i_addR++;
                                i_ketsugou += i_addR;
                                i_pastSR += i_addR;
                                s_mes += "\r\n f05 91 file count: " + i_loopc.ToString();                                
                                s_mes += "    record count: " + i_addR.ToString();
                                s_mes += "    end row count: " + i_endline.ToString();
                                textBox_mes.Text = s_mes;                                

                                textBox_mes.SelectionStart = textBox_mes.Text.Length;
                                textBox_mes.Focus();
                                textBox_mes.ScrollToCaret();

                                // 照合したシートの終わり
                            }
                            // シート探しのforeachの終わり
                        }
                        // この入力ファイルが終わったとき
                        // 入力ｴｸｾﾙをclose
                        oWB_in.Close(Type.Missing, Type.Missing, Type.Missing);

                    }
                    // 入力ﾌｫﾙﾀﾞからﾌｧｲﾙの取り出しforeachのおわり

                }

                //　不要行を調べ削除
                int i_dline = 0;
                int i_outsrow = 1;
                for (int j=1; j <= i_pastER  ; j++)
                {                    
                    oWS_out.Select();
                    
                    var o_cellsdv = oWS_out.Cells[i_outsrow, 1];
                    Excel.Range o_rd = oWS_out.Range[o_cellsdv, o_cellsdv];

                    i_outsrow++;

                    if (o_rd.Value != null)
                    {
                        if (o_rd.Value.GetType() == typeof(System.String))
                        {
                            if (o_rd.Value == "::deleteLINE::")
                            {
                                // 削除
                                i_outsrow = i_outsrow - 1;
                                oWS_out.Rows[i_outsrow].Delete();
                                i_dline++;                                
                            }
                        }
                    }
                }

                // 出力ｴｸｾﾙをsave
                oWS_out.Cells[1, 1].Select();
                oWB_out.Save();

                // 出力ｴｸｾﾙをclose
                oWB_out.Close();

                // ｴｸｾﾙをquit
                oXls.Quit();

                s_mes += "\r\n f05 92 結合行=" + i_ketsugou.ToString();
                s_mes += "\r\n f05 93 削除行=" + i_dline.ToString();
                textBox_mes.Text = s_mes;

                textBox_mes.SelectionStart = textBox_mes.Text.Length;
                textBox_mes.Focus();
                textBox_mes.ScrollToCaret();

                return "OK ";
            }

            catch (Exception ex)
            {
                s_msg0 = s_msg1 + s_msg2 + "\r\n ｴﾗｰ：" + ex.Message;
                if (!msgbox2(s_msg0, "E"))
                {
                    return "ERROR " + s_msg0;
                }
                return "ERROR \r\n" + s_msg0 + "\r\n";
            }
            finally
            {
                if (oXls != null && Marshal.IsComObject(oXls)) oXls.Quit();

                if (oRangeU != null && Marshal.IsComObject(oRangeU)) Marshal.ReleaseComObject(oRangeU);
                if (oRangeIN != null && Marshal.IsComObject(oRangeIN)) Marshal.ReleaseComObject(oRangeIN);
                if (oRangeOUT != null && Marshal.IsComObject(oRangeOUT)) Marshal.ReleaseComObject(oRangeOUT);

                if (oWS_in != null && Marshal.IsComObject(oWS_in)) Marshal.ReleaseComObject(oWS_in);
                if (oWS_out != null && Marshal.IsComObject(oWS_out)) Marshal.ReleaseComObject(oWS_out);

                if (oWB_in != null && Marshal.IsComObject(oWB_in)) Marshal.ReleaseComObject(oWB_in);
                if (oWB_out != null && Marshal.IsComObject(oWB_out)) Marshal.ReleaseComObject(oWB_out);

                if (oXls != null && Marshal.IsComObject(oXls)) Marshal.ReleaseComObject(oXls);

                GC.Collect();
            }
        }

        private void button_help_Click(object sender, EventArgs e)
        {
            // ==== 画面の実行記録クリアと説明再表示

            string s_ans = f3_setsumei();
        }



        // ----------------------------------------------------------------- fgk 2017/7/1
    }
}
