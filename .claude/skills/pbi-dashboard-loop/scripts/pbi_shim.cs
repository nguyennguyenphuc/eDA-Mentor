// Ban thay the toi thieu cua `pbi --json dax execute` cho may khong cai pbi-cli.
//
// Vi sao la .exe chu khong phai .cmd/.ps1: check_numbers.py goi
// subprocess.run(["pbi", ...]) voi shell=False. Tren Windows, CreateProcess chi
// noi them duoi ".exe" khi do PATH -- KHONG doc PATHEXT -- nen pbi.cmd va pbi.bat
// deu khong bao gio duoc tim thay. Ngoai ra chay truc tiep se bo duoc ~1 giay
// khoi dong PowerShell cho MOI luot goi DAX, ma mot vong check_numbers.py goi
// hang tram luot.
//
// Tham so den tay Main() da duoc CLR tach san, nen truy van DAX (tieng Viet co
// dau, ngoac vuong, nhay kep) di qua nguyen ven, khong qua tang bao trich nao.
//
//   pbi --json dax execute "<DAX>"    -> stdout {"rows":[{...}]}
//   pbi --json tmsl "<TMSL JSON>"     -> chay lenh TMSL (refresh)
//   pbi connect                       -> khong lam gi, tra 0
//
// Bien dich: tools/build_pbi_shim.ps1

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.AnalysisServices.AdomdClient;

static class PbiShim
{
    const string Dll = @"C:\Program Files\Microsoft Power BI Desktop\bin\Microsoft.PowerBI.AdomdClient.dll";

    static string FindPort()
    {
        string root = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            @"Microsoft\Power BI Desktop\AnalysisServicesWorkspaces");
        if (!Directory.Exists(root))
            throw new Exception("Khong thay thu muc workspace: " + root);
        var f = new DirectoryInfo(root)
            .GetFiles("msmdsrv.port.txt", SearchOption.AllDirectories)
            .OrderByDescending(x => x.LastWriteTimeUtc)
            .FirstOrDefault();
        if (f == null)
            throw new Exception("Khong thay msmdsrv.port.txt. Power BI Desktop da mo du an chua?");
        // Tep nay ghi UTF-16LE khong BOM.
        string raw = File.ReadAllText(f.FullName, Encoding.Unicode);
        string port = new string(raw.Where(char.IsDigit).ToArray());
        if (port.Length == 0) throw new Exception("Cong rong trong " + f.FullName);
        return port;
    }

    // Liet ke ten catalog cho thong bao loi, de nguoi dung biet dat gi vao
    // PBI_CATALOG ma khong phai di tra.
    static string TenCatalog(DataTable cat)
    {
        var ds = new List<string>();
        foreach (DataRow r in cat.Rows) ds.Add((string)r["CATALOG_NAME"]);
        return string.Join(" | ", ds.ToArray());
    }

    static void Esc(StringBuilder sb, string s)
    {
        sb.Append('"');
        foreach (char c in s)
        {
            if (c == '"' || c == '\\') { sb.Append('\\').Append(c); }
            else if (c == '\n') sb.Append("\\n");
            else if (c == '\r') sb.Append("\\r");
            else if (c == '\t') sb.Append("\\t");
            else if (c < 0x20) sb.AppendFormat("\\u{0:x4}", (int)c);
            else sb.Append(c);
        }
        sb.Append('"');
    }

    static int Main(string[] argv)
    {
        // pbi.exe khong nam canh Microsoft.PowerBI.AdomdClient.dll, va DLL do
        // khong duoc dang ky trong GAC, nen phai tu tro duong den thu muc bin
        // cua Power BI Desktop.
        //
        // Than chuong trinh phai nam trong Run() va Run() phai NoInlining: neu
        // de thang trong Main thi JIT nap assembly ngay luc buoc vao Main, tuc
        // TRUOC khi handler kip dang ky, va van bao FileNotFoundException.
        AppDomain.CurrentDomain.AssemblyResolve += (sender, e) =>
        {
            string name = new System.Reflection.AssemblyName(e.Name).Name;
            string p = Path.Combine(Path.GetDirectoryName(Dll), name + ".dll");
            return File.Exists(p) ? System.Reflection.Assembly.LoadFrom(p) : null;
        };
        return Run(argv);
    }

    [System.Runtime.CompilerServices.MethodImpl(
        System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    static int Run(string[] argv)
    {
        // Ghi stdout bang UTF-8 KHONG BOM.
        //
        // Khong dat cai nay thi Console dung codepage ANSI cua console, va moi
        // ky tu tieng Viet trong KET QUA bi thay bang '?' NGAY TRONG SHIM --
        // mat han, phia Python giai ma kieu gi cung khong cuu lai duoc.
        //
        // Do duoc 02/09/2026: cong so lieu v3 bao 20 phep do, tat ca deu la so
        // sanh chuoi ("Doanh nghiep", "Thap nhat", "Dau ky", "Theo so so",
        // "36-60 thang"); gia tri so cua model deu dung. Trieu chung nhin nhu
        // model sai, that ra la duong ra bi hong.
        //
        // Ghi chu da co tu truoc ve chieu VAO (dung PowerShell 5.1 thi tham so
        // truyen cho exe bi ep ve ANSI) van dung; day la chieu RA, mot loi
        // khac va phai vá rieng.
        Console.OutputEncoding = new UTF8Encoding(false);

        var args = argv.Where(a => a != "--json").ToArray();
        if (args.Length == 0) { Console.Error.WriteLine("pbi_shim: thieu tham so"); return 2; }
        if (args[0] == "connect")
        {
            Console.WriteLine("connect: shim tu do cong tu msmdsrv.port.txt, khong can buoc nay");
            return 0;
        }

        bool tmsl;
        string text;
        if (args[0] == "dax" && args.Length >= 3 && args[1] == "execute") { tmsl = false; text = args[2]; }
        else if (args[0] == "tmsl" && args.Length >= 2) { tmsl = true; text = args[1]; }
        else { Console.Error.WriteLine("pbi_shim: khong ho tro lenh nay"); return 2; }

        try
        {
            string port = FindPort();
            using (var conn = new AdomdConnection("Data Source=localhost:" + port))
            {
                conn.Open();
                // conn.Databases tra ve RONG khi ket noi khong kem Catalog, nen
                // phai hoi qua schema rowset roi ChangeDatabase.
                DataTable cat = conn.GetSchemaDataSet("DBSCHEMA_CATALOGS", null).Tables[0];
                if (cat.Rows.Count == 0) throw new Exception("Instance chua nap database nao");

                // MOT tien trinh msmdsrv chua NHIEU catalog, moi tep Desktop dang mo
                // mot cai. Lay Rows[0] la doan mo. Da gap that 01/09/2026: mo them mot
                // tep khac thi shim tra ve model chuoi cung ung trong khi dang doi
                // chieu bao cao ngan hang, va khong bao gi ca.
                // Nay: mot catalog thi dung luon; nhieu hon thi PHAI chi ro qua bien
                // moi truong PBI_CATALOG, khong doan.
                string muon = Environment.GetEnvironmentVariable("PBI_CATALOG");
                string chon = null;
                if (!string.IsNullOrEmpty(muon))
                {
                    foreach (DataRow r0 in cat.Rows)
                    {
                        string ten = (string)r0["CATALOG_NAME"];
                        if (ten.IndexOf(muon, StringComparison.OrdinalIgnoreCase) >= 0)
                        { chon = ten; break; }
                    }
                    if (chon == null)
                        throw new Exception("Khong catalog nao khop PBI_CATALOG='"
                            + muon + "'. Dang co: " + TenCatalog(cat));
                }
                else if (cat.Rows.Count > 1)
                {
                    throw new Exception("Co " + cat.Rows.Count + " catalog dang mo, phai dat "
                        + "PBI_CATALOG de chi ro. Dang co: " + TenCatalog(cat));
                }
                else chon = (string)cat.Rows[0]["CATALOG_NAME"];
                conn.ChangeDatabase(chon);

                var cmd = conn.CreateCommand();
                cmd.CommandTimeout = 7200;
                cmd.CommandText = text;

                if (tmsl)
                {
                    cmd.ExecuteNonQuery();
                    Console.Out.Write("{\"rows\":[{\"[tmsl]\":\"OK\"}]}");
                    return 0;
                }

                var sb = new StringBuilder("{\"rows\":[");
                using (var r = cmd.ExecuteReader())
                {
                    var cols = new List<string>();
                    for (int i = 0; i < r.FieldCount; i++) cols.Add(r.GetName(i));
                    bool first = true;
                    while (r.Read())
                    {
                        if (!first) sb.Append(',');
                        first = false;
                        sb.Append('{');
                        for (int i = 0; i < cols.Count; i++)
                        {
                            if (i > 0) sb.Append(',');
                            Esc(sb, cols[i]);
                            sb.Append(':');
                            object v = r.GetValue(i);
                            if (v == null || v is DBNull) sb.Append("null");
                            else if (v is bool) sb.Append((bool)v ? "true" : "false");
                            else if (v is double || v is float || v is decimal || v is int
                                  || v is long || v is short || v is byte || v is sbyte
                                  || v is uint || v is ulong || v is ushort)
                                // InvariantCulture: may nay cai vi-VN, khong ep thi
                                // dau thap phan ra dau PHAY va JSON hong ngay.
                                sb.Append(Convert.ToString(v, CultureInfo.InvariantCulture));
                            else if (v is DateTime)
                                Esc(sb, ((DateTime)v).ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture));
                            else Esc(sb, Convert.ToString(v, CultureInfo.InvariantCulture));
                        }
                        sb.Append('}');
                    }
                }
                sb.Append("]}");
                Console.Out.Write(sb.ToString());
                return 0;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("pbi_shim: " + ex.Message);
            return 1;
        }
    }
}
