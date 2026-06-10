using System.Text;
using OpenMcdf;

namespace DocxTemplateEngine.Handlers;

/// <summary>
/// Builds a CFBF (Compound File Binary Format) blob for embedding non-Office
/// files (PDF, TXT, CSV, JSON, ZIP, ...) inside a Word document. Word stores
/// these as OLE "Package" objects and refuses to activate them unless the
/// embedded part is a CFBF container with the Object Packager CLSID and
/// CompObj / ObjInfo / Ole10Native streams. Layout reproduced from a
/// known-good manual embed; see [MS-OLEDS] §2.3.5–2.3.7 and [MS-DOC]
/// ObjectPool stream-naming conventions.
/// </summary>
internal static class OlePackageBuilder
{
    private static readonly Guid PackageClsid = new("0003000C-0000-0000-C000-000000000046");

    private static readonly Encoding AnsiEncoding = InitAnsiEncoding();

    private static Encoding InitAnsiEncoding()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(1252);
    }

    private const string CompObjStreamName = "CompObj";
    private const string ObjInfoStreamName = "ObjInfo";
    private const string Ole10NativeStreamName = "Ole10Native";

    public static byte[] Build(string sourcePath, string displayLabel)
    {
        var fileBytes = File.ReadAllBytes(sourcePath);

        var ms = new MemoryStream();
        using (var root = RootStorage.Create(ms))
        {
            root.CLSID = PackageClsid;

            using (var compObj = root.CreateStream(CompObjStreamName))
            {
                var bytes = BuildCompObjStream();
                compObj.Write(bytes, 0, bytes.Length);
            }

            using (var objInfo = root.CreateStream(ObjInfoStreamName))
            {
                var bytes = new byte[] { 0x40, 0x00, 0x03, 0x00, 0x01, 0x00 };
                objInfo.Write(bytes, 0, bytes.Length);
            }

            using (var oleNative = root.CreateStream(Ole10NativeStreamName))
            {
                var bytes = BuildOle10NativeStream(displayLabel, fileBytes);
                oleNative.Write(bytes, 0, bytes.Length);
            }

            root.Flush();
        }
        return ms.ToArray();
    }

    private static byte[] BuildCompObjStream()
    {
        using var ms = new MemoryStream();
        using var w = new BinaryWriter(ms);

        w.Write((uint)0xFFFE0001);
        w.Write((uint)0x00000A03);
        w.Write((uint)0xFFFFFFFF);
        w.Write(PackageClsid.ToByteArray());

        WriteLengthPrefixedAnsi(w, "OLE Package");
        w.Write((uint)0);
        WriteLengthPrefixedAnsi(w, "Package");

        w.Write((uint)0x71B239F4);
        w.Write((uint)0);
        w.Write((uint)0);
        w.Write((uint)0);

        return ms.ToArray();
    }

    private static byte[] BuildOle10NativeStream(string label, byte[] nativeData)
    {
        var ansi = AnsiEncoding;
        var labelBytes = ansi.GetBytes(label);

        using var payload = new MemoryStream();
        using (var w = new BinaryWriter(payload))
        {
            w.Write((ushort)0x0002);
            WriteAnsiZ(w, labelBytes);
            WriteAnsiZ(w, labelBytes);
            w.Write((uint)0x00030000);
            w.Write((uint)(labelBytes.Length + 1));
            WriteAnsiZ(w, labelBytes);
            w.Write((uint)nativeData.Length);
            w.Write(nativeData);
        }

        var payloadBytes = payload.ToArray();
        using var ms = new MemoryStream();
        using var sw = new BinaryWriter(ms);
        sw.Write((uint)payloadBytes.Length);
        sw.Write(payloadBytes);
        return ms.ToArray();
    }

    private static void WriteLengthPrefixedAnsi(BinaryWriter w, string s)
    {
        var bytes = AnsiEncoding.GetBytes(s);
        w.Write((uint)(bytes.Length + 1));
        w.Write(bytes);
        w.Write((byte)0);
    }

    private static void WriteAnsiZ(BinaryWriter w, byte[] ansiBytes)
    {
        w.Write(ansiBytes);
        w.Write((byte)0);
    }
}
