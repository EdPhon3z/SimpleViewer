using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

static class IconProbe
{
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern uint ExtractIconEx(string szFileName, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    public static Bitmap? GetIcon(string path, int index)
    {
        var handles = new IntPtr[1];
        var count = ExtractIconEx(path, index, handles, null, 1);
        if (count == 0 || handles[0] == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            using var icon = Icon.FromHandle(handles[0]);
            return (Bitmap)icon.ToBitmap().Clone();
        }
        finally
        {
            DestroyIcon(handles[0]);
        }
    }
}

static class Program
{
    private static readonly char[] Map = " .':-=+*#%@".ToCharArray();

    static string ToAscii(Bitmap bmp)
    {
        const int target = 24;
        using var thumb = new Bitmap(target, target);
        using (var g = Graphics.FromImage(thumb))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
            g.DrawImage(bmp, new Rectangle(0, 0, target, target));
        }

        var lines = new List<string>(target);
        for (int y = 0; y < target; y++)
        {
            Span<char> line = stackalloc char[target];
            for (int x = 0; x < target; x++)
            {
                var pixel = thumb.GetPixel(x, y);
                var brightness = (pixel.R + pixel.G + pixel.B) / 3f;
                var index = (int)Math.Round((brightness / 255f) * (Map.Length - 1));
                line[x] = Map[Math.Clamp(index, 0, Map.Length - 1)];
            }
            lines.Add(new string(line));
        }

        return string.Join(Environment.NewLine, lines);
    }

    static void Main(string[] args)
    {
        int start = 0;
        int end = 300;
        string? saveDirectory = null;

        if (args.Length >= 1 && int.TryParse(args[0], out var parsedStart))
        {
            start = parsedStart;
        }
        if (args.Length >= 2 && int.TryParse(args[1], out var parsedEnd))
        {
            end = parsedEnd;
        }
        if (args.Length >= 3)
        {
            saveDirectory = args[2];
            if (!Directory.Exists(saveDirectory))
            {
                Directory.CreateDirectory(saveDirectory);
            }
        }

        var path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "imageres.dll");
        for (int i = start; i <= end; i++)
        {
            using var bmp = IconProbe.GetIcon(path, i);
            if (bmp is null)
            {
                continue;
            }

            Console.WriteLine($"ID {i}");
            Console.WriteLine(ToAscii(bmp));
            Console.WriteLine();

            if (saveDirectory is not null)
            {
                var destination = Path.Combine(saveDirectory, $"icon_{i}.png");
                bmp.Save(destination, ImageFormat.Png);
                Console.WriteLine($"Saved {destination}");
            }
        }
    }
}
