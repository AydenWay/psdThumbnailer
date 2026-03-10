using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace CustomShellExtensions
{
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f")]
    public interface IInitializeWithStream
    {
        void Initialize(IStream pstream, uint grfMode);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("e357fccd-a995-4576-b01f-234630154e96")]
    public interface IThumbnailProvider
    {
        void GetThumbnail(uint cx, out IntPtr phbmp, out uint pdwAlpha);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("Custom.PsdThumbnailProvider")]
    [Guid("YOUR-GUID-HERE")] // <--- PASTE YOUR GUID HERE
    public class PsdThumbnailProvider : IInitializeWithStream, IThumbnailProvider, IDisposable
    {
        private IStream _stream;

        public void Initialize(IStream pstream, uint grfMode)
        {
            _stream = pstream;
        }

        public void GetThumbnail(uint cx, out IntPtr phbmp, out uint pdwAlpha)
        {
            phbmp = IntPtr.Zero;
            pdwAlpha = 0;

            if (_stream == null) return;

            try
            {
                using (var comStream = new ComStreamWrapper(_stream))
                using (var reader = new BinaryReader(comStream))
                {
                    string signature = new string(reader.ReadChars(4));
                    if (signature != "8BPS") return;

                    reader.BaseStream.Seek(22, SeekOrigin.Current);

                    uint colorModeLength = ReadBigEndianUInt32(reader);
                    reader.BaseStream.Seek(colorModeLength, SeekOrigin.Current);

                    uint resourcesLength = ReadBigEndianUInt32(reader);
                    long resourcesEndPosition = reader.BaseStream.Position + resourcesLength;

                    while (reader.BaseStream.Position < resourcesEndPosition)
                    {
                        string resSignature = new string(reader.ReadChars(4));
                        if (resSignature != "8BIM") break;

                        ushort resId = ReadBigEndianUInt16(reader);

                        byte nameLen = reader.ReadByte();
                        int nameSkip = nameLen + (nameLen % 2 == 0 ? 1 : 0);
                        reader.BaseStream.Seek(nameSkip, SeekOrigin.Current);

                        uint resDataLength = ReadBigEndianUInt32(reader);
                        uint paddedDataLength = resDataLength + (resDataLength % 2);

                        if (resId == 1036)
                        {
                            reader.BaseStream.Seek(28, SeekOrigin.Current);
                            int jpegLength = (int)resDataLength - 28;
                            byte[] jpegData = reader.ReadBytes(jpegLength);

                            using (var ms = new MemoryStream(jpegData))
                            using (var bmp = new Bitmap(ms))
                            {
                                phbmp = bmp.GetHbitmap();
                                pdwAlpha = 1; // FIX 1: WTSAT_RGB — tells Explorer the bitmap is valid
                            }
                            return;
                        }
                        else
                        {
                            reader.BaseStream.Seek(paddedDataLength, SeekOrigin.Current);
                        }
                    }
                }
            }
            catch
            {
                // Silently fail if file is locked or malformed
            }
            finally
            {
                // Release the COM stream immediately so file handles are freed
                Marshal.ReleaseComObject(_stream);
                _stream = null;
            }
        }

        public void Dispose()
        {
            if (_stream != null)
            {
                Marshal.ReleaseComObject(_stream);
                _stream = null;
            }
        }

        private uint ReadBigEndianUInt32(BinaryReader reader)
        {
            byte[] bytes = reader.ReadBytes(4);
            Array.Reverse(bytes);
            return BitConverter.ToUInt32(bytes, 0);
        }

        private ushort ReadBigEndianUInt16(BinaryReader reader)
        {
            byte[] bytes = reader.ReadBytes(2);
            Array.Reverse(bytes);
            return BitConverter.ToUInt16(bytes, 0);
        }
    }

    public class ComStreamWrapper : Stream
    {
        private readonly IStream _comStream;
        public ComStreamWrapper(IStream comStream) { _comStream = comStream; }
        public override bool CanRead  { get { return true; } }
        public override bool CanSeek  { get { return true; } }
        public override bool CanWrite { get { return false; } }

        public override long Position
        {
            get { return Seek(0, SeekOrigin.Current); }
            set { Seek(value, SeekOrigin.Begin); }
        }

        // FIX 3: Use COM Stat() instead of throwing, avoids swallowed exceptions
        public override long Length
        {
            get
            {
                System.Runtime.InteropServices.ComTypes.STATSTG stat;
                _comStream.Stat(out stat, 1); // STATFLAG_NONAME
                return stat.cbSize;
            }
        }

        public override void Flush() { }

        // FIX 2: Use a temp buffer and copy to respect the offset parameter
        public override int Read(byte[] buffer, int offset, int count)
        {
            byte[] temp = new byte[count];
            int bytesRead = 0;
            IntPtr ptr = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(int)));
            try
            {
                _comStream.Read(temp, count, ptr);
                bytesRead = Marshal.ReadInt32(ptr);
            }
            finally
            {
                Marshal.FreeCoTaskMem(ptr);
            }
            Array.Copy(temp, 0, buffer, offset, bytesRead);
            return bytesRead;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            long newPosition = 0;
            IntPtr ptr = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(long)));
            try
            {
                _comStream.Seek(offset, (int)origin, ptr);
                newPosition = Marshal.ReadInt64(ptr);
            }
            finally
            {
                Marshal.FreeCoTaskMem(ptr);
            }
            return newPosition;
        }

        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
    }
}
