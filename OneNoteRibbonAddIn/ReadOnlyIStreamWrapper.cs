using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

namespace OneNoteRibbonAddIn
{
    class ReadOnlyIStreamWrapper : IStream
    {
        private readonly MemoryStream _stream;

        public ReadOnlyIStreamWrapper(MemoryStream stream)
        {
            _stream = stream;
        }

        public void Read(byte[] pv, int cb, IntPtr pcbRead)
        {
            Marshal.WriteInt64(pcbRead, _stream.Read(pv, 0, cb));
        }

        public void Write(byte[] pv, int cb, IntPtr pcbWritten)
        {
            Marshal.WriteInt64(pcbWritten, 0L);
            _stream.Write(pv, 0, cb);
            Marshal.WriteInt64(pcbWritten, cb);
        }

        public void Seek(long dlibMove, int dwOrigin, IntPtr plibNewPosition)
        {
            long num;
            Marshal.WriteInt64(plibNewPosition, _stream.Position);
            switch (dwOrigin)
            {
                case 0:
                    num = dlibMove;
                    break;

                case 1:
                    num = _stream.Position + dlibMove;
                    break;

                case 2:
                    num = _stream.Length + dlibMove;
                    break;

                default:
                    return;
            }
            if ((num >= 0L) && (num < _stream.Length))
            {
                _stream.Position = num;
                Marshal.WriteInt64(plibNewPosition, _stream.Position);
            }
        }

        public void SetSize(long libNewSize)
        {
            _stream.SetLength(libNewSize);
        }

        public void CopyTo(IStream pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten)
        {
            throw new NotSupportedException("ReadOnlyIStreamWrapper does not support CopyTo");
        }

        public void Commit(int grfCommitFlags)
        {
            _stream.Flush();
        }

        public void Revert()
        {
            throw new NotSupportedException("Stream does not support CopyTo");
        }

        public void LockRegion(long libOffset, long cb, int dwLockType)
        {
            throw new NotSupportedException("ReadOnlyIStreamWrapper does not support CopyTo");
        }

        public void UnlockRegion(long libOffset, long cb, int dwLockType)
        {
            throw new NotSupportedException("ReadOnlyIStreamWrapper does not support UnlockRegion");
        }

        public void Stat(out STATSTG pstatstg, int grfStatFlag)
        {
            pstatstg = new STATSTG();
            pstatstg.cbSize = _stream.Length;
            if ((grfStatFlag & 1) == 0)
            {
                pstatstg.pwcsName = _stream.ToString();
            }
        }

        public void Clone(out IStream ppstm)
        {
            ppstm = new ReadOnlyIStreamWrapper(_stream);
        }
    }
}
