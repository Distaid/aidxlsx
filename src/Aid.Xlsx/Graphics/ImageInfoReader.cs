#nullable disable

namespace Aid.Xlsx.Graphics
{
    internal abstract class ImageInfoReader
    {
        public bool TryGetInfo(Stream stream, out XlsxPictureInfo info)
        {
            info = default;
            stream.Position = 0;
            if (!CheckHeader(stream))
            {
                stream.Position = 0;
                return false;
            }

            stream.Position = 0;
            info = ReadInfo(stream);
            return true;
        }

        protected abstract bool CheckHeader(Stream stream);

        protected abstract XlsxPictureInfo ReadInfo(Stream stream);
    }
}
