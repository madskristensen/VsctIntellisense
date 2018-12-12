using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VsctCompletion
{
    public static class ExtensionMethods
    {
        private static IVsImageService2 _imageService;

        public static BitmapSource ToBitmap(this ImageMoniker moniker, int size)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            _imageService = _imageService ?? ServiceProvider.GlobalProvider.GetService<SVsImageService, IVsImageService2>();

            var imageAttributes = new ImageAttributes
            {
                Flags = (uint)_ImageAttributesFlags.IAF_RequiredFlags,
                ImageType = (uint)_UIImageType.IT_Bitmap,
                Format = (uint)_UIDataFormat.DF_WPF,
                LogicalHeight = size,
                LogicalWidth = size,
                StructSize = Marshal.SizeOf(typeof(ImageAttributes))
            };

            IVsUIObject result = _imageService.GetImage(moniker, imageAttributes);

            result.get_Data(out object data);
            
            return data as BitmapSource;
        }

    }
}
