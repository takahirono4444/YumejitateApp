namespace YumejitateApp
{
    using System;
    using System.Windows.Forms;

    
    public class FormMovie : Form { }
    public class FormSaveCamera : Form
    {
        public new void Show() { }
    }
    public class FormSnap : Form
    {
        public void FormInit() { }
        public void UpdateSnapShot(IntPtr hCamera, uint dwWidth, uint dwHeight,
                                   uint dwPreviewPixelFormat, byte[] pbyteImageBuffer,
                                   ref uint dwLastErrorNo)
        { }
        public void SaveImage() { }
    }
    public class FormSaveDejicame : Form { }

   
}