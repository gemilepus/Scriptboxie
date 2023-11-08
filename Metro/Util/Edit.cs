using System.IO;

namespace Metro
{
    public class Edit
    {
        public string ModifiedTime;

        public string GetModifiedTime(string path)
        {
            FileInfo mFileInfo = new FileInfo(path);

            return mFileInfo.LastWriteTime.ToString();
        }

        public bool CheckIsModifie(string path)
        {
            if (ModifiedTime.Equals("")) return false;
            return ModifiedTime.Equals(GetModifiedTime(path)) ? false : true;
        }

    }
}
