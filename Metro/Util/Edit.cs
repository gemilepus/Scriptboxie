using System.IO;

namespace Metro
{
    public class Edit
    {
        public string ModifiedTime, FilePath;

        public void StartEdit(string path)
        {
            FilePath = path;
            ModifiedTime = GetModifiedTime(path);
        }

        public string GetModifiedTime(string path)
        {
            FileInfo mFileInfo = new FileInfo(path);

            return mFileInfo.LastWriteTime.ToString();
        }

        public bool CheckIsModifie()
        {
            if (ModifiedTime.Equals("")) return false;
            return ModifiedTime.Equals(GetModifiedTime(FilePath)) ? false : true;
        }

    }
}
