namespace ExcelMerger
{
    public static class ProArgs
    {
        private static string[] args;

        public static string Base;
        public static string Theirs;
        public static string Mine;
        public static string Merged;

        public static string[] Args
        {
            set { args = value;
                CheckPath();
            }
        }

        private static void CheckPath()
        {
            Base = args[0];
            Theirs = args[1];
            Mine = args[2];
            Merged = args[3];
        }
    }
}