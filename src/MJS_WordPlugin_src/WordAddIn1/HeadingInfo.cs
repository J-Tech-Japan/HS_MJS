namespace WordAddIn1
{
    /// <summary>
    /// 見出し情報
    /// </summary>
    public class HeadingInfo
    {
        // 項番
        public string num;

        // タイトル
        public string title;

        // ID
        public string id;

        // Merger to
        public string mergeto = "";

        // Equalsオーバーライド
        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            HeadingInfo c = (HeadingInfo)obj;
            return (num == c.num) && (title == c.title) && (id == c.id) && (mergeto.Replace("(", "").Replace(")", "") == c.mergeto.Replace("(", "").Replace(")", ""));
        }
    }
}
