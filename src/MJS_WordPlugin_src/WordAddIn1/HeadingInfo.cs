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
            return (num == c.num) && (title == c.title) && (id == c.id) && (mergeto.Replace("(", "").Replace(")", "") == c.mergeto.Replace("(", "").Replace(")",""));
        }

        public override int GetHashCode()
        {
            // 例: Title と Level で比較している場合
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + (title != null ? title.GetHashCode() : 0);
                hash = hash * 23 + num.GetHashCode();
                return hash;
            }
        }
    }
}
