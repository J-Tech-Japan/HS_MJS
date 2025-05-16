namespace WordAddIn1
{
    // 比較結果
    public class CheckInfo
    {
        // 旧.項番情報
        public int old1 { get; set; }
        public int old2 { get; set; }
        public int old3 { get; set; }
        public int old4 { get; set; }

        // 新.項番情報
        public int new1 { get; set; }
        public int new2 { get; set; }
        public int new3 { get; set; }
        public int new4 { get; set; }

        private string _old_num = "";

        public string old_num
        {
            get { return _old_num; }
            set
            {
                _old_num = value;
                ParseNum(value, out int n1, out int n2, out int n3, out int n4);
                old1 = n1; old2 = n2; old3 = n3; old4 = n4;
            }
        }

        private string _new_num = "";

        public string new_num
        {
            get { return _new_num; }
            set
            {
                _new_num = value;
                ParseNum(value, out int n1, out int n2, out int n3, out int n4);
                new1 = n1; new2 = n2; new3 = n3; new4 = n4;
            }
        }

        // 共通のパース処理
        private void ParseNum(string value, out int n1, out int n2, out int n3, out int n4)
        {
            n1 = n2 = n3 = n4 = 0;
            if (string.IsNullOrEmpty(value)) return;
            var nums = value.Split('.');
            if (nums.Length > 0) int.TryParse(nums[0], out n1);
            if (nums.Length > 1) int.TryParse(nums[1], out n2);
            if (nums.Length > 2) int.TryParse(nums[2], out n3);
            if (nums.Length > 3) int.TryParse(nums[3], out n4);
        }

        public string old_title { get; set; } = "";
        public string old_id { get; set; } = "";
        public string new_title { get; set; } = "";
        public string new_num_color { get; set; } = "";
        public string new_title_color { get; set; } = "";
        public string new_id { get; set; } = "";
        public string new_id_color { get; set; } = "";
        public string new_id_show { get; set; } = "";
        public string new_id_show_color { get; set; } = "";
        public string diff { get; set; } = "";
        public string diff_color { get; set; } = "";
        public string editshow { get; set; } = "";
        public string editshow_color { get; set; } = "";
        public string edit { get; set; } = "";
        public string edit_color { get; set; } = "";

        // 比較用メソッド（old1～old4の順で比較）
        public int CompareOldTo(CheckInfo other)
        {
            if (old1 != other.old1) return old1.CompareTo(other.old1);
            if (old2 != other.old2) return old2.CompareTo(other.old2);
            if (old3 != other.old3) return old3.CompareTo(other.old3);
            if (old4 != other.old4) return old4.CompareTo(other.old4);
            return 0;
        }
    }
}
