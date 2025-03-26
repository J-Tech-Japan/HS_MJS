namespace WordAddIn1
{
    /// <summary>
    /// 比較結果
    /// </summary>
    public class CheckInfo
    {
        private int _old1 = 0;
        public int old1
        {
            get { return _old1; }
            set { this._old1 = value; }
        }
        private int _old2 = 0;
        public int old2
        {
            get { return _old2; }
            set { this._old2 = value; }
        }
        private int _old3 = 0;
        public int old3
        {
            get { return _old3; }
            set { this._old3 = value; }
        }
        private int _old4 = 0;
        public int old4
        {
            get { return _old4; }
            set { this._old4 = value; }
        }

        private int _new1 = 0;
        public int new1
        {
            get { return _new1; }
            set { this._new1 = value; }
        }
        private int _new2 = 0;
        public int new2
        {
            get { return _new2; }
            set { this._new2 = value; }
        }
        private int _new3 = 0;
        public int new3
        {
            get { return _new3; }
            set { this._new3 = value; }
        }
        private int _new4 = 0;
        public int new4
        {
            get { return _new4; }
            set { this._new4 = value; }
        }

        // 旧.項番
        private string _old_num = "";
        public string old_num
        {
            get { return _old_num; }
            set
            {
                this._old_num = value;

                if (string.IsNullOrEmpty(value))
                {
                    return;
                }

                string[] oldnums = value.Split('.');
                if (oldnums.Length == 4)
                {
                    old1 = int.Parse(oldnums[0]);
                    old2 = int.Parse(oldnums[1]);
                    old3 = int.Parse(oldnums[2]);
                    old4 = int.Parse(oldnums[3]);
                }
                else if (oldnums.Length == 3)
                {
                    old1 = int.Parse(oldnums[0]);
                    old2 = int.Parse(oldnums[1]);
                    old3 = int.Parse(oldnums[2]);
                }
                else if (oldnums.Length == 2)
                {
                    old1 = int.Parse(oldnums[0]);
                    old2 = int.Parse(oldnums[1]);
                }
                else if (oldnums.Length == 1)
                {
                    old1 = int.Parse(oldnums[0]);
                }
            }
        }

        // 旧.タイトル
        private string _old_title = "";
        public string old_title
        {
            get { return _old_title; }
            set { this._old_title = value; }
        }

        // 旧.ID
        private string _old_id = "";
        public string old_id
        {
            get { return _old_id; }
            set { this._old_id = value; }
        }

        // 新.項番
        private string _new_num = "";
        public string new_num
        {
            get { return _new_num; }
            set
            {
                this._new_num = value;

                if (string.IsNullOrEmpty(value))
                {
                    return;
                }

                string[] newnums = value.Split('.');
                if (newnums.Length == 4)
                {
                    new1 = int.Parse(newnums[0]);
                    new2 = int.Parse(newnums[1]);
                    new3 = int.Parse(newnums[2]);
                    new4 = int.Parse(newnums[3]);
                }
                else if (newnums.Length == 3)
                {
                    new1 = int.Parse(newnums[0]);
                    new2 = int.Parse(newnums[1]);
                    new3 = int.Parse(newnums[2]);
                }
                else if (newnums.Length == 2)
                {
                    new1 = int.Parse(newnums[0]);
                    new2 = int.Parse(newnums[1]);
                }
                else if (newnums.Length == 1)
                {
                    new1 = int.Parse(newnums[0]);
                }
            }
        }

        // 新.項番（色）
        private string _new_num_color = "";
        public string new_num_color
        {
            get { return _new_num_color; }
            set { this._new_num_color = value; }
        }

        // 新.タイトル
        private string _new_title = "";
        public string new_title
        {
            get { return _new_title; }
            set { this._new_title = value; }
        }

        // 新.タイトル（色）
        private string _new_title_color = "";
        public string new_title_color
        {
            get { return _new_title_color; }
            set { this._new_title_color = value; }
        }

        // 新.ID
        private string _new_id = "";
        public string new_id
        {
            get { return _new_id; }
            set { this._new_id = value; }
        }

        // 新.ID（色）
        private string _new_id_color = "";
        public string new_id_color
        {
            get { return _new_id_color; }
            set { this._new_id_color = value; }
        }

        // 新.ID（修正候補）
        private string _new_id_show = "";
        public string new_id_show
        {
            get { return _new_id_show; }
            set { this._new_id_show = value; }
        }

        // 新.ID（修正候補）色
        private string _new_id_show_color = "";
        public string new_id_show_color
        {
            get { return _new_id_show_color; }
            set { this._new_id_show_color = value; }
        }

        // 差異内容
        private string _diff = "";
        public string diff
        {
            get { return _diff; }
            set { this._diff = value; }
        }

        // 差異内容（色）
        private string _diff_color = "";
        public string diff_color
        {
            get { return _diff_color; }
            set { this._diff_color = value; }
        }

        // 修正処理（候補）
        private string _editshow = "";
        public string editshow
        {
            get { return _editshow; }
            set { this._editshow = value; }
        }

        // 修正処理（候補）（色）
        private string _editshow_color = "";
        public string editshow_color
        {
            get { return _editshow_color; }
            set { this._editshow_color = value; }
        }

        // 新規追加
        private string _edit = "";
        public string edit
        {
            get { return _edit; }
            set { this._edit = value; }
        }

        // 新規追加（色）
        private string _edit_color = "";
        public string edit_color
        {
            get { return _edit_color; }
            set { this._edit_color = value; }
        }

    }
}
