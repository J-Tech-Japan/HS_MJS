using System.Collections.Generic;
using System.Linq;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private int CheckDocInfo(List<HeadingInfo> oldInfos, List<HeadingInfo> newInfos, out List<CheckInfo> checkResult)
        {
            // 比較結果リスト初期化する
            checkResult = new List<CheckInfo>();
            List<CheckInfo> syoriList = new List<CheckInfo>();
            List<CheckInfo> deleteList = new List<CheckInfo>();
            int returnCode = 0;

            // 一致判定と削除判定
            foreach (HeadingInfo oldInfo in oldInfos)
            {
                bool oldTitleExist = false;
                bool oldIdExist = false;

                foreach (HeadingInfo newInfo in newInfos)
                {
                    // 書誌情報（新）.タイトル＝書誌情報（旧）.タイトルかつ書誌情報（新）.ID＝書誌情報（旧）.IDが存在する場合
                    if (oldInfo.title.Equals(newInfo.title) && oldInfo.id.Equals(newInfo.id))
                    {
                        // 比較結果（一致）を作成する
                        CheckInfo checkInfo = new CheckInfo();
                        // 旧.項番
                        checkInfo.old_num = oldInfo.num;
                        // 旧.タイトル
                        checkInfo.old_title = oldInfo.title;
                        // 旧.ID
                        checkInfo.old_id = oldInfo.id;
                        // 旧.ID結合済
                        if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                        // 新.項番
                        checkInfo.new_num = newInfo.num;
                        // 新.タイトル
                        checkInfo.new_title = newInfo.title;
                        // 新.ID
                        checkInfo.new_id = newInfo.id;
                        // 新.ID結合済
                        if (!newInfo.mergeto.Equals("")) { checkInfo.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                        // 新.ID（修正候補）
                        checkInfo.new_id_show = newInfo.id;
                        // 新.ID（修正候補）結合済
                        if (!newInfo.mergeto.Equals("")) { checkInfo.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }

                        // check merge 
                        if (oldInfo.mergeto.Equals("") && !newInfo.mergeto.Equals(""))
                        {
                            checkInfo.diff = "結合追加";
                            checkInfo.new_id_color = "red";
                            returnCode = 1;
                        }
                        else if (!oldInfo.mergeto.Equals("") && newInfo.mergeto.Equals(""))
                        {
                            checkInfo.diff = "結合解除";
                            checkInfo.new_id_color = "red";
                            returnCode = 1;
                        }

                        syoriList.Add(checkInfo);
                    }

                    // 書誌情報（新）.タイトル＝書誌情報（旧）.タイトル
                    if (oldInfo.title.Equals(newInfo.title))
                    {
                        oldTitleExist = true;
                    }

                    // 書誌情報（新）.ID＝書誌情報（旧）.IDが存在する場合
                    if (oldInfo.id.Equals(newInfo.id))
                    {
                        oldIdExist = true;
                    }
                }

                // 書誌情報（旧）.タイトルと書誌情報（旧）.IDが書誌情報（新）に存在しない場合
                if (!oldTitleExist && !oldIdExist)
                {
                    // 比較結果（削除）を作成する
                    CheckInfo checkInfo = new CheckInfo();
                    // 旧.項番
                    checkInfo.old_num = oldInfo.num;
                    // 旧.タイトル
                    checkInfo.old_title = oldInfo.title;
                    // 旧.ID
                    checkInfo.old_id = oldInfo.id;
                    // 旧.ID結合済
                    if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                    // 差異内容
                    checkInfo.diff = "削除";

                    deleteList.Add(checkInfo);
                }
            }

            // 新規判定
            foreach (HeadingInfo newInfo in newInfos)
            {
                bool newTitleExist = false;
                bool newIdExist = false;

                foreach (HeadingInfo oldInfo in oldInfos)
                {
                    // 書誌情報（新）.ID＝書誌情報（旧）.IDが存在する場合
                    if (oldInfo.id.Equals(newInfo.id))
                    {
                        newIdExist = true;
                    }

                    // 書誌情報（新）.タイトル＝書誌情報（旧）.タイトル
                    if (oldInfo.title.Equals(newInfo.title))
                    {
                        newTitleExist = true;
                    }
                }

                // 書誌情報（新）.タイトルと書誌情報（新）.IDが書誌情報（旧）に存在しない場合
                if (!newTitleExist && !newIdExist)
                {
                    // 比較結果（新規）を作成する
                    CheckInfo checkInfo = new CheckInfo();
                    // 新.項番
                    checkInfo.new_num = newInfo.num;
                    // 新.項番（色）
                    checkInfo.new_num_color = "blue";
                    // 新.タイトル
                    checkInfo.new_title = newInfo.title;
                    // 新.タイトル（色）
                    checkInfo.new_title_color = "blue";
                    // 新.ID
                    checkInfo.new_id = newInfo.id;
                    // 新.ID結合済
                    if (!newInfo.mergeto.Equals("")) { checkInfo.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                    // 新.ID（修正候補）
                    checkInfo.new_id_show = newInfo.id;
                    // 新.ID（修正候補）結合済
                    if (!newInfo.mergeto.Equals("")) { checkInfo.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }
                    // 新.ID（色）
                    checkInfo.new_id_color = "blue";

                    // 差異内容
                    checkInfo.diff = "新規追加";

                    // ＋結合追加
                    if (!newInfo.mergeto.Equals(""))
                    {
                        checkInfo.diff = "新規追加・結合追加";

                    }

                    syoriList.Add(checkInfo);
                }
            }

            // ID不一致判定
            foreach (HeadingInfo newInfo in newInfos)
            {
                foreach (HeadingInfo oldInfo in oldInfos)
                {
                    // リストに存在するか
                    CheckInfo hasOne = syoriList.Where(p => p.new_id.Equals(newInfo.id)).FirstOrDefault();
                    if (hasOne != null)
                    {
                        break;
                    }

                    // 書誌情報（新）.タイトル＝書誌情報（旧）.タイトル
                    if (oldInfo.title.Equals(newInfo.title))
                    {
                        // 書誌情報（新）.ID<>書誌情報（旧）.ID
                        if (!oldInfo.id.Equals(newInfo.id))
                        {
                            // 項番階層
                            string oldNum = oldInfo.num;
                            string newNum = newInfo.num;
                            int oldNumKaisou = oldNum.Split('.').Length;
                            int newNumKaisou = newNum.Split('.').Length;

                            // (旧.見出しレベルが3 階層かつ新.見出しレベルが４階層) 
                            // または　(旧.見出しレベルが4 階層かつ新.見出しレベルが3階層) )の場合
                            if ((oldNumKaisou == 3 && newNumKaisou == 4)
                                || (oldNumKaisou == 4 && newNumKaisou == 3))
                            {
                                // 比較結果（見出しレベル変更）を作成する
                                CheckInfo checkInfo = new CheckInfo();
                                // 旧.項番
                                checkInfo.old_num = oldInfo.num;
                                // 旧.タイトル
                                checkInfo.old_title = oldInfo.title;
                                // 旧.ID
                                checkInfo.old_id = oldInfo.id;
                                // 旧.ID結合済
                                if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                                // 新.項番
                                checkInfo.new_num = newInfo.num;
                                // 新.項番（色）
                                checkInfo.new_num_color = "red";
                                // 新.タイトル
                                checkInfo.new_title = newInfo.title;
                                // 新.ID
                                checkInfo.new_id = newInfo.id;
                                // 新.ID結合済
                                if (!newInfo.mergeto.Equals("")) { checkInfo.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                                // 新.ID（修正候補）
                                checkInfo.new_id_show = newInfo.id;
                                // 新.ID（修正候補）結合済
                                if (!newInfo.mergeto.Equals("")) { checkInfo.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }
                                // 新.ID（色）
                                checkInfo.new_id_color = "red";
                                // 差異内容
                                checkInfo.diff = "見出しレベル変更";

                                syoriList.Add(checkInfo);
                            }
                            else
                            {
                                // 構成変更に伴うID変更
                                bool isHenko = false;
                                if (oldNumKaisou == 4 && newNumKaisou == 4)
                                {
                                    string[] oldids = oldInfo.id.Split('#');
                                    string[] newids = newInfo.id.Split('#');

                                    if (oldids.Length == 2 && newids.Length == 2
                                        && oldids[1].Equals(newids[1]))
                                    {

                                        // 比較結果（構成変更に伴うID変更）を作成する
                                        CheckInfo checkInfo2 = new CheckInfo();
                                        // 旧.項番
                                        checkInfo2.old_num = oldInfo.num;
                                        // 旧.タイトル
                                        checkInfo2.old_title = oldInfo.title;
                                        // 旧.ID
                                        checkInfo2.old_id = oldInfo.id;
                                        // 旧.ID結合済
                                        if (!oldInfo.mergeto.Equals("")) { checkInfo2.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                                        // 新.項番
                                        checkInfo2.new_num = newInfo.num;
                                        // 新.項番（色）
                                        checkInfo2.new_num_color = "red";
                                        // 新.タイトル
                                        checkInfo2.new_title = newInfo.title;
                                        // 新.ID
                                        checkInfo2.new_id = newInfo.id;
                                        // 新.ID結合済
                                        if (!newInfo.mergeto.Equals("")) { checkInfo2.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                                        // 新.ID（修正候補）
                                        checkInfo2.new_id_show = newInfo.id;
                                        // 新.ID（修正候補）結合済
                                        if (!newInfo.mergeto.Equals("")) { checkInfo2.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }

                                        // 新.ID（色）
                                        checkInfo2.new_id_color = "red";
                                        // 差異内容
                                        checkInfo2.diff = "構成変更に伴うID変更";

                                        syoriList.Add(checkInfo2);

                                        isHenko = true;
                                    }

                                }

                                if (!isHenko)
                                {
                                    // 比較結果（ID不一致）を作成する
                                    CheckInfo checkInfo = new CheckInfo();
                                    // 旧.項番
                                    checkInfo.old_num = oldInfo.num;
                                    // 旧.タイトル
                                    checkInfo.old_title = oldInfo.title;
                                    // 旧.ID
                                    checkInfo.old_id = oldInfo.id;
                                    // 旧.ID結合済
                                    if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                                    // 新.項番
                                    checkInfo.new_num = newInfo.num;
                                    // 新.項番（色）
                                    // 旧.項番<>新.項番の場合、赤
                                    if (!oldInfo.num.Equals(newInfo.num))
                                    {
                                        checkInfo.new_num_color = "red";
                                    }
                                    // 新.タイトル
                                    checkInfo.new_title = newInfo.title;
                                    // 新.ID
                                    checkInfo.new_id = newInfo.id;
                                    // 新.ID結合済
                                    if (!newInfo.mergeto.Equals("")) { checkInfo.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                                    // 新.ID（色）
                                    checkInfo.new_id_color = "red";
                                    // 新.ID（修正候補）
                                    checkInfo.new_id_show = oldInfo.id;
                                    // 新/ID（修正候補）結合済
                                    if (!newInfo.mergeto.Equals("")) { checkInfo.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }
                                    // 差異内容
                                    checkInfo.diff = "ID不一致？";
                                    // 差異内容（色）
                                    checkInfo.diff_color = "red";

                                    // 修正処理（候補）
                                    checkInfo.editshow = "旧IDに戻す";

                                    // check merge 
                                    if (oldInfo.mergeto.Equals("") && !newInfo.mergeto.Equals(""))
                                    {
                                        checkInfo.diff = "ID不一致？・結合追加";
                                    }
                                    else if (!oldInfo.mergeto.Equals("") && newInfo.mergeto.Equals(""))
                                    {
                                        checkInfo.diff = "ID不一致？・結合解除";
                                    }

                                    syoriList.Add(checkInfo);

                                    returnCode = 1;
                                }
                            }
                        }
                    }
                }
            }

            // タイトル変更判定
            foreach (HeadingInfo newInfo in newInfos)
            {
                // リストに存在するか
                CheckInfo hasOne = syoriList.Where(p => p.new_id.Equals(newInfo.id)).FirstOrDefault();
                if (hasOne != null)
                {
                    continue;
                }

                foreach (HeadingInfo oldInfo in oldInfos)
                {
                    // 書誌情報（新）.ID＝書誌情報（旧）.IDが存在する場合
                    if (oldInfo.id.Equals(newInfo.id))
                    {
                        // 書誌情報（新）.タイトル<>書誌情報（旧）.タイトル
                        if (!oldInfo.title.Equals(newInfo.title))
                        {
                            // 比較結果（タイトル変更）を作成する
                            CheckInfo checkInfo = new CheckInfo();
                            // 旧.項番
                            checkInfo.old_num = oldInfo.num;
                            // 旧.タイトル
                            checkInfo.old_title = oldInfo.title;
                            // 旧.ID
                            checkInfo.old_id = oldInfo.id;
                            // 旧・ID結合済
                            if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                            // 新.項番
                            checkInfo.new_num = newInfo.num;

                            // 新.項番（色）
                            // 旧.項番<>新.項番の場合、赤
                            if (!oldInfo.num.Equals(newInfo.num))
                            {
                                checkInfo.new_num_color = "red";
                            }

                            // 新.タイトル
                            checkInfo.new_title = newInfo.title;
                            // 新.タイトル（色）
                            checkInfo.new_title_color = "red";
                            // 新.ID
                            checkInfo.new_id = newInfo.id;
                            // 新.ID結合済
                            if (!newInfo.mergeto.Equals("")) { checkInfo.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                            // 新.ID（修正候補）
                            checkInfo.new_id_show = newInfo.id;
                            // 新.ID（修正候補）結合済
                            if (!newInfo.mergeto.Equals("")) { checkInfo.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }

                            // 差異内容
                            checkInfo.diff = "●タイトル変更";

                            // 新規追加
                            checkInfo.edit = "○新規追加";

                            // 新規追加（色）
                            checkInfo.edit_color = "blue";

                            // check merge 
                            if (oldInfo.mergeto.Equals("") && !newInfo.mergeto.Equals(""))
                            {
                                checkInfo.diff = "●タイトル変更・結合追加";
                                checkInfo.new_id_color = "red";
                            }
                            else if (!oldInfo.mergeto.Equals("") && newInfo.mergeto.Equals(""))
                            {
                                checkInfo.diff = "●タイトル変更・結合解除";
                                checkInfo.new_id_color = "red";
                            }

                            syoriList.Add(checkInfo);

                            returnCode = 1;
                        }
                    }
                }
            }

            // 削除再判定
            foreach (HeadingInfo oldInfo in oldInfos)
            {
                var issyori = syoriList.Where(p => p.old_num.Equals(oldInfo.num)).ToList();
                if (issyori != null && issyori.Count > 0)
                {
                    continue;
                }

                var isdelete = deleteList.Where(p => p.old_num.Equals(oldInfo.num)).ToList();
                if (isdelete != null && isdelete.Count > 0)
                {
                    continue;
                }

                // 比較結果（削除）を作成する
                CheckInfo checkInfo = new CheckInfo();
                // 旧.項番
                checkInfo.old_num = oldInfo.num;
                // 旧.タイトル
                checkInfo.old_title = oldInfo.title;
                // 旧.ID
                checkInfo.old_id = oldInfo.id;
                // 旧・ID結合済
                if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                // 差異内容
                checkInfo.diff = "削除";

                deleteList.Add(checkInfo);
            }

            // ソート
            deleteList = deleteList.OrderBy(rec => rec.old1).ThenBy(rec =>
            rec.old2).ThenBy(rec => rec.old3).ThenBy(rec => rec.old4).ToList();

            // ソート
            syoriList = syoriList.OrderBy(rec => rec.new1).ThenBy(rec =>
                rec.new2).ThenBy(rec => rec.new3).ThenBy(rec => rec.new4).ToList();

            if (deleteList.Count > 0)
            {
                int i = 0;
                bool stopFlag = false;

                for (int j = 0; j < syoriList.Count; j++)
                {
                    while (!stopFlag && CheckSortInfo(deleteList[i], syoriList, j))
                    {
                        checkResult.Add(deleteList[i]);
                        i++;

                        if (deleteList.Count == i)
                        {
                            stopFlag = true;
                        }
                    }

                    checkResult.Add(syoriList[j]);
                }

                while (i < deleteList.Count)
                {
                    checkResult.Add(deleteList[i]);
                    i++;
                }
            }
            else
            {
                checkResult = syoriList;
            }
            if (newInfos.Count == oldInfos.Count)
            {
                foreach (HeadingInfo newInfo in newInfos)
                {
                    var checkHeadingInfo = oldInfos.Where(x => x.id == newInfo.id && x.num == newInfo.num && x.mergeto == newInfo.mergeto && x.title == newInfo.title);
                    if (checkHeadingInfo == null)
                    {
                        returnCode = 1;
                        break;
                    }
                }

            }
            else
            {
                returnCode = 1;
            }


            return returnCode;
        }
    }
}
