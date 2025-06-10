using System.Collections.Generic;
using System.Linq;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private int CheckDocInfo(List<HeadingInfo> oldInfos, List<HeadingInfo> newInfos, out List<CheckInfo> checkResult)
        {
            checkResult = new List<CheckInfo>();
            var syoriList = new List<CheckInfo>();
            var deleteList = new List<CheckInfo>();
            int returnCode = 0;

            // 一致・削除判定
            ProcessMatchAndDelete(oldInfos, newInfos, syoriList, deleteList, ref returnCode);
            // 新規追加判定
            ProcessAdditions(oldInfos, newInfos, syoriList);
            // ID不一致・構成変更・見出しレベル変更
            ProcessIdMismatch(oldInfos, newInfos, syoriList, ref returnCode);
            // タイトル変更
            ProcessTitleChange(oldInfos, newInfos, syoriList, ref returnCode);
            // 削除再判定
            ProcessDeleteRecheck(oldInfos, syoriList, deleteList);

            // ソート
            deleteList = deleteList.OrderBy(rec => rec.old1).ThenBy(rec => rec.old2).ThenBy(rec => rec.old3).ThenBy(rec => rec.old4).ToList();
            syoriList = syoriList.OrderBy(rec => rec.new1).ThenBy(rec => rec.new2).ThenBy(rec => rec.new3).ThenBy(rec => rec.new4).ToList();

            MergeResultLists(deleteList, syoriList, out checkResult);

            // 完全一致チェック
            if (newInfos.Count == oldInfos.Count)
            {
                foreach (var newInfo in newInfos)
                {
                    var checkHeadingInfo = oldInfos.Where(x => x.id == newInfo.id && x.num == newInfo.num && x.mergeto == newInfo.mergeto && x.title == newInfo.title);
                    if (!checkHeadingInfo.Any())
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

        // 一致・削除判定
        private void ProcessMatchAndDelete(List<HeadingInfo> oldInfos, List<HeadingInfo> newInfos, List<CheckInfo> syoriList, List<CheckInfo> deleteList, ref int returnCode)
        {
            foreach (var oldInfo in oldInfos)
            {
                bool oldTitleExist = false;
                bool oldIdExist = false;
                foreach (var newInfo in newInfos)
                {
                    if (oldInfo.title.Equals(newInfo.title) && oldInfo.id.Equals(newInfo.id))
                    {
                        var checkInfo = CreateCheckInfoMatch(oldInfo, newInfo);
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
                    if (oldInfo.title.Equals(newInfo.title)) oldTitleExist = true;
                    if (oldInfo.id.Equals(newInfo.id)) oldIdExist = true;
                }
                if (!oldTitleExist && !oldIdExist)
                {
                    var checkInfo = CreateCheckInfoDelete(oldInfo);
                    deleteList.Add(checkInfo);
                }
            }
        }

        // 新規追加判定
        private void ProcessAdditions(List<HeadingInfo> oldInfos, List<HeadingInfo> newInfos, List<CheckInfo> syoriList)
        {
            foreach (var newInfo in newInfos)
            {
                bool newTitleExist = false;
                bool newIdExist = false;
                foreach (var oldInfo in oldInfos)
                {
                    if (oldInfo.id.Equals(newInfo.id)) newIdExist = true;
                    if (oldInfo.title.Equals(newInfo.title)) newTitleExist = true;
                }
                if (!newTitleExist && !newIdExist)
                {
                    var checkInfo = CreateCheckInfoAdd(newInfo);
                    syoriList.Add(checkInfo);
                }
            }
        }

        // ID不一致・構成変更・見出しレベル変更
        private void ProcessIdMismatch(List<HeadingInfo> oldInfos, List<HeadingInfo> newInfos, List<CheckInfo> syoriList, ref int returnCode)
        {
            foreach (var newInfo in newInfos)
            {
                foreach (var oldInfo in oldInfos)
                {
                    if (syoriList.Any(p => p.new_id.Equals(newInfo.id))) break;
                    if (oldInfo.title.Equals(newInfo.title) && !oldInfo.id.Equals(newInfo.id))
                    {
                        int oldNumKaisou = oldInfo.num.Split('.').Length;
                        int newNumKaisou = newInfo.num.Split('.').Length;
                        if ((oldNumKaisou == 3 && newNumKaisou == 4) || (oldNumKaisou == 4 && newNumKaisou == 3))
                        {
                            var checkInfo = CreateCheckInfoLevelChange(oldInfo, newInfo);
                            syoriList.Add(checkInfo);
                        }
                        else
                        {
                            bool isHenko = false;
                            if (oldNumKaisou == 4 && newNumKaisou == 4)
                            {
                                string[] oldids = oldInfo.id.Split('#');
                                string[] newids = newInfo.id.Split('#');
                                if (oldids.Length == 2 && newids.Length == 2 && oldids[1].Equals(newids[1]))
                                {
                                    var checkInfo2 = CreateCheckInfoStructureChange(oldInfo, newInfo);
                                    syoriList.Add(checkInfo2);
                                    isHenko = true;
                                }
                            }
                            if (!isHenko)
                            {
                                var checkInfo = CreateCheckInfoIdMismatch(oldInfo, newInfo);
                                syoriList.Add(checkInfo);
                                returnCode = 1;
                            }
                        }
                    }
                }
            }
        }

        // タイトル変更
        private void ProcessTitleChange(List<HeadingInfo> oldInfos, List<HeadingInfo> newInfos, List<CheckInfo> syoriList, ref int returnCode)
        {
            foreach (var newInfo in newInfos)
            {
                if (syoriList.Any(p => p.new_id.Equals(newInfo.id))) continue;
                foreach (var oldInfo in oldInfos)
                {
                    if (oldInfo.id.Equals(newInfo.id) && !oldInfo.title.Equals(newInfo.title))
                    {
                        var checkInfo = CreateCheckInfoTitleChange(oldInfo, newInfo);
                        syoriList.Add(checkInfo);
                        returnCode = 1;
                    }
                }
            }
        }

        // 削除再判定
        private void ProcessDeleteRecheck(List<HeadingInfo> oldInfos, List<CheckInfo> syoriList, List<CheckInfo> deleteList)
        {
            foreach (var oldInfo in oldInfos)
            {
                if (syoriList.Any(p => p.old_num.Equals(oldInfo.num))) continue;
                if (deleteList.Any(p => p.old_num.Equals(oldInfo.num))) continue;
                var checkInfo = CreateCheckInfoDelete(oldInfo);
                deleteList.Add(checkInfo);
            }
        }

        // 結果リストのマージ
        private void MergeResultLists(List<CheckInfo> deleteList, List<CheckInfo> syoriList, out List<CheckInfo> checkResult)
        {
            checkResult = new List<CheckInfo>();
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
                        if (deleteList.Count == i) stopFlag = true;
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
        }

        // --- CheckInfo生成ヘルパ ---
        private CheckInfo CreateCheckInfoMatch(HeadingInfo oldInfo, HeadingInfo newInfo)
        {
            var checkInfo = new CheckInfo
            {
                old_num = oldInfo.num,
                old_title = oldInfo.title,
                old_id = !string.IsNullOrEmpty(oldInfo.mergeto) ? oldInfo.id + " " + oldInfo.mergeto : oldInfo.id,
                new_num = newInfo.num,
                new_title = newInfo.title,
                new_id = !string.IsNullOrEmpty(newInfo.mergeto) ? newInfo.id + " (" + newInfo.mergeto + ")" : newInfo.id,
                new_id_show = !string.IsNullOrEmpty(newInfo.mergeto) ? newInfo.id + " (" + newInfo.mergeto + ")" : newInfo.id
            };
            return checkInfo;
        }
        private CheckInfo CreateCheckInfoDelete(HeadingInfo oldInfo)
        {
            return new CheckInfo
            {
                old_num = oldInfo.num,
                old_title = oldInfo.title,
                old_id = !string.IsNullOrEmpty(oldInfo.mergeto) ? oldInfo.id + " " + oldInfo.mergeto : oldInfo.id,
                diff = "削除"
            };
        }
        private CheckInfo CreateCheckInfoAdd(HeadingInfo newInfo)
        {
            var merged = !string.IsNullOrEmpty(newInfo.mergeto);
            return new CheckInfo
            {
                new_num = newInfo.num,
                new_num_color = "blue",
                new_title = newInfo.title,
                new_title_color = "blue",
                new_id = merged ? newInfo.id + " (" + newInfo.mergeto + ")" : newInfo.id,
                new_id_show = merged ? newInfo.id + " (" + newInfo.mergeto + ")" : newInfo.id,
                new_id_color = "blue",
                diff = merged ? "新規追加・結合追加" : "新規追加"
            };
        }
        private CheckInfo CreateCheckInfoLevelChange(HeadingInfo oldInfo, HeadingInfo newInfo)
        {
            return new CheckInfo
            {
                old_num = oldInfo.num,
                old_title = oldInfo.title,
                old_id = !string.IsNullOrEmpty(oldInfo.mergeto) ? oldInfo.id + " " + oldInfo.mergeto : oldInfo.id,
                new_num = newInfo.num,
                new_num_color = "red",
                new_title = newInfo.title,
                new_id = !string.IsNullOrEmpty(newInfo.mergeto) ? newInfo.id + " (" + newInfo.mergeto + ")" : newInfo.id,
                new_id_show = !string.IsNullOrEmpty(newInfo.mergeto) ? newInfo.id + " (" + newInfo.mergeto + ")" : newInfo.id,
                new_id_color = "red",
                diff = "見出しレベル変更"
            };
        }
        private CheckInfo CreateCheckInfoStructureChange(HeadingInfo oldInfo, HeadingInfo newInfo)
        {
            return new CheckInfo
            {
                old_num = oldInfo.num,
                old_title = oldInfo.title,
                old_id = !string.IsNullOrEmpty(oldInfo.mergeto) ? oldInfo.id + " " + oldInfo.mergeto : oldInfo.id,
                new_num = newInfo.num,
                new_num_color = "red",
                new_title = newInfo.title,
                new_id = !string.IsNullOrEmpty(newInfo.mergeto) ? newInfo.id + " (" + newInfo.mergeto + ")" : newInfo.id,
                new_id_show = !string.IsNullOrEmpty(newInfo.mergeto) ? newInfo.id + " (" + newInfo.mergeto + ")" : newInfo.id,
                new_id_color = "red",
                diff = "構成変更に伴うID変更"
            };
        }
        private CheckInfo CreateCheckInfoIdMismatch(HeadingInfo oldInfo, HeadingInfo newInfo)
        {
            var checkInfo = new CheckInfo
            {
                old_num = oldInfo.num,
                old_title = oldInfo.title,
                old_id = !string.IsNullOrEmpty(oldInfo.mergeto) ? oldInfo.id + " " + oldInfo.mergeto : oldInfo.id,
                new_num = newInfo.num,
                new_title = newInfo.title,
                new_id = !string.IsNullOrEmpty(newInfo.mergeto) ? newInfo.id + " (" + newInfo.mergeto + ")" : newInfo.id,
                new_id_color = "red",
                new_id_show = !string.IsNullOrEmpty(newInfo.mergeto) ? newInfo.id + " (" + newInfo.mergeto + ")" : oldInfo.id,
                diff = "ID不一致？",
                diff_color = "red",
                editshow = "旧IDに戻す"
            };
            if (!oldInfo.num.Equals(newInfo.num)) checkInfo.new_num_color = "red";
            if (oldInfo.mergeto.Equals("") && !newInfo.mergeto.Equals("")) checkInfo.diff = "ID不一致？・結合追加";
            else if (!oldInfo.mergeto.Equals("") && newInfo.mergeto.Equals("")) checkInfo.diff = "ID不一致？・結合解除";
            return checkInfo;
        }
        private CheckInfo CreateCheckInfoTitleChange(HeadingInfo oldInfo, HeadingInfo newInfo)
        {
            var checkInfo = new CheckInfo
            {
                old_num = oldInfo.num,
                old_title = oldInfo.title,
                old_id = !string.IsNullOrEmpty(oldInfo.mergeto) ? oldInfo.id + " " + oldInfo.mergeto : oldInfo.id,
                new_num = newInfo.num,
                new_title = newInfo.title,
                new_title_color = "red",
                new_id = !string.IsNullOrEmpty(newInfo.mergeto) ? newInfo.id + " (" + newInfo.mergeto + ")" : newInfo.id,
                new_id_show = !string.IsNullOrEmpty(newInfo.mergeto) ? newInfo.id + " (" + newInfo.mergeto + ")" : newInfo.id,
                diff = "●タイトル変更",
                edit = "○新規追加",
                edit_color = "blue"
            };
            if (!oldInfo.num.Equals(newInfo.num)) checkInfo.new_num_color = "red";
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
            return checkInfo;
        }
    }
}
