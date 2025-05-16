using System.Collections.Generic;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private bool CheckSortInfo(CheckInfo old, List<CheckInfo> newInfos, int j)
        {
            // newInfoとの比較
            bool ret = old.CompareOldTo(newInfos[j]) < 0;

            // newInfos[j+1]以降と比較
            for (int k = j + 1; k < newInfos.Count; k++)
            {
                CheckInfo newInfoK = newInfos[k];
                if (string.IsNullOrEmpty(newInfoK.old_id))
                {
                    continue;
                }
                if (old.CompareOldTo(newInfoK) > 0)
                {
                    ret = false;
                }
            }
            return ret;
        }
    }
}
