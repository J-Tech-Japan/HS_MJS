using System.Collections.Generic;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private bool checkSortInfo(CheckInfo old, List<CheckInfo> newInfos, int j)
        {
            bool ret = false;

            CheckInfo newInfo = newInfos[j];

            if (old.old1 < newInfo.old1)
            {
                ret = true;
            }
            else if (old.old1 == newInfo.old1 && old.old2 < newInfo.old2)
            {
                ret = true;
            }
            else if (old.old1 == newInfo.old1 && old.old2 == newInfo.old2 && old.old3 < newInfo.old3)
            {
                ret = true;
            }
            else if (old.old1 == newInfo.old1 && old.old2 == newInfo.old2 && old.old3 == newInfo.old3 && old.old4 < newInfo.old4)
            {
                ret = true;
            }

            for (int k = j + 1; k < newInfos.Count; k++)
            {
                CheckInfo newInfoK = newInfos[k];

                if (string.IsNullOrEmpty(newInfoK.old_id))
                {
                    continue;
                }

                if (old.old1 > newInfoK.old1)
                {
                    ret = false;
                }
                else if (old.old1 == newInfoK.old1 && old.old2 > newInfoK.old2)
                {
                    ret = false;
                }
                else if (old.old1 == newInfoK.old1 && old.old2 == newInfoK.old2 && old.old3 > newInfoK.old3)
                {
                    ret = false;
                }
                else if (old.old1 == newInfoK.old1 && old.old2 == newInfoK.old2 && old.old3 == newInfoK.old3 && old.old4 > newInfoK.old4)
                {
                    ret = false;
                }
            }

            return ret;
        }
    }
}
