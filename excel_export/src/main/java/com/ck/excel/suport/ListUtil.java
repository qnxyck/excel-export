package com.ck.excel.suport;

import java.util.ArrayList;
import java.util.List;

/**
 * List拆分多份
 *
 * @author QianNianXiaoYao
 */
public class ListUtil {

    /**
     * 将 ls 根据 num 大小拆分多份并返回
     *
     * @param ls  ls
     * @param num 大小
     * @param <T> T
     * @return .
     * @author QianNianXiaoYao
     */
    public static <T> List<List<T>> splitList(List<T> ls, int num) {
        List<List<T>> lists = new ArrayList<>();
        int iMax = ls.size() / num;
        if (iMax > 0) {
            for (int i = 0; i < iMax; i++) {
                lists.add(ls.subList(i * num, (i + 1) * num));
            }
            int sInt = iMax * num;
            if (sInt < ls.size()) {
                lists.add(ls.subList(sInt, ls.size()));
            }
        } else {
            lists.add(ls);
        }
        return lists;
    }
}
