package com.july.excel.utils;

import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

/**
 * 集合操作工具
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-06 17:42
 **/
public class CollectionUtils {

    public CollectionUtils() {
    }

    public static <T> boolean isNotEmpty(Collection<T> collection) {
        return !isEmpty(collection);
    }

    public static <T> boolean isEmpty(Collection<T> collection) {
        return null == collection || collection.isEmpty();
    }

    public static boolean isEmpty(Map<?, ?> map) {
        return (map == null || map.isEmpty());
    }

    public static <T> List<T> slice(List<T> source, int from, int to) {
        if (isEmpty(source)) {
            return null;
        } else {
            int size = source.size();
            return to > size ? (List) slice(source.stream(), from, size).collect(Collectors.toList()) : (List) slice(source.stream(), from, to).collect(Collectors.toList());
        }
    }

    public static <T> Stream<T> slice(Stream<T> stream, int startIndex, int endIndex) {
        return ((List) stream.collect(Collectors.toList())).subList(startIndex, endIndex).stream();
    }

    public static <T> List<List<T>> partition(List<T> list, int size) {
        if (list != null && list.size() > 0) {
            return new ArrayList(((Map) IntStream.range(0, list.size()).boxed().collect(Collectors.groupingBy((e) -> {
                return e / size;
            }, Collectors.mapping((e) -> {
                return list.get(e);
            }, Collectors.toList())))).values());
        } else {
            throw new NullPointerException("List can't be null");
        }
    }

    public static <T> List<T> asList(T... a) {
        return (List) (null != a ? Arrays.asList(a) : new ArrayList(0));
    }

    public static <T> List<T> mergeAndSwap(List<T> result1, List<T> result2) {
        if (isEmpty(result1) && isEmpty(result2)) {
            return new ArrayList(0);
        } else if (isEmpty(result1) && isNotEmpty(result2)) {
            return result2;
        } else if (isNotEmpty(result1) && isEmpty(result2)) {
            return result1;
        } else {
            int a = result1.size();
            int b = result2.size();
            int size = a + b;
            List<T> finalResult = new ArrayList(size);
            int i;
            if (a >= b) {
                for (i = 0; i < size; ++i) {
                    if (i > (b << 1) - 1) {
                        finalResult.add(result1.get(i - b));
                    } else if ((i & 1) == 0 && i >> 1 < a) {
                        finalResult.add(result1.get(i >> 1));
                    }

                    if ((i & 1) == 1 && i - 1 >> 1 < b) {
                        finalResult.add(result2.get(i - 1 >> 1));
                    }
                }
            } else {
                for (i = 0; i < size; ++i) {
                    if ((i & 1) == 0 && i >> 1 < a) {
                        finalResult.add(result1.get(i >> 1));
                    }

                    if ((i & 1) == 1 && i >> 1 < a - 1) {
                        finalResult.add(result2.get(i - 1 >> 1));
                    }

                    if (i >= (a << 1) - 1) {
                        finalResult.add(result2.get(i - a));
                    }
                }
            }

            return finalResult;
        }
    }

    public static List<Map<String, Object>> filterEmpty(List<Map<String, Object>> list, String... exceptKeys) {
        List<Map<String, Object>> tempList = new ArrayList();
        Iterator var3 = list.iterator();

        while (var3.hasNext()) {
            Map<String, Object> map = (Map) var3.next();
            int a = map.size();
            int counter = 0;
            Iterator var7 = map.entrySet().iterator();

            while (var7.hasNext()) {
                Map.Entry<String, Object> entry = (Map.Entry) var7.next();
                String[] var9 = exceptKeys;
                int var10 = exceptKeys.length;

                for (int var11 = 0; var11 < var10; ++var11) {
                    String str = var9[var11];
                    if (!((String) entry.getKey()).equals(str)) {
                        Object value = entry.getValue();
                        if (value instanceof String) {
                            if (StringUtils.isEmpty(String.valueOf(value))) {
                                ++counter;
                            }
                        } else if (null == value) {
                            ++counter;
                        }
                    }
                }
            }

            if (counter < a - exceptKeys.length) {
                tempList.add(map);
            }
        }

        return tempList;
    }
}
