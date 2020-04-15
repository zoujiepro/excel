package pub.vie.excel.common.utils;

import java.io.Closeable;
import java.io.IOException;

/**
 * @Descrption :
 * @Author: zoujie
 * @Date: 2020-4-14
 */
public class CommonUtils {

    public static void close(Closeable... closeables){
        if(closeables != null &&  closeables.length > 0){
            for (Closeable closeable : closeables) {
                try {
                    if(closeable != null) {
                        closeable.close();
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static boolean isBlank(CharSequence cs){
        int strLen;
        if (cs != null && (strLen = cs.length()) != 0) {
            for(int i = 0; i < strLen; ++i) {
                if (!Character.isWhitespace(cs.charAt(i))) {
                    return false;
                }
            }

            return true;
        } else {
            return true;
        }
    }

    public static boolean arrayEmpty(Object[] objects) {
        return objects == null || objects.length < 1;
    }
}
