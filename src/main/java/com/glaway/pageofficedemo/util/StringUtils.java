package com.glaway.pageofficedemo.util;

public class StringUtils {

    public static boolean isEmpty(String str){
        if(null == str || str.isEmpty()){
            return true;
        }else
            return false;
    }

    public static boolean isNotEmpty(String str) {
        return !isEmpty(str);
    }
}
