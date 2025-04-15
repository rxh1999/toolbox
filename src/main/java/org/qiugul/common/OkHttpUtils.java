package org.qiugul.common;

import okhttp3.OkHttpClient;

public class OkHttpUtils {
    private static final OkHttpClient client = new OkHttpClient();
    public static final OkHttpClient getClient(){
        return client;
    }
}
