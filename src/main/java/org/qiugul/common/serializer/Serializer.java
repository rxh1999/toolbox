package org.qiugul.common.serializer;

public interface Serializer<T> {
    public String serializeToString(T data) throws Exception;
    public T deserializeFromString(String str) throws Exception;
}
