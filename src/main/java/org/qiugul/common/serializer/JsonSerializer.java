package org.qiugul.common.serializer;

import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.util.Optional;

public class JsonSerializer extends ObjectMapper {

    private final static JsonSerializer INSTANCE = new JsonSerializer();

    static {
        INSTANCE.findAndRegisterModules();  // 自动加载所有模块（如Java 8时间支持）
        INSTANCE.configure(DeserializationFeature.ACCEPT_EMPTY_STRING_AS_NULL_OBJECT, true);

    }

    public static JsonSerializer getInstance(){
        return INSTANCE;
    }


    public static Optional<JsonNode> getNode(JsonNode root, String... nodes){

        if (root == null){
            return Optional.empty();
        }


        JsonNode cur = root;
        for (String node : nodes) {
//            if (node.endsWith(Const.RIGHT_BRACKET)){
//                int lastLeftBracketIdx = node.lastIndexOf(Const.LEFT_BRACKET);
//                String fieldName = node.substring(0, lastLeftBracketIdx);
//                String idx = node.substring(lastLeftBracketIdx + 1, node.length() - 1);
//
//            }
            JsonNode next = cur.get(node);
            if (next == null){
                return Optional.empty();
            }
            cur  = next;
        }

        return Optional.ofNullable(cur);

    }


}
