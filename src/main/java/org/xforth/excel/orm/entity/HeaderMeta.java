package org.xforth.excel.orm.entity;

import org.xforth.excel.orm.exception.HeaderNotMatchException;

import java.util.HashMap;

/**
 * thread unsafe
 */
public final class HeaderMeta {
    private HashMap<String,String> titlePropertyMap;

    public HeaderMeta() {
        titlePropertyMap = new HashMap<>();
    }

    public void put(String headerTitle,String propertyName){
        if(titlePropertyMap.containsKey(headerTitle)){
            throw new HeaderNotMatchException("headerTitle map duplicate propertyName for title:"
                    +headerTitle);
        }else {
            titlePropertyMap.put(headerTitle, propertyName);
        }
    }
    public boolean contains(String headerTitle){
        return titlePropertyMap.containsKey(headerTitle);
    }
    public int size(){
        return titlePropertyMap.size();
    }
    public String getPropertyNameByHeaderTitle(String headerTitle){
        return titlePropertyMap.get(headerTitle);
    }
}
