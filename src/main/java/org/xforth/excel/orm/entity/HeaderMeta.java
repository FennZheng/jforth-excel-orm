package org.xforth.excel.orm.entity;

import org.xforth.excel.orm.exception.HeaderNotMatchException;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

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
    public List<String> getHeaderTitle(){
        List<String> headerTitleList = new ArrayList<String>(titlePropertyMap.size());
        Iterator<String> iterator = titlePropertyMap.keySet().iterator();
        while(iterator.hasNext()){
            headerTitleList.add(iterator.next());
        }
        return headerTitleList;
    }
}
