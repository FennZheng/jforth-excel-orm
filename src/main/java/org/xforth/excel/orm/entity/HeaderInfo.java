package org.xforth.excel.orm.entity;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xforth.excel.orm.exception.HeaderNotMatchException;

import java.util.HashMap;

public final class HeaderInfo {
    private static final Logger logger = LoggerFactory.getLogger(HeaderInfo.class);
    private HashMap<Integer,String> indexTitleMap;
    private HeaderMeta headerMeta;

    public HeaderInfo(HeaderMeta headerMeta) {
        this.headerMeta = headerMeta;
        indexTitleMap = new HashMap<>();
    }
    public void put(int index,String headerTitle){
        if(indexTitleMap.get(index)!=null){
            throw new HeaderNotMatchException("header index: "+index+" duplicate orgin value:"
            +indexTitleMap.get(index)+" new value："+headerTitle);
        }
        if(!headerMeta.contains(headerTitle)){
            if(logger.isDebugEnabled()){
                logger.debug("unMatched header title(ignore) :"+headerTitle);
            }
        }else {
            indexTitleMap.put(index, headerTitle);
        }
    }
    public boolean validate(){
        if(indexTitleMap.size()==headerMeta.size()){
            return true;
        }
        return false;
    }
    public String getPropertyNameByColumnIndex(int index){
        String headerTitle = indexTitleMap.get(index);
        if(headerTitle==null){
            throw new HeaderNotMatchException("headerTitle not found by index:"+index);
        }else{
            headerMeta.getPropertyNameByHeaderTitle(headerTitle);
        }
    }
}
