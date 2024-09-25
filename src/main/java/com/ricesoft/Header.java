package com.ricesoft;


public class Header {

    public Header(String name, Integer index) {
        this.name = name;
        this.index = index;
    }

    String name;
    Integer index;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }

}
