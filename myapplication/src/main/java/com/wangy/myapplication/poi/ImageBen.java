package com.wangy.myapplication.poi;

public class ImageBen {
    // 资源路径
    String src;
    String netPath;
    // 宽度和高度
    int width;
    int height;
    // 边距
    int margionLeft;
    int margionRight;
    int margionTop;
    int margionBootm;
//    int imageType;
    PreviousImageEnum enumType = PreviousImageEnum.DEFULT;

    public ImageBen(String src, String netPath, int width, int height, int margionLeft, int margionRight, int margionTop, int margionBootm) {
        this.src = src;
        this.netPath = netPath;
        this.width = width;
        this.height = height;
        this.margionLeft = margionLeft;
        this.margionRight = margionRight;
        this.margionTop = margionTop;
        this.margionBootm = margionBootm;
    }

    public PreviousImageEnum getEnumType() {
        return enumType;
    }

    public void setEnumType(PreviousImageEnum enumType) {
        this.enumType = enumType;
    }

    public String getSrc() {
        return src;
    }

    public void setSrc(String src) {
        this.src = src;
    }

    public String getNetPath() {
        return netPath;
    }

    public void setNetPath(String netPath) {
        this.netPath = netPath;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public int getMargionLeft() {
        return margionLeft;
    }

    public void setMargionLeft(int margionLeft) {
        this.margionLeft = margionLeft;
    }

    public int getMargionRight() {
        return margionRight;
    }

    public void setMargionRight(int margionRight) {
        this.margionRight = margionRight;
    }

    public int getMargionTop() {
        return margionTop;
    }

    public void setMargionTop(int margionTop) {
        this.margionTop = margionTop;
    }

    public int getMargionBootm() {
        return margionBootm;
    }

    public void setMargionBootm(int margionBootm) {
        this.margionBootm = margionBootm;
    }
}
