package com.wangy.myapplication.poi;

public class Student {
    String name;
    int old;
    String sex;
    double height;
    double body_weight;

    public Student(String name, int old, String sex, double height, double body_weight) {
        this.name = name;
        this.old = old;
        this.sex = sex;
        this.height = height;
        this.body_weight = body_weight;
    }

    public Student() {
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getOld() {
        return old;
    }

    public void setOld(int old) {
        this.old = old;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public double getHeight() {
        return height;
    }

    public void setHeight(double height) {
        this.height = height;
    }

    public double getBody_weight() {
        return body_weight;
    }

    public void setBody_weight(double body_weight) {
        this.body_weight = body_weight;
    }
}
