package com.wangy.myapplication.jxls;

public class Worker {
    private String name;
    private int id;
    private int old;
    private double prices;

    public Worker(String name, int id, int old, double prices) {
        this.name = name;
        this.id = id;
        this.old = old;
        this.prices = prices;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public int getOld() {
        return old;
    }

    public void setOld(int old) {
        this.old = old;
    }

    public double getPrices() {
        return prices;
    }

    public void setPrices(double prices) {
        this.prices = prices;
    }
}
