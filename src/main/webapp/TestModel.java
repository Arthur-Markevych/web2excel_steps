package main.webapp;

import java.util.Arrays;
import java.util.List;

public class TestModel {

    private String imgPath;

    private String name;

    private String options;

    private int amount;

    private double price;

    private double height;

    private double width;

    private double depth;

    private double volume;

    public TestModel() {
    }

    public static List<TestModel> getAll() {
        TestModel m1 = new TestModel("C:\\Users\\Artur_Markevych\\Documents\\Pre_Prod_java_q3q4_2018\\task1 – git pracrice I\\images2\\006.jpg",
                "Chick", "wings", 2, 23.90, 20, 20 ,20);
        TestModel m2 = new TestModel("C:\\Users\\Artur_Markevych\\Documents\\Pre_Prod_java_q3q4_2018\\task1 – git pracrice I\\images2\\004.jpg",
                "Crazy", "Egs", 1, 20.70, 20, 20 ,20);
        TestModel m3 = new TestModel("C:\\Users\\Artur_Markevych\\Documents\\Pre_Prod_java_q3q4_2018\\task1 – git pracrice I\\images2\\005.jpg",
                "Fly hunter", "Gun", 4, 25.50, 20, 20 ,20);
        return Arrays.asList(m1, m2, m3);
    }

    public TestModel(String imgPath, String name, String options, int amount,
                     double price, double height, double width, double depth) {
        this.imgPath = imgPath;
        this.name = name;
        this.options = options;
        this.amount = amount;
        this.price = price;
        this.height = height;
        this.width = width;
        this.depth = depth;
    }

    public String getImgPath() {
        return imgPath;
    }

    public void setImgPath(String imgPath) {
        this.imgPath = imgPath;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getOptions() {
        return options;
    }

    public void setOptions(String options) {
        this.options = options;
    }

    public int getAmount() {
        return amount;
    }

    public void setAmount(int amount) {
        this.amount = amount;
    }

    public double getPrice() {
        return price;
    }

    public void setPrice(double price) {
        this.price = price;
    }

    public double getHeight() {
        return height;
    }

    public void setHeight(double height) {
        this.height = height;
    }

    public double getWidth() {
        return width;
    }

    public void setWidth(double width) {
        this.width = width;
    }

    public double getDepth() {
        return depth;
    }

    public void setDepth(double depth) {
        this.depth = depth;
    }

    public double getVolume() {
        return depth * width * height;
    }

    public void setVolume(double volume) {
        this.volume = volume;
    }


}
