package com.wangy.myapplication.poi;

import android.Manifest;
import android.os.Build;
import android.os.Bundle;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.Toast;

import com.wangy.myapplication.R;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import static android.content.pm.PackageManager.PERMISSION_GRANTED;

public class PoiActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_poi);

        // 1. 注册权限
        initPro();
    }

    /**
     * 给应用授权的方法
     */
    private void initPro() {
        // 继续执行之后的操作
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            if (ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PERMISSION_GRANTED) {
                ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, 1);
            } else {
                Toast.makeText(this, "您已经申请了权限!", Toast.LENGTH_SHORT).show();
            }
        }
    }

    /**
     * 通过poi 的工具类读取内容
     */
    public void exportPoiDataForTemplate(View view) {
        try {
            // 设局

            ArrayList<Student> students = getStudents();
            Map<String, Object[]> data = new HashMap<>();
            Map<String, ArrayList<String>> imageMap = new HashMap<>();
            ArrayList<String> images = new ArrayList<>();
            Object[] datas = new Object[3];
            // 传入的类型
            // 如果传入0 表示的是类型，传入1表示的是传入的是list
            datas[0] = 1;
            // 传入的是数据
            datas[1] = students;
            datas[2] = Student.class;
            data.put("students", datas);
            images.add("/sdcard/11/Pictures/duola.png");
            images.add("/sdcard/11/Pictures/2.png");
            imageMap.put("imags",images);
            PoiUtils.init("/sdcard/11/data.xlsx").writeTab(data,null,null,imageMap, null).
                    copy("/sdcard/11/new_data.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private ArrayList<Student> getStudents() {
        ArrayList<Student> students = new ArrayList<>();
        students.add(new Student("小明", 11, "男", 172, 85));
        students.add(new Student("小李", 21, "女", 162, 85));
        students.add(new Student("小红", 31, "男", 122, 45));
        students.add(new Student("小白", 18, "女", 162, 85));
        students.add(new Student("小黑", 14, "女", 172, 75));
        students.add(new Student("小白", 18, "女", 162, 85));
        students.add(new Student("小黑", 14, "女", 172, 75));
        students.add(new Student("小明", 11, "男", 172, 85));
        students.add(new Student("小李", 21, "女", 162, 85));
        return students;
    }


}