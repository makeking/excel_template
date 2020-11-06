package com.wangy.myapplication;

import android.Manifest;
import android.os.Build;
import android.os.Bundle;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.Toast;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.TableStyle;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;

import static android.content.pm.PackageManager.PERMISSION_GRANTED;

public class MainActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
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



    public void exportDoc(View view) {
        //核心API采用了极简设计，只需要一行代码
        try {

            RowRenderData header = RowRenderData.build(new TextRenderData("000000", "姓名"), new TextRenderData("000000", "学历"));

            RowRenderData row0 = RowRenderData.build("张三", "研究生");
            RowRenderData row1 = RowRenderData.build("李四", "博士");
            RowRenderData row2 = RowRenderData.build("李四  ww 33", "  博士  ddg ert ");
            RowRenderData row3= RowRenderData.build("李四123", "博士321");
            TableStyle tableStyle = new TableStyle();
            tableStyle.setAlign(STJc.CENTER);
            row3.setRowStyle(tableStyle);
            XWPFTemplate.compile("mnt/sdcard/11/数据.docx").render(

                    new HashMap<String, Object>() {
                        //  1. 设置文字
                        {
                            put("bigtitle", "测试");
                            put("smallTitle", "文字");
                            put("smallTitle2", "表格");
                            put("smallTitle3", "图表");
                            put("smallTitle4", "图片");
                            //3. 设置图片
                            put("log", new PictureRenderData(80, 100, "/mnt/sdcard/11/Pictures/duola.png"));
                            put("table", new MiniTableRenderData(header, Arrays.asList(row0, row1,row2,row3)));
                            // 2. 设置表格
                        }
                        //4. 设置图表
                    }
            ).writeToFile("mnt/sdcard/11/新数据.docx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
