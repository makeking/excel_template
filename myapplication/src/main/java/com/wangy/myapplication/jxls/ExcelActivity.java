//package com.wangy.myapplication.jxls;
//
//import android.Manifest;
//import android.annotation.SuppressLint;
//import android.os.Build;
//import android.os.Bundle;
//import android.support.annotation.Nullable;
//import android.support.v4.app.ActivityCompat;
//import android.support.v4.content.ContextCompat;
//import android.support.v7.app.AppCompatActivity;
//import android.view.View;
//import android.widget.Toast;
//
//import com.wangy.myapplication.R;
//
//import org.jxls.common.Context;
//import org.jxls.util.JxlsHelper;
//
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.io.InputStream;
//import java.io.OutputStream;
//import java.math.BigDecimal;
//import java.text.ParseException;
//import java.text.SimpleDateFormat;
//import java.util.ArrayList;
//import java.util.List;
//import java.util.Locale;
//
//import static android.content.pm.PackageManager.PERMISSION_GRANTED;
//
//public class ExcelActivity extends AppCompatActivity {
//    @Override
//    protected void onCreate(@Nullable Bundle savedInstanceState) {
//        super.onCreate(savedInstanceState);
//        setContentView(R.layout.activity_excel);
//        initPro();
//    }
//
//
//    /**
//     * 给应用授权的方法
//     */
//    private void initPro() {
//        // 继续执行之后的操作
//        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
//            if (ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PERMISSION_GRANTED) {
//                ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, 1);
//            } else {
//                Toast.makeText(this, "您已经申请了权限!", Toast.LENGTH_SHORT).show();
//            }
//        }
//    }
//
//    /*
//    根据模板导出excel
//     */
//    @SuppressLint("SdCardPath")
//    public void formattoExcel(View view) {
//        List<Employee> workers = generateSampleWorkerData();
//        InputStream is = null;
//        OutputStream os = null;
//        try {
////             is =  R.class.getResourceAsStream("woker.xls");
//            is = new FileInputStream(new File("/sdcard/11/data.xls"));
//            os = new FileOutputStream("/sdcard/11/new_data.xls");
//
//            Context context = new Context();
//            context.putVar("employees", workers);
//
//            JxlsHelper.getInstance().processTemplate(is, os, context);
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//
//    }
//
//
//    private List<Employee> generateSampleWorkerData() {
//        ArrayList<Employee> employees = new ArrayList<>();
//        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
//        try {
//            employees.add(new Employee("Elsa", dateFormat.parse("1970-Jul-10"), new BigDecimal(1500), new BigDecimal(0.15)));
//            employees.add(new Employee("Oleg", dateFormat.parse("1973-Apr-30"), new BigDecimal(2300), new BigDecimal(0.25)));
//            employees.add(new Employee("Neil", dateFormat.parse("1975-Oct-05"), new BigDecimal(2500), new BigDecimal(0.00)));
//            employees.add(new Employee("Maria", dateFormat.parse("1978-Jan-07"), new BigDecimal(1700), new BigDecimal(0.15)));
//            employees.add(new Employee("John", dateFormat.parse("1969-May-30"), new BigDecimal(2800), new BigDecimal(0.20)));
//        } catch (ParseException e) {
//            e.printStackTrace();
//        }
//
//        return employees;
//    }
//}
