package com.seawolf.fsmobile;
import android.content.Intent;
import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;


public class StartActivity extends AppCompatActivity {

    private static String TAG = "ExcelLog";
    private String mailTo = "";
    private String date = "";
    private String chantier = "";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_start);
        Button writeExcelButton = findViewById(R.id.continueButton);
        writeExcelButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                saveExcelFile();
                goToMain();
            }
        });
    }

    protected void goToMain(){
        Intent intent = new Intent(this,MainActivity.class);
        intent.putExtra("MAILTO", mailTo);
        intent.putExtra("DATE", date);
        intent.putExtra("CHANTIER", chantier);
        startActivity(intent);
    }

    private boolean saveExcelFile() {

        boolean success = false;

        InputStream is = getResources().openRawResource(R.raw.model);


        // Create a workbook using the File System
        HSSFWorkbook wb = null;
        try {
            wb = new HSSFWorkbook(is);
        } catch (IOException e) {
            e.printStackTrace();
        }

        EditText dateText = findViewById(R.id.dateText);
        date =dateText.getText().toString();
        wb.getSheetAt(0).getRow(3).getCell(CellReference.convertColStringToIndex("L")).setCellValue(date);
        EditText mailText = findViewById(R.id.mailToText);
        mailTo = mailText.getText().toString();
        EditText nameForText = findViewById(R.id.nameForText);
        wb.getSheetAt(0).getRow(48).getCell(CellReference.convertColStringToIndex("M")).setCellValue(nameForText.getText().toString());
        EditText nameForHelpTest = findViewById(R.id.nameForHelpTest);
        wb.getSheetAt(0).getRow(48).getCell(CellReference.convertColStringToIndex("V")).setCellValue(nameForHelpTest.getText().toString());
        EditText chantierText = findViewById(R.id.chantierText);
        chantier = chantierText.getText().toString();
        wb.getSheetAt(0).getRow(3).getCell(CellReference.convertColStringToIndex("S")).setCellValue(chantier);
        EditText refText = findViewById(R.id.refText);
        wb.getSheetAt(0).getRow(3).getCell(CellReference.convertColStringToIndex("AE")).setCellValue(refText.getText().toString());
        EditText techText = findViewById(R.id.techText);
        wb.getSheetAt(0).getRow(5).getCell(CellReference.convertColStringToIndex("N")).setCellValue(techText.getText().toString());




        // Create a path where we will place our List of objects on external storage
        File file = new File(getExternalFilesDir(null), "ExcelFsMobile.xls");
        FileOutputStream os = null;

        try {
            os = new FileOutputStream(file);
            wb.write(os);
            Log.w("FileUtils", "Writing file" + file);
            success = true;
        } catch (IOException e) {
            Log.w("FileUtils", "Error writing " + file, e);
        } catch (Exception e) {
            Log.w("FileUtils", "Failed to save file", e);
        } finally {
            try {
                if (null != os)
                    os.close();
            } catch (Exception ex) {
            }
        }
        return success;
    }

}
