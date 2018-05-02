package com.seawolf.fsmobile;

import android.content.Intent;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Spinner;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class DrillActivity extends AppCompatActivity {

    private static final String[]pieces = {"1","2","3","4","5","6","7","8","9"};
    private static final String[]diam = {"50","60","75","90","100","110","125","150","175","200","250"};
    private static final String[]wall = {"/","60", "80", "100","120","140","160","180","200","220","240","260","280","300","350","400"};
    private static final String[]dalle = {"/","60", "80", "100","120","140","160","180","200","220","240","260","280","300","350","400"};
    private static final String[]machines = {"Gr. Groupe", "For. Hydr", "Grd. Hilti","Pte. Hilti"};
    private static final String[]mat = {"Béton", "Béton armé", "Maçonnerie"};
    private static final String[]dem = {"DT","C","V","S","E"};
    private static final String[]floors = {"1","2","3","4","5","6","7","8","9"};

    private int rowNb = 0;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_drill);
        Spinner pieceSpinner = findViewById(R.id.pieceSpinner);
        Spinner diamSpinner = findViewById(R.id.diamSpinner);
        Spinner wallSpinner = findViewById(R.id.wallSpinner);
        Spinner dalleSpinner = findViewById(R.id.dalleSpinner);
        Spinner machinesSpinner = findViewById(R.id.machinesSpinner);
        Spinner matSpinner = findViewById(R.id.matSpinner);
        Spinner demSpinner = findViewById(R.id.demSpinner);
        Spinner floorSpinner = findViewById(R.id.floorSpinner);

        pieceSpinner.setAdapter(new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item, pieces));
        diamSpinner.setAdapter(new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item, diam));
        wallSpinner.setAdapter(new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item, wall));
        dalleSpinner.setAdapter(new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item, dalle));
        machinesSpinner.setAdapter(new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item, machines));
        matSpinner.setAdapter(new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item, mat));
        demSpinner.setAdapter(new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item, dem));
        floorSpinner.setAdapter(new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item, floors));

        Button returnButton = findViewById(R.id.returnButton);
        returnButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                goToMenu(false);
            }
        });

        Button validButton = findViewById(R.id.validButton);
        validButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                saveExcel();
                goToMenu(true);
            }
        });

        Intent intent = getIntent();

        if (intent != null) {
            rowNb = intent.getIntExtra("ROWNB",0);
        }

    }

    private void saveExcel(){

        File filelocation = new File(getExternalFilesDir(null), "ExcelFsMobile.xls");
        InputStream is = null;
        try {
            is = new FileInputStream(filelocation);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }


        // Create a workbook using the File System
        HSSFWorkbook wb = null;
        try {
            wb = new HSSFWorkbook(is);
        } catch (IOException e) {
            e.printStackTrace();
        }

        Spinner mySpinner1=(Spinner) findViewById(R.id.pieceSpinner);
        String text1 = mySpinner1.getSelectedItem().toString();
        wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("B")).setCellValue(text1);
        Spinner mySpinner2=(Spinner) findViewById(R.id.diamSpinner);
        String text2 = mySpinner2.getSelectedItem().toString();
        wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("C")).setCellValue(text2);
        Spinner mySpinner3=(Spinner) findViewById(R.id.wallSpinner);
        String text3 = mySpinner3.getSelectedItem().toString();
        wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("G")).setCellValue(text3);
        Spinner mySpinner4=(Spinner) findViewById(R.id.dalleSpinner);
        String text4 = mySpinner4.getSelectedItem().toString();
        wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("E")).setCellValue(text4);
        Spinner mySpinner5=(Spinner) findViewById(R.id.machinesSpinner);
        String text5 = mySpinner5.getSelectedItem().toString();
        wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("M")).setCellValue(text5);
        Spinner mySpinner7=(Spinner) findViewById(R.id.demSpinner);
        String text7 = mySpinner7.getSelectedItem().toString();
        wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("Z")).setCellValue(text7);
        Spinner mySpinner8=(Spinner) findViewById(R.id.floorSpinner);
        String text8 = mySpinner8.getSelectedItem().toString();
        wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("AA")).setCellValue(text8);
        EditText chantierText = findViewById(R.id.descText);
        wb.getSheetAt(0).getRow(25+rowNb).getCell(CellReference.convertColStringToIndex("B")).setCellValue(chantierText.getText().toString());
        EditText hoursText = findViewById(R.id.hoursText);
        wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("X")).setCellValue(hoursText.getText().toString());
        wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("Y")).setCellValue(hoursText.getText().toString());
        EditText batText = findViewById(R.id.batText);
        wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("AC")).setCellValue(batText.getText().toString());


        Spinner mySpinner6=(Spinner) findViewById(R.id.matSpinner);
        String text6 = mySpinner6.getSelectedItem().toString();
        switch(text6){
            case "Béton":
                wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("J")).setCellValue("X");
                break;
            case "Béton armé":
                wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("I")).setCellValue("X");
                break;
            case "Maçonnerie":
                wb.getSheetAt(0).getRow(13+rowNb).getCell(CellReference.convertColStringToIndex("K")).setCellValue("X");
                break;
            default:
                break;
        }

        // Create a path where we will place our List of objects on external storage
        File file = new File(getExternalFilesDir(null), "ExcelFsMobile.xls");
        FileOutputStream os = null;

        try {
            os = new FileOutputStream(file);
            wb.write(os);
            Log.w("FileUtils", "Writing file" + file);
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

    }

    private void goToMenu(boolean increment){
        Intent intent = new Intent(this,MainActivity.class);
        if(increment){
            intent.putExtra("INCR",rowNb+1);
        }
        startActivity(intent);
    }
}
