package com.seawolf.fsmobile;

import android.app.AlertDialog;
import android.content.ActivityNotFoundException;
import android.content.DialogInterface;
import android.content.Intent;
import android.net.Uri;
import android.os.Environment;
import android.os.StrictMode;
import android.support.v4.content.FileProvider;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import android.app.Activity;
import android.content.Context;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;
import android.widget.Toast;


public class MainActivity extends Activity {
    static String TAG = "ExcelLog";
    private String mailTo = "";
    private String date = "";
    private String chantier = "";
    private int rowNb = 0;

    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        Button drillButton = findViewById(R.id.drillButton);
        drillButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                goToDrill();
            }
        });
        Button homeButton = findViewById(R.id.homeButton);
        homeButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                dialogResetConfirm();
            }
        });

        Button mailButton = findViewById(R.id.mailXLSButton);
        mailButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                sendMail();
            }
        });

        Button openButton = findViewById(R.id.seeXLSButton);
        openButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                displayXLS();
            }
        });


        Intent intent = getIntent();
        int updatedRowNb;
        if (intent != null) {
            mailTo = intent.getStringExtra("MAILTO");
            date = intent.getStringExtra("DATE");
            chantier = intent.getStringExtra("CHANTIER");
            rowNb = intent.getIntExtra("ROWNB",0);
            updatedRowNb = intent.getIntExtra("INCR",0);
            if(updatedRowNb != 0)
                rowNb = updatedRowNb;
        }

    }


    private void goToDrill() {
        Intent intent = new Intent(this, DrillActivity.class);
        intent.putExtra("ROWNB",rowNb);
        startActivity(intent);
    }

    private void goToHome() {
        Intent intent = new Intent(this, StartActivity.class);
        startActivity(intent);
    }

    private void displayXLS(){
        File filelocation = new File(getExternalFilesDir(null), "ExcelFsMobile.xls");
        Uri path = FileProvider.getUriForFile(this,this.getApplicationContext().getPackageName() + ".my.package.name.provider",filelocation);
        Intent i = new Intent(Intent.ACTION_VIEW);
        i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        i.setDataAndType(path, "application/vnd.ms-excel");
        try {
            startActivity(i);
        }
        catch (ActivityNotFoundException e) {
            Toast.makeText(this, "Pas d'application pour lire les excels trouvée", Toast.LENGTH_SHORT).show();
        }
    }

    private void dialogResetConfirm() {
        new AlertDialog.Builder(this)
                .setIcon(android.R.drawable.ic_dialog_alert)
                .setTitle("Attention")
                .setMessage("Revenir à l'accueil produira une nouvelle feuille, êtes-vous sûr ?")
                .setPositiveButton("Oui", new DialogInterface.OnClickListener() {
                    @Override
                    public void onClick(DialogInterface dialog, int which) {
                        goToHome();
                    }

                })
                .setNegativeButton("Non", null)
                .show();
    }

    private void sendMail() {
        File filelocation = new File(getExternalFilesDir(null), "ExcelFsMobile.xls");
        Uri path = FileProvider.getUriForFile(this,this.getApplicationContext().getPackageName() + ".my.package.name.provider",filelocation);
        Intent i = new Intent(Intent.ACTION_SEND);
        i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        i.setType("vnd.android.cursor.dir/email");
        i.putExtra(Intent.EXTRA_EMAIL, new String[]{mailTo});
        i.putExtra(Intent.EXTRA_SUBJECT, chantier + " "+date);
        i.putExtra(Intent.EXTRA_STREAM, path);
        Log.v(getClass().getSimpleName(),path.getPath());
        try {
            startActivity(Intent.createChooser(i, "Envoi du mail ..."));
        } catch (android.content.ActivityNotFoundException ex) {
            Toast.makeText(this, "Pas de client mail installé.", Toast.LENGTH_SHORT).show();
        }
    }



    public static boolean isExternalStorageReadOnly() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState)) {
            return true;
        }
        return false;
    }

    public static boolean isExternalStorageAvailable() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(extStorageState)) {
            return true;
        }
        return false;
    }
}
