package in.kriscent.exceldemoandroid;

import android.database.Cursor;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.widget.Toast;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity {
    private static final String TAG = MainActivity.class.getSimpleName();
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        createExcel();
    }

    /* create excel and write
    * on internal memory
     */
    private void createExcel() {
        File sd = Environment.getExternalStorageDirectory();
        String csvFile = "KLMSEnquiry.xls";

        File directory = new File(sd.getAbsolutePath());
        //create directory if not exist
        if (!directory.isDirectory()) {
            directory.mkdirs();
        }
        try {
            //file path
            File file = new File(directory, csvFile);
            WorkbookSettings wbSettings = new WorkbookSettings();
            wbSettings.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook;
            workbook = Workbook.createWorkbook(file, wbSettings);
            //Excel sheet name. 0 represents first sheet
            WritableSheet sheet = workbook.createSheet("userList", 0);

            sheet.addCell(new Label(0, 0, "UserName")); // column and row
            sheet.addCell(new Label(1, 0, "PhoneNumber"));

            sheet.addCell(new Label(0, 0, "Ram Narayan Raigar"));
            sheet.addCell(new Label(1, 0, "8619713127"));

            sheet.addCell(new Label(0, 1, "Hemant nayak"));
            sheet.addCell(new Label(1, 1, "9001493751"));

            sheet.addCell(new Label(0, 2, "Sonu raghuvanshi"));
            sheet.addCell(new Label(1, 2, "8005883136"));


            //closing cursor
            workbook.write();
            workbook.close();
            Toast.makeText(getApplication(), "Data Exported in a Excel Sheet", Toast.LENGTH_SHORT).show();


        } catch (WriteException | IOException e) {
            Log.e(TAG, e.getMessage());
        }
    }
}

